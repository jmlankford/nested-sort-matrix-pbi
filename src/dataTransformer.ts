/*
 * dataTransformer.ts
 * ------------------
 * Converts a flat Power BI Table DataView into the internal tree model the
 * renderer consumes. Responsibilities:
 *   - Field discovery (which row/value/col role slots are actually populated).
 *   - Client-side grouping into a row hierarchy (RowTreeNode tree).
 *   - Crosstab / pivot computation when column-field slots are populated.
 *   - SUM aggregation for subtotals, column subtotals, grand totals.
 *
 * KNOWN LIMITATION (documented in README): subtotals are computed by SUMMING
 * pre-computed row-level measure values. Non-additive measures (DISTINCTCOUNT,
 * ratios, AVERAGE, etc.) will therefore produce incorrect rolled-up values.
 * Workaround: author a dedicated subtotal DAX measure and map it to a separate
 * value slot.
 *
 * Strict TypeScript: no `any`. The only `unknown` usage is JSON.parse boundaries
 * which are validated before use.
 */

import powerbi from "powerbi-visuals-api";
import DataView = powerbi.DataView;
import DataViewTable = powerbi.DataViewTable;
import DataViewMetadataColumn = powerbi.DataViewMetadataColumn;
import PrimitiveValue = powerbi.PrimitiveValue;
import ISelectionId = powerbi.visuals.ISelectionId;
import DataViewObjects = powerbi.DataViewObjects;

import { valueFormatter } from "powerbi-visuals-utils-formattingutils";

import { VisualSettings } from "./settings";

// Stacked-bucket role names (must match capabilities.json exactly). Each role is
// a single field well that accepts multiple stacked fields, identical to the
// native Power BI matrix field wells.
const ROLE_ROW_FIELDS = "rowFields";
const ROLE_COLUMN_FIELDS = "columnFields";
const ROLE_VALUES = "values";

// Internal key separators. Printable but deliberately unlikely to appear in
// real field values, so composed path/column keys never collide.
const PATH_SEP = "|#|";
const ID_SEP = "::#::";
const SUBTOTAL_TOKEN = "~ST~";
const GRANDTOTAL_TOKEN = "~GT~";
const BLANK_LABEL = "(Blank)";

// ---------------------------------------------------------------------------
// Public model types.
// ---------------------------------------------------------------------------

export type FieldRole = "row" | "value" | "col";

/** Describes a single populated role slot bound to a Table DataView column. */
export interface FieldMeta {
    role: FieldRole;
    /** Dense 0-based position of this field within its bucket, in DataView column order. */
    slotIndex: number;
    /** Index into dataView.table.columns. */
    columnIndex: number;
    queryName: string;
    /** Original display name from the DataView (designer/model supplied). */
    originalName: string;
    /** Effective display name after a session rename (config panel). */
    displayName: string;
    formatString: string;
    isNumeric: boolean;
    /** Per-column objects bag (carries per-slot value formatting & CF). */
    columnObjects: DataViewObjects | undefined;
}

/** One value column in the (possibly pivoted) grid. */
export interface LeafColumn {
    /** Stable id used to key cell values on each node. */
    id: string;
    /** Original value slot index (0..14) — used for CF & number formatting. */
    valueSlotIndex: number;
    /** Path of column-field labels ([] in non-pivot mode). */
    pivotPath: string[];
    isColSubtotal: boolean;
    isColGrandTotal: boolean;
}

export interface RowTreeNode {
    key: string;
    /** Group label for this node (formatted). */
    label: string;
    /** Raw underlying group value, retained for sorting. */
    rawValue: PrimitiveValue | null;
    /** 0-based row hierarchy depth. */
    level: number;
    children: RowTreeNode[];
    isLeaf: boolean;
    isGrandTotal: boolean;
    /** Aggregated measure values keyed by LeafColumn.id. */
    values: { [leafColId: string]: number | null };
    /** Underlying table row indices — populated on leaf nodes only. */
    rowIndices: number[];
    /** Lazily-created selection id for this node (built on demand). */
    selectionId?: ISelectionId;
    parent?: RowTreeNode;
    /** Transient child lookup used only during construction. */
    _childMap?: Map<string, RowTreeNode>;
}

export interface ColumnHeaderCell {
    label: string;
    /** Number of leaf columns this header cell spans. */
    span: number;
    /** Header level (0..colFieldCount-1 for pivot fields). */
    level: number;
    isSubtotal: boolean;
    isGrandTotal: boolean;
}

export interface ColumnHeaderLayout {
    /** One row of header cells per column-field level (empty in non-pivot mode). */
    pivotRows: ColumnHeaderCell[][];
    /** Bottom header row: one cell per leaf column (the measure name). */
    leafRow: LeafColumn[];
}

export interface TransformResult {
    rootNodes: RowTreeNode[];
    grandTotal: RowTreeNode | null;
    leafColumns: LeafColumn[];
    columnHeader: ColumnHeaderLayout;
    activeRowFields: FieldMeta[];
    activeValueFields: FieldMeta[];
    activeColFields: FieldMeta[];
    /** Underlying data row count. */
    rowCount: number;
    hasRowFields: boolean;
    isPivot: boolean;
}

/** Factory injected by the visual so the transformer can build selection ids. */
export type SelectionIdFactory = (rowIndices: number[]) => ISelectionId | undefined;

// ---------------------------------------------------------------------------
// Field discovery.
// ---------------------------------------------------------------------------

interface DiscoveredFields {
    rows: FieldMeta[];
    values: FieldMeta[];
    cols: FieldMeta[];
}

function makeFieldMeta(
    role: FieldRole,
    slotIndex: number,
    columnIndex: number,
    column: DataViewMetadataColumn
): FieldMeta {
    const original = column.displayName != null ? String(column.displayName) : "";
    return {
        role,
        slotIndex,
        columnIndex,
        queryName: column.queryName != null ? String(column.queryName) : `col${columnIndex}`,
        originalName: original,
        displayName: original,
        formatString: column.format != null ? String(column.format) : "",
        isNumeric: !!(column.type && (column.type.numeric || column.type.integer)),
        columnObjects: column.objects
    };
}

/**
 * Discover the populated fields in the Table DataView by reading each column's
 * role membership (column.roles). Each of the three stacked buckets
 * (rowFields / columnFields / values) may hold many fields.
 *
 * The slot index assigned to each field is its dense position WITHIN its bucket
 * in DataView column order. That order reflects how the report developer stacked
 * the fields in the field well, i.e. the designer's intended hierarchy order;
 * the config panel can reorder from this baseline. All downstream logic keys on
 * this slotIndex, so it stays consistent across the visual.
 */
export function discoverFields(dataView: DataView | undefined): DiscoveredFields {
    const result: DiscoveredFields = { rows: [], values: [], cols: [] };
    const table: DataViewTable | undefined = dataView && dataView.table;
    if (!table || !table.columns) {
        return result;
    }

    let rowSlot = 0;
    let colSlot = 0;
    let valSlot = 0;

    table.columns.forEach((column: DataViewMetadataColumn, columnIndex: number) => {
        const roles = column.roles;
        if (!roles) {
            return;
        }
        if (roles[ROLE_ROW_FIELDS]) {
            result.rows.push(makeFieldMeta("row", rowSlot++, columnIndex, column));
        } else if (roles[ROLE_COLUMN_FIELDS]) {
            result.cols.push(makeFieldMeta("col", colSlot++, columnIndex, column));
        } else if (roles[ROLE_VALUES]) {
            result.values.push(makeFieldMeta("value", valSlot++, columnIndex, column));
        }
    });

    // Arrays are already in DataView column order (== bucket stack order).
    return result;
}

// ---------------------------------------------------------------------------
// Formatting helpers.
// ---------------------------------------------------------------------------

function buildFormatters(fields: FieldMeta[]): Map<number, valueFormatter.IValueFormatter> {
    const map = new Map<number, valueFormatter.IValueFormatter>();
    fields.forEach((f) => {
        map.set(
            f.columnIndex,
            valueFormatter.create({ format: f.formatString || undefined })
        );
    });
    return map;
}

function formatLabel(
    raw: PrimitiveValue,
    formatter: valueFormatter.IValueFormatter | undefined
): string {
    if (raw === null || raw === undefined || raw === "") {
        return BLANK_LABEL;
    }
    if (formatter) {
        const text = formatter.format(raw);
        return text != null && text !== "" ? text : BLANK_LABEL;
    }
    return String(raw);
}

function toNumber(raw: PrimitiveValue): number | null {
    if (raw === null || raw === undefined || raw === "") {
        return null;
    }
    const n = typeof raw === "number" ? raw : Number(raw);
    return isNaN(n) ? null : n;
}

// ---------------------------------------------------------------------------
// Column (pivot) computation.
// ---------------------------------------------------------------------------

interface ColTreeNode {
    label: string;
    level: number;
    path: string[];
    children: ColTreeNode[];
    childMap: Map<string, ColTreeNode>;
}

function newColNode(label: string, level: number, path: string[]): ColTreeNode {
    return { label, level, path, children: [], childMap: new Map() };
}

interface ColumnPlan {
    leafColumns: LeafColumn[];
    columnHeader: ColumnHeaderLayout;
    /**
     * For a given data row's column path, returns every LeafColumn.id the row's
     * value contributes to (exact leaf + ancestor subtotals + grand total),
     * for a single value slot. Cached by joined path.
     */
    contributingColumnIds: (colPath: string[], valueSlotIndex: number) => string[];
}

function leafColId(pivotPathKey: string, valueSlotIndex: number): string {
    return `${pivotPathKey}${ID_SEP}v${valueSlotIndex}`;
}

function buildNonPivotPlan(valueFields: FieldMeta[]): ColumnPlan {
    const leafColumns: LeafColumn[] = valueFields.map((vf) => ({
        id: leafColId("", vf.slotIndex),
        valueSlotIndex: vf.slotIndex,
        pivotPath: [],
        isColSubtotal: false,
        isColGrandTotal: false
    }));

    const columnHeader: ColumnHeaderLayout = {
        pivotRows: [],
        leafRow: leafColumns
    };

    return {
        leafColumns,
        columnHeader,
        contributingColumnIds: (_colPath: string[], valueSlotIndex: number) => [
            leafColId("", valueSlotIndex)
        ]
    };
}

function buildPivotPlan(
    table: DataViewTable,
    colFields: FieldMeta[],
    valueFields: FieldMeta[],
    settings: VisualSettings
): ColumnPlan {
    const formatters = buildFormatters(colFields);
    const lastColLevel = colFields.length - 1;

    // 1. Build the column tree from the unique column-field tuples in the data.
    const root = newColNode("", -1, []);
    const rows = table.rows || [];
    for (let r = 0; r < rows.length; r++) {
        const row = rows[r];
        let node = root;
        const path: string[] = [];
        for (let level = 0; level < colFields.length; level++) {
            const cf = colFields[level];
            const label = formatLabel(row[cf.columnIndex], formatters.get(cf.columnIndex));
            path.push(label);
            let child = node.childMap.get(label);
            if (!child) {
                child = newColNode(label, level, path.slice());
                node.childMap.set(label, child);
                node.children.push(child);
            }
            node = child;
        }
    }

    // 2. Sort each level's children ascending by default (column header sorting
    //    is layered on top by the renderer / sort manager).
    const sortChildren = (node: ColTreeNode): void => {
        node.children.sort((a, b) => a.label.localeCompare(b.label, undefined, { numeric: true }));
        node.children.forEach(sortChildren);
    };
    sortChildren(root);

    // 3. Emit leaf columns in render order (DFS, subtotal after each parent group),
    //    plus build the spanning header layout.
    const leafColumns: LeafColumn[] = [];
    const pivotRows: ColumnHeaderCell[][] = [];
    for (let i = 0; i < colFields.length; i++) {
        pivotRows.push([]);
    }

    const pathKey = (path: string[]): string => "c:" + path.join(PATH_SEP);

    const emitValueLeaves = (node: ColTreeNode): void => {
        valueFields.forEach((vf) => {
            leafColumns.push({
                id: leafColId(pathKey(node.path), vf.slotIndex),
                valueSlotIndex: vf.slotIndex,
                pivotPath: node.path.slice(),
                isColSubtotal: false,
                isColGrandTotal: false
            });
        });
    };

    const emitSubtotalLeaves = (node: ColTreeNode): void => {
        valueFields.forEach((vf) => {
            leafColumns.push({
                id: leafColId(pathKey(node.path) + PATH_SEP + SUBTOTAL_TOKEN, vf.slotIndex),
                valueSlotIndex: vf.slotIndex,
                pivotPath: node.path.slice(),
                isColSubtotal: true,
                isColGrandTotal: false
            });
        });
    };

    // Record the leaf-span for header cells as we go.
    const headerCellSpans: { cell: ColumnHeaderCell; startLeaf: number }[] = [];

    const recurse = (node: ColTreeNode): void => {
        node.children.forEach((child) => {
            const startLeaf = leafColumns.length;
            if (child.level === lastColLevel) {
                emitValueLeaves(child);
            } else {
                recurse(child);
                if (settings.subtotals.columnSubtotals) {
                    emitSubtotalLeaves(child);
                }
            }
            const span = leafColumns.length - startLeaf;
            if (span > 0) {
                const cell: ColumnHeaderCell = {
                    label: child.label,
                    span,
                    level: child.level,
                    isSubtotal: false,
                    isGrandTotal: false
                };
                pivotRows[child.level].push(cell);
                headerCellSpans.push({ cell, startLeaf });
            }
        });
    };

    recurse(root);

    // 4. Grand total columns at the far right.
    if (settings.subtotals.grandTotalColumn) {
        valueFields.forEach((vf) => {
            leafColumns.push({
                id: leafColId(GRANDTOTAL_TOKEN, vf.slotIndex),
                valueSlotIndex: vf.slotIndex,
                pivotPath: [],
                isColSubtotal: false,
                isColGrandTotal: true
            });
        });
    }

    const columnHeader: ColumnHeaderLayout = { pivotRows, leafRow: leafColumns };

    // 5. Cache of contributing column ids per row column-path.
    const cache = new Map<string, string[]>();
    const contributingColumnIds = (colPath: string[], valueSlotIndex: number): string[] => {
        const cacheKey = colPath.join(PATH_SEP) + ID_SEP + valueSlotIndex;
        const cached = cache.get(cacheKey);
        if (cached) {
            return cached;
        }
        const ids: string[] = [];
        // Exact leaf.
        ids.push(leafColId(pathKey(colPath), valueSlotIndex));
        // Ancestor subtotals (prefixes shorter than the full path).
        if (settings.subtotals.columnSubtotals) {
            for (let k = 1; k < colPath.length; k++) {
                const prefix = colPath.slice(0, k);
                ids.push(leafColId(pathKey(prefix) + PATH_SEP + SUBTOTAL_TOKEN, valueSlotIndex));
            }
        }
        // Grand total column.
        if (settings.subtotals.grandTotalColumn) {
            ids.push(leafColId(GRANDTOTAL_TOKEN, valueSlotIndex));
        }
        cache.set(cacheKey, ids);
        return ids;
    };

    return { leafColumns, columnHeader, contributingColumnIds };
}

// ---------------------------------------------------------------------------
// Row tree construction + aggregation.
// ---------------------------------------------------------------------------

function newNode(
    key: string,
    label: string,
    rawValue: PrimitiveValue | null,
    level: number,
    parent: RowTreeNode | undefined
): RowTreeNode {
    return {
        key,
        label,
        rawValue,
        level,
        children: [],
        isLeaf: false,
        isGrandTotal: false,
        values: {},
        rowIndices: [],
        parent,
        _childMap: new Map<string, RowTreeNode>()
    };
}

function addInto(target: { [k: string]: number | null }, id: string, value: number | null): void {
    if (value === null) {
        // Preserve a null slot only if nothing else has written a number yet.
        if (!(id in target)) {
            target[id] = null;
        }
        return;
    }
    const existing = target[id];
    target[id] = (existing === null || existing === undefined ? 0 : existing) + value;
}

/** Roll child values up into parents (post-order). */
function rollUp(node: RowTreeNode): void {
    if (node.isLeaf) {
        return;
    }
    node.children.forEach((child) => {
        rollUp(child);
        for (const id in child.values) {
            addInto(node.values, id, child.values[id]);
        }
    });
}

/** Strip transient construction state. */
function finalize(node: RowTreeNode): void {
    node._childMap = undefined;
    node.children.forEach(finalize);
}

/**
 * Collect every underlying data row index beneath a node (used for selecting a
 * whole group). Leaf nodes carry their own indices; groups gather descendants.
 */
export function collectRowIndices(node: RowTreeNode): number[] {
    if (node.isLeaf) {
        return node.rowIndices.slice();
    }
    const out: number[] = [];
    const stack: RowTreeNode[] = [node];
    while (stack.length) {
        const n = stack.pop() as RowTreeNode;
        if (n.isLeaf) {
            for (let i = 0; i < n.rowIndices.length; i++) {
                out.push(n.rowIndices[i]);
            }
        } else {
            for (let i = 0; i < n.children.length; i++) {
                stack.push(n.children[i]);
            }
        }
    }
    return out;
}

// ---------------------------------------------------------------------------
// Main transform.
// ---------------------------------------------------------------------------

export function transform(
    dataView: DataView | undefined,
    activeRowFields: FieldMeta[],
    activeValueFields: FieldMeta[],
    activeColFields: FieldMeta[],
    settings: VisualSettings,
    selectionFactory: SelectionIdFactory
): TransformResult {
    const table: DataViewTable | undefined = dataView && dataView.table;
    const rows = (table && table.rows) || [];
    const isPivot = activeColFields.length > 0;
    const hasRowFields = activeRowFields.length > 0;

    const plan: ColumnPlan = isPivot
        ? buildPivotPlan(table as DataViewTable, activeColFields, activeValueFields, settings)
        : buildNonPivotPlan(activeValueFields);

    const rowFormatters = buildFormatters(activeRowFields);
    const colFormatters = buildFormatters(activeColFields);

    const roots: RowTreeNode[] = [];
    const rootMap = new Map<string, RowTreeNode>();

    // Grand total node always accumulates everything (rendered only if enabled).
    const grandTotal: RowTreeNode = newNode("__grandTotal__", "Grand Total", null, -1, undefined);
    grandTotal.isGrandTotal = true;

    const accumulateRow = (leaf: RowTreeNode, rowIndex: number): void => {
        const row = rows[rowIndex];
        // Determine the column path for this data row (pivot mode).
        let colPath: string[] = [];
        if (isPivot) {
            colPath = activeColFields.map((cf) =>
                formatLabel(row[cf.columnIndex], colFormatters.get(cf.columnIndex))
            );
        }
        activeValueFields.forEach((vf) => {
            const value = toNumber(row[vf.columnIndex]);
            const ids = plan.contributingColumnIds(colPath, vf.slotIndex);
            for (let i = 0; i < ids.length; i++) {
                addInto(leaf.values, ids[i], value);
            }
        });
    };

    if (!hasRowFields) {
        // No row grouping: everything aggregates into a single (grand total) row.
        const synthetic = newNode("__all__", settings.subtotals.labelText || "Total", null, 0, undefined);
        synthetic.isLeaf = true;
        for (let r = 0; r < rows.length; r++) {
            synthetic.rowIndices.push(r);
            accumulateRow(synthetic, r);
        }
        // Mirror values onto the grand total node.
        for (const id in synthetic.values) {
            grandTotal.values[id] = synthetic.values[id];
        }
        finalize(synthetic);
        return {
            rootNodes: [],
            grandTotal,
            leafColumns: plan.leafColumns,
            columnHeader: plan.columnHeader,
            activeRowFields,
            activeValueFields,
            activeColFields,
            rowCount: rows.length,
            hasRowFields: false,
            isPivot
        };
    }

    const lastRowLevel = activeRowFields.length - 1;

    for (let r = 0; r < rows.length; r++) {
        const row = rows[r];
        let list = roots;
        let map = rootMap;
        let parent: RowTreeNode | undefined = undefined;
        let pathKey = "";
        let node: RowTreeNode | undefined = undefined;

        for (let level = 0; level <= lastRowLevel; level++) {
            const rf = activeRowFields[level];
            const raw = row[rf.columnIndex];
            const label = formatLabel(raw, rowFormatters.get(rf.columnIndex));
            pathKey += PATH_SEP + label;

            let child = map.get(label);
            if (!child) {
                child = newNode(pathKey, label, raw, level, parent);
                map.set(label, child);
                list.push(child);
            }
            node = child;
            parent = child;
            list = child.children;
            map = child._childMap as Map<string, RowTreeNode>;

            if (level === lastRowLevel) {
                child.isLeaf = true;
                child.rowIndices.push(r);
            }
        }

        if (node) {
            accumulateRow(node, r);
        }
    }

    // Roll subtotals up the row hierarchy.
    roots.forEach((root) => {
        rollUp(root);
        for (const id in root.values) {
            addInto(grandTotal.values, id, root.values[id]);
        }
    });

    roots.forEach(finalize);

    // Pre-build selection ids for leaf nodes lazily via the factory on demand
    // (renderer requests node.selectionId; we attach a getter-like helper here).
    // We attach selection ids eagerly only for leaves to keep behaviour simple.
    const attachSelection = (node: RowTreeNode): void => {
        if (node.isLeaf) {
            node.selectionId = selectionFactory(node.rowIndices);
        } else {
            node.children.forEach(attachSelection);
        }
    };
    roots.forEach(attachSelection);

    return {
        rootNodes: roots,
        grandTotal,
        leafColumns: plan.leafColumns,
        columnHeader: plan.columnHeader,
        activeRowFields,
        activeValueFields,
        activeColFields,
        rowCount: rows.length,
        hasRowFields: true,
        isPivot
    };
}
