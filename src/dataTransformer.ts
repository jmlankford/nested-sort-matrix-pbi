/*
 * dataTransformer.ts
 * ------------------
 * Converts a Power BI **Matrix** DataView into the internal tree model the
 * renderer consumes. Responsibilities:
 *   - Field discovery (row / column-group / value role slots from the matrix
 *     hierarchies and value sources).
 *   - Walk of dataView.matrix.rows.root into the RowTreeNode hierarchy. The
 *     engine supplies every level's values (including subtotal / grand-total
 *     nodes) at that node's true scope, so the visual performs no aggregation
 *     of its own.
 *   - Crosstab / pivot column layout from dataView.matrix.columns.root.
 *
 * Because the engine evaluates each measure at the true scope of every
 * hierarchy node, level-aware and non-additive measures (DISTINCTCOUNT, ratios,
 * AVERAGE, ISINSCOPE dispatchers, ...) roll up correctly.
 *
 * Value cells may be text (text measures). They are stored raw (number | string
 * | null); numeric-only checks live in the consumers (renderer / CF / sort).
 *
 * Strict TypeScript: no `any`. The only `unknown` usage is JSON.parse boundaries
 * which are validated before use.
 */

import powerbi from "powerbi-visuals-api";
import DataView = powerbi.DataView;
import DataViewMatrix = powerbi.DataViewMatrix;
import DataViewMatrixNode = powerbi.DataViewMatrixNode;
import DataViewHierarchyLevel = powerbi.DataViewHierarchyLevel;
import DataViewMetadataColumn = powerbi.DataViewMetadataColumn;
import PrimitiveValue = powerbi.PrimitiveValue;
import ISelectionId = powerbi.visuals.ISelectionId;
import DataViewObjects = powerbi.DataViewObjects;
import CustomVisualOpaqueIdentity = powerbi.visuals.CustomVisualOpaqueIdentity;

import { valueFormatter } from "powerbi-visuals-utils-formattingutils";

import { VisualSettings } from "./settings";

// Stacked-bucket role names (must match capabilities.json exactly). Each role is
// a single field well that accepts multiple stacked fields, identical to the
// native Power BI matrix field wells.
const ROLE_ROW_FIELDS = "rowFields";
const ROLE_COLUMN_FIELDS = "columnFields";

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

/** A raw value cell from the matrix — text measures are kept as strings. */
export type CellValue = number | string | null;

/** Describes a single populated role slot bound to a Matrix DataView source. */
export interface FieldMeta {
    role: FieldRole;
    /** Dense 0-based position of this field within its bucket, in projection order. */
    slotIndex: number;
    /**
     * Synthetic dense index across all discovered fields. In the matrix world
     * there is no single flat column array, so this is used only as a stable,
     * unique key for internal formatter maps.
     */
    columnIndex: number;
    queryName: string;
    /** Original display name from the DataView (designer/model supplied). */
    originalName: string;
    /** Effective display name after a session rename (config panel). */
    displayName: string;
    formatString: string;
    isNumeric: boolean;
    /** Per-source objects bag (carries per-slot value formatting, CF & subtotal toggles). */
    columnObjects: DataViewObjects | undefined;
}

/** One value column in the (possibly pivoted) grid. */
export interface LeafColumn {
    /** Stable id used to key cell values on each node. */
    id: string;
    /** Original value slot index (0..N) — used for CF & number formatting. */
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
    /**
     * 0-based DISPLAY depth. When no row levels are hidden this equals the matrix
     * hierarchy level; when a level is hidden its nodes are flattened out and the
     * survivors' display depth is compacted, so this can be shallower than the
     * matrix level. All rendering / indentation / sort keys off this.
     */
    level: number;
    /** True matrix hierarchy level (used for host expand/collapse targeting). */
    matrixLevel?: number;
    /**
     * Full matrix ancestor chain (root→node), INCLUDING nodes at hidden levels,
     * so a selection id built via withMatrixNode keeps the complete level path
     * even when intermediate levels are flattened out of the display tree.
     */
    matrixChain?: DataViewMatrixNode[];
    children: RowTreeNode[];
    isLeaf: boolean;
    isGrandTotal: boolean;
    /** Engine-computed measure values keyed by LeafColumn.id (may be text). */
    values: { [leafColId: string]: CellValue };
    /** Lazily-created selection id for this node (built on demand). */
    selectionId?: ISelectionId;
    /** Opaque identity of the backing matrix node (used for selection & expand/collapse). */
    identity?: CustomVisualOpaqueIdentity;
    /** The backing matrix node, retained for Phase 2 selection / expansion services. */
    matrixNode?: DataViewMatrixNode;
    /** Engine expansion flag mirrored from the matrix node (undefined = not expandable). */
    isCollapsed?: boolean;
    parent?: RowTreeNode;
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
    /** Leaf (innermost) row count. */
    rowCount: number;
    hasRowFields: boolean;
    isPivot: boolean;
    /**
     * Every expandable node in the FULL matrix tree — including nodes at HIDDEN
     * levels that are flattened out of rootNodes — each carrying its matrixChain.
     * Expand/Collapse-All walks this by matrix level so it can still reach and
     * expand collapsed hidden levels (whose children must materialise before the
     * flatten can re-parent them). Empty when there are no row fields.
     */
    expandNodes: RowTreeNode[];
}

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
 * Discover the populated fields from the Matrix DataView:
 *   - row fields  from matrix.rows.levels[*].sources (role === rowFields)
 *   - col fields  from matrix.columns.levels[*].sources (role === columnFields)
 *   - value fields from matrix.valueSources (already in projection order)
 *
 * The slot index assigned to each field is its dense position within its bucket
 * in projection order — the same order the engine used to build the hierarchies.
 * Row-field slotIndex therefore equals the matrix hierarchy level, which the
 * downstream per-level subtotal toggles rely on.
 */
export function discoverFields(dataView: DataView | undefined): DiscoveredFields {
    const result: DiscoveredFields = { rows: [], values: [], cols: [] };
    const matrix: DataViewMatrix | undefined = dataView && dataView.matrix;
    if (!matrix) {
        return result;
    }

    let synthIndex = 0;
    let rowSlot = 0;
    let colSlot = 0;
    let valSlot = 0;

    const rowLevels = (matrix.rows && matrix.rows.levels) || [];
    rowLevels.forEach((level: DataViewHierarchyLevel) => {
        (level.sources || []).forEach((src: DataViewMetadataColumn) => {
            if (src.roles && src.roles[ROLE_ROW_FIELDS]) {
                result.rows.push(makeFieldMeta("row", rowSlot++, synthIndex++, src));
            }
        });
    });

    const colLevels = (matrix.columns && matrix.columns.levels) || [];
    colLevels.forEach((level: DataViewHierarchyLevel) => {
        (level.sources || []).forEach((src: DataViewMetadataColumn) => {
            // The innermost column level holds the measures (role === values);
            // only true column-group fields belong in the pivot bucket.
            if (src.roles && src.roles[ROLE_COLUMN_FIELDS]) {
                result.cols.push(makeFieldMeta("col", colSlot++, synthIndex++, src));
            }
        });
    });

    const valueSources = matrix.valueSources || [];
    valueSources.forEach((src: DataViewMetadataColumn) => {
        result.values.push(makeFieldMeta("value", valSlot++, synthIndex++, src));
    });

    return result;
}

// ---------------------------------------------------------------------------
// Formatting helpers.
// ---------------------------------------------------------------------------

function makeFormatter(formatString: string): valueFormatter.IValueFormatter {
    return valueFormatter.create({ format: formatString || undefined });
}

function formatLabel(
    raw: PrimitiveValue | null | undefined,
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

/** Normalize a raw matrix cell into the stored value union (keeps text). */
function normalizeCell(raw: PrimitiveValue | null | undefined): CellValue {
    if (raw === null || raw === undefined) {
        return null;
    }
    if (typeof raw === "number" || typeof raw === "string") {
        return raw;
    }
    // boolean / Date / other — render as text; numeric consumers ignore it.
    return String(raw);
}

function leafColId(pivotPathKey: string, valueSlotIndex: number): string {
    return `${pivotPathKey}${ID_SEP}v${valueSlotIndex}`;
}

// ---------------------------------------------------------------------------
// Column (pivot) plan — built from the matrix columns hierarchy.
// ---------------------------------------------------------------------------

/** Per-column-leaf descriptor recorded in matrix DFS (values-key) order. */
interface ColLeafInfo {
    path: string[];
    isColSubtotal: boolean;
    isColGrandTotal: boolean;
}

interface ColumnPlan {
    leafColumns: LeafColumn[];
    columnHeader: ColumnHeaderLayout;
    /** DFS-ordinal -> column-leaf descriptor; keys match matrix node value keys. */
    ordinalInfo: ColLeafInfo[];
}

function pivotPathKey(path: string[]): string {
    return "c:" + path.join(PATH_SEP);
}

/** Compose the stable leaf-column id from a column descriptor + measure slot. */
function colIdFor(info: ColLeafInfo, valueSlotIndex: number): string {
    if (info.isColGrandTotal) {
        return leafColId(GRANDTOTAL_TOKEN, valueSlotIndex);
    }
    if (info.isColSubtotal) {
        return leafColId(pivotPathKey(info.path) + PATH_SEP + SUBTOTAL_TOKEN, valueSlotIndex);
    }
    if (info.path.length > 0) {
        return leafColId(pivotPathKey(info.path), valueSlotIndex);
    }
    return leafColId("", valueSlotIndex);
}

function buildColumnPlan(
    matrix: DataViewMatrix | undefined,
    colFields: FieldMeta[],
    valueFields: FieldMeta[]
): ColumnPlan {
    const colFieldCount = colFields.length;
    const columns = matrix && matrix.columns;
    const root = columns && columns.root;
    const levels = (columns && columns.levels) || [];

    // Formatter per column-group level (single source per level assumed).
    const levelFormatters: (valueFormatter.IValueFormatter | undefined)[] = [];
    for (let i = 0; i < colFieldCount; i++) {
        const src = levels[i] && levels[i].sources && levels[i].sources[0];
        levelFormatters.push(src ? makeFormatter(src.format != null ? String(src.format) : "") : undefined);
    }

    const ordinalInfo: ColLeafInfo[] = [];
    const pivotRows: ColumnHeaderCell[][] = [];
    for (let i = 0; i < colFieldCount; i++) {
        pivotRows.push([]);
    }

    // Walk the engine column hierarchy. Group nodes sit at levels 0..colFieldCount-1;
    // measure nodes sit at level === colFieldCount and are the DFS leaves whose
    // ordinal position matches the keys in each row node's `values` map.
    const walk = (
        node: DataViewMatrixNode,
        path: string[],
        subtotal: boolean,
        grandTotal: boolean,
        parentIsRoot: boolean
    ): void => {
        const children = node.children || [];
        for (let i = 0; i < children.length; i++) {
            const child = children[i];
            const childLevel = child.level != null ? child.level : path.length;
            const isMeasureLeaf = childLevel >= colFieldCount || !child.children || child.children.length === 0;

            if (isMeasureLeaf && childLevel >= colFieldCount) {
                // A measure leaf — one DFS ordinal, inheriting the parent's descriptor.
                ordinalInfo.push({ path: path.slice(), isColSubtotal: subtotal, isColGrandTotal: grandTotal });
                continue;
            }

            // A column-group node (or, defensively, a leaf when there are no measures).
            const childSubtotal = subtotal || !!child.isSubtotal;
            // A subtotal directly under root spans every column => grand-total column.
            const childGrandTotal = grandTotal || (!!child.isSubtotal && parentIsRoot);
            const label = child.isSubtotal
                ? ""
                : formatLabel(matrixNodeRaw(child), levelFormatters[childLevel]);
            const childPath = child.isSubtotal ? path.slice() : path.concat(label);

            const startLeaf = ordinalInfo.length;
            walk(child, childPath, childSubtotal, childGrandTotal, false);
            const span = ordinalInfo.length - startLeaf;

            if (span > 0 && childLevel < colFieldCount) {
                const cell: ColumnHeaderCell = {
                    label: child.isSubtotal ? "" : label,
                    span,
                    level: childLevel,
                    isSubtotal: !!child.isSubtotal && !childGrandTotal,
                    isGrandTotal: childGrandTotal
                };
                pivotRows[childLevel].push(cell);
            }
        }
    };

    if (root) {
        walk(root, [], false, false, true);
    }

    // Build the leaf columns (visible value fields, in config order) per distinct
    // column position, preserving DFS order.
    const leafColumns: LeafColumn[] = [];
    if (colFieldCount === 0) {
        // Non-pivot: one leaf column per visible value field.
        valueFields.forEach((vf) => {
            leafColumns.push({
                id: leafColId("", vf.slotIndex),
                valueSlotIndex: vf.slotIndex,
                pivotPath: [],
                isColSubtotal: false,
                isColGrandTotal: false
            });
        });
    } else {
        const seen = new Set<string>();
        ordinalInfo.forEach((info) => {
            const posKey =
                (info.isColGrandTotal ? "G" : info.isColSubtotal ? "S" : "N") + "|" + info.path.join(PATH_SEP);
            if (seen.has(posKey)) {
                return;
            }
            seen.add(posKey);
            valueFields.forEach((vf) => {
                leafColumns.push({
                    id: colIdFor(info, vf.slotIndex),
                    valueSlotIndex: vf.slotIndex,
                    pivotPath: info.path.slice(),
                    isColSubtotal: info.isColSubtotal,
                    isColGrandTotal: info.isColGrandTotal
                });
            });
        });
    }

    return {
        leafColumns,
        columnHeader: { pivotRows, leafRow: leafColumns },
        ordinalInfo
    };
}

// ---------------------------------------------------------------------------
// Row tree construction (direct walk of the matrix rows hierarchy).
// ---------------------------------------------------------------------------

/** The raw group value of a matrix node (levelValues preferred over deprecated value). */
function matrixNodeRaw(node: DataViewMatrixNode): PrimitiveValue | null {
    if (node.levelValues && node.levelValues.length > 0 && node.levelValues[0].value != null) {
        return node.levelValues[0].value;
    }
    return node.value != null ? node.value : null;
}

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
        parent
    };
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
    /**
     * Matrix row-hierarchy levels the user has hidden via the config panel. Nodes
     * at these levels are flattened out of the display tree (their children are
     * re-parented up and their subtotal rows suppressed) while remaining in the
     * matrix chain so engine values and selection ids stay correct.
     */
    hiddenRowLevels: Set<number> = new Set()
): TransformResult {
    const matrix: DataViewMatrix | undefined = dataView && dataView.matrix;
    const isPivot = activeColFields.length > 0;
    const hasRowFields = activeRowFields.length > 0;

    const plan = buildColumnPlan(matrix, activeColFields, activeValueFields);

    // Read a matrix node's intersection values into our keyed value map. The keys
    // of node.values are the DFS ordinals of the column-hierarchy leaves; each
    // cell's valueSourceIndex selects the measure slot.
    const readNodeValues = (node: DataViewMatrixNode | undefined): { [id: string]: CellValue } => {
        const out: { [id: string]: CellValue } = {};
        const mv = node && node.values;
        if (!mv) {
            return out;
        }
        for (const k in mv) {
            const cell = mv[k];
            const info = plan.ordinalInfo[Number(k)];
            const descriptor: ColLeafInfo = info || { path: [], isColSubtotal: false, isColGrandTotal: false };
            const measureSlot = cell.valueSourceIndex != null ? cell.valueSourceIndex : 0;
            out[colIdFor(descriptor, measureSlot)] = normalizeCell(cell.value);
        }
        return out;
    };

    // Grand total: the root's own values, or its subtotal child when present.
    const grandTotal: RowTreeNode = newNode("__grandTotal__", "Grand Total", null, -1, undefined);
    grandTotal.isGrandTotal = true;
    const rowRoot: DataViewMatrixNode | undefined = matrix && matrix.rows && matrix.rows.root;
    if (rowRoot) {
        const rootChildren = rowRoot.children || [];
        const rootSubtotal = rootChildren.filter((c) => c.isSubtotal)[0];
        const gtValues = readNodeValues(rootSubtotal || rowRoot);
        grandTotal.values = gtValues;
        grandTotal.matrixNode = rootSubtotal || rowRoot;
        grandTotal.identity = (rootSubtotal || rowRoot).identity;
    }

    if (!hasRowFields || !rowRoot) {
        // No row grouping: the whole grid is a single (grand total) row.
        const synthetic = newNode("__all__", settings.subtotals.labelText || "Total", null, 0, undefined);
        synthetic.isLeaf = true;
        synthetic.values = grandTotal.values;
        synthetic.matrixNode = grandTotal.matrixNode;
        synthetic.identity = grandTotal.identity;
        return {
            rootNodes: [],
            grandTotal,
            leafColumns: plan.leafColumns,
            columnHeader: plan.columnHeader,
            activeRowFields,
            activeValueFields,
            activeColFields,
            rowCount: 0,
            hasRowFields: false,
            isPivot,
            expandNodes: []
        };
    }

    // Row-group level formatters (single source per level).
    const rowLevels = (matrix && matrix.rows && matrix.rows.levels) || [];
    const rowLevelFormatters: (valueFormatter.IValueFormatter | undefined)[] = rowLevels.map((lvl) => {
        const src = lvl.sources && lvl.sources[0];
        return src ? makeFormatter(src.format != null ? String(src.format) : "") : undefined;
    });

    let leafCount = 0;
    // Flat list of every expandable matrix node (visible + hidden level), for
    // Expand/Collapse-All progression by matrix level.
    const expandNodes: RowTreeNode[] = [];

    /**
     * Build the display node(s) for one matrix node. Returns an ARRAY so that a
     * node at a HIDDEN level contributes its (recursively flattened) children in
     * its place rather than a node of its own — the re-parent-to-grandparent step.
     *
     * @param mnode         the matrix node
     * @param parent        the nearest VISIBLE display ancestor (re-parent target)
     * @param displayLevel  compacted 0-based depth among VISIBLE levels
     * @param chain         full matrix ancestor chain root→mnode (hidden included)
     * @param keyPrefix     serialized ancestor path (hidden levels included) so
     *                      re-parented siblings keep unique keys
     */
    const buildNodes = (
        mnode: DataViewMatrixNode,
        parent: RowTreeNode | undefined,
        displayLevel: number,
        chain: DataViewMatrixNode[],
        keyPrefix: string
    ): RowTreeNode[] => {
        const matrixLevel = mnode.level != null ? mnode.level : chain.length;
        const raw = matrixNodeRaw(mnode);
        const label = formatLabel(raw, rowLevelFormatters[matrixLevel]);
        const myChain = chain.concat(mnode);
        // Key path always includes hidden levels, so two invoices re-parented from
        // different (hidden) customers under the same group never collide.
        const myKey = keyPrefix + PATH_SEP + label + ID_SEP + matrixLevel;

        const children = mnode.children || [];
        const realChildren = children.filter((c) => !c.isSubtotal);
        const subtotalChild = children.filter((c) => c.isSubtotal)[0];

        // Record every expandable node (collapsible group) for Expand/Collapse-All,
        // whether or not this level is displayed — a hidden collapsed level must
        // still be reachable so its children can materialise.
        if (mnode.isCollapsed !== undefined) {
            const stub = newNode(myKey, label, raw, displayLevel, parent);
            stub.matrixNode = mnode;
            stub.matrixChain = myChain;
            stub.matrixLevel = matrixLevel;
            stub.isCollapsed = mnode.isCollapsed;
            expandNodes.push(stub);
        }

        if (hiddenRowLevels.has(matrixLevel)) {
            // Hidden level: emit no node of our own; splice this node's real
            // children up to the current display parent at the SAME display level.
            // The subtotal child is dropped (its row is suppressed). A collapsed
            // hidden node has no children yet, so it contributes nothing until the
            // user expands it (via Expand All / drill), which is coherent with the
            // host: the children simply are not in the DataView while collapsed.
            const out: RowTreeNode[] = [];
            realChildren.forEach((child) => {
                out.push(...buildNodes(child, parent, displayLevel, myChain, myKey));
            });
            return out;
        }

        const node = newNode(myKey, label, raw, displayLevel, parent);
        node.identity = mnode.identity;
        node.matrixNode = mnode;
        node.matrixChain = myChain;
        node.matrixLevel = matrixLevel;
        node.isCollapsed = mnode.isCollapsed;

        // Leaf-ness is driven by the host's isCollapsed flag, NOT the projected
        // level count. With expand/collapse active a collapsed group arrives with
        // NO children, so a level-based test would wrongly classify every collapsed
        // group as a leaf and draw no +/- control.
        //   isCollapsed === true  -> collapsed group (expandable)
        //   isCollapsed === false -> expanded group
        //   isCollapsed undefined -> a true leaf (or, defensively, a fully-
        //                            delivered internal node that still has children)
        node.isLeaf = mnode.isCollapsed === undefined && realChildren.length === 0;
        if (node.isLeaf) {
            node.values = readNodeValues(mnode);
            leafCount++;
        } else {
            // Group header shows the engine-computed aggregate carried on the
            // matching subtotal node (true-scope value, NOT a sum of children).
            node.values = subtotalChild ? readNodeValues(subtotalChild) : readNodeValues(mnode);
            realChildren.forEach((child) => {
                node.children.push(...buildNodes(child, node, displayLevel + 1, myChain, myKey));
            });
        }
        return [node];
    };

    const roots: RowTreeNode[] = [];
    (rowRoot.children || [])
        .filter((c) => !c.isSubtotal)
        .forEach((c) => roots.push(...buildNodes(c, undefined, 0, [], "")));

    return {
        rootNodes: roots,
        grandTotal,
        leafColumns: plan.leafColumns,
        columnHeader: plan.columnHeader,
        activeRowFields,
        activeValueFields,
        activeColFields,
        rowCount: leafCount,
        hasRowFields: true,
        isPivot,
        expandNodes
    };
}
