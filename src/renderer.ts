/*
 * renderer.ts
 * -----------
 * D3-driven table rendering: sticky headers (row-field columns + value/pivot
 * column headers), the virtualized row body, and per-cell value formatting and
 * conditional formatting. The renderer owns the single scroll container that
 * scrolls on both axes; headers are kept on screen with position:sticky and the
 * VirtualScroller reserves the header height as a top offset.
 *
 * The renderer is intentionally framework-free (no React/Angular/Vue). D3 is used
 * for building the header DOM; the perf-critical recycled body rows are built
 * with direct DOM calls inside the VirtualScroller's render callback.
 */

import * as d3 from "d3";
import { valueFormatter } from "powerbi-visuals-utils-formattingutils";

import {
    TransformResult,
    RowTreeNode,
    LeafColumn,
    FieldMeta
} from "./dataTransformer";
import {
    VisualSettings,
    ValueFormatSettings,
    SpecificColumnSettings,
    CFSettings,
    DisplayUnit,
    ExpandStyle,
    ColumnAlignment,
    ColumnApplyTo
} from "./settings";
import { SortManager, SortDirection } from "./sortManager";
import { VisualSelectionManager, ClickModifiers } from "./selectionManager";
import { VirtualScroller } from "./virtualScroller";
import { ConditionalFormatter, computeDomains, Domain } from "./conditionalFormatter";

// Column geometry.
const DEFAULT_ROW_FIELD_WIDTH = 160; // default row-field column width (resizable)
const DEFAULT_COL_WIDTH = 120; // default value-column width (resizable, Fix 3)
const MIN_COL_WIDTH = 40;
const HEADER_ROW_HEIGHT = 30;
const DEFAULT_ROW_HEIGHT = 28;
const ASC_ARROW = "▲";
const DESC_ARROW = "▼";

/** A persisted column width entry. `type` distinguishes row-field vs value columns. */
export interface ColumnWidth {
    type: "row" | "val";
    index: number;
    width: number;
}

export interface ThemeColors {
    foreground: string;
    background: string;
    headerBackground: string;
    headerForeground: string;
    rowBorder: string;
    subtotalBackground: string;
    selectionFill: string;
    accent: string;
}

export type DisplayRowKind = "group" | "leaf" | "subtotal" | "grandtotal";

export interface DisplayRow {
    node: RowTreeNode;
    kind: DisplayRowKind;
    level: number;
    /** Whether this row shows aggregated values (group headers when collapsed). */
    showValues: boolean;
    /** Whether this row is selectable (group/leaf only). */
    selectable: boolean;
}

export interface RenderInput {
    transform: TransformResult;
    settings: VisualSettings;
    expanded: Set<string>;
    theme: ThemeColors;
    sort: SortManager;
    selection: VisualSelectionManager;
    valueFormatFor: (slot: number) => ValueFormatSettings;
    specificColumnFor: (slot: number) => SpecificColumnSettings;
    cfFor: (slot: number) => CFSettings;
    leafLabelFor: (col: LeafColumn) => string;
    onRowFieldSort: (level: number, label: string) => void;
    onValueSort: (leafColId: string, label: string) => void;
    onColumnSort: () => void;
    onToggleExpand: (node: RowTreeNode) => void;
    onRowClick: (node: RowTreeNode, mods: ClickModifiers) => void;
    onEmptyClick: () => void;
    /** Right-click on a value column header opens the in-visual CF panel. */
    onColumnRightClick: (slotIndex: number, displayName: string) => void;
    /** Restored persisted column widths (Fix 3D). */
    columnWidths: ColumnWidth[];
    /** Called when the user finishes resizing a column, to persist the widths. */
    onColumnWidthsChanged: (widths: ColumnWidth[]) => void;
}

type ValueFormatFn = (value: number | null) => string;

export class Renderer {
    private readonly scrollEl: HTMLElement;
    private readonly headerEl: HTMLElement;
    private readonly scroller: VirtualScroller<DisplayRow>;

    private readonly cf = new ConditionalFormatter();

    // Per-render state captured for the recycled-row binder.
    private current: RenderInput | null = null;
    private leftOffsets: number[] = [];
    private rowFieldTotalWidth = 0;
    private contentWidth = 0;
    private valueFormatters = new Map<number, ValueFormatFn>();
    private valueFieldBySlot = new Map<number, FieldMeta>();
    private domains = new Map<string, Domain>();
    private headerHeight = HEADER_ROW_HEIGHT;
    private rowHeight = DEFAULT_ROW_HEIGHT;

    // Resizable column-width state (Fix 3). Keyed by leaf-column index.
    private columnWidths = new Map<number, number>();
    // Resizable row-field column widths, keyed by row-field column index.
    private rowFieldWidths = new Map<number, number>();
    private leafWidths: number[] = [];
    private leafLefts: number[] = [];
    private valueRegionWidth = 0;
    private relayoutScheduled = false;
    /** Number of row-field columns rendered (depends on layout mode). */
    private rowFieldColumnCount = 1;

    constructor(host: HTMLElement) {
        this.scrollEl = document.createElement("div");
        this.scrollEl.className = "nsm-scroll";
        host.appendChild(this.scrollEl);

        this.headerEl = document.createElement("div");
        this.headerEl.className = "nsm-header";
        this.scrollEl.appendChild(this.headerEl);

        this.scroller = new VirtualScroller<DisplayRow>(this.scrollEl);
        this.scroller.setRenderRow((el, item, index) => this.bindRow(el, item, index));

        // Empty-space click clears selection.
        this.scrollEl.addEventListener("click", (e) => {
            if (e.target === this.scrollEl || (e.target as HTMLElement) === this.headerEl) {
                if (this.current) {
                    this.current.onEmptyClick();
                }
            }
        });
    }

    public destroy(): void {
        this.scroller.destroy();
        if (this.scrollEl.parentNode) {
            this.scrollEl.parentNode.removeChild(this.scrollEl);
        }
    }

    public setViewportHeight(px: number): void {
        this.scrollEl.style.height = Math.max(0, px) + "px";
    }

    /** Reset scroll to top (data refresh, sort change, structural expand/collapse). */
    public scrollToTop(): void {
        this.scroller.scrollToTop();
    }

    public render(input: RenderInput): void {
        this.current = input;
        // Load persisted column widths from the host. Done only on the public
        // render path (driven by Visual.update) — live resize re-renders go
        // through relayout() and must NOT clobber the in-progress edits with the
        // not-yet-persisted values.
        this.loadColumnWidths(input.columnWidths);
        this.layout();
    }

    /** Re-run geometry + header + body using the current input (no width reload). */
    private layout(): void {
        const input = this.current;
        if (!input) {
            return;
        }
        const t = input.transform;

        // Theme + typography on the scroll root. Font family is applied per
        // element via CSS variables on the visual's root (see Visual.update).
        this.scrollEl.style.background = input.theme.background;
        this.scrollEl.style.color = input.theme.foreground;
        this.scrollEl.style.fontSize = input.settings.grid.fontSize + "px";

        // Geometry.
        this.computeGeometry(t);

        // Value formatters + CF domains.
        this.buildFormatters(input);
        const leafNodes = this.collectLeafNodes(t);
        this.domains = computeDomains(leafNodes, t.leafColumns, input.cfFor);

        // Header.
        this.buildHeader(input);

        // Body.
        const rowHeight =
            input.settings.grid.rowHeightMode === "fixed"
                ? input.settings.grid.rowHeightPx
                : DEFAULT_ROW_HEIGHT;
        this.rowHeight = rowHeight;
        this.scroller.configure(rowHeight, 20);
        // Use the ACTUAL rendered header height for the sticky top offset so body
        // rows never start under the header. Fall back to the computed height if
        // the element hasn't been laid out yet (measured height of 0).
        const measuredHeader = this.headerEl.getBoundingClientRect().height;
        const effectiveHeaderHeight = measuredHeader > 0 ? measuredHeader : this.headerHeight;
        this.scroller.setTopOffset(effectiveHeaderHeight);
        this.scroller.setContentWidth(this.contentWidth);

        const displayRows = this.flatten(input);
        this.scroller.setItems(displayRows);

        // Register selectable nodes (group/leaf) in display order.
        const selectables = displayRows.filter((r) => r.selectable).map((r) => r.node);
        input.selection.setSelectables(selectables);
    }

    // -----------------------------------------------------------------------
    // Column width state + resize interaction (Fix 3).
    // -----------------------------------------------------------------------

    private loadColumnWidths(widths: ColumnWidth[]): void {
        this.columnWidths.clear();
        this.rowFieldWidths.clear();
        if (widths) {
            widths.forEach((w) => {
                if (w && isFinite(w.index) && isFinite(w.width) && w.width >= MIN_COL_WIDTH) {
                    if (w.type === "row") {
                        this.rowFieldWidths.set(w.index, w.width);
                    } else {
                        this.columnWidths.set(w.index, w.width);
                    }
                }
            });
        }
    }

    private getColWidth(colIndex: number): number {
        const w = this.columnWidths.get(colIndex);
        return w !== undefined ? w : DEFAULT_COL_WIDTH;
    }

    private getRowFieldWidth(index: number): number {
        const w = this.rowFieldWidths.get(index);
        return w !== undefined ? w : DEFAULT_ROW_FIELD_WIDTH;
    }

    /**
     * Update one column width and re-lay-out live (no width reload, no persist).
     * A NEGATIVE colIndex addresses a row-field column: row-field column k is
     * encoded as -(k + 1).
     */
    private setColumnWidth(colIndex: number, width: number): void {
        const w = Math.max(MIN_COL_WIDTH, Math.round(width));
        if (colIndex < 0) {
            this.rowFieldWidths.set(Math.abs(colIndex) - 1, w);
        } else {
            this.columnWidths.set(colIndex, w);
        }
        if (this.relayoutScheduled) {
            return;
        }
        this.relayoutScheduled = true;
        requestAnimationFrame(() => {
            this.relayoutScheduled = false;
            // layout() recomputes geometry (leftOffsets/rowFieldTotalWidth) and re-renders.
            this.layout();
        });
    }

    /** Serialize all widths (row-field + value) and hand them to the host. */
    private commitColumnWidths(): void {
        if (!this.current) {
            return;
        }
        const widths: ColumnWidth[] = [];
        this.rowFieldWidths.forEach((width, index) => widths.push({ type: "row", index, width }));
        this.columnWidths.forEach((width, index) => widths.push({ type: "val", index, width }));
        this.current.onColumnWidthsChanged(widths);
    }

    /**
     * Append a drag-to-resize handle to a header cell. `colIndex` is a value
     * leaf-column index, or -(k + 1) to address row-field column k.
     */
    private appendResizeHandle(cellNode: HTMLElement, colIndex: number): void {
        const handle = document.createElement("div");
        handle.className = "nsm-col-resize-handle";
        // A bare click on the handle must not trigger the cell's sort handler.
        handle.addEventListener("click", (e: MouseEvent) => e.stopPropagation());
        handle.addEventListener("mousedown", (event: MouseEvent) => {
            event.preventDefault();
            event.stopPropagation();
            const startX = event.clientX;
            const startWidth =
                colIndex < 0 ? this.getRowFieldWidth(Math.abs(colIndex) - 1) : this.getColWidth(colIndex);
            const onMove = (m: MouseEvent): void => {
                const newWidth = Math.max(MIN_COL_WIDTH, startWidth + (m.clientX - startX));
                this.setColumnWidth(colIndex, newWidth);
            };
            const onUp = (): void => {
                document.removeEventListener("mousemove", onMove);
                document.removeEventListener("mouseup", onUp);
                this.commitColumnWidths();
            };
            document.addEventListener("mousemove", onMove);
            document.addEventListener("mouseup", onUp);
        });
        cellNode.appendChild(handle);
    }

    /**
     * Append a resize handle to a pivot SPANNING header cell. Dragging it
     * distributes the width change proportionally across the leaf columns it
     * spans (indices [leafStart, leafStart + span)).
     */
    private appendPivotResizeHandle(cellNode: HTMLElement, leafStart: number, span: number): void {
        const handle = document.createElement("div");
        handle.className = "nsm-col-resize-handle";
        handle.addEventListener("click", (e: MouseEvent) => e.stopPropagation());
        handle.addEventListener("mousedown", (event: MouseEvent) => {
            event.preventDefault();
            event.stopPropagation();
            const startX = event.clientX;
            const orig: number[] = [];
            let totalWidth = 0;
            for (let s = 0; s < span; s++) {
                const w = this.getColWidth(leafStart + s);
                orig.push(w);
                totalWidth += w;
            }
            if (totalWidth <= 0) {
                return;
            }
            const onMove = (m: MouseEvent): void => {
                const delta = m.clientX - startX;
                const newTotal = Math.max(span * MIN_COL_WIDTH, totalWidth + delta);
                for (let s = 0; s < span; s++) {
                    const newLeafWidth = Math.round((orig[s] / totalWidth) * newTotal);
                    this.setColumnWidth(leafStart + s, newLeafWidth);
                }
            };
            const onUp = (): void => {
                document.removeEventListener("mousemove", onMove);
                document.removeEventListener("mouseup", onUp);
                this.commitColumnWidths();
            };
            document.addEventListener("mousemove", onMove);
            document.addEventListener("mouseup", onUp);
        });
        cellNode.appendChild(handle);
    }

    // -----------------------------------------------------------------------
    // Geometry / formatters.
    // -----------------------------------------------------------------------

    private computeGeometry(t: TransformResult): void {
        // Layout mode determines how many row-field columns are rendered.
        const mode = this.current ? this.current.settings.rowHeaders.layoutMode : "compact";
        const columnCount = mode === "compact" ? 1 : t.hasRowFields ? t.activeRowFields.length : 1;
        this.rowFieldColumnCount = columnCount;

        this.leftOffsets = [];
        let x = 0;
        for (let i = 0; i < columnCount; i++) {
            this.leftOffsets.push(x);
            x += this.getRowFieldWidth(i);
        }
        this.rowFieldTotalWidth = x;

        // Per-leaf-column widths and absolute left offsets (Fix 3). Widths come
        // from the resizable map (default 120px) so they stay aligned between
        // the header cells and the body cells.
        this.leafWidths = [];
        this.leafLefts = [];
        let lx = this.rowFieldTotalWidth;
        for (let i = 0; i < t.leafColumns.length; i++) {
            const w = this.getColWidth(i);
            this.leafLefts.push(lx);
            this.leafWidths.push(w);
            lx += w;
        }
        this.valueRegionWidth = lx - this.rowFieldTotalWidth;
        this.contentWidth = lx;

        // Header height accounts for pivot levels + the measure row.
        const pivotLevels = t.columnHeader.pivotRows.length;
        this.headerHeight = (pivotLevels + 1) * HEADER_ROW_HEIGHT;
    }

    private buildFormatters(input: RenderInput): void {
        this.valueFormatters = new Map<number, ValueFormatFn>();
        this.valueFieldBySlot = new Map<number, FieldMeta>();
        input.transform.activeValueFields.forEach((field) => {
            const vfs = input.valueFormatFor(field.slotIndex);
            this.valueFormatters.set(field.slotIndex, this.makeFormatter(field, vfs));
            this.valueFieldBySlot.set(field.slotIndex, field);
        });
    }

    private makeFormatter(field: FieldMeta, vfs: ValueFormatSettings): ValueFormatFn {
        const unitMap: { [k in DisplayUnit]: number } = {
            none: 0,
            thousands: 1e3,
            millions: 1e6,
            billions: 1e9
        };
        let resolvedFormat = vfs.numberFormat || field.formatString || undefined;

        // If decimals is explicitly set, embed it into the format string directly.
        // valueFormatter ignores opts.precision when opts.format is present,
        // so we must modify the format string itself.
        if (vfs.decimals >= 0 && resolvedFormat) {
            resolvedFormat = this.applyDecimalsToFormat(resolvedFormat, vfs.decimals);
        }

        const opts: valueFormatter.ValueFormatterOptions = {
            format: resolvedFormat
        };

        // Only use precision when there is no format string to override it.
        if (vfs.decimals >= 0 && !resolvedFormat) {
            opts.precision = vfs.decimals;
        }
        const unitVal = unitMap[vfs.unit];
        if (unitVal) {
            opts.value = unitVal;
        }
        const formatter = valueFormatter.create(opts);
        const prefix = vfs.prefix || "";
        const suffix = vfs.suffix || "";
        return (value: number | null): string => {
            if (value === null || value === undefined || isNaN(value)) {
                return "";
            }
            return prefix + formatter.format(value) + suffix;
        };
    }

    /**
     * Embed a decimal count directly into a Power BI format string.
     * valueFormatter.create() ignores opts.precision when opts.format is set,
     * so we must modify the format string itself to control decimal places.
     * Handles multi-section format strings (positive;negative;zero;text).
     */
    private applyDecimalsToFormat(fmt: string, decimals: number): string {
        const decimalSuffix = decimals > 0 ? "." + "0".repeat(decimals) : "";
        return fmt
            .split(";")
            .map((section) => {
                // Strip any existing decimal section (e.g. .00, .000, .##, .0#)
                const stripped = section.replace(/\.[\d#,]*/g, "");
                if (decimals === 0) {
                    return stripped;
                }
                // Re-insert decimal section after the last digit placeholder (0 or #)
                // before any suffix characters (%, space, currency symbols, etc.)
                const match = stripped.match(/^(.*[0#])([^0#]*)$/);
                if (match) {
                    return match[1] + decimalSuffix + match[2];
                }
                // No digit placeholder found — append decimal suffix
                return stripped + decimalSuffix;
            })
            .join(";");
    }

    private collectLeafNodes(t: TransformResult): RowTreeNode[] {
        const out: RowTreeNode[] = [];
        const stack: RowTreeNode[] = t.hasRowFields ? t.rootNodes.slice() : [];
        if (!t.hasRowFields && t.grandTotal) {
            out.push(t.grandTotal);
            return out;
        }
        while (stack.length) {
            const n = stack.pop() as RowTreeNode;
            if (n.isLeaf) {
                out.push(n);
            } else {
                for (let i = 0; i < n.children.length; i++) {
                    stack.push(n.children[i]);
                }
            }
        }
        return out;
    }

    // -----------------------------------------------------------------------
    // Flatten tree -> display rows given expand state.
    // -----------------------------------------------------------------------

    private flatten(input: RenderInput): DisplayRow[] {
        const t = input.transform;
        const rows: DisplayRow[] = [];

        if (!t.hasRowFields) {
            if (t.grandTotal) {
                rows.push({
                    node: t.grandTotal,
                    kind: "grandtotal",
                    level: 0,
                    showValues: true,
                    selectable: true
                });
            }
            return rows;
        }

        const subtotalOn = (level: number): boolean => {
            if (!input.settings.subtotals.rowSubtotals) {
                return false;
            }
            const lv = input.settings.subtotals.levels;
            return level < lv.length ? lv[level] : true;
        };

        const mode = input.settings.rowHeaders.layoutMode;

        const walk = (node: RowTreeNode): void => {
            if (node.isLeaf) {
                rows.push({
                    node,
                    kind: "leaf",
                    level: node.level,
                    showValues: true,
                    selectable: true
                });
                return;
            }
            const expanded = input.expanded.has(node.key);
            const stEnabled = subtotalOn(node.level);

            // In tabular mode, group rows are suppressed — leaf rows carry the full
            // ancestor label chain across columns, and subtotals serve as group
            // summaries. Rendering group rows in tabular mode creates visual
            // duplication because the group aggregate appears both in the group row
            // and the subtotal row. Recurse into children only; still emit the
            // subtotal so each group has its summary line.
            if (mode === "tabular") {
                node.children.forEach(walk);
                if (stEnabled) {
                    rows.push({
                        node,
                        kind: "subtotal",
                        level: node.level,
                        showValues: true,
                        selectable: false
                    });
                }
                return;
            }

            rows.push({
                node,
                kind: "group",
                level: node.level,
                // Group header rows always show their own rolled-up aggregated
                // values (from rollUp() in dataTransformer). The subtotal row
                // beneath an expanded group is an additive labeled row — it does
                // not replace the group header's own value display.
                showValues: true,
                selectable: true
            });
            if (expanded) {
                node.children.forEach(walk);
                if (stEnabled) {
                    rows.push({
                        node,
                        kind: "subtotal",
                        level: node.level,
                        showValues: true,
                        selectable: false
                    });
                }
            }
        };

        t.rootNodes.forEach(walk);

        if (input.settings.subtotals.grandTotalRow && t.grandTotal) {
            rows.push({
                node: t.grandTotal,
                kind: "grandtotal",
                level: 0,
                showValues: true,
                selectable: false
            });
        }
        return rows;
    }

    // -----------------------------------------------------------------------
    // Header construction (D3).
    // -----------------------------------------------------------------------

    private buildHeader(input: RenderInput): void {
        const t = input.transform;
        const theme = input.theme;
        const headerSel = d3.select(this.headerEl);
        headerSel.selectAll("*").remove();
        this.headerEl.style.height = this.headerHeight + "px";
        this.headerEl.style.width = this.contentWidth + "px";
        this.headerEl.style.background = theme.headerBackground;
        this.headerEl.style.color = theme.headerForeground;

        const colHdr = input.settings.columnHeaders;

        // --- Row-field header columns. ---
        // In tabular mode only level 0 stays a frozen sticky corner; levels 1+
        // become regular scrollable column headers on the bottom header row.
        const isTabular = input.settings.rowHeaders.layoutMode === "tabular";
        const tabularLeafTop = t.columnHeader.pivotRows.length * HEADER_ROW_HEIGHT;
        const count = this.rowFieldColumnCount;
        for (let level = 0; level < count; level++) {
            if (isTabular && level > 0) {
                // Levels 1+ in tabular mode: regular non-sticky header cell,
                // positioned like a value column header (scrolls horizontally).
                const cell = headerSel
                    .append("div")
                    .attr("class", "nsm-hcell nsm-hcell-tabular-rowfield")
                    .style("position", "absolute")
                    .style("left", this.leftOffsets[level] + "px")
                    .style("top", tabularLeafTop + "px") // same row as value headers
                    .style("width", this.getRowFieldWidth(level) + "px")
                    .style("height", HEADER_ROW_HEIGHT + "px")
                    .style("line-height", HEADER_ROW_HEIGHT + "px")
                    .style("font-weight", colHdr.bold ? "700" : "400")
                    .style("font-size", colHdr.fontSize + "px");

                const field: FieldMeta | undefined = t.hasRowFields ? t.activeRowFields[level] : undefined;
                cell.append("span").attr("class", "nsm-hlabel").text(field ? field.displayName : "");
                // No sort handle on tabular non-frozen columns for now.
                const cellNodeTabular = cell.node() as HTMLElement;
                if (cellNodeTabular) this.appendResizeHandle(cellNodeTabular, -(level + 1));
                continue;
            }

            // Level 0 always, or all levels in non-tabular modes: sticky corner.
            const cell = headerSel
                .append("div")
                .attr("class", "nsm-hcell nsm-hcell-rowfield")
                .style("left", this.leftOffsets[level] + "px")
                .style("width", this.getRowFieldWidth(level) + "px")
                .style("height", this.headerHeight + "px")
                .style("line-height", this.headerHeight + "px")
                .style("background", theme.headerBackground)
                .style("font-weight", colHdr.bold ? "700" : "400")
                .style("font-size", colHdr.fontSize + "px");

            const field: FieldMeta | undefined = t.hasRowFields ? t.activeRowFields[level] : undefined;
            const label = field ? field.displayName : "";
            cell.append("span").attr("class", "nsm-hlabel").text(label);

            if (field && colHdr.showSortArrows) {
                const dir = input.sort.directionForRowField(level);
                cell.append("span")
                    .attr("class", "nsm-sortarrow")
                    .text(this.arrow(dir));
            }

            if (field) {
                cell.style("cursor", "pointer").on("click", (event: MouseEvent) => {
                    event.stopPropagation();
                    input.onRowFieldSort(level, field.displayName);
                });
            }

            // Drag-to-resize handle. Row-field column k is encoded as -(k + 1).
            const cellNode = cell.node() as HTMLElement;
            this.appendResizeHandle(cellNode, -(level + 1));
        }

        // --- Value / pivot column headers. ---
        // Pivot spanning rows. A spanning cell's width is the sum of the actual
        // (resizable) widths of the leaf columns it covers, so it stays aligned.
        const pivotRows = t.columnHeader.pivotRows;
        pivotRows.forEach((cells, level) => {
            let leafIdx = 0;
            cells.forEach((c) => {
                const startLeaf = leafIdx;
                let w = 0;
                for (let s = 0; s < c.span; s++) {
                    w += this.leafWidths[leafIdx + s] || DEFAULT_COL_WIDTH;
                }
                const left = this.leafLefts[leafIdx] !== undefined ? this.leafLefts[leafIdx] : this.rowFieldTotalWidth;
                const cell = headerSel
                    .append("div")
                    .attr("class", "nsm-hcell nsm-hcell-pivot")
                    .style("left", left + "px")
                    .style("top", level * HEADER_ROW_HEIGHT + "px")
                    .style("width", w + "px")
                    .style("height", HEADER_ROW_HEIGHT + "px")
                    .style("line-height", HEADER_ROW_HEIGHT + "px")
                    .style("font-weight", colHdr.bold ? "700" : "400")
                    .style("font-size", colHdr.fontSize + "px");
                cell.append("span").attr("class", "nsm-hlabel").text(c.label);
                cell.style("cursor", "pointer").on("click", (event: MouseEvent) => {
                    event.stopPropagation();
                    input.onColumnSort();
                });

                // Resize handle distributes the delta proportionally across the
                // leaf columns this pivot group spans.
                const cellNode = cell.node() as HTMLElement;
                this.appendPivotResizeHandle(cellNode, startLeaf, c.span);

                leafIdx += c.span;
            });
        });

        // Leaf measure header row (bottom-most header line).
        const leafTop = pivotRows.length * HEADER_ROW_HEIGHT;
        t.leafColumns.forEach((col, colIndex) => {
            const cell = headerSel
                .append("div")
                .attr("class", "nsm-hcell nsm-hcell-value")
                .classed("nsm-hcell-subtotal", col.isColSubtotal)
                .classed("nsm-hcell-grandtotal", col.isColGrandTotal)
                .style("left", this.leafLefts[colIndex] + "px")
                .style("top", leafTop + "px")
                .style("width", this.leafWidths[colIndex] + "px")
                .style("height", HEADER_ROW_HEIGHT + "px")
                .style("line-height", HEADER_ROW_HEIGHT + "px")
                .style("font-weight", colHdr.bold ? "700" : "400")
                .style("font-size", colHdr.fontSize + "px");

            // Per-measure value formatting applied to the leaf header cell.
            const hvfs = input.valueFormatFor(col.valueSlotIndex);
            cell
                .style("font-family", hvfs.fontFamily)
                .style("font-size", hvfs.fontSize + "px")
                .style("font-weight", hvfs.bold ? "700" : "400")
                .style("font-style", hvfs.italic ? "italic" : "normal");
            if (hvfs.backgroundColor) {
                cell.style("background", hvfs.backgroundColor);
            }
            if (hvfs.textColor) {
                cell.style("color", hvfs.textColor);
            }

            // Specific column final override for the header cell.
            const hsc = input.specificColumnFor(col.valueSlotIndex);
            if (this.specificColumnApplies(hsc.applyTo, "header")) {
                if (hsc.backgroundColor) {
                    cell.style("background", hsc.backgroundColor);
                }
                if (hsc.textColor) {
                    cell.style("color", hsc.textColor);
                }
                if (hsc.alignment === "left") {
                    cell.style("justify-content", "flex-start");
                } else if (hsc.alignment === "center") {
                    cell.style("justify-content", "center");
                } else if (hsc.alignment === "right") {
                    cell.style("justify-content", "flex-end");
                }
            }

            cell.append("span").attr("class", "nsm-hlabel").text(input.leafLabelFor(col));

            if (colHdr.showSortArrows) {
                const dir = input.sort.directionForValue(col.id);
                cell.append("span").attr("class", "nsm-sortarrow").text(this.arrow(dir));
            }
            cell.style("cursor", "pointer").on("click", (event: MouseEvent) => {
                event.stopPropagation();
                input.onValueSort(col.id, input.leafLabelFor(col));
            });

            // Drag-to-resize handle on each value (non-row-field) header cell.
            const cellNode = cell.node() as HTMLElement;
            this.appendResizeHandle(cellNode, colIndex);

            // Right-click opens the in-visual conditional-formatting panel.
            cellNode.addEventListener("contextmenu", (e: MouseEvent) => {
                e.preventDefault();
                e.stopPropagation();
                input.onColumnRightClick(col.valueSlotIndex, input.leafLabelFor(col));
            });
        });
    }

    private arrow(dir: SortDirection | undefined): string {
        if (dir === "asc") {
            return " " + ASC_ARROW;
        }
        if (dir === "desc") {
            return " " + DESC_ARROW;
        }
        return "";
    }

    // -----------------------------------------------------------------------
    // Recycled-row binder.
    // -----------------------------------------------------------------------

    private bindRow(el: HTMLElement, row: DisplayRow, index: number): void {
        const input = this.current;
        if (!input) {
            return;
        }
        const t = input.transform;
        const theme = input.theme;

        // Reset element.
        el.className = "nsm-row nsm-row-" + row.kind;
        el.textContent = "";
        el.style.height = this.rowHeight + "px";
        el.style.width = this.contentWidth + "px";

        // Row background (alternating / subtotal / selection).
        let bg = theme.background;
        if (row.kind === "subtotal" || row.kind === "grandtotal") {
            bg = theme.subtotalBackground;
        } else if (input.settings.alternateRows.show && index % 2 === 1) {
            bg = input.settings.alternateRows.color;
        }
        const selected = row.selectable && input.selection.isSelected(row.node);
        if (selected) {
            bg = theme.selectionFill;
        }
        el.style.background = bg;
        el.style.borderBottom = "1px solid " + theme.rowBorder;
        el.style.opacity = row.selectable && input.selection.isDimmed(row.node) ? "0.45" : "1";

        // --- Row-field cells (absolutely positioned within the row). ---
        const mode = input.settings.rowHeaders.layoutMode;
        const columnCount = this.rowFieldColumnCount;
        const labelLevel = row.kind === "grandtotal" ? 0 : row.level;
        // Repeat ancestor labels on the first visible row (outline/tabular).
        const repeat =
            input.settings.rowHeaders.repeatRowHeaders &&
            (mode === "outline" || mode === "tabular") &&
            this.scroller.getFirstVisibleIndex() === index;
        const isTabular = mode === "tabular";
        for (let c = 0; c < columnCount; c++) {
            const cell = document.createElement("div");
            // In tabular mode only level 0 is frozen; levels 1+ are regular cells.
            const isFrozen = !isTabular || c === 0;
            cell.className = isFrozen
                ? "nsm-cell nsm-cell-rowfield"
                : "nsm-cell nsm-cell-rowfield-tabular";

            cell.style.position = "absolute";
            cell.style.left = this.leftOffsets[c] + "px";
            cell.style.top = "0";
            cell.style.height = "100%";
            cell.style.width = this.getRowFieldWidth(c) + "px";

            // Only the frozen column gets z-index elevation; non-frozen tabular
            // cells behave like ordinary scrolling body cells.
            if (isFrozen) {
                cell.style.zIndex = "20";
            }

            if (mode === "tabular") {
                if (c === labelLevel) {
                    // Own label: always show
                    this.fillRowHeaderCell(cell, row, input, 0, c);
                } else if (c < labelLevel) {
                    // Ancestor label: show only when repeatRowHeaders is on,
                    // OR when this row is the first occurrence under that ancestor group.
                    const shouldShow =
                        input.settings.rowHeaders.repeatRowHeaders ||
                        this.isFirstInGroup(row.node, c);
                    if (shouldShow) {
                        this.fillRowHeaderCell(cell, row, input, 0, c);
                    }
                    // When shouldShow is false, leave the cell blank —
                    // fillRowHeaderCell is not called, so the cell renders empty.
                }
                // c > labelLevel: fillRowHeaderCell already handles this (returns early),
                // but we skip the call entirely for clarity.
            } else if (mode === "outline") {
                // Label only at the row's own level column; repeat ancestors on
                // the first visible row when enabled.
                if (c === labelLevel) {
                    this.fillRowHeaderCell(cell, row, input, 0);
                } else if (repeat && c < labelLevel) {
                    this.fillRowHeaderCell(cell, row, input, 0, c);
                }
            } else {
                // Compact: all labels in column 0, indented by depth.
                if (c === 0) {
                    const indent = row.level * input.settings.rowHeaders.indentPerLevel;
                    this.fillRowHeaderCell(cell, row, input, indent);
                }
            }
            el.appendChild(cell);
        }

        // --- Value cells (always shown; one absolutely-positioned cell each). ---
        t.leafColumns.forEach((col, colIndex) => {
            el.appendChild(this.buildValueCell(col, colIndex, row, input, index));
        });

        // Row click (selection) for selectable rows.
        if (row.selectable) {
            el.style.cursor = "pointer";
            el.onclick = (e: MouseEvent) => {
                e.stopPropagation();
                input.onRowClick(row.node, {
                    ctrlKey: e.ctrlKey,
                    metaKey: e.metaKey,
                    shiftKey: e.shiftKey
                });
            };
        } else {
            el.style.cursor = "default";
            el.onclick = null;
        }
    }

    private fillRowHeaderCell(
        cell: HTMLElement,
        row: DisplayRow,
        input: RenderInput,
        indent: number,
        levelOverride?: number
    ): void {
        cell.style.paddingLeft = 8 + indent + "px";
        cell.style.fontWeight =
            row.kind === "subtotal" || row.kind === "grandtotal" || input.settings.rowHeaders.bold
                ? "700"
                : "400";
        cell.style.fontSize = input.settings.rowHeaders.fontSize + "px";

        // Tabular: a row has no value at a level deeper than itself — leave blank.
        if (levelOverride !== undefined && levelOverride > row.level) {
            return;
        }

        // An ancestor cell (tabular/repeat) shows an ancestor group's label and
        // never an expand/collapse control.
        const isAncestorCell = levelOverride !== undefined && levelOverride < row.level;
        const isOwnLevelCell = !isAncestorCell;

        // Expand/collapse button — only on the row's own-level group cell.
        if (row.kind === "group" && isOwnLevelCell) {
            const ec = input.settings.expandCollapse;
            if (ec.show) {
                const btn = document.createElement("span");
                btn.className = "nsm-expand-btn";
                const expanded = input.expanded.has(row.node.key);
                btn.textContent = this.expandIcon(ec.style, expanded);
                btn.style.fontSize = ec.buttonSize + "px";
                btn.style.color = ec.buttonColor || input.theme.foreground;
                btn.onclick = (e: MouseEvent) => {
                    e.stopPropagation();
                    input.onToggleExpand(row.node);
                };
                cell.appendChild(btn);
            } else {
                const spacer = document.createElement("span");
                spacer.className = "nsm-chevron-spacer";
                cell.appendChild(spacer);
            }
        } else if (row.kind === "leaf" && isOwnLevelCell) {
            const spacer = document.createElement("span");
            spacer.className = "nsm-chevron-spacer";
            cell.appendChild(spacer);
        }

        const label = document.createElement("span");
        label.className = "nsm-rowlabel";
        if (isAncestorCell) {
            label.textContent = this.ancestorLabel(row.node, levelOverride as number);
        } else if (row.kind === "subtotal") {
            label.textContent =
                (input.settings.subtotals.labelText || "Total") + " — " + row.node.label;
        } else if (row.kind === "grandtotal") {
            label.textContent = input.transform.hasRowFields
                ? "Grand Total"
                : input.settings.subtotals.labelText || "Total";
        } else {
            label.textContent = row.node.label;
        }
        label.title = label.textContent || "";
        cell.appendChild(label);
    }

    /** The expand/collapse glyph for a style + state. */
    private expandIcon(style: ExpandStyle, expanded: boolean): string {
        if (style === "chevron") {
            return expanded ? "⌄" : "›";
        }
        if (style === "triangle") {
            return expanded ? "▼" : "▶";
        }
        // plusMinus (default)
        return expanded ? "−" : "+";
    }

    /** Walk the ancestor chain to the node at a given depth and return its label. */
    private ancestorLabel(node: RowTreeNode, level: number): string {
        let n: RowTreeNode | undefined = node;
        while (n && n.level > level) {
            n = n.parent;
        }
        return n ? n.label : "";
    }

    /**
     * Returns true if this node is the first-in-group at the given ancestor level.
     * Walks the ancestor chain from the node up to ancestorLevel, checking at each
     * step whether the node is the first child of its parent. If any step is not
     * first-child, this is not the first occurrence under that ancestor group.
     * Works correctly with virtual scrolling — requires no external state.
     */
    private isFirstInGroup(node: RowTreeNode, ancestorLevel: number): boolean {
        let current: RowTreeNode | undefined = node;
        while (current && current.level > ancestorLevel) {
            const p: RowTreeNode | undefined = current.parent;
            if (!p) return true;
            if (p.children[0] !== current) return false;
            current = p;
        }
        return true;
    }

    private buildValueCell(
        col: LeafColumn,
        colIndex: number,
        row: DisplayRow,
        input: RenderInput,
        rowIndex: number
    ): HTMLElement {
        const cell = document.createElement("div");
        cell.className = "nsm-cell nsm-cell-value";
        const width = this.leafWidths[colIndex] !== undefined ? this.leafWidths[colIndex] : DEFAULT_COL_WIDTH;
        cell.style.position = "absolute";
        cell.style.left = this.leafLefts[colIndex] + "px";
        cell.style.top = "0";
        cell.style.height = "100%";
        cell.style.width = width + "px";

        const vfs = input.valueFormatFor(col.valueSlotIndex);
        const sc = input.specificColumnFor(col.valueSlotIndex);
        const scApplies = this.specificColumnApplies(sc.applyTo, row.kind);
        const isAlt = rowIndex % 2 === 1;

        const raw = row.node.values[col.id];
        const value = raw === undefined ? null : raw;

        // Number text — re-formatted when the specific column overrides unit/decimals.
        let text: string;
        if (scApplies && (sc.decimals !== -1 || sc.unit !== "none")) {
            const field = this.valueFieldBySlot.get(col.valueSlotIndex);
            const overriddenVfs: ValueFormatSettings = {
                ...vfs,
                unit: sc.unit !== "none" ? sc.unit : vfs.unit,
                decimals: sc.decimals !== -1 ? sc.decimals : vfs.decimals
            };
            const fmt = field
                ? this.makeFormatter(field, overriddenVfs)
                : this.valueFormatters.get(col.valueSlotIndex);
            text = fmt ? fmt(value) : value === null ? "" : String(value);
        } else {
            const formatter = this.valueFormatters.get(col.valueSlotIndex);
            text = formatter ? formatter(value) : value === null ? "" : String(value);
        }

        const isTotalRow =
            row.kind === "subtotal" || row.kind === "grandtotal" || col.isColSubtotal || col.isColGrandTotal;

        // --- Value formatting base styling (per-measure font + colors). ---
        cell.style.fontFamily = vfs.fontFamily;
        cell.style.fontSize = vfs.fontSize + "px";
        const baseTextColor = isAlt && vfs.altTextColor ? vfs.altTextColor : vfs.textColor;
        const baseBgColor = isAlt && vfs.altBackgroundColor ? vfs.altBackgroundColor : vfs.backgroundColor;
        if (baseBgColor) {
            cell.style.background = baseBgColor;
        }
        if (baseTextColor) {
            cell.style.color = baseTextColor;
        }

        // --- Conditional formatting (overrides value-format colors). ---
        const cfSettings = input.cfFor(col.valueSlotIndex);
        const applyCf =
            cfSettings.cfType !== "none" &&
            (!isTotalRow || input.settings.subtotals.applyCfToTotals || cfSettings.applyToTotals);

        if (applyCf) {
            const fmt = this.cf.format(value, cfSettings, this.domains.get(col.id), isTotalRow);
            if (fmt) {
                if (fmt.background) {
                    cell.style.background = fmt.background;
                }
                if (fmt.fontColor) {
                    cell.style.color = fmt.fontColor;
                }
                if (fmt.dataBar) {
                    const b = fmt.dataBar;
                    const start = b.negative ? b.axisPct - b.widthPct : b.axisPct;
                    const end = b.negative ? b.axisPct : b.axisPct + b.widthPct;
                    const s = Math.max(0, Math.min(100, start));
                    const e = Math.max(0, Math.min(100, end));
                    cell.style.background = `linear-gradient(to right, transparent ${s}%, ${b.color} ${s}%, ${b.color} ${e}%, transparent ${e}%)`;
                }
                if (fmt.icon) {
                    const ic = document.createElement("span");
                    ic.className = "nsm-cf-icon";
                    ic.textContent = fmt.icon.glyph;
                    ic.style.color = fmt.icon.color;
                    cell.appendChild(ic);
                }
            }
        }

        // --- Field value CF: color read from a referenced measure column's hex. ---
        if (cfSettings.cfType === "fieldValue" && cfSettings.fieldValueSlot >= 0) {
            const fvCol = input.transform.leafColumns.find(
                (c) => c.valueSlotIndex === cfSettings.fieldValueSlot
            );
            if (fvCol) {
                const hexRaw = row.node.values[fvCol.id];
                if (typeof hexRaw === "string" && /^#[0-9A-Fa-f]{6}$/.test(hexRaw)) {
                    if (cfSettings.fieldValueApplyAs !== "font") {
                        cell.style.background = hexRaw;
                    }
                    if (cfSettings.fieldValueApplyAs !== "background") {
                        cell.style.color = hexRaw;
                    }
                }
            }
        }

        // --- Specific column: final override of colors (after value format + CF). ---
        if (scApplies) {
            if (sc.backgroundColor) {
                cell.style.background = sc.backgroundColor;
            }
            if (sc.textColor) {
                cell.style.color = sc.textColor;
            }
        }

        const span = document.createElement("span");
        span.className = "nsm-cellvalue";
        // Bold from value formatting, or subtotal/grand-total rows. Applied inline
        // on the span because .nsm-cellvalue { font-weight: 400 } is more specific.
        const bold = vfs.bold || row.kind === "subtotal" || row.kind === "grandtotal";
        span.style.fontWeight = bold ? "700" : "400";
        if (vfs.italic) {
            span.style.fontStyle = "italic";
        }
        if (vfs.textWrap) {
            span.style.whiteSpace = "normal";
        }
        // Specific column alignment overrides the default right-alignment.
        if (scApplies && sc.alignment !== "auto") {
            this.applyAlignment(cell, span, sc.alignment);
        }
        span.textContent = text;
        cell.appendChild(span);
        return cell;
    }

    /** Whether a specific-column setting applies to a given row/header kind. */
    private specificColumnApplies(applyTo: ColumnApplyTo, kind: DisplayRowKind | "header"): boolean {
        switch (applyTo) {
            case "all":
                return true;
            case "values":
                return kind === "leaf";
            case "header":
                return kind === "header";
            case "subtotals":
                return kind === "subtotal";
            case "grandTotal":
                return kind === "grandtotal";
            default:
                return false;
        }
    }

    /** Apply a specific-column alignment to a value cell + its value span. */
    private applyAlignment(cell: HTMLElement, span: HTMLElement, alignment: ColumnAlignment): void {
        if (alignment === "left") {
            span.style.textAlign = "left";
            cell.style.justifyContent = "flex-start";
        } else if (alignment === "center") {
            span.style.textAlign = "center";
            cell.style.justifyContent = "center";
        } else if (alignment === "right") {
            span.style.textAlign = "right";
            cell.style.justifyContent = "flex-end";
        }
    }
}
