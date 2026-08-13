/*
 * visual.ts
 * ---------
 * Main IVisual implementation for the Nested Sort Matrix. Orchestrates the data
 * transform, nested sort, selection, status bar, config panel, and renderer, and
 * owns visual-lifetime session state (expand/collapse, sort, selection, config).
 *
 * Format pane: implemented via getFormattingModel() (the modern formatting
 * model). Per-measure Values and Specific Column cards expose ConstantOrRule
 * color pickers so Power BI's native fx conditional-formatting dialog drives the
 * per-column color properties. The capabilities.json object schema is unchanged
 * (it still defines persistence); only the display layer was migrated.
 */

import "./styles/visual.less";

import powerbi from "powerbi-visuals-api";
import IVisual = powerbi.extensibility.visual.IVisual;
import VisualConstructorOptions = powerbi.extensibility.visual.VisualConstructorOptions;
import VisualUpdateOptions = powerbi.extensibility.visual.VisualUpdateOptions;
import IVisualHost = powerbi.extensibility.visual.IVisualHost;
import DataView = powerbi.DataView;
import DataViewMatrixNode = powerbi.DataViewMatrixNode;
import ISelectionManager = powerbi.extensibility.ISelectionManager;
import ISelectionId = powerbi.visuals.ISelectionId;
import Selector = powerbi.data.Selector;
import FormattingModel = powerbi.visuals.FormattingModel;
import FormattingCard = powerbi.visuals.FormattingCard;
import FormattingGroup = powerbi.visuals.FormattingGroup;
import SimpleVisualFormattingSlice = powerbi.visuals.SimpleVisualFormattingSlice;
import FormattingComponent = powerbi.visuals.FormattingComponent;
import FormattingDescriptor = powerbi.visuals.FormattingDescriptor;
import VisualEnumerationInstanceKinds = powerbi.VisualEnumerationInstanceKinds;

// The "whole visual" selector. Power BI accepts null at runtime; the typed
// Selector interface is non-nullable, so we cast a single shared constant.
const NULL_SELECTOR: Selector = null as unknown as Selector;

import {
    parseVisualSettings,
    parseValueFormat,
    parseSpecificColumn,
    parseCF,
    VisualSettings,
    ValueFormatSettings,
    SpecificColumnSettings,
    CFSettings,
    DEFAULTS
} from "./settings";
import {
    discoverFields,
    transform,
    TransformResult,
    RowTreeNode,
    FieldMeta,
    LeafColumn
} from "./dataTransformer";
import { SortManager } from "./sortManager";
import { VisualSelectionManager } from "./selectionManager";
import { StatusBar } from "./statusBar";
import { ConfigPanel, ConfigModel, SlotEntry, ConfigRole } from "./configPanel";
import { Renderer, RenderInput, ThemeColors, ColumnWidth, ScrollAnchor, LevelExpandState } from "./renderer";
import { ContextMenu } from "./contextMenu";
import { copyText, ManualCopyPanel } from "./clipboard";

interface PersistedSlot {
    i: number; // slot index
    v: boolean; // visible
}
interface PersistedConfig {
    rows: PersistedSlot[];
    values: PersistedSlot[];
    cols: PersistedSlot[];
}

export class Visual implements IVisual {
    private readonly host: IVisualHost;
    private readonly target: HTMLElement;
    private readonly contentEl: HTMLElement;
    private readonly landingEl: HTMLElement;

    private readonly renderer: Renderer;
    private readonly statusBar: StatusBar;
    private readonly configPanel: ConfigPanel;
    private readonly contextMenu: ContextMenu;
    private readonly manualCopy: ManualCopyPanel;
    /** Value-column leaf ids currently selected for a column copy. */
    private selectedColumnIds = new Set<string>();
    /** Document-level Ctrl+C / Escape handler (stored so it can be detached). */
    private readonly onKeyDown = (e: KeyboardEvent): void => {
        // Let native copy work inside text fields (rename inputs, manual-copy box).
        const tag = (e.target as HTMLElement | null)?.tagName;
        const inTextField = tag === "INPUT" || tag === "TEXTAREA";
        if (!inTextField && (e.ctrlKey || e.metaKey) && (e.key === "c" || e.key === "C")) {
            if (this.selectedColumnIds.size > 0) {
                e.preventDefault();
                this.copyColumns();
            } else if (this.selection.hasSelection()) {
                e.preventDefault();
                this.copyRows();
            }
        } else if (e.key === "Escape" && !inTextField) {
            this.clearCopySelection();
        }
    };
    private readonly sort = new SortManager();
    private readonly selection: VisualSelectionManager;
    private readonly hostSelectionManager: ISelectionManager;

    // Session state.
    private settings: VisualSettings = parseVisualSettings(undefined);
    private config: ConfigModel = { rows: [], values: [], cols: [] };
    private schemaSignature = "";
    // Prefix-tolerant schema-change detection (Item 1): expand/collapse changes only
    // the DEPTH of projected row levels (the row-field list grows/shrinks as a
    // prefix), which must NOT count as a schema change or sort/selection/config
    // would reset on every drill.
    private prevRowNames: string[] | null = null;
    private prevValSig = "";
    private prevColSig = "";
    /** Leaf row count of the previous host update. A change that is NOT a schema
     *  change and NOT an expand/collapse (which anchors scroll) means the row set
     *  changed via filter / slicer / data refresh — reset the scroller to the top
     *  so it can't paint recycled rows from a scrollTop beyond the new content. */
    private prevRowCount = -1;

    // Native expand/collapse (Phase 2). Expansion state is owned by the host and
    // reflected in each matrix node's isCollapsed flag — the visual keeps no local
    // "expanded" set. bulkOp drives Expand All / Collapse All across the async,
    // per-level host round-trips; pendingAnchor preserves scroll across toggles.
    private bulkOp: "expand" | "collapse" | null = null;
    private bulkGuard = 0;
    private bulkLastLevel = -1;
    private pendingAnchor: ScrollAnchor | null = null;

    // Latest update artefacts (kept for callbacks & enumeration).
    private dataView: DataView | undefined = undefined;
    private lastTransform: TransformResult | null = null;
    private activeRowFields: FieldMeta[] = [];
    private activeValueFields: FieldMeta[] = [];
    private valueFormatBySlot = new Map<number, ValueFormatSettings>();
    private specificColumnBySlot = new Map<number, SpecificColumnSettings>();
    private cfBySlot = new Map<number, CFSettings>();
    /** Slots with locally-applied CF awaiting persistence confirmation.
     *  Key: slotIndex. Value: JSON of the CFSettings we expect parseCF to
     *  eventually return once the host merges the persisted objects. */
    private cfPendingConfirm = new Map<number, string>();
    private valueNameBySlot = new Map<number, string>();
    /** Restored/edited column widths, persisted via persistProperties (Fix 3D). */
    private columnWidths: ColumnWidth[] = [];

    constructor(options?: VisualConstructorOptions) {
        // The generated visual plugin passes `options` as optional; guard so the
        // rest of the constructor sees a fully-defined value under strict null
        // checks.
        if (!options) {
            throw new Error("VisualConstructorOptions are required.");
        }
        this.host = options.host;
        this.target = options.element;
        this.target.classList.add("nsm-root");
        // Focusable so the visual receives Ctrl+C / Escape key events for copy.
        if (!this.target.hasAttribute("tabindex")) {
            this.target.setAttribute("tabindex", "0");
        }

        this.hostSelectionManager = this.host.createSelectionManager();

        this.contentEl = document.createElement("div");
        this.contentEl.className = "nsm-content";
        this.target.appendChild(this.contentEl);

        this.landingEl = document.createElement("div");
        this.landingEl.className = "nsm-landing";
        this.landingEl.style.display = "none";
        const landingInner = document.createElement("div");
        landingInner.className = "nsm-landing-inner";
        const landingTitle = document.createElement("h3");
        landingTitle.textContent = "Nested Sort Matrix";
        const landingText = document.createElement("p");
        landingText.textContent =
            "Add fields to Row Fields, Values, and (optionally) Column Fields to begin.";
        landingInner.appendChild(landingTitle);
        landingInner.appendChild(landingText);
        this.landingEl.appendChild(landingInner);
        this.target.appendChild(this.landingEl);

        this.renderer = new Renderer(this.contentEl);

        this.selection = new VisualSelectionManager(
            this.hostSelectionManager,
            () => this.renderBodyOnly(),
            () => this.settings.crossFilter.show
        );

        this.statusBar = new StatusBar(this.target, {
            onSetup: () => this.openConfig(),
            onToggleExpandAll: (expand) => this.toggleExpandAll(expand)
        });

        this.configPanel = new ConfigPanel(this.target, {
            onApply: (model) => this.applyConfig(model, true),
            onReset: () => this.buildDefaultConfig(),
            onClose: () => undefined
        });

        this.contextMenu = new ContextMenu(this.target);
        this.manualCopy = new ManualCopyPanel(this.target);

        // Keyboard: Ctrl/Cmd+C copies the current selection (columns take priority
        // over rows when both exist); Escape clears row + column selection. Attached
        // at document level (scoped to the visual's own sandboxed iframe) so it fires
        // whenever the visual has focus, without needing a specific element focused.
        document.addEventListener("keydown", this.onKeyDown, true);
    }

    // -----------------------------------------------------------------------
    // Update lifecycle.
    // -----------------------------------------------------------------------

    public update(options: VisualUpdateOptions): void {
        const dataView: DataView | undefined =
            options.dataViews && options.dataViews.length ? options.dataViews[0] : undefined;
        this.dataView = dataView;

        const metadataObjects = dataView && dataView.metadata ? dataView.metadata.objects : undefined;
        this.settings = parseVisualSettings(metadataObjects);

        // Per-element font families (Fix 2): expose each as a CSS variable on the
        // root container; the stylesheet maps each variable to its element type.
        this.target.style.setProperty("--nsm-row-font", this.settings.rowHeaders.rowFontFamily);
        this.target.style.setProperty("--nsm-col-font", this.settings.columnHeaders.columnFontFamily);
        this.target.style.setProperty("--nsm-val-font", this.settings.grid.valueFontFamily);

        // Theme background as a CSS variable so .nsm-row has a real, opaque
        // background that the frozen row-header column inherits (Fix 2) — this
        // prevents value cells from showing through during horizontal scroll.
        const paletteBg = (this.host.colorPalette as powerbi.extensibility.ISandboxExtendedColorPalette)
            .background;
        this.target.style.setProperty(
            "--nsm-background",
            paletteBg && paletteBg.value ? paletteBg.value : "#ffffff"
        );

        // Restore persisted column widths (Fix 3D) before the first render pass.
        this.columnWidths = this.parseColumnWidths(this.settings.columnWidths);

        const discovered = discoverFields(dataView);

        // Landing page when nothing to show.
        const hasAnything =
            discovered.rows.length + discovered.values.length + discovered.cols.length > 0;
        if (!hasAnything) {
            this.landingEl.style.display = "flex";
            this.contentEl.style.display = "none";
            this.statusBar.render({ visible: false, sortText: "", rowCount: 0, allExpanded: false });
            return;
        }
        this.landingEl.style.display = "none";
        this.contentEl.style.display = "block";

        // Detect a structural (schema) change == full data refresh. Expand/collapse
        // grows or shrinks the projected row-field list as a PREFIX; that is not a
        // schema change. A real rebind changes the value/column fields or breaks the
        // row prefix relationship (different field or order at a shared position).
        const rowNames = discovered.rows.map((f) => f.queryName);
        const valSig = discovered.values.map((f) => f.queryName).join(",");
        const colSig = discovered.cols.map((f) => f.queryName).join(",");
        const isPrefix = (short: string[], long: string[]): boolean =>
            short.every((x, i) => long[i] === x);
        const rowsCompatible =
            this.prevRowNames !== null &&
            (rowNames.length <= this.prevRowNames.length
                ? isPrefix(rowNames, this.prevRowNames)
                : isPrefix(this.prevRowNames, rowNames));
        const schemaChanged =
            this.prevRowNames === null ||
            !rowsCompatible ||
            valSig !== this.prevValSig ||
            colSig !== this.prevColSig;
        this.prevRowNames = rowNames;
        this.prevValSig = valSig;
        this.prevColSig = colSig;
        this.schemaSignature = `r:${rowNames.join(",")}|v:${valSig}|c:${colSig}`;

        if (schemaChanged) {
            // Full refresh: reset volatile session state. The ACTIVE SORT is
            // deliberately NOT reset here — adding or removing a bound field must
            // never change the active sort (Item B); reconcile() keeps it and only
            // refreshes its label, so a field add leaves the current sort untouched
            // (and leaves the visual unsorted when nothing was active).
            this.selection.reset();
            this.selectedColumnIds = new Set();
            this.bulkOp = null;
            this.pendingAnchor = null;
            this.config = this.reconcileConfig(discovered, this.parsePersisted(this.settings.configState));
        } else {
            // Cross-filter / value update: keep order/visibility/renames; just
            // ensure config still covers the (unchanged) slot set.
            this.config = this.reconcileConfig(discovered, this.configToPersisted(this.config));
        }

        // Resolve active fields from config (visible + ordered) and apply renames.
        const activeRowFields = this.activeFields(discovered.rows, this.config.rows, "row");
        const activeValueFields = this.activeFields(discovered.values, this.config.values, "value");
        const activeColFields = this.activeFields(discovered.cols, this.config.cols, "col");
        this.activeRowFields = activeRowFields;
        this.activeValueFields = activeValueFields;

        // Per-row-field subtotal toggles, read from each field's per-column
        // `subtotals.levelEnabled` object (Fix 2). Index == hierarchy level.
        this.settings.subtotals.levels = activeRowFields.map((f) =>
            this.readLevelEnabled(f.columnObjects)
        );

        // Per-slot settings maps.
        this.valueFormatBySlot.clear();
        this.specificColumnBySlot.clear();
        this.cfBySlot.clear();
        this.valueNameBySlot.clear();
        activeValueFields.forEach((f) => {
            this.valueFormatBySlot.set(f.slotIndex, parseValueFormat(f.columnObjects));
            this.specificColumnBySlot.set(f.slotIndex, parseSpecificColumn(f.columnObjects));
            const parsed = parseCF(f.columnObjects);
            const pending = this.cfPendingConfirm.get(f.slotIndex);
            if (pending !== undefined) {
                if (JSON.stringify(parsed) === pending) {
                    // Persistence round-trip complete — parsed now matches local state.
                    this.cfPendingConfirm.delete(f.slotIndex);
                    this.cfBySlot.set(f.slotIndex, parsed);
                } else {
                    // Host DataView not yet carrying the persisted CF — keep local state.
                    this.cfBySlot.set(f.slotIndex, JSON.parse(pending));
                }
            } else {
                this.cfBySlot.set(f.slotIndex, parsed);
            }
            this.valueNameBySlot.set(f.slotIndex, f.displayName);
        });

        // Transform.
        const result = transform(
            dataView,
            activeRowFields,
            activeValueFields,
            activeColFields,
            this.settings,
            this.hiddenRowLevels()
        );
        this.lastTransform = result;

        // Attach matrix-node selection ids to leaf rows (Phase 2 Part E).
        this.attachSelectionIds(result);

        // Reconcile sort labels & apply nested sort.
        this.sort.reconcile(activeRowFields, result.leafColumns, (c) => this.leafLabel(c));
        this.sort.applyNestedSort(result);

        // Layout: reserve status bar height.
        const viewportH = options.viewport.height;
        const viewportW = options.viewport.width;
        this.contentEl.style.width = viewportW + "px";
        const statusH = this.settings.statusBar.show ? this.statusBar.height() : 0;
        this.renderer.setViewportHeight(viewportH - statusH);
        this.contentEl.style.width = viewportW + "px";

        // Render grid.
        this.renderer.render(this.buildRenderInput(result));

        // Scroll behaviour for this host update. These paths are mutually
        // exclusive; a header SORT never reaches update() (it re-renders in place
        // via applySortToggle and keeps raw scrollTop), so sorting is unaffected.
        if (schemaChanged) {
            // Rebind: reset to the top.
            this.renderer.scrollToTop();
        } else if (this.pendingAnchor) {
            // Expand / collapse: keep the anchored node fixed in the viewport.
            this.renderer.restoreAnchor(this.pendingAnchor);
            if (!this.bulkOp) {
                this.pendingAnchor = null;
            }
        } else if (result.rowCount !== this.prevRowCount) {
            // The row set changed with no schema change and no expand/collapse in
            // flight — i.e. a cross-filter, slicer, or data refresh. The old
            // scrollTop may exceed the new (often shorter) content, leaving the
            // scroller painting recycled rows from the previous set until the user
            // scrolls. Reset to the top so the sizer height, scrollTop, and
            // rendered window all reflect the new row set immediately.
            this.renderer.scrollToTop();
        }
        this.prevRowCount = result.rowCount;

        // Status bar.
        this.statusBar.render({
            visible: this.settings.statusBar.show,
            sortText: this.sort.getStackText(80),
            rowCount: result.rowCount,
            allExpanded: this.computeAllExpanded(result),
            debug: this.sort.getSortDiag()
        });

        // Continue a bulk Expand All / Collapse All in progress (Part B global).
        if (this.bulkOp) {
            this.driveBulk();
        }
    }

    private buildRenderInput(result: TransformResult): RenderInput {
        const theme = this.computeTheme();
        // Expose the theme accent so the column resize-handle hover matches it.
        this.target.style.setProperty("--nsm-theme-accent", theme.accent);
        return {
            transform: result,
            settings: this.settings,
            theme,
            sort: this.sort,
            selection: this.selection,
            valueFormatFor: (slot) => this.valueFormatBySlot.get(slot) || DEFAULTS.valueFormat,
            specificColumnFor: (slot) => this.specificColumnBySlot.get(slot) || DEFAULTS.specificColumn,
            cfFor: (slot) => this.cfBySlot.get(slot) || DEFAULTS.cf,
            leafLabelFor: (col) => this.leafLabel(col),
            onRowFieldSort: (level, label) =>
                this.applySortToggle(() => this.sort.toggleRowField(level, label)),
            onValueSort: (leafColId, label) =>
                this.applySortToggle(() => this.sort.toggleValue(leafColId, label)),
            onColumnSort: () => this.applySortToggle(() => this.sort.toggleColumnDirection()),
            onToggleExpand: (node) => this.toggleNode(node),
            onToggleLevel: (level) => this.toggleLevel(level),
            levelExpandState: (level) => this.levelExpandState(level),
            onRowClick: (node, mods) => this.selection.handleRowClick(node, mods),
            onRowContextMenu: (node, x, y) => this.showRowContextMenu(node, x, y),
            onEmptyClick: () => this.selection.clearSelection(),
            columnWidths: this.columnWidths,
            onColumnWidthsChanged: (widths) => this.persistColumnWidths(widths),
            selectedColumnIds: this.selectedColumnIds,
            onColumnHeaderContext: (colId, ctrl, x, y) =>
                this.handleColumnHeaderContext(colId, ctrl, x, y)
        };
    }

    /**
     * Apply a header-sort toggle and re-render, preserving the RAW scroll offset
     * rather than pinning a node. A sort keeps the row count and row height, so the
     * total content height is unchanged and the same pixel offset lands the viewport
     * at the same fractional position (scrolled 25% down stays 25% down), regardless
     * of which rows now occupy it. (Expand/collapse still uses node-key anchoring via
     * pendingAnchor + restoreAnchor, because there the row set changes.)
     */
    private applySortToggle(toggle: () => void): void {
        const scrollTop = this.renderer.getScrollTop();
        toggle();
        this.rerender(false);
        this.renderer.setScrollTop(scrollTop);
    }

    /** Re-run the sort + render path using the cached tree (no host update). */
    private rerender(resetScroll: boolean): void {
        // Sorting only needs the cached tree, NOT the live dataView — the old
        // `!this.dataView` guard could silently no-op a sort click after a
        // viewport-only update. A render exception must never freeze interaction,
        // so it is caught and the status bar still refreshes.
        if (!this.lastTransform) {
            return;
        }
        try {
            this.sort.applyNestedSort(this.lastTransform);
            this.renderer.render(this.buildRenderInput(this.lastTransform));
            if (resetScroll) {
                this.renderer.scrollToTop();
            }
        } catch {
            /* keep the visual responsive even if a render pass throws */
        }
        this.statusBar.render({
            visible: this.settings.statusBar.show,
            sortText: this.sort.getStackText(80),
            rowCount: this.lastTransform.rowCount,
            allExpanded: this.computeAllExpanded(this.lastTransform),
            debug: this.sort.getSortDiag()
        });
    }

    /** Re-render the body only (selection change) without recomputing layout. */
    private renderBodyOnly(): void {
        if (!this.lastTransform) {
            return;
        }
        this.renderer.render(this.buildRenderInput(this.lastTransform));
    }

    // -----------------------------------------------------------------------
    // Native expand/collapse (Phase 2 Parts A/B).
    // -----------------------------------------------------------------------

    /**
     * Toggle a single group node via the host expand/collapse API. The host
     * refetches and re-delivers the DataView with the node's isCollapsed flag
     * flipped, which drives the next render.
     */
    private toggleNode(node: RowTreeNode): void {
        const id = this.buildMatrixSelectionId(node);
        if (!id) {
            return;
        }
        this.pendingAnchor = this.renderer.captureAnchor();
        try {
            void this.hostSelectionManager.toggleExpandCollapse(id);
        } catch {
            this.pendingAnchor = null;
        }
    }

    /**
     * Right-click on a row: open the visual's own minimal copy menu. Cell-level
     * Copy is not available from the host context menu for custom visuals, so we
     * build TSV ourselves. Right-click does NOT alter the row selection or emit a
     * cross-filter; "Copy selection" copies the current selection, or just the
     * right-clicked row when nothing is selected.
     */
    private showRowContextMenu(node: RowTreeNode, x: number, y: number): void {
        this.contextMenu.open(
            [{ label: "Copy selection", action: () => this.copyRows(node) }],
            x,
            y
        );
    }

    /**
     * Right-click on a value column header: select the column for copy (Ctrl adds
     * to a multi-column set; plain replaces), repaint the highlight, and open the
     * copy menu. Left-click still sorts, so right-click is free for this.
     */
    private handleColumnHeaderContext(colId: string, ctrl: boolean, x: number, y: number): void {
        if (ctrl) {
            this.selectedColumnIds.add(colId);
        } else if (!(this.selectedColumnIds.size === 1 && this.selectedColumnIds.has(colId))) {
            this.selectedColumnIds = new Set([colId]);
        }
        this.rerender(false); // repaint header + cell highlight
        const many = this.selectedColumnIds.size > 1;
        this.contextMenu.open(
            [{ label: many ? "Copy columns" : "Copy column", action: () => this.copyColumns() }],
            x,
            y
        );
    }

    /** Copy the selected rows (or the right-clicked row) as TSV. */
    private copyRows(contextNode?: RowTreeNode): void {
        let keys = this.selection.getSelectedKeys();
        if (keys.size === 0 && contextNode) {
            keys = new Set([contextNode.key]);
        }
        const tsv = this.renderer.buildRowCopyTSV(keys);
        if (!tsv) {
            this.statusBar.flash("Nothing to copy");
            return;
        }
        this.doCopy(tsv.text, tsv.count);
    }

    /** Copy the selected value columns (deepest rendered level rows) as TSV. */
    private copyColumns(): void {
        // TEMP (Item B): surface the copy-state diagnostic instead of the normal
        // "Copied N rows" confirmation so the failure mode is visible on device.
        const diag = this.renderer.columnCopyDiag(this.selectedColumnIds);
        const tsv = this.renderer.buildColumnCopyTSV(this.selectedColumnIds);
        if (tsv) {
            copyText(tsv.text, (t) => this.manualCopy.show(t));
        }
        this.statusBar.flash(diag, 6000);
    }

    /** Write text to the clipboard and flash a transient status confirmation. */
    private doCopy(text: string, rowCount: number): void {
        const result = copyText(text, (t) => this.manualCopy.show(t));
        if (result === "copied") {
            this.statusBar.flash(`Copied ${rowCount} row${rowCount === 1 ? "" : "s"}`);
        }
    }

    /** Escape: clear row and column copy selection and dismiss any popups. */
    private clearCopySelection(): void {
        this.contextMenu.close();
        this.manualCopy.close();
        const hadCols = this.selectedColumnIds.size > 0;
        this.selectedColumnIds = new Set();
        this.selection.clearSelection();
        if (hadCols) {
            this.rerender(false);
        }
    }

    /** Toggle an entire hierarchy level (row-field header control, Part B). */
    private toggleLevel(level: number): void {
        const state = this.levelExpandState(level);
        if (state === "none") {
            return;
        }
        // To expand the level, pass a currently-collapsed node at that level; to
        // collapse it, pass an expanded one, so the host toggles the intended way.
        const wantCollapsedNode = state === "expand";
        const node = this.findLevelNode(level, wantCollapsedNode);
        if (!node) {
            return;
        }
        const id = this.buildMatrixSelectionId(node);
        if (!id) {
            return;
        }
        this.pendingAnchor = this.renderer.captureAnchor();
        try {
            void this.hostSelectionManager.toggleExpandCollapse(id, true /* entireLevel */);
        } catch {
            this.pendingAnchor = null;
        }
    }

    /** Expand All / Collapse All from the status bar (Part B global). */
    private toggleExpandAll(expand: boolean): void {
        this.pendingAnchor = this.renderer.captureAnchor();
        this.bulkOp = expand ? "expand" : "collapse";
        this.bulkGuard = 0;
        this.bulkLastLevel = -1;
        this.driveBulk();
    }

    /**
     * Advance a bulk expand/collapse. Because entire-level toggles are async and
     * deeper levels only materialize after their parent expands, this issues one
     * level toggle per host round-trip and is re-entered from update() until the
     * whole tree reaches the target state.
     */
    private driveBulk(): void {
        if (!this.bulkOp || !this.lastTransform) {
            return;
        }
        if (this.bulkGuard++ > 64) {
            this.bulkOp = null;
            this.pendingAnchor = null;
            return;
        }
        const target =
            this.bulkOp === "expand"
                ? this.shallowestCollapsedLevel()
                : this.deepestExpandedLevel();
        // No progress possible, or done.
        if (target < 0 || target === this.bulkLastLevel) {
            this.bulkOp = null;
            this.pendingAnchor = null;
            return;
        }
        const node = this.findExpandNodeAtMatrixLevel(target, this.bulkOp === "expand");
        if (!node) {
            this.bulkOp = null;
            this.pendingAnchor = null;
            return;
        }
        const id = this.buildMatrixSelectionId(node);
        if (!id) {
            this.bulkOp = null;
            this.pendingAnchor = null;
            return;
        }
        this.bulkLastLevel = target;
        try {
            void this.hostSelectionManager.toggleExpandCollapse(id, true /* entireLevel */);
        } catch {
            this.bulkOp = null;
            this.pendingAnchor = null;
        }
    }

    // ---- Expand/collapse tree helpers ----

    /** Build an ISelectionId for a matrix node by chaining withMatrixNode over
     *  its ancestor path (root-most first), per the expand/collapse API. */
    private buildMatrixSelectionId(node: RowTreeNode): ISelectionId | undefined {
        const matrix = this.dataView && this.dataView.matrix;
        if (!matrix || !matrix.rows) {
            return undefined;
        }
        const levels = matrix.rows.levels;
        // Prefer the full matrix ancestor chain captured at transform time — it
        // includes nodes at HIDDEN levels that the display parent chain skips, so
        // the selection id keeps the complete level path. Fall back to walking the
        // display parents when a node has no captured chain (e.g. synthetic nodes).
        let matrixNodes: DataViewMatrixNode[];
        if (node.matrixChain && node.matrixChain.length > 0) {
            matrixNodes = node.matrixChain;
        } else {
            matrixNodes = [];
            let n: RowTreeNode | undefined = node;
            while (n) {
                if (!n.matrixNode) {
                    return undefined;
                }
                matrixNodes.unshift(n.matrixNode);
                n = n.parent;
            }
        }
        try {
            let builder = this.host.createSelectionIdBuilder();
            for (let i = 0; i < matrixNodes.length; i++) {
                builder = builder.withMatrixNode(matrixNodes[i], levels);
            }
            return builder.createSelectionId();
        } catch {
            return undefined;
        }
    }

    /**
     * Attach a matrix selection id to EVERY row node (Item 2). Leaf, group and
     * subtotal rows are all selectable; a group's id (ancestor-chained
     * withMatrixNode ending at the group) scopes cross-filter to that whole group.
     */
    private attachSelectionIds(result: TransformResult): void {
        const walk = (node: RowTreeNode): void => {
            node.selectionId = this.buildMatrixSelectionId(node);
            node.children.forEach(walk);
        };
        result.rootNodes.forEach(walk);
    }

    /** True if no expandable node anywhere is still collapsed. Uses the full
     *  expandable-node list (incl. hidden levels), not just the display tree. */
    private computeAllExpanded(result: TransformResult | null): boolean {
        if (!result || result.expandNodes.length === 0) {
            return false;
        }
        return !result.expandNodes.some((n) => n.isCollapsed === true);
    }

    /** Aggregate expand/collapse state of a hierarchy level for its header control. */
    private levelExpandState(level: number): LevelExpandState {
        if (!this.lastTransform) {
            return "none";
        }
        let present = false;
        let anyCollapsed = false;
        let anyExpanded = false;
        const walk = (node: RowTreeNode): void => {
            if (node.level === level && node.isCollapsed !== undefined) {
                present = true;
                if (node.isCollapsed) {
                    anyCollapsed = true;
                } else {
                    anyExpanded = true;
                }
            }
            if (node.level < level) {
                node.children.forEach(walk);
            }
        };
        this.lastTransform.rootNodes.forEach(walk);
        if (!present) {
            return "none";
        }
        // If any node is collapsed, offer "expand"; otherwise offer "collapse".
        return anyCollapsed || !anyExpanded ? "expand" : "collapse";
    }

    /** Find a node at a level, preferring collapsed (or expanded) per the flag. */
    private findLevelNode(level: number, wantCollapsed: boolean): RowTreeNode | undefined {
        if (!this.lastTransform) {
            return undefined;
        }
        let fallback: RowTreeNode | undefined;
        let found: RowTreeNode | undefined;
        const walk = (node: RowTreeNode): void => {
            if (found) {
                return;
            }
            if (node.level === level && node.isCollapsed !== undefined) {
                if (!fallback) {
                    fallback = node;
                }
                if (node.isCollapsed === wantCollapsed) {
                    found = node;
                    return;
                }
            }
            if (node.level < level) {
                node.children.forEach(walk);
            }
        };
        this.lastTransform.rootNodes.forEach(walk);
        return found || fallback;
    }

    /**
     * Shallowest MATRIX level that still has a collapsed expandable node (or -1).
     * Bulk expand/collapse walks the full expandable-node list — including hidden
     * levels — by matrix level, so Expand All can reach and expand a collapsed
     * hidden level whose children have not yet materialised.
     */
    private shallowestCollapsedLevel(): number {
        const nodes = this.lastTransform ? this.lastTransform.expandNodes : [];
        let best = -1;
        for (const n of nodes) {
            const ml = n.matrixLevel != null ? n.matrixLevel : n.level;
            if (n.isCollapsed === true && (best < 0 || ml < best)) {
                best = ml;
            }
        }
        return best;
    }

    /** Deepest MATRIX level that still has an expanded expandable node (or -1). */
    private deepestExpandedLevel(): number {
        const nodes = this.lastTransform ? this.lastTransform.expandNodes : [];
        let best = -1;
        for (const n of nodes) {
            const ml = n.matrixLevel != null ? n.matrixLevel : n.level;
            if (n.isCollapsed === false && ml > best) {
                best = ml;
            }
        }
        return best;
    }

    /** A representative expandable node at a MATRIX level (bulk expand/collapse),
     *  preferring one whose collapsed-state matches, drawn from the full list. */
    private findExpandNodeAtMatrixLevel(
        matrixLevel: number,
        wantCollapsed: boolean
    ): RowTreeNode | undefined {
        const nodes = this.lastTransform ? this.lastTransform.expandNodes : [];
        let fallback: RowTreeNode | undefined;
        for (const n of nodes) {
            const ml = n.matrixLevel != null ? n.matrixLevel : n.level;
            if (ml !== matrixLevel || n.isCollapsed === undefined) {
                continue;
            }
            if (!fallback) {
                fallback = n;
            }
            if (n.isCollapsed === wantCollapsed) {
                return n;
            }
        }
        return fallback;
    }

    // -----------------------------------------------------------------------
    // Config (panel) handling.
    // -----------------------------------------------------------------------

    private openConfig(): void {
        this.configPanel.open(this.config);
    }

    private applyConfig(model: ConfigModel, persist: boolean): void {
        this.config = model;
        if (persist) {
            this.persistConfig(model);
        }
        // Renames change display names -> re-run the full update path on next
        // host update; for immediate feedback, re-derive active fields now.
        if (this.dataView) {
            const discovered = discoverFields(this.dataView);
            const activeRowFields = this.activeFields(discovered.rows, this.config.rows, "row");
            const activeValueFields = this.activeFields(discovered.values, this.config.values, "value");
            const activeColFields = this.activeFields(discovered.cols, this.config.cols, "col");
            this.activeRowFields = activeRowFields;
            this.activeValueFields = activeValueFields;
            // Per-VISIBLE-row-level subtotal toggles must track the current
            // visibility set so a hidden level's subtotal row is suppressed and the
            // toggle array stays aligned to the compacted display levels.
            this.settings.subtotals.levels = activeRowFields.map((f) =>
                this.readLevelEnabled(f.columnObjects)
            );
            this.valueNameBySlot.clear();
            this.valueFormatBySlot.clear();
            this.specificColumnBySlot.clear();
            // CF is deliberately NOT re-read here: this.dataView is the last
            // DataView the host delivered, so parseCF would clobber a newer
            // in-memory CF applied by the CF panel (saveCfState writes cfBySlot
            // directly). CF is never edited via the config panel, so leaving
            // cfBySlot untouched is always safe; the next host update() re-reads
            // it from fresh column objects.
            activeValueFields.forEach((f) => {
                this.valueNameBySlot.set(f.slotIndex, f.displayName);
                this.valueFormatBySlot.set(f.slotIndex, parseValueFormat(f.columnObjects));
                this.specificColumnBySlot.set(f.slotIndex, parseSpecificColumn(f.columnObjects));
            });
            const result = transform(
                this.dataView,
                activeRowFields,
                activeValueFields,
                activeColFields,
                this.settings,
                this.hiddenRowLevels()
            );
            this.lastTransform = result;
            this.attachSelectionIds(result);
            this.sort.reconcile(activeRowFields, result.leafColumns, (c) => this.leafLabel(c));
            this.sort.applyNestedSort(result);
            this.renderer.render(this.buildRenderInput(result));
            this.renderer.scrollToTop();
        }
    }

    private buildDefaultConfig(): ConfigModel {
        const discovered = discoverFields(this.dataView);
        const make = (fields: FieldMeta[], role: ConfigRole): SlotEntry[] =>
            fields.map((f) => ({
                role,
                slotIndex: f.slotIndex,
                originalName: f.originalName,
                visible: true,
                rename: null
            }));
        return {
            rows: make(discovered.rows, "row"),
            values: make(discovered.values, "value"),
            cols: make(discovered.cols, "col")
        };
    }

    private persistConfig(model: ConfigModel): void {
        const persisted = this.configToPersisted(model);
        const props: { [k: string]: powerbi.DataViewPropertyValue } = {
            configState: JSON.stringify(persisted)
        };
        this.host.persistProperties({
            merge: [
                {
                    objectName: "general",
                    selector: NULL_SELECTOR,
                    properties: props
                }
            ]
        });
    }

    /** Parse the persisted column-widths JSON into a validated array (Fix 3D). */
    private parseColumnWidths(json: string): ColumnWidth[] {
        if (!json) {
            return [];
        }
        try {
            const parsed: unknown = JSON.parse(json);
            if (!Array.isArray(parsed)) {
                return [];
            }
            const out: ColumnWidth[] = [];
            parsed.forEach((entry) => {
                const e = entry as { type?: unknown; index?: unknown; width?: unknown };
                if (typeof e.index === "number" && typeof e.width === "number") {
                    // Older persisted entries have no `type` — treat them as value columns.
                    const type = e.type === "row" ? "row" : "val";
                    out.push({ type, index: e.index, width: e.width });
                }
            });
            return out;
        } catch {
            return [];
        }
    }

    /** Persist column widths and keep the in-memory copy in sync (Fix 3D). */
    private persistColumnWidths(widths: ColumnWidth[]): void {
        this.columnWidths = widths;
        const props: { [k: string]: powerbi.DataViewPropertyValue } = {
            columnWidths: JSON.stringify(widths)
        };
        this.host.persistProperties({
            merge: [
                {
                    objectName: "general",
                    selector: NULL_SELECTOR,
                    properties: props
                }
            ]
        });
    }

    private parsePersisted(json: string): PersistedConfig | null {
        if (!json) {
            return null;
        }
        try {
            const parsed = JSON.parse(json) as PersistedConfig;
            if (parsed && parsed.rows && parsed.values && parsed.cols) {
                return parsed;
            }
            return null;
        } catch {
            return null;
        }
    }

    private configToPersisted(model: ConfigModel): PersistedConfig {
        const map = (entries: SlotEntry[]): PersistedSlot[] =>
            entries.map((e) => ({ i: e.slotIndex, v: e.visible }));
        return { rows: map(model.rows), values: map(model.values), cols: map(model.cols) };
    }

    /**
     * Reconcile a config against the currently-discovered fields: honour the
     * persisted order/visibility where slots still exist, append newly-added
     * slots (visible), and drop slots that are no longer populated. Renames from
     * the existing in-memory config are preserved when the slot survives.
     */
    private reconcileConfig(
        discovered: { rows: FieldMeta[]; values: FieldMeta[]; cols: FieldMeta[] },
        persisted: PersistedConfig | null
    ): ConfigModel {
        const renameLookup = (role: ConfigRole, slot: number): string | null => {
            const list =
                role === "row" ? this.config.rows : role === "value" ? this.config.values : this.config.cols;
            const found = list.find((e) => e.slotIndex === slot);
            return found ? found.rename : null;
        };

        const build = (
            fields: FieldMeta[],
            order: PersistedSlot[] | undefined,
            role: ConfigRole
        ): SlotEntry[] => {
            const bySlot = new Map<number, FieldMeta>();
            fields.forEach((f) => bySlot.set(f.slotIndex, f));
            const result: SlotEntry[] = [];
            const used = new Set<number>();

            if (order) {
                order.forEach((p) => {
                    const f = bySlot.get(p.i);
                    if (f && !used.has(p.i)) {
                        used.add(p.i);
                        result.push({
                            role,
                            slotIndex: f.slotIndex,
                            originalName: f.originalName,
                            visible: p.v,
                            rename: renameLookup(role, f.slotIndex)
                        });
                    }
                });
            }
            // Append any not covered by the order (new fields), in slot order.
            fields.forEach((f) => {
                if (!used.has(f.slotIndex)) {
                    result.push({
                        role,
                        slotIndex: f.slotIndex,
                        originalName: f.originalName,
                        visible: true,
                        rename: renameLookup(role, f.slotIndex)
                    });
                }
            });
            return result;
        };

        return {
            rows: build(discovered.rows, persisted ? persisted.rows : undefined, "row"),
            values: build(discovered.values, persisted ? persisted.values : undefined, "value"),
            cols: build(discovered.cols, persisted ? persisted.cols : undefined, "col")
        };
    }

    /** Resolve the ordered, visible FieldMeta list for a role, applying renames. */
    /**
     * Matrix row levels the user has hidden. A row field's slotIndex equals its
     * matrix hierarchy level (row fields are never reordered — hierarchy order is
     * fixed by the field well), so a hidden entry's slotIndex is the matrix level
     * to flatten out of the display tree. Visibility is a DISPLAY concern only: it
     * never changes the projection/schema signature, so toggling it does not reset
     * sort / selection / expansion state.
     */
    private hiddenRowLevels(): Set<number> {
        const hidden = new Set<number>();
        this.config.rows.forEach((e) => {
            if (!e.visible) {
                hidden.add(e.slotIndex);
            }
        });
        return hidden;
    }

    private activeFields(fields: FieldMeta[], entries: SlotEntry[], _role: ConfigRole): FieldMeta[] {
        const bySlot = new Map<number, FieldMeta>();
        fields.forEach((f) => bySlot.set(f.slotIndex, f));
        const out: FieldMeta[] = [];
        entries.forEach((e) => {
            if (!e.visible) {
                return;
            }
            const f = bySlot.get(e.slotIndex);
            if (!f) {
                return;
            }
            // Apply session rename by cloning the meta (do not mutate source).
            out.push({ ...f, displayName: e.rename != null ? e.rename : f.originalName });
        });
        return out;
    }

    // -----------------------------------------------------------------------
    // Helpers.
    // -----------------------------------------------------------------------

    private leafLabel(col: LeafColumn): string {
        return this.valueNameBySlot.get(col.valueSlotIndex) || `Value ${col.valueSlotIndex + 1}`;
    }

    // -----------------------------------------------------------------------
    // Theme.
    // -----------------------------------------------------------------------

    private computeTheme(): ThemeColors {
        const palette = this.host.colorPalette as powerbi.extensibility.ISandboxExtendedColorPalette;
        const bg = this.readPaletteColor(palette.background, "#FFFFFF");
        const fg = this.readPaletteColor(palette.foreground, "#252423");
        const accent = this.firstThemeColor(palette, "#118DFF");

        const isHighContrast = !!palette.isHighContrast;
        if (isHighContrast) {
            return {
                foreground: fg,
                background: bg,
                headerBackground: bg,
                headerForeground: fg,
                rowBorder: fg,
                subtotalBackground: bg,
                selectionFill: accent,
                accent
            };
        }

        return {
            foreground: fg,
            background: bg,
            headerBackground: blend(bg, fg, 0.06),
            headerForeground: fg,
            rowBorder: blend(bg, fg, 0.12),
            subtotalBackground: blend(bg, fg, 0.08),
            selectionFill: blend(bg, accent, 0.22),
            accent
        };
    }

    private readPaletteColor(
        slot: { value?: string } | undefined,
        fallback: string
    ): string {
        if (slot && typeof slot.value === "string" && slot.value) {
            return slot.value;
        }
        return fallback;
    }

    private firstThemeColor(
        palette: powerbi.extensibility.ISandboxExtendedColorPalette,
        fallback: string
    ): string {
        try {
            const c = palette.getColor("0");
            if (c && c.value) {
                return c.value;
            }
        } catch {
            /* ignore */
        }
        return fallback;
    }

    // -----------------------------------------------------------------------
    // Format pane (modern formatting model via getFormattingModel).
    // -----------------------------------------------------------------------

    public getFormattingModel(): FormattingModel {
        const s = this.settings;
        const CR = VisualEnumerationInstanceKinds.ConstantOrRule;

        // Metadata selector for per-measure (per-column) properties.
        const sel = (queryName: string): powerbi.data.Selector =>
            ({ metadata: queryName } as powerbi.data.Selector);

        // Local descriptor helper. Global descriptors pass selector = null.
        const desc = (
            objectName: string,
            propertyName: string,
            selector: powerbi.data.Selector | null = null,
            instanceKind?: VisualEnumerationInstanceKinds
        ): FormattingDescriptor => ({
            objectName,
            propertyName,
            selector: selector ?? undefined,
            instanceKind
        });

        // A globally-unique slice uid: object|property|selector.
        const keyOf = (d: FormattingDescriptor): string => {
            const m = d.selector as { metadata?: string } | undefined;
            return d.objectName + "|" + d.propertyName + "|" + (m && m.metadata ? m.metadata : "g");
        };

        const toggle = (displayName: string, d: FormattingDescriptor, value: boolean): SimpleVisualFormattingSlice => ({
            uid: keyOf(d),
            displayName,
            control: { type: FormattingComponent.ToggleSwitch, properties: { descriptor: d, value } }
        });
        const num = (displayName: string, d: FormattingDescriptor, value: number): SimpleVisualFormattingSlice => ({
            uid: keyOf(d),
            displayName,
            control: { type: FormattingComponent.NumUpDown, properties: { descriptor: d, value } }
        });
        const text = (displayName: string, d: FormattingDescriptor, value: string): SimpleVisualFormattingSlice => ({
            uid: keyOf(d),
            displayName,
            control: { type: FormattingComponent.TextInput, properties: { descriptor: d, value, placeholder: "" } }
        });
        const color = (displayName: string, d: FormattingDescriptor, value: string | null): SimpleVisualFormattingSlice => ({
            uid: keyOf(d),
            displayName,
            control: { type: FormattingComponent.ColorPicker, properties: { descriptor: d, value: { value: value || "" } } }
        });
        const dropdown = (displayName: string, d: FormattingDescriptor, value: string): SimpleVisualFormattingSlice => ({
            uid: keyOf(d),
            displayName,
            control: { type: FormattingComponent.Dropdown, properties: { descriptor: d, value } }
        });

        const cards: FormattingCard[] = [];

        // ---- Grid ----
        {
            const d = [
                desc("grid", "rowHeightMode"),
                desc("grid", "rowHeightPx"),
                desc("grid", "fontSize"),
                desc("grid", "valueFontFamily")
            ];
            cards.push({
                uid: "card-grid",
                displayName: "Grid",
                groups: [
                    {
                        uid: "grid-g",
                        displayName: "",
                        slices: [
                            dropdown("Row height", d[0], s.grid.rowHeightMode),
                            num("Fixed row height (px)", d[1], s.grid.rowHeightPx),
                            num("Cell font size", d[2], s.grid.fontSize),
                            dropdown("Value cell font", d[3], s.grid.valueFontFamily)
                        ]
                    }
                ],
                revertToDefaultDescriptors: d
            });
        }

        // ---- Row Headers ----
        {
            const d = [
                desc("rowHeaders", "bold"),
                desc("rowHeaders", "fontSize"),
                desc("rowHeaders", "indentPerLevel"),
                desc("rowHeaders", "rowFontFamily"),
                desc("rowHeaders", "wrapHeaderText"),
                desc("rowHeaders", "headerAlignment")
            ];
            cards.push({
                uid: "card-rowHeaders",
                displayName: "Row Headers",
                groups: [
                    {
                        uid: "rowHeaders-g",
                        displayName: "",
                        slices: [
                            toggle("Bold", d[0], s.rowHeaders.bold),
                            num("Font size", d[1], s.rowHeaders.fontSize),
                            num("Indent per level (px)", d[2], s.rowHeaders.indentPerLevel),
                            dropdown("Font family", d[3], s.rowHeaders.rowFontFamily),
                            toggle("Wrap header text", d[4], s.rowHeaders.wrapHeaderText),
                            dropdown("Header alignment", d[5], s.rowHeaders.headerAlignment)
                        ]
                    }
                ],
                revertToDefaultDescriptors: d
            });
        }

        // ---- Layout ----
        {
            const d = [desc("layoutOptions", "layoutMode"), desc("layoutOptions", "repeatRowHeaders")];
            cards.push({
                uid: "card-layoutOptions",
                displayName: "Layout",
                groups: [
                    {
                        uid: "layoutOptions-g",
                        displayName: "",
                        slices: [
                            dropdown("Layout mode", d[0], s.rowHeaders.layoutMode),
                            toggle("Repeat row headers", d[1], s.rowHeaders.repeatRowHeaders)
                        ]
                    }
                ],
                revertToDefaultDescriptors: d
            });
        }

        // ---- Column Headers ----
        {
            const d = [
                desc("columnHeaders", "bold"),
                desc("columnHeaders", "fontSize"),
                desc("columnHeaders", "showSortArrows"),
                desc("columnHeaders", "columnFontFamily"),
                desc("columnHeaders", "wrapHeaderText"),
                desc("columnHeaders", "headerAlignment")
            ];
            cards.push({
                uid: "card-columnHeaders",
                displayName: "Column Headers",
                groups: [
                    {
                        uid: "columnHeaders-g",
                        displayName: "",
                        slices: [
                            toggle("Bold", d[0], s.columnHeaders.bold),
                            num("Font size", d[1], s.columnHeaders.fontSize),
                            toggle("Show sort arrows", d[2], s.columnHeaders.showSortArrows),
                            dropdown("Font family", d[3], s.columnHeaders.columnFontFamily),
                            toggle("Wrap header text", d[4], s.columnHeaders.wrapHeaderText),
                            dropdown("Header alignment", d[5], s.columnHeaders.headerAlignment)
                        ]
                    }
                ],
                revertToDefaultDescriptors: d
            });
        }

        // ---- Expand/Collapse Buttons ----
        {
            const d = [
                desc("expandCollapse", "show"),
                desc("expandCollapse", "buttonSize"),
                desc("expandCollapse", "buttonColor"),
                desc("expandCollapse", "style")
            ];
            cards.push({
                uid: "card-expandCollapse",
                displayName: "Expand/Collapse Buttons",
                groups: [
                    {
                        uid: "expandCollapse-g",
                        displayName: "",
                        slices: [
                            toggle("Show expand/collapse buttons", d[0], s.expandCollapse.show),
                            num("Button size (px)", d[1], s.expandCollapse.buttonSize),
                            color("Button color", d[2], s.expandCollapse.buttonColor),
                            dropdown("Button style", d[3], s.expandCollapse.style)
                        ]
                    }
                ],
                revertToDefaultDescriptors: d
            });
        }

        // ---- Subtotals & Totals ----
        {
            const gd = [
                desc("subtotals", "rowSubtotals"),
                desc("subtotals", "columnSubtotals"),
                desc("subtotals", "grandTotalRow"),
                desc("subtotals", "grandTotalColumn"),
                desc("subtotals", "labelText"),
                desc("subtotals", "applyCfToTotals")
            ];
            const groups: FormattingGroup[] = [
                {
                    uid: "subtotals-g",
                    displayName: "",
                    slices: [
                        toggle("Row subtotals", gd[0], s.subtotals.rowSubtotals),
                        toggle("Column subtotals", gd[1], s.subtotals.columnSubtotals),
                        toggle("Grand total row", gd[2], s.subtotals.grandTotalRow),
                        toggle("Grand total column", gd[3], s.subtotals.grandTotalColumn),
                        text("Subtotal label", gd[4], s.subtotals.labelText),
                        toggle("Apply CF to totals", gd[5], s.subtotals.applyCfToTotals)
                    ]
                }
            ];
            const revert: FormattingDescriptor[] = gd.slice();
            // When row subtotals are on, add one toggle group per active row field.
            if (s.subtotals.rowSubtotals) {
                this.activeRowFields.forEach((field, idx) => {
                    const enabled = s.subtotals.levels[idx];
                    const ld = desc("subtotals", "levelEnabled", sel(field.queryName));
                    revert.push(ld);
                    groups.push({
                        uid: "subtotals-level-" + field.queryName,
                        displayName: field.displayName,
                        slices: [toggle("Show subtotal", ld, enabled !== undefined ? enabled : true)]
                    });
                });
            }
            cards.push({
                uid: "card-subtotals",
                displayName: "Subtotals & Totals",
                groups,
                revertToDefaultDescriptors: revert
            });
        }

        // ---- Alternate Row Color ----
        {
            const d = [desc("alternateRows", "show"), desc("alternateRows", "color")];
            cards.push({
                uid: "card-alternateRows",
                displayName: "Alternate Row Color",
                groups: [
                    {
                        uid: "alternateRows-g",
                        displayName: "",
                        slices: [
                            toggle("Show", d[0], s.alternateRows.show),
                            color("Color", d[1], s.alternateRows.color)
                        ]
                    }
                ],
                revertToDefaultDescriptors: d
            });
        }

        // ---- Status Bar ----
        {
            const d = [desc("statusBar", "show")];
            cards.push({
                uid: "card-statusBar",
                displayName: "Status Bar",
                groups: [{ uid: "statusBar-g", displayName: "", slices: [toggle("Show", d[0], s.statusBar.show)] }],
                revertToDefaultDescriptors: d
            });
        }

        // ---- Cross-Filter ----
        {
            const d = [desc("crossFilter", "show")];
            cards.push({
                uid: "card-crossFilter",
                displayName: "Cross-Filter",
                groups: [{ uid: "crossFilter-g", displayName: "", slices: [toggle("Enable", d[0], s.crossFilter.show)] }],
                revertToDefaultDescriptors: d
            });
        }

        // ---- Values (one card, one group per measure) ----
        {
            const valuesGroups: FormattingGroup[] = [];
            const valuesRevert: FormattingDescriptor[] = [];
            this.activeValueFields.forEach((f) => {
                const v = this.valueFormatBySlot.get(f.slotIndex) || DEFAULTS.valueFormat;
                const m = sel(f.queryName);

                const dFontFamily = desc("valueFormatting", "fontFamily", m);
                const dFontSize = desc("valueFormatting", "fontSize", m);
                const dBold = desc("valueFormatting", "bold", m);
                const dItalic = desc("valueFormatting", "italic", m);
                // These four enable the native fx CF dialog (gradient / rules / field value).
                const dTextColor = desc("valueFormatting", "textColor", m, CR);
                const dBackgroundColor = desc("valueFormatting", "backgroundColor", m, CR);
                const dAltTextColor = desc("valueFormatting", "altTextColor", m, CR);
                const dAltBackgroundColor = desc("valueFormatting", "altBackgroundColor", m, CR);
                const dTextWrap = desc("valueFormatting", "textWrap", m);
                const dNumberFormat = desc("valueFormatting", "numberFormat", m);
                const dDecimals = desc("valueFormatting", "decimals", m);
                const dUnit = desc("valueFormatting", "unit", m);
                const dPrefix = desc("valueFormatting", "prefix", m);
                const dSuffix = desc("valueFormatting", "suffix", m);

                valuesGroups.push({
                    uid: "valueFormatting-group-" + f.slotIndex,
                    displayName: f.displayName,
                    slices: [
                        // Typography
                        dropdown("Font", dFontFamily, v.fontFamily),
                        num("Font size", dFontSize, v.fontSize),
                        toggle("Bold", dBold, v.bold),
                        toggle("Italic", dItalic, v.italic),
                        // Colors
                        color("Font color", dTextColor, v.textColor),
                        color("Background color", dBackgroundColor, v.backgroundColor),
                        color("Alternate font color", dAltTextColor, v.altTextColor),
                        color("Alternate background color", dAltBackgroundColor, v.altBackgroundColor),
                        // Display
                        toggle("Text wrap", dTextWrap, v.textWrap),
                        text("Format", dNumberFormat, v.numberFormat),
                        num("Decimal places", dDecimals, v.decimals),
                        dropdown("Display unit", dUnit, v.unit),
                        text("Prefix", dPrefix, v.prefix),
                        text("Suffix", dSuffix, v.suffix)
                    ]
                });
                valuesRevert.push(
                    dFontFamily,
                    dFontSize,
                    dBold,
                    dItalic,
                    dTextColor,
                    dBackgroundColor,
                    dAltTextColor,
                    dAltBackgroundColor,
                    dTextWrap,
                    dNumberFormat,
                    dDecimals,
                    dUnit,
                    dPrefix,
                    dSuffix
                );
            });
            cards.push({
                uid: "valueFormatting-card",
                displayName: "Values",
                groups: valuesGroups,
                revertToDefaultDescriptors: valuesRevert
            });
        }

        // ---- Specific Column (one card, one group per measure) ----
        {
            const scGroups: FormattingGroup[] = [];
            const scRevert: FormattingDescriptor[] = [];
            this.activeValueFields.forEach((f) => {
                const sc = this.specificColumnBySlot.get(f.slotIndex) || DEFAULTS.specificColumn;
                const m = sel(f.queryName);

                const dApplyTo = desc("specificColumn", "applyTo", m);
                const dTextColor = desc("specificColumn", "textColor", m, CR);
                const dBackgroundColor = desc("specificColumn", "backgroundColor", m, CR);
                const dAlignment = desc("specificColumn", "alignment", m);
                const dUnit = desc("specificColumn", "unit", m);
                const dDecimals = desc("specificColumn", "decimals", m);

                scGroups.push({
                    uid: "specificColumn-group-" + f.slotIndex,
                    displayName: f.displayName,
                    slices: [
                        dropdown("Apply to", dApplyTo, sc.applyTo),
                        color("Font color", dTextColor, sc.textColor),
                        color("Background color", dBackgroundColor, sc.backgroundColor),
                        dropdown("Alignment", dAlignment, sc.alignment),
                        dropdown("Display unit", dUnit, sc.unit),
                        num("Decimal places", dDecimals, sc.decimals)
                    ]
                });
                scRevert.push(dApplyTo, dTextColor, dBackgroundColor, dAlignment, dUnit, dDecimals);
            });
            cards.push({
                uid: "specificColumn-card",
                displayName: "Specific Column",
                groups: scGroups,
                revertToDefaultDescriptors: scRevert
            });
        }

        // NOTE: The "cfSettings" section is intentionally NOT emitted here. The
        // four ConstantOrRule color pickers above (Values + Specific Column)
        // surface Power BI's native fx conditional-formatting dialog. When the
        // user sets a rule there, Power BI evaluates it against the DataView and
        // writes the RESOLVED color back into the SAME per-column object property
        // slot (e.g. valueFormatting.textColor) through dataView.metadata.objects
        // — the exact path settings.ts already parses and the renderer already
        // applies. So no settings.ts or renderer changes are needed for CF values
        // to take effect. (cfSettings stays in capabilities.json only for backward
        // persistence compatibility.)

        return { cards };
    }

    /** Read a row field's per-column `subtotals.levelEnabled` toggle (default true). */
    private readLevelEnabled(columnObjects: powerbi.DataViewObjects | undefined): boolean {
        const sub = columnObjects ? columnObjects["subtotals"] : undefined;
        const le = sub ? sub["levelEnabled"] : undefined;
        return typeof le === "boolean" ? le : true;
    }


    public destroy(): void {
        this.renderer.destroy();
        this.statusBar.destroy();
        this.configPanel.destroy();
        this.contextMenu.close();
        this.manualCopy.close();
        document.removeEventListener("keydown", this.onKeyDown, true);
    }
}

// ---------------------------------------------------------------------------
// Small color blend helper (theme derivation only — not a hardcoded theme color).
// ---------------------------------------------------------------------------

function blend(base: string, mix: string, amount: number): string {
    const a = hexToRgb(base);
    const b = hexToRgb(mix);
    const r = Math.round(a.r + (b.r - a.r) * amount);
    const g = Math.round(a.g + (b.g - a.g) * amount);
    const bl = Math.round(a.b + (b.b - a.b) * amount);
    return rgbToHex(r, g, bl);
}

function hexToRgb(hex: string): { r: number; g: number; b: number } {
    let h = (hex || "").trim().replace("#", "");
    if (h.length === 3) {
        h = h[0] + h[0] + h[1] + h[1] + h[2] + h[2];
    }
    const num = parseInt(h, 16);
    if (h.length !== 6 || isNaN(num)) {
        return { r: 255, g: 255, b: 255 };
    }
    return { r: (num >> 16) & 255, g: (num >> 8) & 255, b: num & 255 };
}

function rgbToHex(r: number, g: number, b: number): string {
    const h = (n: number): string => {
        const c = Math.max(0, Math.min(255, n)).toString(16);
        return c.length === 1 ? "0" + c : c;
    };
    return `#${h(r)}${h(g)}${h(b)}`;
}
