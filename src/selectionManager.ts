/*
 * selectionManager.ts
 * -------------------
 * Thin wrapper around the host ISelectionManager. Owns the visual's selection
 * state (which row nodes are selected), implements multi-select (Ctrl/Cmd),
 * range-select (Shift), toggle-off (click an already-selected single row), and
 * clear-on-empty-space. Cross-filter emission is gated by the format-pane
 * Cross-Filter toggle.
 *
 * Selection ids are built from Table DataView row identities. A single row node
 * may map to several underlying table rows (e.g. when column fields split a
 * row-field combination across pivot columns), so selecting a node emits the
 * selection ids for ALL of its descendant data rows.
 */

import powerbi from "powerbi-visuals-api";
import ISelectionManager = powerbi.extensibility.ISelectionManager;
import ISelectionId = powerbi.visuals.ISelectionId;

import { RowTreeNode, collectRowIndices } from "./dataTransformer";

export interface ClickModifiers {
    ctrlKey: boolean;
    metaKey: boolean;
    shiftKey: boolean;
}

export type RowIdFactory = (rowIndex: number) => ISelectionId | undefined;

export class VisualSelectionManager {
    private selected = new Set<string>();
    /** Ordered, flattened list of currently selectable row nodes (for range select). */
    private selectables: RowTreeNode[] = [];
    private keyToNode = new Map<string, RowTreeNode>();
    private lastClickedKey: string | null = null;
    private idForRow: RowIdFactory = () => undefined;

    constructor(
        private readonly selectionManager: ISelectionManager,
        private readonly onChange: () => void,
        private readonly isEnabled: () => boolean
    ) {}

    /** Provide the per-update row-index -> ISelectionId factory. */
    public setRowIdFactory(factory: RowIdFactory): void {
        this.idForRow = factory;
    }

    /** Provide the ordered list of selectable nodes for the current render. */
    public setSelectables(nodes: RowTreeNode[]): void {
        this.selectables = nodes;
        this.keyToNode.clear();
        nodes.forEach((n) => this.keyToNode.set(n.key, n));
        // Drop selection entries that no longer exist (e.g. after collapse/refresh).
        const stale: string[] = [];
        this.selected.forEach((k) => {
            if (!this.keyToNode.has(k)) {
                stale.push(k);
            }
        });
        stale.forEach((k) => this.selected.delete(k));
    }

    public hasSelection(): boolean {
        return this.selected.size > 0;
    }

    public isSelected(node: RowTreeNode): boolean {
        return this.selected.has(node.key);
    }

    /** A node should be dimmed if there is a selection and it is not part of it. */
    public isDimmed(node: RowTreeNode): boolean {
        return this.selected.size > 0 && !this.selected.has(node.key);
    }

    public reset(): void {
        this.selected.clear();
        this.lastClickedKey = null;
        this.selectables = [];
        this.keyToNode.clear();
        // Best-effort clear of host selection.
        try {
            this.selectionManager.clear();
        } catch {
            /* host may not be ready during teardown */
        }
    }

    // -----------------------------------------------------------------------
    // Interaction.
    // -----------------------------------------------------------------------

    public handleRowClick(node: RowTreeNode, mods: ClickModifiers): void {
        if (!this.isEnabled()) {
            return;
        }
        // Only leaf-level DATA rows may build/emit an ISelectionId. Group header
        // rows (which have children) and subtotal/grand-total rows must never
        // cross-filter — otherwise expanding/collapsing or clicking a group would
        // filter the report and make other rows disappear.
        if (!node.isLeaf) {
            return;
        }
        const key = node.key;

        if (mods.shiftKey && this.lastClickedKey && this.keyToNode.has(this.lastClickedKey)) {
            this.selectRange(this.lastClickedKey, key);
        } else if (mods.ctrlKey || mods.metaKey) {
            if (this.selected.has(key)) {
                this.selected.delete(key);
            } else {
                this.selected.add(key);
            }
            this.lastClickedKey = key;
        } else {
            // Plain click: clicking the sole selected row clears it; otherwise
            // selection becomes just this row.
            if (this.selected.size === 1 && this.selected.has(key)) {
                this.selected.clear();
                this.lastClickedKey = null;
            } else {
                this.selected.clear();
                this.selected.add(key);
                this.lastClickedKey = key;
            }
        }

        this.syncHost();
        this.onChange();
    }

    public clearSelection(): void {
        if (this.selected.size === 0) {
            return;
        }
        this.selected.clear();
        this.lastClickedKey = null;
        this.syncHost();
        this.onChange();
    }

    private selectRange(fromKey: string, toKey: string): void {
        const fromIdx = this.selectables.findIndex((n) => n.key === fromKey);
        const toIdx = this.selectables.findIndex((n) => n.key === toKey);
        if (fromIdx < 0 || toIdx < 0) {
            return;
        }
        const lo = Math.min(fromIdx, toIdx);
        const hi = Math.max(fromIdx, toIdx);
        this.selected.clear();
        for (let i = lo; i <= hi; i++) {
            this.selected.add(this.selectables[i].key);
        }
        this.lastClickedKey = toKey;
    }

    /** Push the current selection to the host as cross-filter input. */
    private syncHost(): void {
        const ids: ISelectionId[] = [];
        this.selected.forEach((key) => {
            const node = this.keyToNode.get(key);
            if (!node) {
                return;
            }
            const rowIndices = collectRowIndices(node);
            for (let i = 0; i < rowIndices.length; i++) {
                const id = this.idForRow(rowIndices[i]);
                if (id) {
                    ids.push(id);
                }
            }
        });

        // Replace host selection with exactly our set.
        try {
            this.selectionManager.clear();
            if (ids.length > 0) {
                // multiSelect=true so the whole array is applied as one selection.
                void this.selectionManager.select(ids, true);
            }
        } catch {
            /* swallow host errors so the UI stays responsive */
        }
    }
}
