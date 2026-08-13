/*
 * selectionManager.ts
 * -------------------
 * Thin wrapper around the host ISelectionManager. Owns the visual's selection
 * state (which row nodes are selected), implements multi-select (Ctrl/Cmd),
 * range-select (Shift), toggle-off (click an already-selected single row), and
 * clear-on-empty-space. Cross-filter emission is gated by the format-pane
 * Cross-Filter toggle.
 *
 * Selection ids are built from Matrix DataView node identities (Phase 2): each
 * leaf RowTreeNode carries its own ISelectionId, built by the visual via
 * host.createSelectionIdBuilder().withMatrixNode(...) over its ancestor chain.
 * Selecting a leaf node emits exactly that node's selection id.
 */

import powerbi from "powerbi-visuals-api";
import ISelectionManager = powerbi.extensibility.ISelectionManager;
import ISelectionId = powerbi.visuals.ISelectionId;

import { RowTreeNode } from "./dataTransformer";

export interface ClickModifiers {
    ctrlKey: boolean;
    metaKey: boolean;
    shiftKey: boolean;
}

export class VisualSelectionManager {
    private selected = new Set<string>();
    /** Ordered, flattened list of currently selectable row nodes (for range select). */
    private selectables: RowTreeNode[] = [];
    private keyToNode = new Map<string, RowTreeNode>();
    private lastClickedKey: string | null = null;

    constructor(
        private readonly selectionManager: ISelectionManager,
        private readonly onChange: () => void,
        private readonly isEnabled: () => boolean
    ) {}

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

    /** Snapshot of the currently selected row-node keys (for clipboard copy). */
    public getSelectedKeys(): Set<string> {
        return new Set(this.selected);
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
        // Every row is selectable (Phase 2 Item 2). Selecting a group/subtotal row
        // cross-filters to that group's entire scope (all descendants); selecting a
        // child level narrows to that child's scope. Selection ids come from each
        // node's matrix identity (ancestor-chained withMatrixNode), so a group's id
        // matches the report scope of that group. A node without a resolvable
        // selection id (e.g. grand total) simply contributes nothing to the filter.
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
            if (!node || !node.selectionId) {
                return;
            }
            ids.push(node.selectionId);
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
