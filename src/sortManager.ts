/*
 * sortManager.ts
 * --------------
 * Owns nested (hierarchical) sort state and applies it scoped WITHIN parent
 * groups — the primary differentiator from the stock matrix. A value-column
 * sort orders the children of every group at every level by that column within
 * each parent's scope (not a global flat sort). A row-field sort orders the
 * groups at one specific hierarchy level within their parent scope.
 *
 * Click cycle per column: none -> ascending -> descending -> none.
 * Sorting a higher-level row field resets sort state on all deeper levels
 * (Excel PivotTable behaviour).
 *
 * Sort state lives in this instance for the visual's lifetime, so it survives
 * cross-filter updates and re-renders. The visual resets it on a full data
 * refresh (schema change).
 */

import {
    TransformResult,
    RowTreeNode,
    LeafColumn,
    FieldMeta
} from "./dataTransformer";
import powerbi from "powerbi-visuals-api";
import PrimitiveValue = powerbi.PrimitiveValue;

export type SortDirection = "asc" | "desc";

interface SortEntry {
    kind: "rowField" | "value";
    /** Row hierarchy level (rowField entries only). */
    level: number;
    /** Leaf column id (value entries only). */
    leafColId: string;
    /** Display label for the sort-stack summary. */
    label: string;
    direction: SortDirection;
}

const ASC_ARROW = "↑";
const DESC_ARROW = "↓";

export class SortManager {
    private stack: SortEntry[] = [];
    /** Direction for default pivot column ordering (label sort of col fields). */
    private columnDirection: SortDirection = "asc";

    public reset(): void {
        this.stack = [];
        this.columnDirection = "asc";
    }

    public isEmpty(): boolean {
        return this.stack.length === 0;
    }

    // -----------------------------------------------------------------------
    // Toggle handlers (called from header clicks).
    // -----------------------------------------------------------------------

    /**
     * Cycle the sort on a row-field column at `level`. Resets sorts on all
     * deeper levels and clears value sorts (which are inner-scoped).
     */
    public toggleRowField(level: number, label: string): void {
        // Find an existing entry for this exact level.
        const idx = this.stack.findIndex((e) => e.kind === "rowField" && e.level === level);
        const current = idx >= 0 ? this.stack[idx].direction : undefined;

        // Reset deeper-level row sorts and all value sorts.
        this.stack = this.stack.filter(
            (e) => e.kind === "rowField" && e.level < level
        );

        const next = this.nextDirection(current);
        if (next) {
            this.stack.push({ kind: "rowField", level, leafColId: "", label, direction: next });
        }
    }

    /** Cycle the sort on a value column (applies within every parent scope). */
    public toggleValue(leafColId: string, label: string): void {
        const idx = this.stack.findIndex((e) => e.kind === "value" && e.leafColId === leafColId);
        const current = idx >= 0 ? this.stack[idx].direction : undefined;

        if (idx >= 0) {
            this.stack.splice(idx, 1);
        }

        const next = this.nextDirection(current);
        if (next) {
            this.stack.push({ kind: "value", level: -1, leafColId, label, direction: next });
        }
    }

    /** Toggle the default ascending/descending ordering of pivot columns. */
    public toggleColumnDirection(): void {
        this.columnDirection = this.columnDirection === "asc" ? "desc" : "asc";
    }

    public getColumnDirection(): SortDirection {
        return this.columnDirection;
    }

    private nextDirection(current: SortDirection | undefined): SortDirection | undefined {
        if (current === undefined) {
            return "asc";
        }
        if (current === "asc") {
            return "desc";
        }
        return undefined; // desc -> none
    }

    // -----------------------------------------------------------------------
    // Query helpers (for header arrow rendering).
    // -----------------------------------------------------------------------

    public directionForRowField(level: number): SortDirection | undefined {
        const e = this.stack.find((x) => x.kind === "rowField" && x.level === level);
        return e ? e.direction : undefined;
    }

    public directionForValue(leafColId: string): SortDirection | undefined {
        const e = this.stack.find((x) => x.kind === "value" && x.leafColId === leafColId);
        return e ? e.direction : undefined;
    }

    public priorityForValue(leafColId: string): number {
        return this.stack.findIndex((x) => x.kind === "value" && x.leafColId === leafColId);
    }

    // -----------------------------------------------------------------------
    // Apply nested sort to a transform result (mutates children arrays).
    // -----------------------------------------------------------------------

    public applyNestedSort(result: TransformResult): void {
        if (!result.hasRowFields) {
            this.applyColumnOrder(result);
            return;
        }
        // Level 0 == top-level roots.
        this.sortSiblings(result.rootNodes, 0);
        result.rootNodes.forEach((n) => this.sortRecursive(n));
        this.applyColumnOrder(result);
    }

    private sortRecursive(node: RowTreeNode): void {
        if (node.isLeaf || node.children.length === 0) {
            return;
        }
        const childLevel = node.level + 1;
        this.sortSiblings(node.children, childLevel);
        node.children.forEach((c) => this.sortRecursive(c));
    }

    private sortSiblings(siblings: RowTreeNode[], level: number): void {
        const applicable = this.stack.filter(
            (e) => e.kind === "value" || (e.kind === "rowField" && e.level === level)
        );
        if (applicable.length === 0) {
            return; // preserve first-seen order
        }
        siblings.sort((a, b) => {
            for (let i = 0; i < applicable.length; i++) {
                const e = applicable[i];
                let c = 0;
                if (e.kind === "value") {
                    c = compareNullableNumber(a.values[e.leafColId], b.values[e.leafColId]);
                } else {
                    c = comparePrimitive(a.rawValue, b.rawValue, a.label, b.label);
                }
                if (e.direction === "desc") {
                    c = -c;
                }
                if (c !== 0) {
                    return c;
                }
            }
            return 0;
        });
    }

    /** Re-order leaf columns when the user flips pivot column direction. */
    private applyColumnOrder(result: TransformResult): void {
        if (!result.isPivot || this.columnDirection === "asc") {
            return; // ascending is the transformer's default build order
        }
        // Descending: reverse the pivot-leaf ordering while keeping each group's
        // value-slot blocks intact and leaving grand-total columns at the end.
        const grandTotals = result.leafColumns.filter((c) => c.isColGrandTotal);
        const body = result.leafColumns.filter((c) => !c.isColGrandTotal);
        body.reverse();
        result.leafColumns = body.concat(grandTotals);
        result.columnHeader.leafRow = result.leafColumns;
        result.columnHeader.pivotRows.forEach((row) => row.reverse());
    }

    // -----------------------------------------------------------------------
    // Sort-stack summary text for the status bar.
    // -----------------------------------------------------------------------

    public getStackText(maxChars: number = 80): string {
        if (this.stack.length === 0) {
            return "";
        }
        const parts = this.stack.map(
            (e) => `${e.label} ${e.direction === "asc" ? ASC_ARROW : DESC_ARROW}`
        );
        let text = parts.join(" → ");
        if (text.length > maxChars) {
            text = text.substring(0, Math.max(0, maxChars - 1)) + "…";
        }
        return text;
    }

    /**
     * Rebuild any stale labels after a rename / re-transform so the stack text
     * stays in sync with current field display names. Entries whose target no
     * longer exists are dropped.
     */
    public reconcile(
        activeRowFields: FieldMeta[],
        leafColumns: LeafColumn[],
        leafLabel: (col: LeafColumn) => string
    ): void {
        const rowByLevel = new Map<number, FieldMeta>();
        activeRowFields.forEach((f, idx) => rowByLevel.set(idx, f));
        const leafById = new Map<string, LeafColumn>();
        leafColumns.forEach((c) => leafById.set(c.id, c));

        this.stack = this.stack.filter((e) => {
            if (e.kind === "rowField") {
                const field = rowByLevel.get(e.level);
                if (!field) {
                    return false;
                }
                e.label = field.displayName;
                return true;
            }
            const col = leafById.get(e.leafColId);
            if (!col) {
                return false;
            }
            e.label = leafLabel(col);
            return true;
        });
    }
}

// ---------------------------------------------------------------------------
// Comparison primitives.
// ---------------------------------------------------------------------------

function compareNullableNumber(a: number | null | undefined, b: number | null | undefined): number {
    const an = a === null || a === undefined;
    const bn = b === null || b === undefined;
    if (an && bn) {
        return 0;
    }
    if (an) {
        return 1; // nulls sort last in ascending
    }
    if (bn) {
        return -1;
    }
    return (a as number) - (b as number);
}

function comparePrimitive(
    aRaw: PrimitiveValue | null,
    bRaw: PrimitiveValue | null,
    aLabel: string,
    bLabel: string
): number {
    const aNum = typeof aRaw === "number";
    const bNum = typeof bRaw === "number";
    if (aNum && bNum) {
        return (aRaw as number) - (bRaw as number);
    }
    if (aRaw instanceof Date && bRaw instanceof Date) {
        return aRaw.getTime() - bRaw.getTime();
    }
    return aLabel.localeCompare(bLabel, undefined, { numeric: true, sensitivity: "base" });
}
