/*
 * sortManager.ts
 * --------------
 * Owns the visual's SINGLE active sort and applies it scoped WITHIN parent
 * groups at EVERY hierarchy level — stock Power BI matrix behaviour.
 *
 * There is exactly one active sort at a time (or none). Clicking any column
 * header — a value column OR a row-field column — makes that column the active
 * sort, replacing whatever was active. There is no priority stack.
 *   - A value sort orders every group's children (at every level) by that value
 *     column's aggregate, within each parent's scope.
 *   - A row-field sort orders every level's siblings by their own row label
 *     (rawValue), within each parent's scope; the indicator shows on the clicked
 *     header only.
 * Click cycle on the SAME column: ascending -> descending -> clear. Clicking a
 * different column starts fresh at ascending.
 *
 * Sort state lives for the visual's lifetime and survives cross-filter updates,
 * re-renders, AND field add/remove (it is only cleared by reset()).
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

interface ActiveSort {
    kind: "rowField" | "value";
    /** Row hierarchy level of the clicked header (rowField only; -1 for value). */
    level: number;
    /** Leaf column id (value only; "" for rowField). */
    leafColId: string;
    /** Display label for the status-bar summary. */
    label: string;
    direction: SortDirection;
}

const ASC_ARROW = "↑";
const DESC_ARROW = "↓";

export class SortManager {
    /** The single active sort, or null when nothing is sorted. */
    private active: ActiveSort | null = null;
    /** Direction for default pivot column ordering (label sort of col fields). */
    private columnDirection: SortDirection = "asc";

    public reset(): void {
        this.active = null;
        this.columnDirection = "asc";
    }

    public isEmpty(): boolean {
        return this.active === null;
    }

    // -----------------------------------------------------------------------
    // Toggle handlers (called from header clicks). One active sort at a time.
    // -----------------------------------------------------------------------

    /** Make (or cycle) a row-field column the active sort. */
    public toggleRowField(level: number, label: string): void {
        if (this.active && this.active.kind === "rowField" && this.active.level === level) {
            const next = this.cycle(this.active.direction);
            this.active = next ? { ...this.active, direction: next } : null;
        } else {
            this.active = { kind: "rowField", level, leafColId: "", label, direction: "asc" };
        }
    }

    /** Make (or cycle) a value column the active sort. */
    public toggleValue(leafColId: string, label: string): void {
        if (this.active && this.active.kind === "value" && this.active.leafColId === leafColId) {
            const next = this.cycle(this.active.direction);
            this.active = next ? { ...this.active, direction: next } : null;
        } else {
            this.active = { kind: "value", level: -1, leafColId, label, direction: "asc" };
        }
    }

    /** Toggle the default ascending/descending ordering of pivot columns. */
    public toggleColumnDirection(): void {
        this.columnDirection = this.columnDirection === "asc" ? "desc" : "asc";
    }

    public getColumnDirection(): SortDirection {
        return this.columnDirection;
    }

    /** Same-column click cycle: asc -> desc -> clear. */
    private cycle(current: SortDirection): SortDirection | undefined {
        return current === "asc" ? "desc" : undefined;
    }

    // -----------------------------------------------------------------------
    // Query helpers (for header arrow rendering) — only the active column.
    // -----------------------------------------------------------------------

    public directionForRowField(level: number): SortDirection | undefined {
        return this.active && this.active.kind === "rowField" && this.active.level === level
            ? this.active.direction
            : undefined;
    }

    public directionForValue(leafColId: string): SortDirection | undefined {
        return this.active && this.active.kind === "value" && this.active.leafColId === leafColId
            ? this.active.direction
            : undefined;
    }

    // -----------------------------------------------------------------------
    // Apply the active sort to a transform result (mutates children arrays).
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

    private sortSiblings(siblings: RowTreeNode[], _level: number): void {
        const a = this.active;
        if (!a) {
            return; // no active sort -> preserve engine order
        }
        siblings.sort((x, y) => {
            let c =
                a.kind === "value"
                    ? compareNullableNumber(x.values[a.leafColId], y.values[a.leafColId])
                    : comparePrimitive(x.rawValue, y.rawValue, x.label, y.label);
            if (a.direction === "desc") {
                c = -c;
            }
            return c;
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
    // Active-sort summary text for the status bar.
    // -----------------------------------------------------------------------

    public getStackText(maxChars: number = 80): string {
        if (!this.active) {
            return "";
        }
        const arrow = this.active.direction === "asc" ? ASC_ARROW : DESC_ARROW;
        let text = `${this.active.label} ${arrow}`;
        if (text.length > maxChars) {
            text = text.substring(0, Math.max(0, maxChars - 1)) + "…";
        }
        return text;
    }

    /**
     * Refresh the active sort's display label after a rename / re-transform. The
     * active sort is KEPT even when its target column/level is temporarily absent
     * (collapsed hierarchy) or after a field add/remove — it re-applies when the
     * target returns; only reset() clears it.
     */
    public reconcile(
        activeRowFields: FieldMeta[],
        leafColumns: LeafColumn[],
        leafLabel: (col: LeafColumn) => string
    ): void {
        if (!this.active) {
            return;
        }
        if (this.active.kind === "rowField") {
            const field = activeRowFields[this.active.level];
            if (field) {
                this.active.label = field.displayName;
            }
        } else {
            const col = leafColumns.find((c) => c.id === this.active!.leafColId);
            if (col) {
                this.active.label = leafLabel(col);
            }
        }
    }
}

// ---------------------------------------------------------------------------
// Comparison primitives.
// ---------------------------------------------------------------------------

/**
 * Layered, DETERMINISTIC text collation. `sensitivity: "base"` folds case and
 * accents (and numeric collation treats "1.0" and "1" as equal), so two visibly
 * distinct labels can compare equal and be left in engine order by the stable
 * sort — which is why row-field sorts appeared to reorder only the apex level
 * while deeper siblings (that happened to differ only by case/accent/format)
 * stayed put. Fall through base -> variant -> raw code point so distinct strings
 * always order.
 */
function collateLabel(a: string, b: string): number {
    const base = a.localeCompare(b, undefined, { numeric: true, sensitivity: "base" });
    if (base !== 0) {
        return base;
    }
    const variant = a.localeCompare(b, undefined, { numeric: true, sensitivity: "variant" });
    if (variant !== 0) {
        return variant;
    }
    return a < b ? -1 : a > b ? 1 : 0;
}

function compareNullableNumber(
    a: number | string | null | undefined,
    b: number | string | null | undefined
): number {
    const an = a === null || a === undefined;
    const bn = b === null || b === undefined;
    if (an && bn) {
        return 0;
    }
    // Blank/null is treated as negative infinity for ordering: it sorts FIRST in
    // ascending order (and, since the caller negates for desc, LAST in descending).
    if (an) {
        return -1;
    }
    if (bn) {
        return 1;
    }
    // Both numbers: numeric compare. Otherwise (text measures) compare lexically;
    // an empty string naturally sorts before any non-empty value.
    if (typeof a === "number" && typeof b === "number") {
        return a - b;
    }
    return collateLabel(String(a), String(b));
}

function comparePrimitive(
    aRaw: PrimitiveValue | null,
    bRaw: PrimitiveValue | null,
    aLabel: string,
    bLabel: string
): number {
    // True numeric / date fields order by their underlying value (2 before 10;
    // chronological), but ONLY when the values actually differ — an equal raw
    // must still fall through to the label so distinct labels never tie.
    if (typeof aRaw === "number" && typeof bRaw === "number" && aRaw !== bRaw) {
        return aRaw - bRaw;
    }
    if (aRaw instanceof Date && bRaw instanceof Date && aRaw.getTime() !== bRaw.getTime()) {
        return aRaw.getTime() - bRaw.getTime();
    }
    // Everything else — text fields, mixed types, or equal raw values — orders by
    // the DISPLAYED row label at every hierarchy level, deterministically.
    return collateLabel(aLabel, bLabel);
}
