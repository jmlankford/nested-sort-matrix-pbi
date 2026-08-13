/*
 * cfTypes.ts
 * ----------
 * Shared conditional-formatting type definitions consumed by the CF evaluator
 * (conditionalFormatter.ts) and the per-measure settings parser (settings.ts).
 *
 * These types were previously declared in the in-visual CF panel module; that
 * panel has been removed (CF is now edited via the format-pane fx dialog), but
 * the types remain so persisted `cfSettings` from existing reports still parse
 * and render.
 */

export type CfType = "none" | "colorScale" | "rules" | "fieldValue";
export type CfApplyTo = "background" | "font" | "both";
export type RuleOperator =
    | ">="
    | ">"
    | "<="
    | "<"
    | "="
    | "!="
    | "between"
    | "isBlank"
    | "isNotBlank";

export interface CfRule {
    operator: RuleOperator;
    value1: number | string;
    value2?: number | string; // "between" second bound only
    hasAnd?: boolean; // whether a second AND condition exists
    operator2?: RuleOperator; // second AND condition operator
    andValue1?: number | string; // second AND condition value
    andValue2?: number | string; // second AND condition "between" bound
    color: string; // hex string e.g. "#FF6B6B"
}

export interface CfColorScale {
    lowColor: string;
    useMid: boolean;
    midColor: string;
    highColor: string;
    basis: "value" | "percent";
}

export interface CfFieldValue {
    measureSlotIndex: number; // slotIndex of the measure whose value is a hex string
    applyAs: CfApplyTo;
}
