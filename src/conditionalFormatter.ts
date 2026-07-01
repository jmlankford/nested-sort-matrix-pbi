/*
 * conditionalFormatter.ts
 * -----------------------
 * Per-cell conditional formatting computation for the four supported types:
 *   - Color Scale (min/mid/max gradient; value or percent-of-range basis)
 *   - Rules (up to 5 ordered comparisons, first match wins)
 *   - Data Bars (in-cell horizontal bar with auto/fixed min-max and axis mode)
 *   - Icon Sets (up to 5 tiers chosen by threshold)
 *
 * Domains (min/max per leaf column) are precomputed from DATA cells only so the
 * gradient/bar scale is not skewed by subtotal/grand-total rows.
 */

import {
    CFSettings,
    CFRule,
    CFIconTier,
    IconShape
} from "./settings";
import { CfRule, CfApplyTo, RuleOperator } from "./cfPanel";
import { LeafColumn, RowTreeNode } from "./dataTransformer";

export interface DataBarFormat {
    /** 0..100 width of the bar as a percent of the cell. */
    widthPct: number;
    color: string;
    /** 0..100 position of the zero axis within the cell. */
    axisPct: number;
    negative: boolean;
}

export interface IconFormat {
    glyph: string;
    color: string;
}

export interface CellFormat {
    background?: string;
    fontColor?: string;
    dataBar?: DataBarFormat;
    icon?: IconFormat;
}

export interface Domain {
    min: number;
    max: number;
}

export type CFForSlot = (slotIndex: number) => CFSettings;

// ---------------------------------------------------------------------------
// Domain computation.
// ---------------------------------------------------------------------------

/**
 * Compute min/max per leaf column from DATA leaf nodes only, restricted to
 * columns whose value slot uses a domain-dependent CF type (colorScale/dataBars).
 */
export function computeDomains(
    leafNodes: RowTreeNode[],
    leafColumns: LeafColumn[],
    cfForSlot: CFForSlot
): Map<string, Domain> {
    const domains = new Map<string, Domain>();
    const relevant = leafColumns.filter((c) => {
        const t = cfForSlot(c.valueSlotIndex).cfType;
        return t === "colorScale" || t === "dataBars";
    });
    if (relevant.length === 0) {
        return domains;
    }
    relevant.forEach((c) => domains.set(c.id, { min: Number.POSITIVE_INFINITY, max: Number.NEGATIVE_INFINITY }));

    for (let i = 0; i < leafNodes.length; i++) {
        const values = leafNodes[i].values;
        for (let j = 0; j < relevant.length; j++) {
            const col = relevant[j];
            const v = values[col.id];
            if (v === null || v === undefined || isNaN(v)) {
                continue;
            }
            const d = domains.get(col.id) as Domain;
            if (v < d.min) {
                d.min = v;
            }
            if (v > d.max) {
                d.max = v;
            }
        }
    }

    // Normalize empty domains.
    domains.forEach((d, key) => {
        if (!isFinite(d.min) || !isFinite(d.max)) {
            domains.set(key, { min: 0, max: 0 });
        }
    });
    return domains;
}

// ---------------------------------------------------------------------------
// Main per-cell entry point.
// ---------------------------------------------------------------------------

export class ConditionalFormatter {
    /**
     * Compute the cell format for a value. Returns null when no formatting
     * applies. `isTotalRow` cells are formatted only when applyToTotals is set.
     */
    public format(
        value: number | null,
        cf: CFSettings,
        domain: Domain | undefined,
        isTotalRow: boolean
    ): CellFormat | null {
        if (cf.cfType === "none") {
            return null;
        }
        if (isTotalRow && !cf.applyToTotals) {
            return null;
        }
        if (value === null || value === undefined || isNaN(value)) {
            return null;
        }

        switch (cf.cfType) {
            case "colorScale":
                return this.colorScale(value, cf, domain);
            case "rules":
                // Prefer the in-visual panel's rulesV2 when present; otherwise
                // fall back to the legacy rulesJson array for compatibility.
                if (cf.rulesV2 && cf.rulesV2.length > 0) {
                    return this.rulesV2(value, cf.rulesV2, cf.cfApplyTo, cf.defaultColor);
                }
                return this.rules(value, cf.rules);
            case "dataBars":
                return this.dataBars(value, cf, domain);
            case "icons":
                return this.icons(value, cf.icons);
            case "fieldValue":
                // Field value CF is resolved in renderer.buildValueCell, which has
                // access to the hex string in the referenced measure column.
                return null;
            default:
                return null;
        }
    }

    /**
     * Evaluate the in-visual panel rules (CfRule[]) top-to-bottom, first match
     * wins. The matched rule's single color is applied as background/font/both
     * per applyTo; when no rule matches, the default color (if set) is used.
     */
    private rulesV2(
        value: number,
        rules: CfRule[],
        applyTo: CfApplyTo,
        defaultColor: string
    ): CellFormat | null {
        for (let i = 0; i < rules.length; i++) {
            if (evalRuleV2(value, rules[i])) {
                return applyColorAs(rules[i].color, applyTo);
            }
        }
        if (defaultColor && /^#[0-9A-Fa-f]{6}$/.test(defaultColor)) {
            return applyColorAs(defaultColor, applyTo);
        }
        return null;
    }

    private colorScale(value: number, cf: CFSettings, domain: Domain | undefined): CellFormat | null {
        if (!domain || domain.max === domain.min) {
            return { background: cf.csLowColor };
        }
        let t = (value - domain.min) / (domain.max - domain.min);
        t = Math.max(0, Math.min(1, t));
        // Percent basis is identical here because t is already a fraction of the
        // range; "value" basis maps raw value position — both reduce to t for a
        // linear scale, kept distinct for forward-compatibility.
        let color: string;
        if (cf.csUseMid) {
            if (t < 0.5) {
                color = hexLerp(cf.csLowColor, cf.csMidColor, t / 0.5);
            } else {
                color = hexLerp(cf.csMidColor, cf.csHighColor, (t - 0.5) / 0.5);
            }
        } else {
            color = hexLerp(cf.csLowColor, cf.csHighColor, t);
        }
        return { background: color, fontColor: contrastColor(color) };
    }

    private rules(value: number, rules: CFRule[]): CellFormat | null {
        for (let i = 0; i < rules.length; i++) {
            const r = rules[i];
            if (evalRule(value, r)) {
                const fmt: CellFormat = {};
                if (r.fill) {
                    fmt.background = r.fill;
                }
                if (r.fontColor) {
                    fmt.fontColor = r.fontColor;
                }
                return fmt;
            }
        }
        return null;
    }

    private dataBars(value: number, cf: CFSettings, domain: Domain | undefined): CellFormat | null {
        let min = cf.barMinMode === "fixed" ? cf.barMin : domain ? domain.min : 0;
        let max = cf.barMaxMode === "fixed" ? cf.barMax : domain ? domain.max : 0;
        // Ensure the axis logic works when data spans negatives.
        if (min > 0) {
            min = 0;
        }
        if (max < 0) {
            max = 0;
        }
        const span = max - min;
        if (span <= 0) {
            return null;
        }

        const negative = value < 0;
        let axisPct: number;
        if (cf.barAxis === "left") {
            axisPct = 0;
        } else if (cf.barAxis === "middle") {
            axisPct = 50;
        } else {
            // auto: axis at the zero position within [min,max]
            axisPct = (-min / span) * 100;
        }

        const valuePct = (Math.abs(value) / span) * 100;
        return {
            dataBar: {
                widthPct: Math.max(0, Math.min(100, valuePct)),
                color: negative ? cf.barNegativeColor : cf.barPositiveColor,
                axisPct: Math.max(0, Math.min(100, axisPct)),
                negative
            }
        };
    }

    private icons(value: number, tiers: CFIconTier[]): CellFormat | null {
        if (!tiers || tiers.length === 0) {
            return null;
        }
        // Choose the tier with the greatest threshold that the value meets.
        const sorted = tiers.slice().sort((a, b) => a.threshold - b.threshold);
        let chosen: CFIconTier | undefined = undefined;
        for (let i = 0; i < sorted.length; i++) {
            if (value >= sorted[i].threshold) {
                chosen = sorted[i];
            }
        }
        if (!chosen) {
            chosen = sorted[0];
        }
        return { icon: iconGlyph(chosen.icon) };
    }
}

// ---------------------------------------------------------------------------
// Rule evaluation.
// ---------------------------------------------------------------------------

function evalRule(value: number, r: CFRule): boolean {
    switch (r.operator) {
        case ">":
            return value > r.value1;
        case ">=":
            return value >= r.value1;
        case "<":
            return value < r.value1;
        case "<=":
            return value <= r.value1;
        case "=":
            return value === r.value1;
        case "between": {
            const lo = Math.min(r.value1, r.value2);
            const hi = Math.max(r.value1, r.value2);
            return value >= lo && value <= hi;
        }
        default:
            return false;
    }
}

/** Coerce a rule operand (number | string) to a finite number, or NaN. */
function toNum(v: number | string | undefined): number {
    if (v === undefined || v === null || v === "") {
        return NaN;
    }
    return typeof v === "number" ? v : Number(v);
}

/** Evaluate an in-visual CfRule against a (non-null) numeric cell value. */
function evalRuleV2(value: number, r: CfRule): boolean {
    if (!evalSingleCondition(value, r.operator, r.value1, r.value2)) {
        return false;
    }
    if (r.hasAnd && r.operator2) {
        return evalSingleCondition(value, r.operator2, r.andValue1 ?? 0, r.andValue2);
    }
    return true;
}

function evalSingleCondition(
    value: number,
    operator: RuleOperator,
    value1: number | string | undefined,
    value2?: number | string
): boolean {
    switch (operator) {
        case ">":
            return value > toNum(value1);
        case ">=":
            return value >= toNum(value1);
        case "<":
            return value < toNum(value1);
        case "<=":
            return value <= toNum(value1);
        case "=":
            return value === toNum(value1);
        case "!=":
            return value !== toNum(value1);
        case "between": {
            const a = toNum(value1);
            const b = toNum(value2);
            const lo = Math.min(a, b);
            const hi = Math.max(a, b);
            return value >= lo && value <= hi;
        }
        case "isBlank":
            return false;
        case "isNotBlank":
            return true;
        default:
            return false;
    }
}

/** Build a CellFormat applying a single color as background/font/both. */
function applyColorAs(color: string, applyTo: CfApplyTo): CellFormat {
    const fmt: CellFormat = {};
    if (applyTo !== "font") {
        fmt.background = color;
    }
    if (applyTo !== "background") {
        fmt.fontColor = color;
    }
    return fmt;
}

// ---------------------------------------------------------------------------
// Icon glyphs + colors.
// ---------------------------------------------------------------------------

const ICON_MAP: { [k in IconShape]: IconFormat } = {
    arrowUp: { glyph: "▲", color: "#3BA755" },
    arrowDown: { glyph: "▼", color: "#D64550" },
    arrowRight: { glyph: "▶", color: "#E0A800" },
    trafficGreen: { glyph: "●", color: "#3BA755" },
    trafficYellow: { glyph: "●", color: "#E0A800" },
    trafficRed: { glyph: "●", color: "#D64550" },
    flag: { glyph: "⚑", color: "#D64550" },
    circle: { glyph: "●", color: "#4C8BF5" },
    triangle: { glyph: "▲", color: "#E0A800" },
    diamond: { glyph: "◆", color: "#7B61FF" }
};

function iconGlyph(shape: IconShape): IconFormat {
    return ICON_MAP[shape] || ICON_MAP.circle;
}

// ---------------------------------------------------------------------------
// Color math.
// ---------------------------------------------------------------------------

interface RGB {
    r: number;
    g: number;
    b: number;
}

function parseHex(hex: string): RGB {
    let h = (hex || "").trim();
    if (h.charAt(0) === "#") {
        h = h.substring(1);
    }
    if (h.length === 3) {
        h = h.charAt(0) + h.charAt(0) + h.charAt(1) + h.charAt(1) + h.charAt(2) + h.charAt(2);
    }
    if (h.length !== 6) {
        return { r: 255, g: 255, b: 255 };
    }
    const num = parseInt(h, 16);
    if (isNaN(num)) {
        return { r: 255, g: 255, b: 255 };
    }
    return { r: (num >> 16) & 255, g: (num >> 8) & 255, b: num & 255 };
}

function toHex(c: RGB): string {
    const h = (n: number): string => {
        const s = Math.max(0, Math.min(255, Math.round(n))).toString(16);
        return s.length === 1 ? "0" + s : s;
    };
    return `#${h(c.r)}${h(c.g)}${h(c.b)}`;
}

function hexLerp(a: string, b: string, t: number): string {
    const ca = parseHex(a);
    const cb = parseHex(b);
    return toHex({
        r: ca.r + (cb.r - ca.r) * t,
        g: ca.g + (cb.g - ca.g) * t,
        b: ca.b + (cb.b - ca.b) * t
    });
}

/** Pick black or white text for legibility against a background fill. */
function contrastColor(bg: string): string {
    const c = parseHex(bg);
    // Relative luminance (sRGB approximation).
    const lum = (0.299 * c.r + 0.587 * c.g + 0.114 * c.b) / 255;
    return lum > 0.6 ? "#000000" : "#FFFFFF";
}
