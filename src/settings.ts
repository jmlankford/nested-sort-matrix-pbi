/*
 * settings.ts
 * -----------
 * All format-pane property definitions for the Nested Sort Matrix visual, plus
 * the strongly-typed parsing layer that converts a Power BI DataView's `objects`
 * bags into typed settings the rest of the visual consumes.
 *
 * The format-pane MODEL (the cards/groups/slices shown in the pane) is built
 * dynamically in visual.ts#getFormattingModel because the Values and Conditional
 * Formatting sections are per-slot (one card per active measure). This file owns
 * the property names, defaults, enums, and the read-back logic.
 *
 * Per-slot value formatting and conditional formatting are stored as per-column
 * object instances (DataView column.objects), keyed by the measure column. The
 * generic (visual-level) sections are stored on dataView.metadata.objects.
 */

import powerbi from "powerbi-visuals-api";
import DataViewObjects = powerbi.DataViewObjects;
import DataViewObject = powerbi.DataViewObject;
import { CfRule, CfApplyTo } from "./cfPanel";

// ---------------------------------------------------------------------------
// Slot capacity constants — shared across the whole visual.
// ---------------------------------------------------------------------------
export const MAX_ROW_FIELDS = 10;
export const MAX_VALUE_FIELDS = 15;
export const MAX_COL_FIELDS = 5;

// Capabilities role-name prefixes (must match capabilities.json exactly).
export const ROW_ROLE_PREFIX = "rowField";
export const VALUE_ROLE_PREFIX = "valueField";
export const COL_ROLE_PREFIX = "colField";

// ---------------------------------------------------------------------------
// Enumerations used by the settings.
// ---------------------------------------------------------------------------
export type RowHeightMode = "auto" | "fixed";
export type LayoutMode = "compact" | "outline" | "tabular";
export type ExpandStyle = "plusMinus" | "chevron" | "triangle";
export type DisplayUnit = "none" | "thousands" | "millions" | "billions";
export type ColumnApplyTo = "all" | "values" | "header" | "subtotals" | "grandTotal";
export type ColumnAlignment = "auto" | "left" | "center" | "right";
export type CFType = "none" | "colorScale" | "rules" | "dataBars" | "icons" | "fieldValue";
export type ColorScaleBasis = "value" | "percent";
export type BarAxis = "auto" | "left" | "middle";
export type MinMaxMode = "auto" | "fixed";
export type RuleOperator = ">" | ">=" | "<" | "<=" | "=" | "between";
export type IconShape =
    | "arrowUp"
    | "arrowDown"
    | "arrowRight"
    | "trafficGreen"
    | "trafficYellow"
    | "trafficRed"
    | "flag"
    | "circle"
    | "triangle"
    | "diamond";

// ---------------------------------------------------------------------------
// Settings interfaces.
// ---------------------------------------------------------------------------
export interface GridSettings {
    rowHeightMode: RowHeightMode;
    rowHeightPx: number;
    fontSize: number;
    /** Value-cell font family (global grid setting). */
    valueFontFamily: string;
}

export interface RowHeaderSettings {
    bold: boolean;
    fontSize: number;
    indentPerLevel: number;
    /** Row-header font family. */
    rowFontFamily: string;
    /** Row layout mode (compact / outline / tabular). */
    layoutMode: LayoutMode;
    /** Repeat ancestor labels on the first visible row when scrolled. */
    repeatRowHeaders: boolean;
}

export interface ColumnHeaderSettings {
    bold: boolean;
    fontSize: number;
    showSortArrows: boolean;
    /** Column-header font family. */
    columnFontFamily: string;
}

export interface ExpandCollapseSettings {
    show: boolean;
    buttonSize: number;
    /** Empty string means "use theme foreground". */
    buttonColor: string;
    style: ExpandStyle;
}

export interface SubtotalSettings {
    rowSubtotals: boolean;
    /** Per-level overrides, index 0 == Level 1 ... index 9 == Level 10. */
    levels: boolean[];
    columnSubtotals: boolean;
    grandTotalRow: boolean;
    grandTotalColumn: boolean;
    labelText: string;
    applyCfToTotals: boolean;
}

export interface AlternateRowSettings {
    show: boolean;
    color: string;
}

export interface StatusBarSettings {
    show: boolean;
}

export interface CrossFilterSettings {
    show: boolean;
}

/** Per-slot value formatting (font, color, and number formatting). */
export interface ValueFormatSettings {
    fontFamily: string;
    fontSize: number;
    bold: boolean;
    italic: boolean;
    textColor: string | null;
    backgroundColor: string | null;
    altTextColor: string | null;
    altBackgroundColor: string | null;
    textWrap: boolean;
    numberFormat: string; // explicit .NET/d3 format override; empty == inherit
    decimals: number; // -1 == use format string default
    unit: DisplayUnit;
    prefix: string;
    suffix: string;
}

/** Per-slot "specific column" final override layer. */
export interface SpecificColumnSettings {
    applyTo: ColumnApplyTo;
    textColor: string | null;
    backgroundColor: string | null;
    alignment: ColumnAlignment;
    unit: DisplayUnit;
    decimals: number; // -1 == no override
}

export interface CFRule {
    operator: RuleOperator;
    value1: number;
    value2: number; // used only for "between"
    fill: string;
    fontColor: string;
}

export interface CFIconTier {
    icon: IconShape;
    threshold: number; // lower bound for this tier (>=)
}

/** Per-slot conditional formatting. */
export interface CFSettings {
    cfType: CFType;
    applyToTotals: boolean;
    // Color scale
    csLowColor: string;
    csUseMid: boolean;
    csMidColor: string;
    csHighColor: string;
    csBasis: ColorScaleBasis;
    // Rules
    rules: CFRule[];
    // Data bars
    barPositiveColor: string;
    barNegativeColor: string;
    barAxis: BarAxis;
    barMinMode: MinMaxMode;
    barMin: number;
    barMaxMode: MinMaxMode;
    barMax: number;
    // Icons
    icons: CFIconTier[];
    // In-visual CF panel additions.
    cfApplyTo: CfApplyTo;
    defaultColor: string;
    fieldValueSlot: number;
    fieldValueApplyAs: CfApplyTo;
    rulesV2: CfRule[];
}

export interface VisualSettings {
    grid: GridSettings;
    rowHeaders: RowHeaderSettings;
    columnHeaders: ColumnHeaderSettings;
    expandCollapse: ExpandCollapseSettings;
    subtotals: SubtotalSettings;
    alternateRows: AlternateRowSettings;
    statusBar: StatusBarSettings;
    crossFilter: CrossFilterSettings;
    /** Persisted config-panel state (JSON string), read from `general.configState`. */
    configState: string;
    /** Persisted column widths (JSON string), read from `general.columnWidths`. */
    columnWidths: string;
}

// ---------------------------------------------------------------------------
// Defaults. These are theme-neutral; renderer overlays theme colors at runtime.
// ---------------------------------------------------------------------------
export const DEFAULTS = {
    grid: {
        rowHeightMode: "auto" as RowHeightMode,
        rowHeightPx: 28,
        fontSize: 12,
        // Default matches the first font-family dropdown option (capabilities.json).
        valueFontFamily: "Segoe UI"
    },
    rowHeaders: {
        bold: false,
        fontSize: 11,
        indentPerLevel: 16,
        // Defaults match the first font-family dropdown option (capabilities.json).
        rowFontFamily: "Segoe UI",
        layoutMode: "compact" as LayoutMode,
        repeatRowHeaders: false
    },
    columnHeaders: {
        bold: true,
        fontSize: 13,
        showSortArrows: true,
        columnFontFamily: "Segoe UI"
    },
    expandCollapse: {
        show: true,
        buttonSize: 12,
        // Empty == use theme foreground at render time.
        buttonColor: "",
        style: "plusMinus" as ExpandStyle
    },
    subtotals: {
        rowSubtotals: true,
        levels: [true, true, true, true, true, true, true, true, true, true],
        columnSubtotals: true,
        grandTotalRow: true,
        grandTotalColumn: true,
        labelText: "Total",
        applyCfToTotals: false
    },
    alternateRows: {
        show: false,
        color: "#F5F5F5"
    },
    statusBar: {
        show: true
    },
    crossFilter: {
        show: true
    },
    valueFormat: {
        fontFamily: "Segoe UI",
        fontSize: 12,
        bold: false,
        italic: false,
        textColor: null,
        backgroundColor: null,
        altTextColor: null,
        altBackgroundColor: null,
        textWrap: false,
        numberFormat: "",
        decimals: -1,
        unit: "none" as DisplayUnit,
        prefix: "",
        suffix: ""
    } as ValueFormatSettings,
    specificColumn: {
        applyTo: "all" as ColumnApplyTo,
        textColor: null,
        backgroundColor: null,
        alignment: "auto" as ColumnAlignment,
        unit: "none" as DisplayUnit,
        decimals: -1
    } as SpecificColumnSettings,
    cf: {
        cfType: "none" as CFType,
        applyToTotals: false,
        csLowColor: "#FFFFFF",
        csUseMid: false,
        csMidColor: "#FFEB84",
        csHighColor: "#63BE7B",
        csBasis: "value" as ColorScaleBasis,
        rules: [] as CFRule[],
        barPositiveColor: "#4C8BF5",
        barNegativeColor: "#E66C5C",
        barAxis: "auto" as BarAxis,
        barMinMode: "auto" as MinMaxMode,
        barMin: 0,
        barMaxMode: "auto" as MinMaxMode,
        barMax: 0,
        icons: [] as CFIconTier[],
        cfApplyTo: "background" as CfApplyTo,
        defaultColor: "",
        fieldValueSlot: -1,
        fieldValueApplyAs: "background" as CfApplyTo,
        rulesV2: [] as CfRule[]
    } as CFSettings
};

// ---------------------------------------------------------------------------
// Low-level typed getters over DataViewObjects.
// These avoid coupling to a specific utils version and keep strict typing.
// ---------------------------------------------------------------------------
function getObject(
    objects: DataViewObjects | undefined,
    objectName: string
): DataViewObject | undefined {
    if (!objects) {
        return undefined;
    }
    return objects[objectName];
}

export function getBool(
    objects: DataViewObjects | undefined,
    objectName: string,
    propertyName: string,
    defaultValue: boolean
): boolean {
    const obj = getObject(objects, objectName);
    if (obj && typeof obj[propertyName] === "boolean") {
        return obj[propertyName] as boolean;
    }
    return defaultValue;
}

export function getNumber(
    objects: DataViewObjects | undefined,
    objectName: string,
    propertyName: string,
    defaultValue: number
): number {
    const obj = getObject(objects, objectName);
    if (obj && obj[propertyName] != null) {
        const raw = obj[propertyName];
        const n = typeof raw === "number" ? raw : Number(raw);
        if (!isNaN(n)) {
            return n;
        }
    }
    return defaultValue;
}

export function getText(
    objects: DataViewObjects | undefined,
    objectName: string,
    propertyName: string,
    defaultValue: string
): string {
    const obj = getObject(objects, objectName);
    if (obj && obj[propertyName] != null) {
        return String(obj[propertyName]);
    }
    return defaultValue;
}

export function getEnum<T extends string>(
    objects: DataViewObjects | undefined,
    objectName: string,
    propertyName: string,
    defaultValue: T
): T {
    const obj = getObject(objects, objectName);
    if (obj && obj[propertyName] != null) {
        return String(obj[propertyName]) as T;
    }
    return defaultValue;
}

export function getFill(
    objects: DataViewObjects | undefined,
    objectName: string,
    propertyName: string,
    defaultValue: string
): string {
    const obj = getObject(objects, objectName);
    if (obj && obj[propertyName] != null) {
        const raw = obj[propertyName] as powerbi.Fill | string;
        if (typeof raw === "string") {
            return raw;
        }
        const fill = raw as powerbi.Fill;
        if (fill && fill.solid && fill.solid.color) {
            return fill.solid.color as string;
        }
    }
    return defaultValue;
}

// ---------------------------------------------------------------------------
// Top-level parsing.
// ---------------------------------------------------------------------------
export function parseVisualSettings(metadataObjects: DataViewObjects | undefined): VisualSettings {
    const levels: boolean[] = [];
    for (let i = 0; i < MAX_ROW_FIELDS; i++) {
        levels.push(
            getBool(metadataObjects, "subtotals", `level${i + 1}`, DEFAULTS.subtotals.levels[i])
        );
    }

    return {
        grid: {
            rowHeightMode: getEnum<RowHeightMode>(
                metadataObjects,
                "grid",
                "rowHeightMode",
                DEFAULTS.grid.rowHeightMode
            ),
            rowHeightPx: getNumber(metadataObjects, "grid", "rowHeightPx", DEFAULTS.grid.rowHeightPx),
            fontSize: getNumber(metadataObjects, "grid", "fontSize", DEFAULTS.grid.fontSize),
            valueFontFamily: getText(
                metadataObjects,
                "grid",
                "valueFontFamily",
                DEFAULTS.grid.valueFontFamily
            )
        },
        rowHeaders: {
            bold: getBool(metadataObjects, "rowHeaders", "bold", DEFAULTS.rowHeaders.bold),
            fontSize: getNumber(metadataObjects, "rowHeaders", "fontSize", DEFAULTS.rowHeaders.fontSize),
            indentPerLevel: getNumber(
                metadataObjects,
                "rowHeaders",
                "indentPerLevel",
                DEFAULTS.rowHeaders.indentPerLevel
            ),
            rowFontFamily: getText(
                metadataObjects,
                "rowHeaders",
                "rowFontFamily",
                DEFAULTS.rowHeaders.rowFontFamily
            ),
            layoutMode: getEnum<LayoutMode>(
                metadataObjects,
                "layoutOptions",
                "layoutMode",
                DEFAULTS.rowHeaders.layoutMode
            ),
            repeatRowHeaders: getBool(
                metadataObjects,
                "layoutOptions",
                "repeatRowHeaders",
                DEFAULTS.rowHeaders.repeatRowHeaders
            )
        },
        columnHeaders: {
            bold: getBool(metadataObjects, "columnHeaders", "bold", DEFAULTS.columnHeaders.bold),
            fontSize: getNumber(
                metadataObjects,
                "columnHeaders",
                "fontSize",
                DEFAULTS.columnHeaders.fontSize
            ),
            showSortArrows: getBool(
                metadataObjects,
                "columnHeaders",
                "showSortArrows",
                DEFAULTS.columnHeaders.showSortArrows
            ),
            columnFontFamily: getText(
                metadataObjects,
                "columnHeaders",
                "columnFontFamily",
                DEFAULTS.columnHeaders.columnFontFamily
            )
        },
        expandCollapse: {
            show: getBool(metadataObjects, "expandCollapse", "show", DEFAULTS.expandCollapse.show),
            buttonSize: getNumber(
                metadataObjects,
                "expandCollapse",
                "buttonSize",
                DEFAULTS.expandCollapse.buttonSize
            ),
            buttonColor: getFill(
                metadataObjects,
                "expandCollapse",
                "buttonColor",
                DEFAULTS.expandCollapse.buttonColor
            ),
            style: getEnum<ExpandStyle>(
                metadataObjects,
                "expandCollapse",
                "style",
                DEFAULTS.expandCollapse.style
            )
        },
        subtotals: {
            rowSubtotals: getBool(
                metadataObjects,
                "subtotals",
                "rowSubtotals",
                DEFAULTS.subtotals.rowSubtotals
            ),
            levels,
            columnSubtotals: getBool(
                metadataObjects,
                "subtotals",
                "columnSubtotals",
                DEFAULTS.subtotals.columnSubtotals
            ),
            grandTotalRow: getBool(
                metadataObjects,
                "subtotals",
                "grandTotalRow",
                DEFAULTS.subtotals.grandTotalRow
            ),
            grandTotalColumn: getBool(
                metadataObjects,
                "subtotals",
                "grandTotalColumn",
                DEFAULTS.subtotals.grandTotalColumn
            ),
            labelText: getText(metadataObjects, "subtotals", "labelText", DEFAULTS.subtotals.labelText),
            applyCfToTotals: getBool(
                metadataObjects,
                "subtotals",
                "applyCfToTotals",
                DEFAULTS.subtotals.applyCfToTotals
            )
        },
        alternateRows: {
            show: getBool(metadataObjects, "alternateRows", "show", DEFAULTS.alternateRows.show),
            color: getFill(metadataObjects, "alternateRows", "color", DEFAULTS.alternateRows.color)
        },
        statusBar: {
            show: getBool(metadataObjects, "statusBar", "show", DEFAULTS.statusBar.show)
        },
        crossFilter: {
            show: getBool(metadataObjects, "crossFilter", "show", DEFAULTS.crossFilter.show)
        },
        configState: getText(metadataObjects, "general", "configState", ""),
        columnWidths: getText(metadataObjects, "general", "columnWidths", "")
    };
}

// ---------------------------------------------------------------------------
// Per-slot parsing (read from a measure column's own objects bag).
// ---------------------------------------------------------------------------
/** Read a fill color that may be unset; returns null when not present. */
export function getFillNullable(
    objects: DataViewObjects | undefined,
    objectName: string,
    propertyName: string
): string | null {
    const obj = getObject(objects, objectName);
    if (obj && obj[propertyName] != null) {
        const raw = obj[propertyName] as powerbi.Fill | string;
        if (typeof raw === "string") {
            return raw;
        }
        const fill = raw as powerbi.Fill;
        if (fill && fill.solid && fill.solid.color) {
            return fill.solid.color as string;
        }
    }
    return null;
}

export function parseValueFormat(columnObjects: DataViewObjects | undefined): ValueFormatSettings {
    return {
        fontFamily: getText(columnObjects, "valueFormatting", "fontFamily", DEFAULTS.valueFormat.fontFamily),
        fontSize: getNumber(columnObjects, "valueFormatting", "fontSize", DEFAULTS.valueFormat.fontSize),
        bold: getBool(columnObjects, "valueFormatting", "bold", DEFAULTS.valueFormat.bold),
        italic: getBool(columnObjects, "valueFormatting", "italic", DEFAULTS.valueFormat.italic),
        textColor: getFillNullable(columnObjects, "valueFormatting", "textColor"),
        backgroundColor: getFillNullable(columnObjects, "valueFormatting", "backgroundColor"),
        altTextColor: getFillNullable(columnObjects, "valueFormatting", "altTextColor"),
        altBackgroundColor: getFillNullable(columnObjects, "valueFormatting", "altBackgroundColor"),
        textWrap: getBool(columnObjects, "valueFormatting", "textWrap", DEFAULTS.valueFormat.textWrap),
        numberFormat: getText(columnObjects, "valueFormatting", "numberFormat", DEFAULTS.valueFormat.numberFormat),
        decimals: getNumber(columnObjects, "valueFormatting", "decimals", DEFAULTS.valueFormat.decimals),
        unit: getEnum<DisplayUnit>(columnObjects, "valueFormatting", "unit", DEFAULTS.valueFormat.unit),
        prefix: getText(columnObjects, "valueFormatting", "prefix", DEFAULTS.valueFormat.prefix),
        suffix: getText(columnObjects, "valueFormatting", "suffix", DEFAULTS.valueFormat.suffix)
    };
}

export function parseSpecificColumn(columnObjects: DataViewObjects | undefined): SpecificColumnSettings {
    return {
        applyTo: getEnum<ColumnApplyTo>(
            columnObjects,
            "specificColumn",
            "applyTo",
            DEFAULTS.specificColumn.applyTo
        ),
        textColor: getFillNullable(columnObjects, "specificColumn", "textColor"),
        backgroundColor: getFillNullable(columnObjects, "specificColumn", "backgroundColor"),
        alignment: getEnum<ColumnAlignment>(
            columnObjects,
            "specificColumn",
            "alignment",
            DEFAULTS.specificColumn.alignment
        ),
        unit: getEnum<DisplayUnit>(columnObjects, "specificColumn", "unit", DEFAULTS.specificColumn.unit),
        decimals: getNumber(columnObjects, "specificColumn", "decimals", DEFAULTS.specificColumn.decimals)
    };
}

function safeParseArray<T>(json: string): T[] {
    if (!json) {
        return [];
    }
    try {
        const parsed: unknown = JSON.parse(json);
        return Array.isArray(parsed) ? (parsed as T[]) : [];
    } catch {
        return [];
    }
}

export function parseCF(columnObjects: DataViewObjects | undefined): CFSettings {
    const rulesJson = getText(columnObjects, "cfSettings", "rulesJson", "");
    const iconsJson = getText(columnObjects, "cfSettings", "iconsJson", "");

    return {
        cfType: getEnum<CFType>(columnObjects, "cfSettings", "cfType", DEFAULTS.cf.cfType),
        applyToTotals: getBool(
            columnObjects,
            "cfSettings",
            "applyToTotals",
            DEFAULTS.cf.applyToTotals
        ),
        csLowColor: getFill(columnObjects, "cfSettings", "csLowColor", DEFAULTS.cf.csLowColor),
        csUseMid: getBool(columnObjects, "cfSettings", "csUseMid", DEFAULTS.cf.csUseMid),
        csMidColor: getFill(columnObjects, "cfSettings", "csMidColor", DEFAULTS.cf.csMidColor),
        csHighColor: getFill(columnObjects, "cfSettings", "csHighColor", DEFAULTS.cf.csHighColor),
        csBasis: getEnum<ColorScaleBasis>(
            columnObjects,
            "cfSettings",
            "csBasis",
            DEFAULTS.cf.csBasis
        ),
        rules: safeParseArray<CFRule>(rulesJson).slice(0, 5),
        barPositiveColor: getFill(
            columnObjects,
            "cfSettings",
            "barPositiveColor",
            DEFAULTS.cf.barPositiveColor
        ),
        barNegativeColor: getFill(
            columnObjects,
            "cfSettings",
            "barNegativeColor",
            DEFAULTS.cf.barNegativeColor
        ),
        barAxis: getEnum<BarAxis>(columnObjects, "cfSettings", "barAxis", DEFAULTS.cf.barAxis),
        barMinMode: getEnum<MinMaxMode>(
            columnObjects,
            "cfSettings",
            "barMinMode",
            DEFAULTS.cf.barMinMode
        ),
        barMin: getNumber(columnObjects, "cfSettings", "barMin", DEFAULTS.cf.barMin),
        barMaxMode: getEnum<MinMaxMode>(
            columnObjects,
            "cfSettings",
            "barMaxMode",
            DEFAULTS.cf.barMaxMode
        ),
        barMax: getNumber(columnObjects, "cfSettings", "barMax", DEFAULTS.cf.barMax),
        icons: safeParseArray<CFIconTier>(iconsJson).slice(0, 5),
        // In-visual CF panel fields.
        cfApplyTo: getText(columnObjects, "cfSettings", "cfApplyTo", "background") as CfApplyTo,
        defaultColor: getText(columnObjects, "cfSettings", "defaultColor", DEFAULTS.cf.defaultColor),
        fieldValueSlot: getNumber(columnObjects, "cfSettings", "fieldValueSlot", DEFAULTS.cf.fieldValueSlot),
        fieldValueApplyAs: getText(
            columnObjects,
            "cfSettings",
            "fieldValueApplyAs",
            "background"
        ) as CfApplyTo,
        rulesV2: ((): CfRule[] => {
            const raw = getText(columnObjects, "cfSettings", "rulesV2", "[]");
            try {
                const parsed: unknown = JSON.parse(raw);
                return Array.isArray(parsed) ? (parsed as CfRule[]) : [];
            } catch {
                return [];
            }
        })()
    };
}
