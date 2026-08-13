# Nested Sort Matrix — Power BI Custom Visual

A production-grade Power BI custom visual that extends the stock matrix with true nested/hierarchical sorting, a runtime configuration panel, and a comprehensive format pane. Built for Power BI Report Server (September 2025) and Power BI Desktop.

The visual is backed by the **Matrix DataView**: the engine evaluates every measure at the true scope of each hierarchy node (leaf, group, subtotal, and grand total), exactly as the stock matrix does. There is no client-side aggregation, so level-aware and non-additive measures roll up correctly. Native host expand/collapse drives which levels are materialized.

---

## Features

### Core
- **Nested/hierarchical sorting** — sort is scoped within parent groups, not a global flat sort. Clicking Revenue sorts salesmen within each region independently.
- **Three row layout modes** — Compact (single column, indented), Outline (one column per level), Tabular (all levels as flat columns in the header)
- **Expand/collapse row groups** — +/− controls with configurable style (Plus/Minus, Chevron, Triangle), size, and color
- **Row virtualization** — only renders rows in the current viewport + 20-row buffer; handles thousands of leaf rows without performance degradation
- **Pivot/column fields mode** — optional crosstab with spanning column group headers

### Configuration Panel
Opened via ⚙ Setup in the status bar. Three tabs:
- **Rows** — toggle level visibility, session rename. Hierarchy **order** comes from the field well (the Matrix DataView bakes hierarchy order into the query, so runtime reordering is not possible — see [Runtime hierarchy reordering](#runtime-hierarchy-reordering)). Hiding a level flattens it out of the display: its children move up under the level above, with their engine-computed values intact.
- **Values** — reorder value measures via drag-and-drop, toggle visibility, session rename (these are display-layer and survive the Matrix migration).
- **Column Fields** — toggle visibility (order fixed by the field well; renaming a pivot field has no visible effect since pivot headers show data values).

Configuration persists via `persistProperties` across report saves and reloads. Session renames reset on data refresh. Toggling level visibility is a display-only change — it does not alter the query, so it never resets sort, selection, or expansion state.

### Sorting
- Click a column header once: ascending. Again: descending. Third time: clear.
- There is a **single active sort** at a time; clicking a different header (a value column or a row-field column) replaces it — matching the stock matrix.
- The active sort is applied **within each parent's scope at every hierarchy level**, not as a global flat sort.
- The active sort is shown in the status bar, e.g. `Sorted: Revenue ↓`.
- It survives cross-filter updates, re-renders, and field add/remove.

### Subtotals and Totals
- Row subtotals per hierarchy level, toggleable globally and per named field
- Column subtotals (when pivot mode active)
- Grand total row and column
- Configurable label text
- Conditional formatting optionally applied to totals

### Conditional Formatting
Configured by **right-clicking a value column header**, which opens the in-visual CF panel. Four types:
- **Color Scale** — min/mid/max gradient with hex color inputs and percent-of-range basis option
- **Rules** — multiple rules with AND compound conditions. Operators: >=, >, <=, <, =, !=, between, is blank, is not blank. First match wins.
- **Data Bars** — horizontal bars in cell background, configurable positive/negative colors
- **Field Value** — a bound measure returns a hex string (e.g. `#FF6B6B`) applied as background and/or font color

### Value Formatting (per measure)
- Font family, size, bold, italic
- Font color and background color (with native Power BI `fx` CF dialog support)
- Alternate row font and background colors
- Text wrap
- Format string (Power BI format strings, e.g. `#,##0.00`)
- Decimal places, display unit (thousands/millions/billions), prefix, suffix

### Specific Column (per measure)
Final override layer applied after value formatting and conditional formatting:
- Apply to: All / Values / Header / Subtotals / Grand Total
- Font color, background color, alignment, display unit, decimal places

### Column and Row Resizing
- Drag the right edge of any column header to resize
- Row field columns independently resizable
- Pivot group headers resize proportionally across their spanned leaf columns
- All widths persisted via `persistProperties`

### Cross-Filtering
- Click a leaf row to cross-filter other visuals on the page
- Ctrl+Click: multi-select
- Shift+Click: range select
- Click empty space to clear selection
- Toggleable via format pane

---

## Field Wells

| Bucket | Role | Notes |
|---|---|---|
| Row Fields | Grouping | Stack multiple fields; order sets hierarchy |
| Column Fields | Grouping | Optional; activates pivot/crosstab mode |
| Values | Measure | Up to ~15 measures recommended |

Field well order sets the row/column hierarchy. Under the Matrix DataView this order is fixed by the query, so the configuration panel offers row-level **visibility** and **rename** but not reordering. Only the **Values** tab reorders. For runtime hierarchy reordering, use the field-parameter pattern described in [Runtime hierarchy reordering](#runtime-hierarchy-reordering).

---

## Format Pane Sections

| Section | Contents |
|---|---|
| Grid | Row height mode, fixed height px, cell font size, value cell font |
| Row Headers | Bold, font size, wrap header text, indent per level, font family |
| Layout | Layout mode (Compact/Outline/Tabular), repeat row headers |
| Column Headers | Bold, font size, show sort arrows, wrap header text, font family |
| Expand/Collapse Buttons | Show/hide, button size, button color, style |
| Subtotals & Totals | Row/column subtotals, grand total row/column, label text, apply CF to totals, per-field level toggles |
| Alternate Row Color | On/off, color |
| Status Bar | Show/hide |
| Cross-Filter | Enable/disable |
| Values | Per-measure: typography, colors (with fx button), display settings |
| Specific Column | Per-measure: apply-to scope, color, alignment, unit, decimals |

---

## Subtotals and Parent Values

Subtotals and parent (group) row values are **engine-computed at every hierarchy level**, straight from the Matrix DataView. The visual never sums child rows. Level-aware and non-additive measures therefore work exactly as they do in the stock matrix — including `ISINSCOPE` dispatchers, `DISTINCTCOUNT`, ratios, averages, and running totals. A group row shows the measure evaluated at that group's scope, not a rollup of its children.

The query request for subtotal nodes is **decoupled** from the render toggle: the engine is always asked for subtotal nodes (so parent rows always have their values), while the **Subtotals & Totals → Row subtotals** toggle only controls whether the subtotal *rows* are drawn. Turning subtotal rows off never blanks parent values.

---

## Measure Design for This Visual

Because the engine evaluates each measure at every hierarchy node, a measure written for a **parent** grain (e.g. a customer-level balance) will also be evaluated at a **finer** grain (e.g. per invoice). If it returns a non-blank value there, it produces "ghost" rows at the leaf level and forces the engine to compute rows you never intended — hurting both correctness and query performance.

**Guard parent-grain measures with `ISINSCOPE`** against the leaf-level column so they return `BLANK()` below their intended grain:

```dax
Customer Balance =
IF (
    ISINSCOPE ( Invoices[InvoiceNumber] ),
    BLANK (),                                   -- below the customer grain: no value
    CALCULATE ( SUM ( Invoices[Balance] ) )     -- at customer grain and above
)
```

A measure that is naturally additive and correct at every grain needs no guard. The `ISINSCOPE` guard is specifically for measures whose meaning is tied to a particular parent level.

---

## Text-Valued Measures

Measures that return **text** render correctly at every level — for example `FIRSTNONBLANK`/`LASTNONBLANK` aggregations or `SELECTEDVALUE` over a text column:

```dax
Status = SELECTEDVALUE ( Invoices[Status], "(multiple)" )
```

Text values are carried through the transform untouched (no numeric coercion); sorting on a text column falls back to a locale-aware string comparison.

---

## Runtime Hierarchy Reordering

The Matrix DataView bakes hierarchy order into the query, so the visual cannot reorder row levels at runtime. To let report consumers reorder the hierarchy live, use the **field parameter** pattern (verified working on Power BI Report Server, September 2025):

1. Create one **field parameter** per hierarchy slot (e.g. three parameters for a three-level hierarchy), each containing the **full** list of candidate fields.
2. Bind the field parameters to **Row Fields** in slot order (parameter 1, then 2, then 3).
3. Add each field parameter as a **single-select** slicer.

Each slicer then chooses which field occupies that hierarchy slot, giving arbitrary three-level reordering at report runtime — outside the visual, at the report layer where the query is built. Three-level arbitrary reordering has been verified on PBIRS September 2025.

---

## Data Reduction

The visual declares **no explicit `dataReductionAlgorithm`** on the matrix rows — the host manages segmentation with its default, hierarchy-aware reduction (this matches Microsoft's reference matrix and is what keeps nesting correct). For very large hierarchies the host may segment the data; `fetchMoreData` is **not currently implemented**, so extremely large expansions may be capped by the host's default window. If you hit a cap, reduce cardinality upstream (filters, aggregations) or collapse levels.

---

## Other Notes

### CF Rules — Compound Conditions
To filter a value range (e.g. between 1000 and 9999), use the `+ AND` button within a single rule row. Do not create two separate rules — each rule is evaluated independently and "first match wins" means the second rule will catch values that failed the first.

### Cell-Level Copy
Custom visuals cannot offer cell-level **Copy value** / **Copy selection** — those context-menu commands are exclusive to first-party visuals (a platform limitation). Right-clicking a row opens the standard custom-visual context menu (Include/Exclude/Show as a table); right-clicking a **value column header** opens the conditional-formatting panel.

### Session Renames
Field renames in the configuration panel are session-only and reset on data refresh. They are cosmetic display overrides, not persistent metadata changes.

### `getFormattingModel()` API
This visual uses `getFormattingModel()` for format pane rendering, which is the current recommended API for Power BI Report Server September 2025.

---

## Build Instructions

### Prerequisites
- Node.js 16+
- Power BI Visuals Tools: `npm install -g powerbi-visuals-tools`

### Install dependencies
```bash
npm install
```

### Type-check
```bash
npx tsc --noEmit
```

### Package
```bash
npx pbiviz package
```

Output: `dist/nestedSortMatrix.*.pbiviz`

### Lint
```bash
npm run lint
```

---

## Import to Power BI Report Server

1. Open the PBIRS web portal
2. Navigate to a report in Power BI Desktop connected to PBIRS
3. In the Visualizations pane, click **...** → **Import a visual from a file**
4. Select `dist/nestedSortMatrix.*.pbiviz`
5. Click **Add** when prompted

To update an existing import: remove the visual from the canvas, delete the existing entry from the org visuals list, and reimport the new `.pbiviz`. Version bumps in `pbiviz.json` ensure PBIRS treats the update as a new visual and does not serve cached capabilities.

---

## Import to Power BI Desktop

1. In the Visualizations pane, click **...** → **Import a visual from a file**
2. Select `dist/nestedSortMatrix.*.pbiviz`
3. Click **Add** when prompted

For full conditional formatting dialog support (`fx` button on color pickers), Power BI Desktop with `getFormattingModel()` support is required. September 2025 PBIRS meets this requirement.

---

## Project Structure
/
├── src/
│   ├── visual.ts              # IVisual lifecycle, format pane, persistProperties
│   ├── dataTransformer.ts     # Matrix DataView → RowTreeNode tree, pivot, level flatten
│   ├── renderer.ts            # D3 rendering, headers, rows, cells, resize handles
│   ├── virtualScroller.ts     # Viewport row virtualization, DOM node recycling
│   ├── configPanel.ts         # Setup panel overlay, HTML5 drag-and-drop
│   ├── cfPanel.ts             # Conditional formatting panel, rules builder
│   ├── statusBar.ts           # Status bar rendering and interactions
│   ├── sortManager.ts         # Nested sort state, scoped sort application
│   ├── conditionalFormatter.ts # CF evaluation: color scale, rules, data bars, icons
│   ├── selectionManager.ts    # ISelectionManager wrapper, multi/range select
│   ├── settings.ts            # Format pane settings classes and DataView parsers
│   └── styles/
│       └── visual.less        # All visual styles
├── capabilities.json          # Data roles, dataViewMappings, format pane objects
├── pbiviz.json                # Visual metadata and API version
├── package.json
├── tsconfig.json
├── .eslintrc.json
└── README.md

---

## Development Notes

### Adding a new format pane property
1. Add the property to `capabilities.json` under the relevant object
2. Add the TypeScript field and default to `settings.ts`
3. Add the getter in `parseVisualSettings()` in `settings.ts`
4. Emit the property in `getFormattingModel()` in `visual.ts`
5. Read and apply in `renderer.ts`

### Updating and pushing to GitHub
After any build, commit and push:
```bash
git add -A
git commit -m "describe your change"
git push
```

Or use the helper script:
```bash
./push.sh "describe your change"
```

---

## Version History

| Version | Notes |
|---|---|
| 1.0.0.0 | Initial build |
| 1.0.1.0 | Data binding fix (stacked field wells), layout modes, CF panel, getFormattingModel() migration |
| 1.0.2.0 | CF panel fixes (visibility, click-outside handling, defaultColor, AND compound rules), tabular group row suppression |
| 3.0.0.0 | **Matrix DataView migration.** Replaced Table DataView + client-side grouping with the Matrix DataView: engine-computed values at every hierarchy level (level-aware and non-additive measures now correct), native host expand/collapse with per-level and Expand/Collapse-All controls, scroll anchoring, single-active-sort model, query subtotals decoupled from the subtotal-row render toggle, row-level hiding as a display flatten, header word wrap, raw scroll preservation on sort, and text-measure support. Runtime hierarchy reordering moves to the report-layer field-parameter pattern. |
