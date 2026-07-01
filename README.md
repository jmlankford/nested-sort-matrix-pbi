# Nested Sort Matrix — Power BI Custom Visual

A production-grade Power BI custom visual that extends the stock matrix with true nested/hierarchical sorting, a runtime configuration panel, and a comprehensive format pane. Built for Power BI Report Server (September 2025) and Power BI Desktop.

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
- **Rows** — reorder row fields via drag-and-drop, toggle visibility, session rename
- **Values** — reorder value measures, toggle visibility, session rename
- **Column Fields** — reorder pivot fields, toggle visibility, session rename (hidden when no column fields are bound)

Configuration persists via `persistProperties` across report saves and reloads. Session renames reset on data refresh.

### Sorting
- Click a column header once: ascending. Again: descending. Third time: clear.
- Sort is scoped within parent groups at every level
- Sorting a parent level resets child level sorts (matches Excel PivotTable behavior)
- Sort stack shown in status bar: `Sorted: Region ↑ → Revenue ↓`

### Subtotals and Totals
- Row subtotals per hierarchy level, toggleable globally and per named field
- Column subtotals (when pivot mode active)
- Grand total row and column
- Configurable label text
- Conditional formatting optionally applied to totals

### Conditional Formatting
Configured via right-click on any value column header. Four types:
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

Field well order sets the baseline hierarchy. The configuration panel overrides display order without changing the field well.

---

## Format Pane Sections

| Section | Contents |
|---|---|
| Grid | Row height mode, fixed height px, cell font size, value cell font |
| Row Headers | Bold, font size, indent per level, font family |
| Layout | Layout mode (Compact/Outline/Tabular), repeat row headers |
| Column Headers | Bold, font size, show sort arrows, font family |
| Expand/Collapse Buttons | Show/hide, button size, button color, style |
| Subtotals & Totals | Row/column subtotals, grand total row/column, label text, apply CF to totals, per-field level toggles |
| Alternate Row Color | On/off, color |
| Status Bar | Show/hide |
| Cross-Filter | Enable/disable |
| Values | Per-measure: typography, colors (with fx button), display settings |
| Specific Column | Per-measure: apply-to scope, color, alignment, unit, decimals |

---

## Known Limitations

### Non-Additive Measure Subtotals
The visual uses Table DataView and computes subtotals by summing child row values. This produces incorrect results for non-additive measures (DISTINCTCOUNT, ratio measures, running totals).

**Workaround:** Create a pre-computed DAX measure for the subtotal value and bind it to a separate value slot. Example:
```dax
Customer Count Subtotal = CALCULATE(DISTINCTCOUNT(Orders[CustomerID]))
```
Bind this alongside your leaf-level measure and use Specific Column to hide it at the Values level (`Apply to: Header` only), or use conditional formatting to suppress display at leaf level.

### CF Rules — Compound Conditions
To filter a value range (e.g. between 1000 and 9999), use the `+ AND` button within a single rule row. Do not create two separate rules — each rule is evaluated independently and "first match wins" means the second rule will catch values that failed the first.

### Field Well Order vs. Config Panel Order
Changing field order in the visualization pane field well resets the config panel to the new baseline. The config panel cannot write back to the field well — this is a platform constraint of the Power BI custom visual API. Use the field well for default order, the config panel for display overrides.

### Session Renames
Field renames in the configuration panel are session-only and reset on data refresh. They are cosmetic display overrides, not persistent metadata changes.

### `getFormattingModel()` API
This visual uses `getFormattingModel()` for format pane rendering, which is the current recommended API for Power BI Report Server September 2025. The legacy `enumerateObjectInstances` audit warning that appeared in earlier builds is resolved.

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
│   ├── dataTransformer.ts     # Table DataView → RowTreeNode tree, pivot, subtotals
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
