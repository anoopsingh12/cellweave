# CellWeave

> 🧵 **The powerful, zero-bloat JavaScript library for creating feature-complete Excel (.xlsx) spreadsheets**

[![npm version](https://img.shields.io/npm/v/cellweave.svg)](https://www.npmjs.com/package/cellweave)
[![license](https://img.shields.io/npm/l/cellweave.svg)](LICENSE)
[![Tests](https://img.shields.io/badge/tests-69%20passing-brightgreen.svg)](#testing)

CellWeave is a production-ready JavaScript library that generates `.xlsx` files from Node.js **and** the browser. It is built from the ground up following **SOLID principles**, with a fully modular, extensible architecture — no legacy code, no XML template strings buried in logic.

---

## Table of Contents

- [Features](#features)
- [Installation](#installation)
- [Quick Start](#quick-start)
- [API Reference](#api-reference)
  - [CellWeave (Workbook)](#cellweave-workbook)
  - [WorksheetBuilder](#worksheetbuilder)
  - [CellBuilder](#cellbuilder)
- [Feature Showcase](#feature-showcase)
  - [Styling](#styling)
  - [Formulas](#formulas)
  - [Merges & Freeze Panes](#merges--freeze-panes)
  - [Conditional Formatting](#conditional-formatting)
  - [Data Validation](#data-validation)
  - [Excel Tables](#excel-tables)
  - [Charts](#charts)
  - [Images](#images)
  - [Page Setup, Headers & Footers](#page-setup-headers--footers)
  - [Sheet Protection](#sheet-protection)
  - [Groupings & Outlines](#groupings--outlines)
  - [Rich Text](#rich-text)
  - [Defined Names](#defined-names)
  - [Multiple Sheets](#multiple-sheets)
- [SOLID Architecture](#solid-architecture)
- [Browser Usage](#browser-usage)
- [Testing](#testing)
- [Contributing](#contributing)
- [License](#license)

---

## Features

| Category | Features |
|----------|----------|
| **Cells** | Values (string, number, boolean, date), Formulas, Array Formulas, Rich Text, Hyperlinks, Comments |
| **Styles** | Font (name, size, bold, italic, underline, color, strikethrough, super/subscript), Fill (solid, pattern, gradient), Border (all 13 border styles), Alignment (H/V, wrap, rotate, indent, shrink), Number Formats (built-in + custom), Cell locking/hiding |
| **Sheets** | Multiple sheets, Tab colors, Freeze panes, Auto filter, Zoom, Gridlines toggle, Right-to-left, Hidden sheets, Sheet copy/move/rename |
| **Layout** | Merge/unmerge cells, Row height, Column width, Best-fit columns, Hidden rows/columns, Row/column grouping & outlines |
| **Data Features** | Conditional Formatting (color scale, data bar, icon sets, cellIs, expression, top10, etc.), Data Validation (list dropdown, number range, date, custom formula) |
| **Objects** | Excel Tables (ListObject) with styles, headers, totals row, calculated columns; Charts (bar, column, line, pie, doughnut, scatter, area, radar); Images (PNG, JPEG, BMP, GIF, TIFF); Sparklines |
| **Page** | Page setup (A4/Letter/Legal, landscape/portrait, scale, fit-to-page), Margins, Headers/Footers (odd/even/first, codes), Page breaks (row & column), Print area |
| **Workbook** | Defined names, 1900/1904 date system, Core properties (title, author, company, keywords…), Active sheet |
| **Protection** | Sheet protection (password-hashed XOR), Cell lock/hide attributes, Workbook protection options |
| **Output** | `Buffer` (Node.js), `Blob` (browser), `base64` string, file save (`saveAs`), static `fromArray` / `fromObjects` |

---

## Installation

```bash
npm install cellweave
```

**Dependencies:** only [`fflate`](https://github.com/101arrowz/fflate) (fast, pure-JS zip).

---

## Quick Start

```js
const { CellWeave, cell } = require('cellweave');

// Create a workbook
const wb = new CellWeave({ title: 'My Report', author: 'Alice' });

// Add a sheet
const sheet = wb.addSheet('Sales');

// Write headers
sheet.writeHeaders('A1', ['Product', 'Revenue', 'Growth']);

// Write data with fluent cell builder
sheet.writeCell('A2', cell().value('Widget A').bold());
sheet.writeCell('B2', cell().value(125000).formatCurrency('$'));
sheet.writeCell('C2', cell().value(0.12).formatPercent());

// Auto filter, freeze header row
sheet.autoFilter('A1:C1').freezeRows(1);

// Save
wb.saveAs('./sales.xlsx');
```

---

## API Reference

### CellWeave (Workbook)

```js
const wb = new CellWeave(properties?)
```

**Constructor options** (`IWorkbookProperties`):

| Option | Type | Description |
|--------|------|-------------|
| `title` | `string` | Document title |
| `author` | `string` | Author name |
| `company` | `string` | Company name |
| `subject` | `string` | Subject |
| `keywords` | `string` | Search keywords |
| `description` | `string` | Description |
| `created` | `Date` | Creation timestamp |
| `date1904` | `boolean` | Use 1904 date system (Mac legacy) |

**Sheet management:**

```js
const sheet = wb.addSheet('Name', insertAt?)   // → WorksheetBuilder
const sheet = wb.getSheet('Name' | index)       // → WorksheetBuilder
wb.removeSheet('Name' | index)
wb.renameSheet('Old', 'New')
wb.moveSheet('Name', targetIndex)
const copy = wb.copySheet('Name', 'Copy Name')  // → WorksheetBuilder
wb.setActiveSheet('Name' | index)

wb.sheetCount   // number
wb.sheetNames   // string[]
```

**Defined names:**

```js
wb.defineRange('TotalRevenue', "Sheet1!$B$2:$B$100")
wb.defineRange('LocalRange', "Sheet1!$A$1", 'Sheet1') // sheet-scoped
```

**Output:**

```js
wb.toBuffer()     // → Uint8Array (.xlsx bytes)
wb.saveAs('./output.xlsx')
wb.toBase64()     // → string
wb.toBlob()       // → Blob (browser)
```

**Static helpers:**

```js
// Convert 2D array directly
const buf = CellWeave.fromArray(data, { sheetName: 'Data', headers: ['A','B','C'] });

// Convert object array
const buf = CellWeave.fromObjects(records, { columns: ['name','age'] });

// Quick cell builder
const data = CellWeave.cell().bold().value('Hello').build();
```

---

### WorksheetBuilder

Returned by `wb.addSheet()` / `wb.getSheet()`. All methods are chainable.

#### Writing Cells

```js
// Write a single value
sheet.write('A1', 42)
sheet.write('B1', { value: 'Hello', style: { font: { bold: true } } })
sheet.write(rowIndex, colIndex, value)

// Write with a CellBuilder or plain ICellData
sheet.writeCell('A1', cell().bold().value('Title'))
sheet.writeCell('A1', { value: 'Title', style: { font: { bold: true } } })

// Write 2D array
sheet.writeArray('A1', [[1, 2], [3, 4]])
sheet.writeArray(rowIdx, colIdx, data, style?)

// Write header row (auto-styled blue)
sheet.writeHeaders('A1', ['Name', 'Age', 'Score'])
sheet.writeHeaders(rowIdx, colIdx, headers, style?)

// Write formulas across a row
sheet.writeFormulas(rowIdx, startCol, ['SUM(A1:A10)', 'AVERAGE(B1:B10)'])
```

#### Row & Column Configuration

```js
sheet.rowHeight(rowIndex, 30)       // height in points
sheet.colWidth(colIndex, 20)        // width in character units
sheet.hideRow(rowIndex)
sheet.hideCol(colIndex)
sheet.showRow(rowIndex)
sheet.showCol(colIndex)
sheet.colBestFit(colIndex)
sheet.rowOutlineLevel(rowIndex, level)
sheet.colOutlineLevel(colIndex, level)
```

#### Merges

```js
sheet.merge('A1:C3')
sheet.unmerge('A1:C3')
```

#### Freeze Panes

```js
sheet.freezeRows(1)           // freeze top 1 row
sheet.freezeCols(2)           // freeze left 2 columns
sheet.freezeAt(rows, cols)    // freeze both
```

#### Auto Filter

```js
sheet.autoFilter('A1:E1')     // add auto filter
sheet.removeAutoFilter()
```

#### Conditional Formatting

```js
// Color scale (min=red, mid=yellow, max=green by default)
sheet.colorScale('A1:A100', ['#F8696B', '#FFEB84', '#63BE7B'])

// Data bar
sheet.dataBar('B1:B100', '#5B9BD5')

// Highlight cells > value
sheet.highlightGreaterThan('C1:C100', 1000, {
  fill: { patternType: 'solid', fgColor: 'C6EFCE' },
  font: { bold: true, color: '375623' },
})

// Full conditional formatting rule
sheet.conditionalFormat({
  range: { startRow: 0, startCol: 0, endRow: 9, endCol: 0 },
  rules: [{
    type: 'top10', priority: 1,
    style: { fill: { patternType: 'solid', fgColor: 'FFFF00' } },
  }],
})
```

#### Data Validation

```js
// Dropdown list
sheet.dropdownList('B1:B100', ['Option A', 'Option B', 'Option C'])

// Number range
sheet.validateNumberRange('C1:C100', 0, 100)

// Full control
sheet.dataValidation({
  range: { startRow: 0, startCol: 4, endRow: 99, endCol: 4 },
  type: 'custom',
  formula1: 'ISNUMBER(FIND("@",E1))',
  showErrorAlert: true,
  errorStyle: 'stop',
  errorTitle: 'Invalid Email',
  error: 'Must contain @',
})
```

#### Excel Tables

```js
sheet.addTable({
  name: 'SalesTable',
  displayName: 'Sales_Table',
  range: { startRow: 0, startCol: 0, endRow: 10, endCol: 4 },
  tableStyle: 'TableStyleMedium9', // or Light1..21, Dark1..11
  showRowStripes: true,
  headerRow: true,
  totalsRow: true,
  columns: [
    { name: 'Product' },
    { name: 'Revenue', totalsRowFunction: 'sum' },
    { name: 'Quantity', totalsRowFunction: 'sum' },
    { name: 'Margin', formula: '[Revenue]/([Revenue]+[Cost])' },
  ],
})
```

#### Charts

```js
sheet.addChart({
  type: 'column',     // 'bar'|'line'|'pie'|'doughnut'|'scatter'|'area'|'radar'
  title: 'Quarterly Sales',
  series: [
    {
      name: 'Product A',
      valuesRange:     { startRow: 1, startCol: 1, endRow: 10, endCol: 1 },
      categoriesRange: { startRow: 1, startCol: 0, endRow: 10, endCol: 0 },
    },
  ],
  position: { row: 12, col: 0, rowSpan: 15, colSpan: 10 },
  showDataLabels: true,
})
```

#### Images

```js
sheet.addImage({
  data: fs.readFileSync('./logo.png'),
  type: 'png',                     // 'jpeg'|'png'|'gif'|'bmp'
  position: { row: 0, col: 6, rowSpan: 6, colSpan: 4 },
})
// Or base64:
sheet.addImage({ data: base64String, type: 'jpeg', position: { ... } })
```

#### Sparklines

```js
sheet.addSparkline({
  type: 'column',   // 'line'|'column'|'winLoss'
  dataRange:     { startRow: 1, startCol: 1, endRow: 1, endCol: 12 },
  locationRange: { startRow: 1, startCol: 13, endRow: 1, endCol: 13 },
  showHighPoint: true,
  showLowPoint: true,
})
```

#### Row/Column Grouping (Outlines)

```js
sheet.groupRows(1, 5)              // group rows 2–6 (0-based)
sheet.groupRows(1, 5, true)        // collapsed by default
sheet.groupCols(0, 2)              // group columns A–C
```

#### Page Setup

```js
sheet.landscape()           // or .portrait()
sheet.paperA4()             // or .paperLetter()
sheet.fitToPage(1, 0)       // fit width=1 page, unlimited height
sheet.printScale(85)        // 85%
sheet.pageMargins({ top: 0.5, bottom: 0.5, left: 0.5, right: 0.5 })
sheet.pageBreakAfterRow(30) // insert row page break
sheet.pageBreakAfterCol(10) // insert column page break
sheet.headerFooter({
  oddHeader:  '&L&"Calibri,Bold"Company Name&R&[Date]',
  oddFooter:  '&CPage &P of &N',
  differentFirst: true,
  firstHeader: '&C&"Calibri,Bold"Title Page',
})
```

#### Sheet View

```js
sheet.zoom(90)
sheet.hideGridLines()
sheet.showFormulas()
sheet.rightToLeft()
sheet.tabColor('#FF6600')
```

#### Sheet Protection

```js
sheet.protect('mypassword', {
  selectLockedCells:   true,
  selectUnlockedCells: true,
  formatCells:         false,  // disallow
  insertRows:          false,
})
sheet.unprotect()
```

---

### CellBuilder

Create via `cell()` or `new CellBuilder()`. All methods are chainable. Call `.build()` to get the `ICellData` object.

#### Value & Type

```js
cell().value(42)
cell().value(new Date())
cell().value('Text')
cell().value(true)
cell().formula('SUM(A1:A10)')
cell().arrayFormula('MMULT(A1:B2,C1:D2)', { startRow:0,startCol:5,endRow:1,endCol:6 })
cell().type('inlineStr')     // force a specific OOXML cell type
cell().comment('This is a note')
cell().hyperlink('https://example.com')
```

#### Rich Text

```js
cell().richText([
  { text: 'Bold',   font: { bold: true, size: 14 } },
  { text: ' and ',  font: {} },
  { text: 'Red!',   font: { color: 'FF0000', italic: true } },
])
```

#### Font

```js
cell().fontName('Calibri').fontSize(12)
cell().bold().italic().underline('double').strikethrough()
cell().color('#1A237E')         // font color
cell().superscript()
cell().subscript()
```

#### Fill

```js
cell().bgColor('#BBDEFB')       // solid fill
cell().fillPattern('lightGrid', '#000000', '#FFFFFF')
cell().gradient('linear', 90, [
  { position: 0, color: '#7B1FA2' },
  { position: 1, color: '#E91E63' },
])
```

#### Borders

```js
cell().allBorders('thin', '#000000')
cell().borderTop('thick', '#FF0000')
cell().borderBottom('double')
cell().borderLeft('dashed', '#0000FF')
cell().borderRight('medium')
```

#### Alignment

```js
cell().alignH('center')        // 'left'|'center'|'right'|'justify'|'fill'|'distributed'
cell().alignV('middle')        // 'top'|'center'|'bottom'|'justify'|'distributed'
cell().wrapText()
cell().shrinkToFit()
cell().rotate(45)              // -90 to 90 degrees
cell().indent(2)
```

#### Number Formats

```js
cell().formatNumber(2, true)               // #,##0.00
cell().formatCurrency('$', 2)             // $#,##0.00
cell().formatAccounting('€', 2)           // accounting format
cell().formatPercent(1)                   // 0.0%
cell().formatDate()                       // yyyy-mm-dd
cell().formatTime()                       // hh:mm:ss
cell().formatDateTime()                   // yyyy-mm-dd hh:mm:ss
cell().formatScientific(3)                // 0.000E+00
cell().formatFraction()                   // # ?/?
cell().formatText()                       // @
cell().formatCode('#,##0.000;[Red]-#,##0.000')  // raw format code
```

#### Build

```js
const cellData = cell().value('Hello').bold().bgColor('#BBDEFB').build();
// Returns: { value: 'Hello', style: { font: { bold: true }, fill: { ... } } }
```

---

## Feature Showcase

### Styling

```js
sheet.writeCell('A1', cell()
  .value('Dashboard')
  .bold().fontSize(20)
  .color('#FFFFFF').bgColor('#0D47A1')
  .alignH('center').alignV('center')
  .allBorders('medium', '#000000')
);
```

### Formulas

```js
// Simple formula
sheet.write(0, 3, { formula: 'SUM(A1:C1)' });

// Array formula
sheet.write(0, 4, { 
  formula: 'MMULT(A1:B2,C1:D2)',
  isArrayFormula: true,
  arrayFormulaRange: { startRow: 0, startCol: 4, endRow: 1, endCol: 5 },
});
```

### Merges & Freeze Panes

```js
sheet.writeCell('A1', cell().value('Q1 Report').bold().fontSize(18).alignH('center'));
sheet.merge('A1:F1').rowHeight(0, 40);
sheet.freezeAt(2, 1); // freeze 2 rows + 1 column
```

### Conditional Formatting

```js
// Traffic light icon set
sheet.conditionalFormat({
  range: { startRow: 1, startCol: 2, endRow: 100, endCol: 2 },
  rules: [{
    type: 'iconSet', priority: 1,
    iconSet: {
      iconSet: '3TrafficLights1',
      cfvo: [{ type: 'num', val: 0 }, { type: 'num', val: 33 }, { type: 'num', val: 67 }],
    },
  }],
});
```

### Multiple Sheets

```js
const wb = new CellWeave({ title: 'Annual Report' });
const summary  = wb.addSheet('Summary').tabColor('#2196F3');
const q1       = wb.addSheet('Q1').tabColor('#4CAF50');
const q2       = wb.addSheet('Q2').tabColor('#FF9800');

// Cross-sheet formula
summary.write(1, 1, { formula: "Q1!B2+Q2!B2" });

wb.defineRange('AllRevenue', "Q1!$B$2:$B$100");
wb.saveAs('./annual.xlsx');
```

### Quick Helpers

```js
const { fromArray, fromObjects } = require('cellweave');

// From 2D array
const buf = fromArray(
  [['Alice', 28], ['Bob', 35]],
  { headers: ['Name', 'Age'], sheetName: 'Staff' }
);

// From object array
const buf = fromObjects([
  { product: 'A', revenue: 50000 },
  { product: 'B', revenue: 30000 },
]);
```

---

## SOLID Architecture

CellWeave is built from the ground up adhering to SOLID principles:

| Principle | Implementation |
|-----------|---------------|
| **S** – Single Responsibility | Each class has one job: `Cell` stores data, `StyleRegistry` deduplicates styles, `XlsxWriter` assembles the zip, `WorksheetXmlWriter` generates XML |
| **O** – Open/Closed | `XlsxWriter` delegates to separate writer modules. New chart types or features are added by extending writers, not modifying existing ones |
| **L** – Liskov Substitution | `WorksheetBuilder` and `CellBuilder` are designed so any part is substitutable — builders work on the same model interfaces |
| **I** – Interface Segregation | Over 35 granular JSDoc `@typedef` interfaces (see `src/interfaces/index.js`) — consumers only depend on the types they need |
| **D** – Dependency Inversion | `CellWeave` (the facade) depends on abstract `Workbook` model, not on `XlsxWriter` directly. Writers depend on the model, not on each other |

**Module tree:**

```
src/
├── interfaces/         ← ISP: 35+ type contracts
├── utils/
│   ├── addressUtils.js ← SRP: A1/R1C1, dates, colors
│   └── xmlBuilder.js   ← SRP: XML generation only
├── core/
│   ├── StyleRegistry.js      ← SRP: style deduplication
│   └── SharedStringTable.js  ← SRP: string deduplication
├── models/
│   ├── Cell.js         ← SRP: cell data model
│   ├── Worksheet.js    ← SRP: sheet data model
│   └── Workbook.js     ← SRP: workbook container
├── builders/
│   ├── CellBuilder.js       ← ISP + fluent API
│   └── WorksheetBuilder.js  ← OCP: wraps model, adds features
├── writers/
│   ├── XlsxWriter.js        ← DIP: orchestrates writers
│   ├── WorksheetXmlWriter.js
│   ├── WorkbookXmlWriter.js
│   ├── StylesXmlWriter.js
│   ├── SharedStringsXmlWriter.js
│   ├── ChartXmlWriter.js
│   └── TableXmlWriter.js
└── CellWeave.js        ← Facade
```

---

## Browser Usage

CellWeave works in modern browsers via bundlers (Webpack, Vite, Rollup):

```html
<script type="module">
import { CellWeave, cell } from 'cellweave';

const wb = new CellWeave({ title: 'Browser Report' });
wb.addSheet('Data').write('A1', 'Hello from browser!');

// Download
const blob = wb.toBlob();
const url = URL.createObjectURL(blob);
const a = document.createElement('a');
a.href = url; a.download = 'report.xlsx'; a.click();
</script>
```

---

## Testing

```bash
npm test              # run all tests with coverage
npm run test:watch    # watch mode
npm run example       # generate all example xlsx files to dist/
```

**Test coverage:** 69 tests across all modules.

---

## Contributing

1. Fork the repo
2. Create a feature branch: `git checkout -b feature/pivot-tables`
3. Follow SOLID principles — add new features via new modules, not by modifying existing ones
4. Write tests for any new module in `tests/`
5. Run `npm test` — all tests must pass
6. Open a PR

---

## License

MIT © CellWeave Contributors
