/**
 * CellWeave – Comprehensive Usage Example
 * Demonstrates all major features of the library.
 */

'use strict';

const path = require('path');
const {
  CellWeave,
  CellBuilder,
  cell,
  fromArray,
  fromObjects,
} = require('../index');

const OUTPUT_DIR = path.join(__dirname, '../dist');
require('fs').mkdirSync(OUTPUT_DIR, { recursive: true });

// ═══════════════════════════════════════════════════════════════════
// 1. BASIC USAGE
// ═══════════════════════════════════════════════════════════════════

console.log('📊 Building Basic Example…');
{
  const wb = new CellWeave({ title: 'Basic Example', author: 'CellWeave Demo' });
  const sheet = wb.addSheet('Basics');

  // Simple values
  sheet.write('A1', 'Hello, CellWeave!');
  sheet.write('A2', 42);
  sheet.write('A3', true);
  sheet.write('A4', new Date());
  sheet.write(4, 0, { formula: 'SUM(A2:A2)', value: 42 }); // A5

  wb.saveAs(path.join(OUTPUT_DIR, '01_basic.xlsx'));
  console.log('  ✓ dist/01_basic.xlsx');
}

// ═══════════════════════════════════════════════════════════════════
// 2. RICH STYLING
// ═══════════════════════════════════════════════════════════════════

console.log('🎨 Building Styling Example…');
{
  const wb = new CellWeave({ title: 'Styling Demo' });
  const sheet = wb.addSheet('Styles');

  // Font styles
  sheet.writeCell('A1', cell().value('Bold Text').bold().fontSize(16).color('#1A237E'));
  sheet.writeCell('A2', cell().value('Italic').italic().fontName('Georgia').fontSize(13));
  sheet.writeCell('A3', cell().value('Underlined').underline('double').color('#B71C1C'));
  sheet.writeCell('A4', cell().value('Strikethrough').strikethrough());
  sheet.writeCell('A5', cell().value('Superscript: x²').superscript());

  // Background colors
  sheet.writeCell('B1', cell().value('Red BG').bgColor('#FFCDD2').bold());
  sheet.writeCell('B2', cell().value('Blue BG').bgColor('#BBDEFB').italic());
  sheet.writeCell('B3', cell().value('Green BG').bgColor('#C8E6C9'));

  // Gradient fill
  sheet.writeCell('C1', cell().value('Gradient').gradient('linear', 90, [
    { position: 0, color: '#7B1FA2' },
    { position: 1, color: '#E91E63' },
  ]).color('#FFFFFF').bold());

  // Borders
  sheet.writeCell('D1', cell().value('All Borders').allBorders('medium', '#000000'));
  sheet.writeCell('D2', cell().value('Double Border').borderBottom('double', '#FF0000'));

  // Alignment
  sheet.writeCell('E1', cell().value('Centered').alignH('center').alignV('center'));
  sheet.writeCell('E2', cell().value('Wrapped long text that goes beyond the cell').wrapText());
  sheet.writeCell('E3', cell().value('Rotated').rotate(45).alignH('center'));
  sheet.colWidth(4, 25).rowHeight(1, 30);

  // Number formats
  sheet.writeCell('F1', cell().value(1234567.89).formatNumber(2, true));
  sheet.writeCell('F2', cell().value(0.1523).formatPercent(1));
  sheet.writeCell('F3', cell().value(new Date()).formatDate());
  sheet.writeCell('F4', cell().value(99999).formatCurrency('$', 2));
  sheet.writeCell('F5', cell().value(99999).formatAccounting('€', 2));
  sheet.writeCell('F6', cell().value(1234).formatScientific(3));
  sheet.writeCell('F7', cell().value(0.625).formatFraction());

  // Rich text
  sheet.writeCell('G1', cell().richText([
    { text: 'Bold ', font: { bold: true, size: 14 } },
    { text: 'and ', font: {} },
    { text: 'Colored!', font: { color: 'E91E63', italic: true } },
  ]));

  sheet.colWidth(0, 18).colWidth(1, 15).colWidth(2, 15).colWidth(3, 15).colWidth(5, 18).colWidth(6, 25);

  wb.saveAs(path.join(OUTPUT_DIR, '02_styling.xlsx'));
  console.log('  ✓ dist/02_styling.xlsx');
}

// ═══════════════════════════════════════════════════════════════════
// 3. MERGES, FREEZE PANES, SHEET PROTECTION
// ═══════════════════════════════════════════════════════════════════

console.log('🔒 Building Merges & Freeze Example…');
{
  const wb = new CellWeave({ title: 'Merges & Freeze' });
  const sheet = wb.addSheet('Sheet1');

  // Merged header
  sheet.writeCell('A1', cell().value('Annual Sales Report 2024').bold().fontSize(18).alignH('center').bgColor('#1565C0').color('#FFFFFF'));
  sheet.merge('A1:H1').rowHeight(0, 40);

  // Sub-headers
  ['Q1', 'Q2', 'Q3', 'Q4'].forEach((q, i) => {
    sheet.writeCell(`${String.fromCharCode(66 + i * 2)}2`,
      cell().value(q).bold().alignH('center').bgColor('#1E88E5').color('#FFFFFF'));
    sheet.merge(`${String.fromCharCode(66 + i * 2)}2:${String.fromCharCode(67 + i * 2)}2`);
  });

  // Data
  const products = ['Widget A', 'Widget B', 'Widget C', 'Gadget X'];
  products.forEach((p, i) => {
    sheet.write(2 + i, 0, { value: p, style: { font: { bold: true } } });
    for (let q = 0; q < 4; q++) {
      const val = Math.round(1000 + Math.random() * 5000);
      sheet.write(2 + i, 1 + q, { value: val, style: { numberFormat: { type: 'currency', currencySymbol: '$', decimalPlaces: 0 } } });
    }
  });

  sheet.freezeAt(2, 1); // Freeze header rows and first column
  sheet.colWidth(0, 15);
  for (let i = 1; i <= 4; i++) sheet.colWidth(i, 14);

  wb.saveAs(path.join(OUTPUT_DIR, '03_merges_freeze.xlsx'));
  console.log('  ✓ dist/03_merges_freeze.xlsx');
}

// ═══════════════════════════════════════════════════════════════════
// 4. FORMULAS
// ═══════════════════════════════════════════════════════════════════

console.log('📐 Building Formulas Example…');
{
  const wb = new CellWeave({ title: 'Formulas Demo' });
  const sheet = wb.addSheet('Formulas');

  sheet.writeHeaders('A1', ['Month', 'Revenue', 'Expenses', 'Profit', 'Margin%']);

  const months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun'];
  months.forEach((month, i) => {
    const row = i + 2;
    const rev = 10000 + Math.round(Math.random() * 5000);
    const exp = 6000 + Math.round(Math.random() * 2000);
    sheet.write(row - 1, 0, { value: month });
    sheet.write(row - 1, 1, { value: rev, style: { numberFormat: { type: 'currency', currencySymbol: '$', decimalPlaces: 0 } } });
    sheet.write(row - 1, 2, { value: exp, style: { numberFormat: { type: 'currency', currencySymbol: '$', decimalPlaces: 0 } } });
    sheet.write(row - 1, 3, { formula: `B${row}-C${row}`, style: { numberFormat: { type: 'currency', currencySymbol: '$', decimalPlaces: 0 } } });
    sheet.write(row - 1, 4, { formula: `D${row}/B${row}`, style: { numberFormat: { type: 'percentage', decimalPlaces: 1 } } });
  });

  // Summary row
  const totalRow = months.length + 1;
  sheet.write(totalRow, 0, { value: 'TOTAL', style: { font: { bold: true } } });
  sheet.write(totalRow, 1, { formula: `SUM(B2:B${totalRow})`, style: { font: { bold: true }, numberFormat: { type: 'currency', currencySymbol: '$', decimalPlaces: 0 } } });
  sheet.write(totalRow, 2, { formula: `SUM(C2:C${totalRow})`, style: { font: { bold: true }, numberFormat: { type: 'currency', currencySymbol: '$', decimalPlaces: 0 } } });
  sheet.write(totalRow, 3, { formula: `SUM(D2:D${totalRow})`, style: { font: { bold: true }, numberFormat: { type: 'currency', currencySymbol: '$', decimalPlaces: 0 } } });
  sheet.write(totalRow, 4, { formula: `D${totalRow + 1}/B${totalRow + 1}`, style: { font: { bold: true }, numberFormat: { type: 'percentage', decimalPlaces: 1 } } });

  sheet.colWidth(0, 10).colWidth(1, 14).colWidth(2, 14).colWidth(3, 14).colWidth(4, 12);
  sheet.autoFilter('A1:E1');
  sheet.freezeRows(1);

  wb.saveAs(path.join(OUTPUT_DIR, '04_formulas.xlsx'));
  console.log('  ✓ dist/04_formulas.xlsx');
}

// ═══════════════════════════════════════════════════════════════════
// 5. CONDITIONAL FORMATTING
// ═══════════════════════════════════════════════════════════════════

console.log('🌈 Building Conditional Formatting Example…');
{
  const wb = new CellWeave({ title: 'Conditional Formatting' });
  const sheet = wb.addSheet('CF Demo');

  sheet.writeHeaders('A1', ['Score', 'Status', 'Revenue', 'Growth']);
  const data = [88, 45, 92, 67, 78, 55, 99, 33, 71, 84];
  data.forEach((score, i) => {
    sheet.write(i + 1, 0, { value: score });
    sheet.write(i + 1, 1, { value: score >= 70 ? 'PASS' : 'FAIL' });
    sheet.write(i + 1, 2, { value: score * 100 });
    sheet.write(i + 1, 3, { value: (score - 50) / 100 });
  });

  // Color scale
  sheet.colorScale('A2:A11', ['#F8696B', '#FFEB84', '#63BE7B']);

  // Highlight fails
  sheet.conditionalFormat({
    range: { startRow: 1, startCol: 1, endRow: 10, endCol: 1 },
    rules: [{
      type: 'cellIs', operator: 'equal', formula: ['"FAIL"'], priority: 1,
      style: { font: { bold: true, color: 'FFFFFF' }, fill: { patternType: 'solid', fgColor: 'C00000' } },
    }],
  });

  // Data bars
  sheet.dataBar('C2:C11', '#5B9BD5');

  // Percentage format
  for (let i = 1; i <= 10; i++) {
    sheet.write(i, 3, { value: data[i - 1] / 100 - 0.5, style: { numberFormat: { type: 'percentage', decimalPlaces: 0 } } });
  }
  sheet.highlightGreaterThan('D2:D11', 0.3, {
    font: { bold: true, color: '375623' },
    fill: { patternType: 'solid', fgColor: 'C6EFCE' },
  });

  sheet.colWidth(0, 10).colWidth(1, 10).colWidth(2, 14).colWidth(3, 10);

  wb.saveAs(path.join(OUTPUT_DIR, '05_conditional_formatting.xlsx'));
  console.log('  ✓ dist/05_conditional_formatting.xlsx');
}

// ═══════════════════════════════════════════════════════════════════
// 6. DATA VALIDATION
// ═══════════════════════════════════════════════════════════════════

console.log('✅ Building Data Validation Example…');
{
  const wb = new CellWeave({ title: 'Data Validation' });
  const sheet = wb.addSheet('Validation');

  sheet.writeHeaders('A1', ['Name', 'Department', 'Score (1-100)', 'Date Joined', 'Email', 'Status']);

  // Dropdown validation for Department
  sheet.dropdownList('B2:B100', ['Engineering', 'Marketing', 'Sales', 'HR', 'Finance', 'Operations']);

  // Number range validation for Score
  sheet.validateNumberRange('C2:C100', 1, 100);

  // Date validation
  sheet.dataValidation({
    range: { startRow: 1, startCol: 3, endRow: 99, endCol: 3 },
    type: 'date',
    operator: 'between',
    formula1: '2020-01-01',
    formula2: '2030-12-31',
    showErrorAlert: true,
    errorStyle: 'warning',
    errorTitle: 'Date Range',
    error: 'Please enter a date between 2020 and 2030',
    showInputMessage: true,
    promptTitle: 'Date',
    prompt: 'Enter the joining date',
    allowBlank: true,
  });

  // Custom formula validation (email-like)
  sheet.dataValidation({
    range: { startRow: 1, startCol: 4, endRow: 99, endCol: 4 },
    type: 'custom',
    formula1: 'ISNUMBER(FIND("@",E2))',
    showErrorAlert: true,
    errorStyle: 'stop',
    errorTitle: 'Invalid Email',
    error: 'Please enter a valid email address containing @',
    allowBlank: true,
  });

  // Status dropdown
  sheet.dropdownList('F2:F100', ['Active', 'Inactive', 'On Leave', 'Terminated']);

  sheet.colWidth(0, 15).colWidth(1, 15).colWidth(2, 16).colWidth(3, 14).colWidth(4, 22).colWidth(5, 12);
  sheet.autoFilter('A1:F1').freezeRows(1);

  wb.saveAs(path.join(OUTPUT_DIR, '06_data_validation.xlsx'));
  console.log('  ✓ dist/06_data_validation.xlsx');
}

// ═══════════════════════════════════════════════════════════════════
// 7. TABLES
// ═══════════════════════════════════════════════════════════════════

console.log('📋 Building Tables Example…');
{
  const wb = new CellWeave({ title: 'Excel Tables' });
  const sheet = wb.addSheet('Tables');

  const headers = ['Product ID', 'Product Name', 'Category', 'Unit Price', 'Stock', 'Total Value'];
  sheet.writeHeaders('A1', headers);

  const products = [
    { id: 'P001', name: 'Laptop Pro', cat: 'Electronics', price: 1299, stock: 45 },
    { id: 'P002', name: 'Desk Chair', cat: 'Furniture',   price: 349,  stock: 120 },
    { id: 'P003', name: 'Monitor 4K', cat: 'Electronics', price: 699,  stock: 80 },
    { id: 'P004', name: 'Keyboard',   cat: 'Electronics', price: 149,  stock: 200 },
    { id: 'P005', name: 'Desk Lamp',  cat: 'Furniture',   price: 79,   stock: 300 },
  ];

  products.forEach((p, i) => {
    const r = i + 1;
    sheet.write(r, 0, { value: p.id });
    sheet.write(r, 1, { value: p.name });
    sheet.write(r, 2, { value: p.cat });
    sheet.write(r, 3, { value: p.price, style: { numberFormat: { type: 'currency', currencySymbol: '$', decimalPlaces: 2 } } });
    sheet.write(r, 4, { value: p.stock });
    sheet.write(r, 5, { formula: `D${r + 1}*E${r + 1}`, style: { numberFormat: { type: 'currency', currencySymbol: '$', decimalPlaces: 2 } } });
  });

  // Add as Excel Table
  sheet.addTable({
    name: 'ProductTable',
    displayName: 'Product_Table',
    range: { startRow: 0, startCol: 0, endRow: products.length, endCol: 5 },
    tableStyle: 'TableStyleMedium9',
    showRowStripes: true,
    headerRow: true,
    totalsRow: true,
    columns: headers.map((h, i) => ({
      name: h,
      totalsRowFunction: i === 4 ? 'sum' : i === 5 ? 'sum' : 'none',
    })),
  });

  sheet.colWidth(0, 12).colWidth(1, 18).colWidth(2, 14).colWidth(3, 12).colWidth(4, 10).colWidth(5, 14);
  sheet.freezeRows(1);

  wb.saveAs(path.join(OUTPUT_DIR, '07_tables.xlsx'));
  console.log('  ✓ dist/07_tables.xlsx');
}

// ═══════════════════════════════════════════════════════════════════
// 8. CHARTS
// ═══════════════════════════════════════════════════════════════════

console.log('📈 Building Charts Example…');
{
  const wb = new CellWeave({ title: 'Charts Demo' });
  const sheet = wb.addSheet('Charts');

  // Data for charts
  sheet.writeHeaders('A1', ['Quarter', 'Product A', 'Product B', 'Product C']);
  const chartData = [
    ['Q1', 12000, 8000, 5000],
    ['Q2', 15000, 9500, 7000],
    ['Q3', 11000, 11000, 8500],
    ['Q4', 18000, 13000, 10000],
  ];
  sheet.writeArray(1, 0, chartData);

  // Column chart
  sheet.addChart({
    type: 'column',
    title: 'Quarterly Revenue by Product',
    series: [
      { name: 'Product A', valuesRange: { startRow: 1, startCol: 1, endRow: 4, endCol: 1 }, categoriesRange: { startRow: 1, startCol: 0, endRow: 4, endCol: 0 } },
      { name: 'Product B', valuesRange: { startRow: 1, startCol: 2, endRow: 4, endCol: 2 }, categoriesRange: { startRow: 1, startCol: 0, endRow: 4, endCol: 0 } },
      { name: 'Product C', valuesRange: { startRow: 1, startCol: 3, endRow: 4, endCol: 3 }, categoriesRange: { startRow: 1, startCol: 0, endRow: 4, endCol: 0 } },
    ],
    position: { row: 6, col: 0, rowSpan: 15, colSpan: 10 },
    showDataLabels: true,
  });

  // Line chart  
  sheet.addChart({
    type: 'line',
    title: 'Revenue Trend',
    series: [
      { name: 'Product A', valuesRange: { startRow: 1, startCol: 1, endRow: 4, endCol: 1 } },
    ],
    position: { row: 6, col: 11, rowSpan: 15, colSpan: 10 },
  });

  // Pie chart
  sheet.addChart({
    type: 'pie',
    title: 'Q4 Market Share',
    series: [
      { name: 'Market Share', valuesRange: { startRow: 4, startCol: 1, endRow: 4, endCol: 3 }, categoriesRange: { startRow: 0, startCol: 1, endRow: 0, endCol: 3 } },
    ],
    position: { row: 22, col: 0, rowSpan: 15, colSpan: 10 },
    showDataLabels: true,
  });

  sheet.colWidth(0, 10).colWidth(1, 12).colWidth(2, 12).colWidth(3, 12);

  wb.saveAs(path.join(OUTPUT_DIR, '08_charts.xlsx'));
  console.log('  ✓ dist/08_charts.xlsx');
}

// ═══════════════════════════════════════════════════════════════════
// 9. PAGE SETUP, HEADERS, FOOTERS, GROUPING
// ═══════════════════════════════════════════════════════════════════

console.log('📄 Building Page Setup Example…');
{
  const wb = new CellWeave({ title: 'Page Setup Demo' });
  const sheet = wb.addSheet('Print Layout');

  // Title
  sheet.writeCell('A1', cell().value('Annual Report').bold().fontSize(20).alignH('center'));
  sheet.merge('A1:I1').rowHeight(0, 50);

  // Data
  for (let i = 1; i <= 50; i++) {
    sheet.write(i, 0, { value: `Item ${i}` });
    for (let j = 1; j <= 8; j++) {
      sheet.write(i, j, { value: Math.round(Math.random() * 1000) });
    }
  }

  sheet
    .landscape()
    .paperA4()
    .fitToPage(1, 0)
    .printScale(90)
    .pageMargins({ top: 0.5, bottom: 0.5, left: 0.5, right: 0.5, header: 0.3, footer: 0.3 })
    .headerFooter({
      oddHeader:  '&L&"Calibri,Bold"Annual Report&C&[Date]&R&[Time]',
      oddFooter:  '&LConfidential&CPage &P of &N&R© CellWeave 2024',
      firstHeader: '&C&"Calibri,Bold"ANNUAL REPORT - CONFIDENTIAL',
      differentFirst: true,
    })
    .pageBreakAfterRow(25)
    .freezeRows(1)
    .groupRows(1, 10)
    .groupRows(11, 25)
    .groupCols(1, 4)
    .tabColor('#007ACC');

  wb.saveAs(path.join(OUTPUT_DIR, '09_page_setup.xlsx'));
  console.log('  ✓ dist/09_page_setup.xlsx');
}

// ═══════════════════════════════════════════════════════════════════
// 10. MULTIPLE SHEETS & DEFINED NAMES
// ═══════════════════════════════════════════════════════════════════

console.log('📚 Building Multi-Sheet Example…');
{
  const wb = new CellWeave({ title: 'Multi-Sheet Report', author: 'Finance Team' });

  // Summary sheet
  const summary = wb.addSheet('Summary').tabColor('#2196F3');
  summary.writeCell('A1', cell().value('COMPANY DASHBOARD').bold().fontSize(18).alignH('center').bgColor('#0D47A1').color('#FFFFFF'));
  summary.merge('A1:E1').rowHeight(0, 45);
  summary.writeHeaders('A2', ['Metric', 'Q1', 'Q2', 'Q3', 'Q4']);
  ['Revenue', 'Expenses', 'Profit', 'Headcount'].forEach((metric, i) => {
    summary.write(2 + i, 0, { value: metric, style: { font: { bold: true } } });
    for (let q = 0; q < 4; q++) {
      summary.write(2 + i, 1 + q, { formula: `${['Revenue','Expenses','Profit','Config'][i]}!B${q + 2}` });
    }
  });

  // Regional sheets
  ['North', 'South', 'East', 'West'].forEach((region, ri) => {
    const sheet = wb.addSheet(region).tabColor(['#E91E63', '#4CAF50', '#FF9800', '#9C27B0'][ri]);
    sheet.writeHeaders('A1', ['Product', 'Revenue', 'Expenses', 'Profit']);
    for (let i = 0; i < 5; i++) {
      sheet.write(i + 1, 0, { value: `Product ${i + 1}` });
      const rev = Math.round(10000 + Math.random() * 40000);
      const exp = Math.round(rev * (0.4 + Math.random() * 0.2));
      sheet.write(i + 1, 1, { value: rev, style: { numberFormat: { type: 'currency', currencySymbol: '$', decimalPlaces: 0 } } });
      sheet.write(i + 1, 2, { value: exp, style: { numberFormat: { type: 'currency', currencySymbol: '$', decimalPlaces: 0 } } });
      sheet.write(i + 1, 3, { formula: `B${i + 2}-C${i + 2}`, style: { numberFormat: { type: 'currency', currencySymbol: '$', decimalPlaces: 0 } } });
    }
    sheet.autoFilter('A1:D1').freezeRows(1).colWidth(0, 14).colWidth(1, 14).colWidth(2, 14).colWidth(3, 14);
  });

  // Config sheet (hidden-ish)
  const config = wb.addSheet('Config');
  config.write('A1', { value: 'Configuration Sheet' });

  // Defined names
  wb.defineRange('RevenueData', "North!$B$2:$B$6");
  wb.defineRange('ExpenseData', "South!$C$2:$C$6");

  // Move/copy sheet demo
  wb.moveSheet('Summary', 0);

  wb.saveAs(path.join(OUTPUT_DIR, '10_multi_sheet.xlsx'));
  console.log('  ✓ dist/10_multi_sheet.xlsx');
}

// ═══════════════════════════════════════════════════════════════════
// 11. QUICK HELPERS: fromArray & fromObjects
// ═══════════════════════════════════════════════════════════════════

console.log('⚡ Building Quick Helpers Example…');
{
  // fromArray
  const arr = fromArray(
    [['Alice', 28, 'Engineering'], ['Bob', 35, 'Marketing'], ['Carol', 31, 'Sales']],
    { sheetName: 'Staff', headers: ['Name', 'Age', 'Department'] }
  );
  require('fs').writeFileSync(path.join(OUTPUT_DIR, '11a_from_array.xlsx'), arr);

  // fromObjects
  const objs = fromObjects([
    { product: 'CellWeave', version: '1.0.0', downloads: 50000 },
    { product: 'NodeJS',    version: '20.0',  downloads: 1000000 },
  ]);
  require('fs').writeFileSync(path.join(OUTPUT_DIR, '11b_from_objects.xlsx'), objs);
  console.log('  ✓ dist/11a_from_array.xlsx');
  console.log('  ✓ dist/11b_from_objects.xlsx');
}

console.log('\n✅ All examples generated in dist/ directory!');
