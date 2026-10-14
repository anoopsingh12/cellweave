/**
 * @fileoverview CellWeave – Comprehensive Test Suite
 */

'use strict';

const { CellWeave, CellBuilder } = require('../index');
const path = require('path');
const fs   = require('fs');

const OUT = path.join(__dirname, '../dist');
if (!fs.existsSync(OUT)) fs.mkdirSync(OUT, { recursive: true });

// ── Address Utilities ─────────────────────────────────────────────────────────

describe('addressUtils', () => {
  const { colIndexToLetter, colLetterToIndex, encodeCell, decodeCell, encodeRange, decodeRange, hexToArgb, dateToSerial, serialToDate } = require('../src/utils/addressUtils');

  test('colIndexToLetter: 0→A, 25→Z, 26→AA', () => {
    expect(colIndexToLetter(0)).toBe('A');
    expect(colIndexToLetter(25)).toBe('Z');
    expect(colIndexToLetter(26)).toBe('AA');
    expect(colIndexToLetter(701)).toBe('ZZ');
  });

  test('colLetterToIndex: A→0, Z→25, AA→26', () => {
    expect(colLetterToIndex('A')).toBe(0);
    expect(colLetterToIndex('Z')).toBe(25);
    expect(colLetterToIndex('AA')).toBe(26);
    expect(colLetterToIndex('ZZ')).toBe(701);
  });

  test('encodeCell: row=0,col=0 → A1', () => {
    expect(encodeCell(0, 0)).toBe('A1');
    expect(encodeCell(9, 2)).toBe('C10');
    expect(encodeCell(0, 0, true, true)).toBe('$A$1');
  });

  test('decodeCell: A1 → {row:0, col:0}', () => {
    const r = decodeCell('A1');
    expect(r).toEqual({ row: 0, col: 0, absRow: false, absCol: false });
    const r2 = decodeCell('$B$3');
    expect(r2).toEqual({ row: 2, col: 1, absRow: true, absCol: true });
  });

  test('encodeRange / decodeRange', () => {
    expect(encodeRange(0, 0, 2, 2)).toBe('A1:C3');
    const r = decodeRange('A1:C3');
    expect(r).toMatchObject({ startRow: 0, startCol: 0, endRow: 2, endCol: 2 });
  });

  test('decodeRange with sheet name', () => {
    const r = decodeRange("Sheet1!A1:B2");
    expect(r.sheet).toBe('Sheet1');
    expect(r.startRow).toBe(0);
    expect(r.endCol).toBe(1);
  });

  test('hexToArgb', () => {
    expect(hexToArgb('#FF0000')).toBe('FFFF0000');
    expect(hexToArgb('#FF0000', 'AA')).toBe('AAFF0000');
  });

  test('dateToSerial / serialToDate roundtrip', () => {
    const date = new Date(Date.UTC(2024, 0, 1));
    const serial = dateToSerial(date);
    const back = serialToDate(serial);
    expect(back.getUTCFullYear()).toBe(2024);
    expect(back.getUTCMonth()).toBe(0);
    expect(back.getUTCDate()).toBe(1);
  });

  test('invalid cell ref throws', () => {
    expect(() => decodeCell('123')).toThrow();
    expect(() => decodeCell('')).toThrow();
  });
});

// ── StyleRegistry ─────────────────────────────────────────────────────────────

describe('StyleRegistry', () => {
  const { StyleRegistry } = require('../src/core/StyleRegistry');

  test('deduplicates fonts', () => {
    const reg = new StyleRegistry();
    const a = reg.registerFont({ name: 'Calibri', size: 11, bold: true });
    const b = reg.registerFont({ name: 'Calibri', size: 11, bold: true });
    expect(a).toBe(b);
    const c = reg.registerFont({ name: 'Arial', size: 12 });
    expect(c).not.toBe(a);
  });

  test('deduplicates fills', () => {
    const reg = new StyleRegistry();
    const a = reg.registerFill({ patternType: 'solid', fgColor: 'FF0000' });
    const b = reg.registerFill({ patternType: 'solid', fgColor: 'FF0000' });
    expect(a).toBe(b);
  });

  test('registers custom number format starting at 164', () => {
    const reg = new StyleRegistry();
    const id = reg.registerNumberFormat('#,##0.000');
    expect(id).toBeGreaterThanOrEqual(164);
  });

  test('resolves builtin number format codes', () => {
    const reg = new StyleRegistry();
    const id = reg.resolveNumberFormat({ formatCode: '0.00' });
    expect(id).toBe(2); // builtin
  });

  test('registerCellStyle returns 0 for empty style', () => {
    const reg = new StyleRegistry();
    // seed default
    reg.getCellFormats(); // force init check
    const id = reg.registerCellStyle(null);
    expect(id).toBe(0);
  });
});

// ── SharedStringTable ─────────────────────────────────────────────────────────

describe('SharedStringTable', () => {
  const { SharedStringTable } = require('../src/core/SharedStringTable');

  test('deduplicates strings', () => {
    const sst = new SharedStringTable();
    const a = sst.addString('Hello');
    const b = sst.addString('Hello');
    expect(a).toBe(b);
    expect(sst.uniqueCount).toBe(1);
    expect(sst.totalCount).toBe(2);
  });

  test('different strings get different indices', () => {
    const sst = new SharedStringTable();
    const a = sst.addString('Alpha');
    const b = sst.addString('Beta');
    expect(a).not.toBe(b);
  });

  test('rich text stored separately', () => {
    const sst = new SharedStringTable();
    sst.addString('plain');
    sst.addRichText([{ text: 'bold', font: { bold: true } }]);
    expect(sst.uniqueCount).toBe(2);
  });
});

// ── Cell Model ────────────────────────────────────────────────────────────────

describe('Cell', () => {
  const { Cell } = require('../src/models/Cell');

  test('resolvedType for string', () => {
    const c = new Cell(0, 0, { value: 'hello' });
    expect(c.resolvedType()).toBe('s');
  });

  test('resolvedType for number', () => {
    const c = new Cell(0, 0, { value: 42 });
    expect(c.resolvedType()).toBe('n');
  });

  test('resolvedType for boolean', () => {
    const c = new Cell(0, 0, { value: true });
    expect(c.resolvedType()).toBe('b');
  });

  test('resolvedType for date', () => {
    const c = new Cell(0, 0, { value: new Date() });
    expect(c.resolvedType()).toBe('n'); // dates stored as numbers
  });

  test('resolvedType for formula', () => {
    const c = new Cell(0, 0, { formula: 'SUM(A1:A10)' });
    expect(c.resolvedType()).toBe('str');
  });

  test('numericValue for boolean', () => {
    expect(new Cell(0, 0, { value: true }).numericValue()).toBe(1);
    expect(new Cell(0, 0, { value: false }).numericValue()).toBe(0);
  });

  test('address returns A1 notation', () => {
    const c = new Cell(0, 0);
    expect(c.address).toBe('A1');
    const c2 = new Cell(9, 2);
    expect(c2.address).toBe('C10');
  });

  test('merge returns new Cell with merged data', () => {
    const c = new Cell(0, 0, { value: 'original' });
    const merged = c.merge({ value: 'updated', formula: 'A1+1' });
    expect(merged.value).toBe('updated');
    expect(merged.formula).toBe('A1+1');
    expect(c.value).toBe('original'); // immutable
  });
});

// ── Worksheet Model ───────────────────────────────────────────────────────────

describe('Worksheet', () => {
  const { Worksheet } = require('../src/models/Worksheet');

  test('setCell / getCell by A1 and indices', () => {
    const ws = new Worksheet('Test', 1);
    ws.setCell('B2', { value: 99 });
    expect(ws.getCell('B2').value).toBe(99);
    expect(ws.getCell(1, 1).value).toBe(99);
  });

  test('writeArray fills correct cells', () => {
    const ws = new Worksheet('Test', 1);
    ws.writeArray(0, 0, [[1, 2], [3, 4]]);
    expect(ws.getCell(0, 0).value).toBe(1);
    expect(ws.getCell(0, 1).value).toBe(2);
    expect(ws.getCell(1, 0).value).toBe(3);
    expect(ws.getCell(1, 1).value).toBe(4);
  });

  test('mergeCells adds to merges array', () => {
    const ws = new Worksheet('Test', 1);
    ws.mergeCells('A1:C3');
    expect(ws.merges.length).toBe(1);
    expect(ws.merges[0]).toMatchObject({ startRow: 0, startCol: 0, endRow: 2, endCol: 2 });
  });

  test('overlapping merge throws', () => {
    const ws = new Worksheet('Test', 1);
    ws.mergeCells('A1:C3');
    expect(() => ws.mergeCells('B2:D4')).toThrow();
  });

  test('setAutoFilter', () => {
    const ws = new Worksheet('Test', 1);
    ws.setAutoFilter('A1:D1');
    expect(ws.autoFilter).not.toBeNull();
    expect(ws.autoFilter.range.startCol).toBe(0);
  });

  test('dimensions are tracked', () => {
    const ws = new Worksheet('Test', 1);
    ws.setCell(5, 10, { value: 'x' });
    expect(ws.dimensions).toMatchObject({ endRow: 5, endCol: 10 });
  });

  test('row and col properties', () => {
    const ws = new Worksheet('Test', 1);
    ws.setRowProperties(2, { height: 30, hidden: false });
    expect(ws.rowProps.get(2).height).toBe(30);
    ws.setColProperties(3, { width: 20 });
    expect(ws.colProps.get(3).width).toBe(20);
  });
});

// ── Workbook Model ────────────────────────────────────────────────────────────

describe('Workbook', () => {
  const { Workbook } = require('../src/models/Workbook');

  test('addSheet creates and returns Worksheet', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Alpha');
    expect(ws.name).toBe('Alpha');
    expect(wb.sheetCount).toBe(1);
  });

  test('getSheet by name and index', () => {
    const wb = new Workbook();
    wb.addSheet('First');
    wb.addSheet('Second');
    expect(wb.getSheet('First').name).toBe('First');
    expect(wb.getSheet(1).name).toBe('Second');
  });

  test('duplicate sheet name throws', () => {
    const wb = new Workbook();
    wb.addSheet('Data');
    expect(() => wb.addSheet('Data')).toThrow();
  });

  test('removeSheet removes by name', () => {
    const wb = new Workbook();
    wb.addSheet('A');
    wb.addSheet('B');
    wb.removeSheet('A');
    expect(wb.sheetCount).toBe(1);
  });

  test('renameSheet works', () => {
    const wb = new Workbook();
    wb.addSheet('Old');
    wb.renameSheet('Old', 'New');
    expect(wb.getSheet('New').name).toBe('New');
  });

  test('moveSheet reorders correctly', () => {
    const wb = new Workbook();
    wb.addSheet('A');
    wb.addSheet('B');
    wb.addSheet('C');
    wb.moveSheet('C', 0);
    expect(wb.sheets[0].name).toBe('C');
  });

  test('definedNames are added and removed', () => {
    const wb = new Workbook();
    wb.addDefinedName({ name: 'MyRange', refers: 'Sheet1!$A$1:$C$3' });
    expect(wb.definedNames.length).toBe(1);
    wb.removeDefinedName('MyRange');
    expect(wb.definedNames.length).toBe(0);
  });

  test('sheet name is sanitized (max 31 chars)', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('A'.repeat(40));
    expect(ws.name.length).toBeLessThanOrEqual(31);
  });
});

// ── CellBuilder ───────────────────────────────────────────────────────────────

describe('CellBuilder', () => {
  test('basic value and font', () => {
    const data = new CellBuilder()
      .value('Hello')
      .bold()
      .fontSize(14)
      .color('#FF0000')
      .build();
    expect(data.value).toBe('Hello');
    expect(data.style.font.bold).toBe(true);
    expect(data.style.font.size).toBe(14);
    expect(data.style.font.color).toBe('FF0000');
  });

  test('number format currency', () => {
    const data = new CellBuilder().value(1234.5).formatCurrency('£', 2).build();
    expect(data.style.numberFormat.type).toBe('currency');
    expect(data.style.numberFormat.currencySymbol).toBe('£');
  });

  test('alignment wrap text', () => {
    const data = new CellBuilder().value('long text').wrapText().alignH('center').build();
    expect(data.style.alignment.wrapText).toBe(true);
    expect(data.style.alignment.horizontal).toBe('center');
  });

  test('borders', () => {
    const data = new CellBuilder().allBorders('thin', '#000000').build();
    expect(data.style.border.top.style).toBe('thin');
    expect(data.style.border.bottom.style).toBe('thin');
  });

  test('gradient fill', () => {
    const data = new CellBuilder()
      .gradient('linear', 90, [{ position: 0, color: '#FF0000' }, { position: 1, color: '#0000FF' }])
      .build();
    expect(data.style.fill.gradient.type).toBe('linear');
    expect(data.style.fill.gradient.stops.length).toBe(2);
  });

  test('formula builder', () => {
    const data = new CellBuilder().formula('SUM(A1:A10)').build();
    expect(data.formula).toBe('SUM(A1:A10)');
  });

  test('hyperlink', () => {
    const data = new CellBuilder().value('Click here').hyperlink('https://example.com').build();
    expect(data.hyperlink).toBe('https://example.com');
  });

  test('rich text', () => {
    const data = new CellBuilder()
      .richText([{ text: 'Bold ', font: { bold: true } }, { text: 'Normal' }])
      .build();
    expect(data.richText.length).toBe(2);
    expect(data.richText[0].font.bold).toBe(true);
  });
});

// ── WorksheetBuilder ─────────────────────────────────────────────────────────

describe('WorksheetBuilder', () => {
  const { WorksheetBuilder } = require('../src/builders/WorksheetBuilder');
  const { Worksheet } = require('../src/models/Worksheet');

  function makeBuilder() {
    return new WorksheetBuilder(new Worksheet('TestSheet', 1));
  }

  test('write by A1 ref', () => {
    const b = makeBuilder();
    b.write('A1', 42);
    expect(b.sheet.getCell('A1').value).toBe(42);
  });

  test('writeHeaders creates header cells', () => {
    const b = makeBuilder();
    b.writeHeaders('A1', ['Name', 'Age', 'Score']);
    expect(b.sheet.getCell('A1').value).toBe('Name');
    expect(b.sheet.getCell('C1').value).toBe('Score');
  });

  test('merge chains', () => {
    const b = makeBuilder();
    b.write('A1', 'Title').merge('A1:C1');
    expect(b.sheet.merges.length).toBe(1);
  });

  test('freezeRows sets freeze panes', () => {
    const b = makeBuilder();
    b.freezeRows(1);
    expect(b.sheet.sheetView.freezePanes.rows).toBe(1);
  });

  test('autoFilter chains', () => {
    const b = makeBuilder();
    b.autoFilter('A1:D1');
    expect(b.sheet.autoFilter).not.toBeNull();
  });

  test('dropdownList adds data validation', () => {
    const b = makeBuilder();
    b.dropdownList('A1:A10', ['Yes', 'No', 'Maybe']);
    expect(b.sheet.dataValidations.length).toBe(1);
    expect(b.sheet.dataValidations[0].formula1).toBe('"Yes,No,Maybe"');
  });

  test('colorScale adds conditional formatting', () => {
    const b = makeBuilder();
    b.colorScale('A1:A10');
    expect(b.sheet.conditionalFormattings.length).toBe(1);
  });

  test('colWidth and rowHeight', () => {
    const b = makeBuilder();
    b.colWidth(0, 25).rowHeight(0, 30);
    expect(b.sheet.colProps.get(0).width).toBe(25);
    expect(b.sheet.rowProps.get(0).height).toBe(30);
  });

  test('protect sets protection', () => {
    const b = makeBuilder();
    b.protect('secret');
    expect(b.sheet.protection.sheet).toBe(true);
    expect(b.sheet.protection.password).toBe('secret');
  });

  test('tabColor sets color', () => {
    const b = makeBuilder();
    b.tabColor('#FF0000');
    expect(b.sheet.tabColor).toBe('#FF0000');
  });

  test('addChart stores chart config', () => {
    const b = makeBuilder();
    b.addChart({ type: 'line', title: 'Sales', series: [] });
    expect(b.sheet.charts.length).toBe(1);
  });

  test('addTable stores table config', () => {
    const b = makeBuilder();
    b.addTable({ name: 'SalesTable', range: { startRow: 0, startCol: 0, endRow: 9, endCol: 3 } });
    expect(b.sheet.tables.length).toBe(1);
  });

  test('groupRows adds grouping', () => {
    const b = makeBuilder();
    b.groupRows(1, 5);
    expect(b.sheet.groupings.length).toBe(1);
    expect(b.sheet.groupings[0].type).toBe('row');
  });

  test('pageBreakAfterRow stores break', () => {
    const b = makeBuilder();
    b.pageBreakAfterRow(10);
    expect(b.sheet.rowPageBreaks).toContain(10);
  });
});

// ── Full XLSX Generation ─────────────────────────────────────────────────────

describe('CellWeave full xlsx generation', () => {
  test('creates a minimal xlsx buffer', () => {
    const cw = new CellWeave({ title: 'Test Workbook' });
    const s = cw.addSheet('Sheet1');
    s.write('A1', 'Hello World');
    const buf = cw.toBuffer();
    expect(buf).toBeInstanceOf(Uint8Array);
    expect(buf.length).toBeGreaterThan(1000);
    // ZIP magic bytes
    expect(buf[0]).toBe(0x50);
    expect(buf[1]).toBe(0x4B);
  });

  test('fromArray static method', () => {
    const buf = CellWeave.fromArray([[1, 2], [3, 4]], { headers: ['A', 'B'] });
    expect(buf).toBeInstanceOf(Uint8Array);
    expect(buf[0]).toBe(0x50);
  });

  test('fromObjects static method', () => {
    const buf = CellWeave.fromObjects([{ name: 'Alice', age: 30 }, { name: 'Bob', age: 25 }]);
    expect(buf).toBeInstanceOf(Uint8Array);
  });

  test('multiple sheets', () => {
    const cw = new CellWeave();
    cw.addSheet('Data');
    cw.addSheet('Summary');
    cw.addSheet('Config');
    const buf = cw.toBuffer();
    expect(buf[0]).toBe(0x50);
  });

  test('saveAs writes a file', () => {
    const outPath = path.join(OUT, 'test_minimal.xlsx');
    const cw = new CellWeave({ title: 'Test' });
    cw.addSheet('Sheet1').write('A1', 'CellWeave');
    cw.saveAs(outPath);
    expect(fs.existsSync(outPath)).toBe(true);
    const stat = fs.statSync(outPath);
    expect(stat.size).toBeGreaterThan(1000);
  });

  test('toBase64 returns valid base64', () => {
    const cw = new CellWeave();
    cw.addSheet('S').write('A1', 'test');
    const b64 = cw.toBase64();
    expect(typeof b64).toBe('string');
    expect(b64.length).toBeGreaterThan(100);
    // Verify it can be decoded back
    const decoded = Buffer.from(b64, 'base64');
    expect(decoded[0]).toBe(0x50);
    expect(decoded[1]).toBe(0x4B);
  });

  test('workbook with all features produces valid zip', () => {
    const cw = new CellWeave({ title: 'Full Feature Test', author: 'Tester' });
    const s = cw.addSheet('Report');

    // Headers + data
    s.writeHeaders('A1', ['Product', 'Q1', 'Q2', 'Q3', 'Q4']);
    s.writeArray(1, 0, [
      ['Widget A', 1000, 1200, 900, 1500],
      ['Widget B', 800, 950, 1100, 1300],
      ['Widget C', 600, 700, 650, 800],
    ]);

    // Styling
    s.writeCell('A1', new CellBuilder().value('Product').bold().bgColor('#4472C4').color('#FFFFFF').build());
    s.colWidth(0, 20).colWidth(1, 12).colWidth(2, 12);
    s.rowHeight(0, 25);
    s.merge('F1:H1');
    s.freezeRows(1);
    s.autoFilter('A1:E1');
    s.colorScale('B2:E4');
    s.dataBar('B2:E4');
    s.dropdownList('A10:A20', ['Widget A', 'Widget B', 'Widget C']);
    s.landscape().paperA4();
    s.tabColor('#FF6600');
    s.addChart({ type: 'column', title: 'Quarterly Sales', series: [{ name: 'Q1', valuesRange: { startRow: 1, startCol: 1, endRow: 3, endCol: 1 } }], position: { row: 6, col: 0, rowSpan: 15, colSpan: 8 } });
    s.addTable({ name: 'SalesTable', displayName: 'Sales_Table', range: { startRow: 0, startCol: 0, endRow: 3, endCol: 4 }, tableStyle: 'TableStyleMedium9' });
    s.groupRows(1, 3);
    s.pageBreakAfterRow(5);
    s.headerFooter({ oddHeader: '&CQuarterly Report', oddFooter: '&CPage &P of &N' });

    // Defined names
    cw.defineRange('SalesData', "Report!$B$2:$E$4");

    const buf = cw.toBuffer();
    expect(buf[0]).toBe(0x50);
    expect(buf[1]).toBe(0x4B);

    // Save for manual inspection
    fs.writeFileSync(path.join(OUT, 'test_full_features.xlsx'), buf);
  });
});
