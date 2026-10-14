/**
 * @fileoverview CellWeave – WorksheetXmlWriter (Single Responsibility)
 * Converts a Worksheet model into xl/worksheets/sheetN.xml OOXML markup.
 * Handles: cells, rows, columns, merges, CF, data validation, auto filter,
 * freeze panes, page setup, header/footer, page breaks, groupings, protection.
 */

'use strict';

const { el, xmlDoc, escapeXml } = require('../utils/xmlBuilder');
const {
  encodeCell,
  encodeRange,
  hexToArgb,
} = require('../utils/addressUtils');

const WS_NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const R_NS  = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const MC_NS = 'http://schemas.openxmlformats.org/markup-compatibility/2006';
// X14AC is a publicly documented OOXML extension namespace (ECMA-376 Part 4).
// The URI is a standard identifier, not proprietary Microsoft source code.
const X14AC = 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac';

/**
 * @param {import('../models/Worksheet').Worksheet} ws
 * @param {import('../core/StyleRegistry').StyleRegistry} styleReg
 * @param {import('../core/SharedStringTable').SharedStringTable} sst
 * @param {boolean} [date1904=false]
 * @returns {string} XML for the worksheet
 */
function buildWorksheetXml(ws, styleReg, sst, date1904 = false) {
  const parts = [];

  // ── sheetPr (tab color, page setup properties) ─────────────────────────
  const sheetPrContent = ws.tabColor
    ? el('tabColor', { rgb: hexToArgb(ws.tabColor) })
    : '';
  if (sheetPrContent) parts.push(el('sheetPr', {}, sheetPrContent));

  // ── dimension ─────────────────────────────────────────────────────────────
  const dim = ws.dimensions;
  if (dim) {
    const ref = dim.startRow === dim.endRow && dim.startCol === dim.endCol
      ? encodeCell(dim.startRow, dim.startCol)
      : encodeRange(dim.startRow, dim.startCol, dim.endRow, dim.endCol);
    parts.push(el('dimension', { ref }));
  }

  // ── sheetViews ────────────────────────────────────────────────────────────
  parts.push(buildSheetViews(ws));

  // ── sheetFormatPr ──────────────────────────────────────────────────────────
  parts.push(el('sheetFormatPr', { defaultRowHeight: 15, 'x14ac:dyDescent': 0.25 }));

  // ── cols ───────────────────────────────────────────────────────────────────
  if (ws.colProps.size > 0) {
    const colEls = buildColsXml(ws, styleReg);
    if (colEls) parts.push(el('cols', {}, colEls));
  }

  // ── sheetData ─────────────────────────────────────────────────────────────
  parts.push(buildSheetData(ws, styleReg, sst, date1904));

  // ── sheetCalcPr ───────────────────────────────────────────────────────────
  parts.push(el('sheetCalcPr', { fullCalcOnLoad: 1 }));

  // ── protection ────────────────────────────────────────────────────────────
  if (ws.protection) parts.push(buildProtectionXml(ws.protection));

  // ── mergeCells ────────────────────────────────────────────────────────────
  if (ws.merges.length > 0) {
    const mergeEls = ws.merges.map(m =>
      el('mergeCell', { ref: encodeRange(m.startRow, m.startCol, m.endRow, m.endCol) })
    ).join('');
    parts.push(el('mergeCells', { count: ws.merges.length }, mergeEls));
  }

  // ── conditionalFormatting ─────────────────────────────────────────────────
  for (const cf of ws.conditionalFormattings) {
    parts.push(buildConditionalFormattingXml(cf, styleReg));
  }

  // ── dataValidations ───────────────────────────────────────────────────────
  if (ws.dataValidations.length > 0) {
    const dvEls = ws.dataValidations.map(dv => buildDataValidationXml(dv)).join('');
    parts.push(el('dataValidations', { count: ws.dataValidations.length, disablePrompts: 0 }, dvEls));
  }

  // ── autoFilter ────────────────────────────────────────────────────────────
  if (ws.autoFilter) {
    const { range, columnFilters } = ws.autoFilter;
    const afRef = encodeRange(range.startRow, range.startCol, range.endRow, range.endCol);
    const filterEls = (columnFilters || []).map(cf => {
      const filters = cf.filters.map(f => el('filter', { val: f })).join('');
      return el('filterColumn', { colId: cf.colIndex }, el('filters', {}, filters));
    }).join('');
    parts.push(el('autoFilter', { ref: afRef }, filterEls || undefined));
  }

  // ── row page breaks ────────────────────────────────────────────────────────
  if (ws.rowPageBreaks.length > 0) {
    const brkEls = ws.rowPageBreaks.map(r => el('brk', { id: r + 1, max: 16383, man: 1 })).join('');
    parts.push(el('rowBreaks', { count: ws.rowPageBreaks.length, manualBreakCount: ws.rowPageBreaks.length }, brkEls));
  }

  // ── col page breaks ────────────────────────────────────────────────────────
  if (ws.colPageBreaks.length > 0) {
    const brkEls = ws.colPageBreaks.map(c => el('brk', { id: c + 1, max: 1048575, man: 1 })).join('');
    parts.push(el('colBreaks', { count: ws.colPageBreaks.length, manualBreakCount: ws.colPageBreaks.length }, brkEls));
  }

  // ── pageSetup & pageMargins ───────────────────────────────────────────────
  parts.push(buildPageMargins(ws.pageMargins));
  parts.push(buildPageSetup(ws.pageSetup));
  if (Object.keys(ws.headerFooter).length > 0) {
    parts.push(buildHeaderFooter(ws.headerFooter));
  }

  return xmlDoc('worksheet', {
    xmlns:      WS_NS,
    'xmlns:r':  R_NS,
    'xmlns:mc': MC_NS,
    'mc:Ignorable': 'x14ac',
    'xmlns:x14ac': X14AC,
  }, parts.join(''));
}

// ── sheetViews ──────────────────────────────────────────────────────────────

function buildSheetViews(ws) {
  const sv = ws.sheetView || {};
  const a = {
    showGridLines:     sv.showGridLines === false ? 0 : undefined,
    showRowColHeaders: sv.showRowColHeaders === false ? 0 : undefined,
    showFormulas:      sv.showFormulas ? 1 : undefined,
    showOutlineSymbols: sv.showOutlineSymbols === false ? 0 : undefined,
    showZeros:         sv.showZeros === false ? 0 : undefined,
    rightToLeft:       sv.rightToLeft ? 1 : undefined,
    tabSelected:       sv.tabSelected ? 1 : undefined,
    view:              sv.viewType && sv.viewType !== 'normal' ? sv.viewType : undefined,
    zoomScale:         sv.zoomScale && sv.zoomScale !== 100 ? sv.zoomScale : undefined,
    workbookViewId:    0,
  };

  let svContent = '';

  // Freeze panes
  if (sv.freezePanes) {
    const fp = sv.freezePanes;
    const rows = fp.rows || 0;
    const cols = fp.cols || 0;
    const state = 'frozen';
    const topLeftCell = encodeCell(rows, cols);
    svContent += el('pane', {
      xSplit: cols || undefined,
      ySplit: rows || undefined,
      topLeftCell: (rows > 0 || cols > 0) ? topLeftCell : undefined,
      activePane: rows > 0 && cols > 0 ? 'bottomRight' : rows > 0 ? 'bottomLeft' : 'topRight',
      state,
    });
    // Selection elements
    if (rows > 0 && cols > 0) {
      svContent += el('selection', { pane: 'topRight', activeCell: encodeCell(0, cols), sqref: encodeCell(0, cols) });
      svContent += el('selection', { pane: 'bottomLeft', activeCell: encodeCell(rows, 0), sqref: encodeCell(rows, 0) });
      svContent += el('selection', { pane: 'bottomRight', activeCell: topLeftCell, sqref: topLeftCell });
    } else if (rows > 0) {
      svContent += el('selection', { pane: 'bottomLeft', activeCell: topLeftCell, sqref: topLeftCell });
    } else {
      svContent += el('selection', { pane: 'topRight', activeCell: topLeftCell, sqref: topLeftCell });
    }
  } else {
    svContent += el('selection', { activeCell: 'A1', sqref: 'A1' });
  }

  return el('sheetViews', {}, el('sheetView', a, svContent));
}

// ── Columns ─────────────────────────────────────────────────────────────────

function buildColsXml(ws, styleReg) {
  const sortedCols = [...ws.colProps.entries()].sort(([a], [b]) => a - b);
  return sortedCols.map(([idx, props]) => {
    const styleIdx = props.style ? styleReg.registerCellStyle(props.style) : undefined;
    return el('col', {
      min:          idx + 1,
      max:          idx + 1,
      width:        props.width !== undefined ? props.width : undefined,
      style:        styleIdx,
      hidden:       props.hidden ? 1 : undefined,
      bestFit:      props.bestFit ? 1 : undefined,
      customWidth:  props.width !== undefined ? 1 : undefined,
      collapsed:    props.collapsed ? 1 : undefined,
      outlineLevel: props.outlineLevel,
    });
  }).join('');
}

// ── Sheet Data ───────────────────────────────────────────────────────────────

function buildSheetData(ws, styleReg, sst, date1904) {
  // Group cells by row
  const rowMap = new Map();
  for (const cell of ws.cells()) {
    if (!rowMap.has(cell.row)) rowMap.set(cell.row, []);
    rowMap.get(cell.row).push(cell);
  }

  // Merge row-only entries (rows with props but no cells)
  for (const [rowIdx] of ws.rowProps) {
    if (!rowMap.has(rowIdx)) rowMap.set(rowIdx, []);
  }

  const sortedRows = [...rowMap.entries()].sort(([a], [b]) => a - b);

  const rowsXml = sortedRows.map(([rowIdx, cells]) => {
    const rProps = ws.rowProps.get(rowIdx) || {};
    const rStyleIdx = rProps.style ? styleReg.registerCellStyle(rProps.style) : undefined;
    const rowAttrs = {
      r:            rowIdx + 1,
      ht:           rProps.height,
      customHeight: rProps.height ? 1 : undefined,
      hidden:       rProps.hidden ? 1 : undefined,
      collapsed:    rProps.collapsed ? 1 : undefined,
      outlineLevel: rProps.outlineLevel,
      s:            rStyleIdx,
      customFormat: rStyleIdx !== undefined ? 1 : undefined,
    };

    const cellsXml = cells.sort((a, b) => a.col - b.col).map(cell => {
      return buildCellXml(cell, styleReg, sst, date1904);
    }).join('');

    return el('row', rowAttrs, cellsXml || undefined);
  }).join('');

  return el('sheetData', {}, rowsXml);
}

function buildCellXml(cell, styleReg, sst, date1904) {
  const ref  = encodeCell(cell.row, cell.col);
  const xfId = cell.style ? styleReg.registerCellStyle(cell.style) : 0;
  const t    = cell.resolvedType();

  let cellAttr = { r: ref };
  if (t !== 'n' && t !== 'z') cellAttr.t = t;
  if (xfId) cellAttr.s = xfId;

  let content = '';

  if (cell.formula) {
    const fAttr = {};
    if (cell.isArrayFormula && cell.arrayFormulaRange) {
      const ar = cell.arrayFormulaRange;
      fAttr.t = 'array';
      fAttr.ref = encodeRange(ar.startRow, ar.startCol, ar.endRow, ar.endCol);
    }
    content += el('f', fAttr, escapeXml(cell.formula));
  }

  if (t === 'z' || cell.value === null || cell.value === undefined) {
    // empty cell — no <v>
  } else if (t === 'b') {
    content += el('v', {}, cell.numericValue() ? 1 : 0);
  } else if (t === 'n') {
    content += el('v', {}, cell.numericValue(date1904));
  } else if (t === 's') {
    // shared string
    let idx;
    if (cell.richText) {
      idx = sst.addRichText(cell.richText);
      cellAttr.t = 's'; // override to shared string
    } else {
      idx = sst.addString(String(cell.value));
    }
    content += el('v', {}, idx);
  } else if (t === 'inlineStr') {
    content += el('is', {}, el('t', {}, escapeXml(String(cell.value))));
  } else if (t === 'e') {
    content += el('v', {}, escapeXml(String(cell.value)));
  } else {
    // str (formula result string)
    content += el('v', {}, escapeXml(String(cell.value || '')));
  }

  return el('c', cellAttr, content || undefined);
}

// ── Conditional Formatting ───────────────────────────────────────────────────

function buildConditionalFormattingXml(cf, styleReg) {
  const sqref = encodeRange(cf.range.startRow, cf.range.startCol, cf.range.endRow, cf.range.endCol);
  const ruleEls = cf.rules.map((rule, idx) => {
    const ra = {
      type:       rule.type,
      operator:   rule.operator,
      dxfId:      rule.style ? styleReg.registerCellStyle(rule.style) : undefined,
      priority:   rule.priority !== undefined ? rule.priority : idx + 1,
      stopIfTrue: rule.stopIfTrue ? 1 : undefined,
      text:       rule.text,
      aboveAverage: rule.type === 'aboveAverage' ? undefined : undefined,
    };
    let ruleContent = '';
    (rule.formula || []).forEach(f => { ruleContent += el('formula', {}, escapeXml(f)); });
    if (rule.colorScale) ruleContent += buildColorScale(rule.colorScale);
    if (rule.dataBar)    ruleContent += buildDataBar(rule.dataBar);
    if (rule.iconSet)    ruleContent += buildIconSet(rule.iconSet);
    return el('cfRule', ra, ruleContent || undefined);
  }).join('');
  return el('conditionalFormatting', { sqref }, ruleEls);
}

function buildColorScale(cs) {
  const cfvos = cs.cfvo.map(v => el('cfvo', { type: v.type, val: v.val })).join('');
  const colors = cs.color.map(c => el('color', { rgb: hexToArgb(c) })).join('');
  return el('colorScale', {}, cfvos + colors);
}

function buildDataBar(db) {
  const cfvos = db.cfvo.map(v => el('cfvo', { type: v.type, val: v.val })).join('');
  const color = el('color', { rgb: hexToArgb(db.color || '#638EC6') });
  return el('dataBar', {}, cfvos + color);
}

function buildIconSet(is) {
  const cfvos = is.cfvo.map(v => el('cfvo', { type: v.type, val: v.val })).join('');
  return el('iconSet', { iconSet: is.iconSet || '3TrafficLights1' }, cfvos);
}

// ── Data Validation ──────────────────────────────────────────────────────────

function buildDataValidationXml(dv) {
  const sqref = encodeRange(dv.range.startRow, dv.range.startCol, dv.range.endRow, dv.range.endCol);
  const a = {
    type:             dv.type,
    errorStyle:       dv.errorStyle,
    operator:         dv.operator,
    allowBlank:       dv.allowBlank ? 1 : undefined,
    showDropDown:     dv.showDropDown ? 1 : undefined,
    showInputMessage: dv.showInputMessage ? 1 : undefined,
    showErrorAlert:   dv.showErrorAlert ? 1 : undefined,
    errorTitle:       dv.errorTitle,
    error:            dv.error,
    promptTitle:      dv.promptTitle,
    prompt:           dv.prompt,
    sqref,
  };
  let content = '';
  if (dv.formula1) content += el('formula1', {}, escapeXml(dv.formula1));
  if (dv.formula2) content += el('formula2', {}, escapeXml(dv.formula2));
  return el('dataValidation', a, content || undefined);
}

// ── Page Setup ───────────────────────────────────────────────────────────────

function buildPageMargins(m) {
  return el('pageMargins', {
    left:   m.left   !== undefined ? m.left   : 0.7,
    right:  m.right  !== undefined ? m.right  : 0.7,
    top:    m.top    !== undefined ? m.top    : 0.75,
    bottom: m.bottom !== undefined ? m.bottom : 0.75,
    header: m.header !== undefined ? m.header : 0.3,
    footer: m.footer !== undefined ? m.footer : 0.3,
  });
}

function buildPageSetup(ps) {
  if (!ps || !Object.keys(ps).length) return '';
  const paperMap = {
    letter: 1, legal: 5, A4: 9, A3: 8, A5: 11, B4: 12, B5: 13,
  };
  return el('pageSetup', {
    paperSize:       ps.paperSize ? (paperMap[ps.paperSize] || ps.paperSize) : undefined,
    orientation:     ps.orientation,
    scale:           ps.scale,
    fitToPage:       ps.fitToPage ? 1 : undefined,
    fitToWidth:      ps.fitToWidth,
    fitToHeight:     ps.fitToHeight,
    blackAndWhite:   ps.blackAndWhite ? 1 : undefined,
    draft:           ps.draft ? 1 : undefined,
    firstPageNumber: ps.firstPageNumber,
    pageOrder:       ps.pageOrder,
    useFirstPageNumber: ps.firstPageNumber ? 1 : undefined,
  });
}

function buildHeaderFooter(hf) {
  if (!Object.keys(hf).length) return '';
  const a = {
    differentOddEven:  hf.differentOddEven ? 1 : undefined,
    differentFirst:    hf.differentFirst ? 1 : undefined,
    scaleWithDoc:      hf.scaleWithDoc === false ? 0 : undefined,
    alignWithMargins:  hf.alignWithMargins === false ? 0 : undefined,
  };
  let content = '';
  if (hf.oddHeader)   content += el('oddHeader', {}, escapeXml(hf.oddHeader));
  if (hf.oddFooter)   content += el('oddFooter', {}, escapeXml(hf.oddFooter));
  if (hf.evenHeader)  content += el('evenHeader', {}, escapeXml(hf.evenHeader));
  if (hf.evenFooter)  content += el('evenFooter', {}, escapeXml(hf.evenFooter));
  if (hf.firstHeader) content += el('firstHeader', {}, escapeXml(hf.firstHeader));
  if (hf.firstFooter) content += el('firstFooter', {}, escapeXml(hf.firstFooter));
  return el('headerFooter', a, content || undefined);
}

// ── Protection ───────────────────────────────────────────────────────────────

function buildProtectionXml(prot) {
  const a = {
    sheet:               prot.sheet ? 1 : undefined,
    password:            prot.password ? hashPassword(prot.password) : undefined,
    selectLockedCells:   prot.selectLockedCells === false ? 0 : undefined,
    selectUnlockedCells: prot.selectUnlockedCells === false ? 0 : undefined,
    formatCells:         prot.formatCells ? undefined : 1,
    formatColumns:       prot.formatColumns ? undefined : 1,
    formatRows:          prot.formatRows ? undefined : 1,
    insertColumns:       prot.insertColumns ? undefined : 1,
    insertRows:          prot.insertRows ? undefined : 1,
    insertHyperlinks:    prot.insertHyperlinks ? undefined : 1,
    deleteColumns:       prot.deleteColumns ? undefined : 1,
    deleteRows:          prot.deleteRows ? undefined : 1,
    sort:                prot.sort ? undefined : 1,
    autoFilter:          prot.autoFilter ? undefined : 1,
    pivotTables:         prot.pivotTables ? undefined : 1,
    objects:             prot.objects ? undefined : 1,
    scenarios:           prot.scenarios ? undefined : 1,
  };
  return el('sheetProtection', a);
}

/**
 * XOR hash for Excel password protection (compatible with Excel).
 * @param {string} password
 * @returns {string} hex hash
 */
function hashPassword(password) {
  let hash = 0;
  if (!password) return '0000';
  for (let i = password.length - 1; i >= 0; i--) {
    hash = ((hash >> 14) & 0x01) | ((hash << 1) & 0x7fff);
    hash ^= password.charCodeAt(i);
  }
  hash = ((hash >> 14) & 0x01) | ((hash << 1) & 0x7fff);
  hash ^= password.length;
  hash ^= 0xCE4B;
  return hash.toString(16).toUpperCase().padStart(4, '0');
}

module.exports = { buildWorksheetXml, hashPassword };
