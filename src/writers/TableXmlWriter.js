/**
 * @fileoverview CellWeave – TableXmlWriter (Single Responsibility)
 * Generates xl/tables/tableN.xml for Excel Tables (ListObject).
 */

'use strict';

const { el, xmlDoc, escapeXml } = require('../utils/xmlBuilder');
const { encodeRange, encodeCell } = require('../utils/addressUtils');

const TBL_NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';

/**
 * @param {import('../interfaces').ITable} table
 * @param {number} tableId
 * @returns {string}
 */
function buildTableXml(table, tableId) {
  const ref = encodeRange(table.range.startRow, table.range.startCol, table.range.endRow, table.range.endCol);

  const columns = (table.columns || []);
  const colCount = table.range.endCol - table.range.startCol + 1;

  const tableColumnsXml = el('tableColumns', { count: columns.length || colCount },
    (columns.length ? columns : Array.from({ length: colCount }, (_, i) => ({ name: `Column${i + 1}` }))).map((col, idx) => {
      const colAttr = {
        id:   idx + 1,
        name: escapeXml(col.name || `Column${idx + 1}`),
        totalsRowFunction: col.totalsRowFunction,
        totalsRowLabel:    col.totalsRowLabel,
      };
      const colContent = col.formula ? el('calculatedColumnFormula', {}, escapeXml(col.formula)) : '';
      return el('tableColumn', colAttr, colContent || undefined);
    }).join('')
  );

  const tableStyleInfo = el('tableStyleInfo', {
    name:              table.tableStyle || 'TableStyleMedium9',
    showFirstColumn:   table.showFirstColumn   ? 1 : 0,
    showLastColumn:    table.showLastColumn     ? 1 : 0,
    showRowStripes:    table.showRowStripes !== false ? 1 : 0,
    showColumnStripes: table.showColumnStripes  ? 1 : 0,
  });

  return xmlDoc('table', {
    xmlns:       TBL_NS,
    id:          tableId,
    name:        escapeXml(table.name),
    displayName: escapeXml(table.displayName || table.name),
    ref,
    headerRowCount: table.headerRow === false ? 0 : 1,
    totalsRowCount: table.totalsRow ? 1 : 0,
  }, tableColumnsXml + tableStyleInfo);
}

module.exports = { buildTableXml };
