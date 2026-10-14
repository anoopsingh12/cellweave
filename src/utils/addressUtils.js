/**
 * @fileoverview CellWeave – Address Utilities
 * Handles A1/R1C1 notation parsing and encoding (Single Responsibility Principle)
 */

'use strict';

/**
 * Convert a 0-based column index to an Excel column letter(s).
 * @param {number} col - 0-based column index
 * @returns {string} e.g. 0 → "A", 25 → "Z", 26 → "AA"
 */
function colIndexToLetter(col) {
  if (col < 0) throw new RangeError(`Column index must be >= 0, got ${col}`);
  let result = '';
  let n = col;
  do {
    result = String.fromCharCode(65 + (n % 26)) + result;
    n = Math.floor(n / 26) - 1;
  } while (n >= 0);
  return result;
}

/**
 * Convert an Excel column letter string to a 0-based index.
 * @param {string} letters - e.g. "A", "AA", "XFD"
 * @returns {number}
 */
function colLetterToIndex(letters) {
  if (!letters || typeof letters !== 'string') throw new TypeError('Column letters must be a non-empty string');
  const upper = letters.toUpperCase();
  let result = 0;
  for (let i = 0; i < upper.length; i++) {
    const charCode = upper.charCodeAt(i) - 64;
    if (charCode < 1 || charCode > 26) throw new RangeError(`Invalid column letter: ${letters}`);
    result = result * 26 + charCode;
  }
  return result - 1;
}

/**
 * Encode a 0-based {row, col} to A1 notation.
 * @param {number} row
 * @param {number} col
 * @param {boolean} [absRow=false] - prepend $ to row
 * @param {boolean} [absCol=false] - prepend $ to col
 * @returns {string} e.g. "A1", "$B$3"
 */
function encodeCell(row, col, absRow = false, absCol = false) {
  return `${absCol ? '$' : ''}${colIndexToLetter(col)}${absRow ? '$' : ''}${row + 1}`;
}

/**
 * Parse A1 notation to a 0-based {row, col, absRow, absCol} object.
 * @param {string} ref - e.g. "A1", "$B$3", "AA100"
 * @returns {{row:number, col:number, absRow:boolean, absCol:boolean}}
 */
function decodeCell(ref) {
  const match = /^(\$?)([A-Za-z]{1,3})(\$?)(\d+)$/.exec(ref.trim());
  if (!match) throw new SyntaxError(`Invalid cell reference: "${ref}"`);
  return {
    col: colLetterToIndex(match[2]),
    row: parseInt(match[4], 10) - 1,
    absCol: match[1] === '$',
    absRow: match[3] === '$',
  };
}

/**
 * Encode a range to A1 notation.
 * @param {number} startRow
 * @param {number} startCol
 * @param {number} endRow
 * @param {number} endCol
 * @returns {string} e.g. "A1:C3"
 */
function encodeRange(startRow, startCol, endRow, endCol) {
  return `${encodeCell(startRow, startCol)}:${encodeCell(endRow, endCol)}`;
}

/**
 * Parse a range string like "A1:C3" or "Sheet1!A1:C3".
 * @param {string} ref
 * @returns {{sheet?:string, startRow:number, startCol:number, endRow:number, endCol:number}}
 */
function decodeRange(ref) {
  let sheetName;
  let rangePart = ref.trim();

  // Strip sheet reference
  const sheetMatch = /^(?:'([^']+)'|([^!]+))!(.+)$/.exec(rangePart);
  if (sheetMatch) {
    sheetName = sheetMatch[1] || sheetMatch[2];
    rangePart = sheetMatch[3];
  }

  const parts = rangePart.split(':');
  if (parts.length === 1) {
    const cell = decodeCell(parts[0]);
    return { sheet: sheetName, startRow: cell.row, startCol: cell.col, endRow: cell.row, endCol: cell.col };
  }
  const start = decodeCell(parts[0]);
  const end = decodeCell(parts[1]);
  return {
    sheet: sheetName,
    startRow: Math.min(start.row, end.row),
    startCol: Math.min(start.col, end.col),
    endRow: Math.max(start.row, end.row),
    endCol: Math.max(start.col, end.col),
  };
}

/**
 * Check if two ranges overlap.
 * @param {import('../interfaces').ICellRange} a
 * @param {import('../interfaces').ICellRange} b
 * @returns {boolean}
 */
function rangesOverlap(a, b) {
  return !(
    a.endRow < b.startRow ||
    b.endRow < a.startRow ||
    a.endCol < b.startCol ||
    b.endCol < a.startCol
  );
}

/**
 * Check if a cell is within a range.
 * @param {number} row
 * @param {number} col
 * @param {import('../interfaces').ICellRange} range
 * @returns {boolean}
 */
function cellInRange(row, col, range) {
  return row >= range.startRow && row <= range.endRow && col >= range.startCol && col <= range.endCol;
}

/**
 * Convert column width in characters to EMUs (used in xlsx XML).
 * @param {number} chars
 * @returns {number}
 */
function charWidthToEmu(chars) {
  return Math.round(chars * 9525);
}

/**
 * Convert row height in points to EMUs.
 * @param {number} pts
 * @returns {number}
 */
function ptToEmu(pts) {
  return Math.round(pts * 12700);
}

/**
 * Parse a number as a date serial (days since 1900-01-01 or 1904-01-01).
 * @param {number} serial
 * @param {boolean} [date1904=false]
 * @returns {Date}
 */
function serialToDate(serial, date1904 = false) {
  const epoch = date1904 ? Date.UTC(1904, 0, 1) : Date.UTC(1900, 0, 0);
  // Excel quirk: 1900 system counts Feb 29 1900 (which didn't exist)
  const adjusted = date1904 ? serial : serial > 59 ? serial - 1 : serial;
  return new Date(epoch + adjusted * 86400000);
}

/**
 * Convert a Date to an Excel date serial number.
 * @param {Date} date
 * @param {boolean} [date1904=false]
 * @returns {number}
 */
function dateToSerial(date, date1904 = false) {
  const epoch = date1904 ? Date.UTC(1904, 0, 1) : Date.UTC(1900, 0, 0);
  const days = (date.getTime() - epoch) / 86400000;
  return date1904 ? days : days > 59 ? days + 1 : days;
}

/**
 * Sanitize a string for use as an XML text node (escape special chars).
 * @param {string} str
 * @returns {string}
 */
function escapeXml(str) {
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

/**
 * Sanitize a sheet name (max 31 chars, no special chars).
 * @param {string} name
 * @returns {string}
 */
function sanitizeSheetName(name) {
  return name.replace(/[:\\/?\[\]*]/g, '_').substring(0, 31);
}

/**
 * Convert a hex color to argb format (e.g. "#FF0000" -> "FFFF0000").
 * @param {string} hex
 * @param {string} [alpha='FF']
 * @returns {string}
 */
function hexToArgb(hex, alpha = 'FF') {
  const cleaned = hex.replace('#', '').toUpperCase();
  if (cleaned.length === 6) return `${alpha}${cleaned}`;
  if (cleaned.length === 8) return cleaned;
  throw new RangeError(`Invalid hex color: ${hex}`);
}

/**
 * Generate a unique id string.
 * @returns {string}
 */
function generateId() {
  return Math.random().toString(36).substring(2, 11) + Date.now().toString(36);
}

module.exports = {
  colIndexToLetter,
  colLetterToIndex,
  encodeCell,
  decodeCell,
  encodeRange,
  decodeRange,
  rangesOverlap,
  cellInRange,
  charWidthToEmu,
  ptToEmu,
  serialToDate,
  dateToSerial,
  escapeXml,
  sanitizeSheetName,
  hexToArgb,
  generateId,
};
