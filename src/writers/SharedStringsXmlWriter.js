/**
 * @fileoverview CellWeave – SharedStringsXmlWriter (Single Responsibility)
 * Generates xl/sharedStrings.xml from SharedStringTable.
 */

'use strict';

const { el, xmlDoc, escapeXml } = require('../utils/xmlBuilder');

const NS = {
  xmlns: 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
};

/**
 * @param {import('../core/SharedStringTable').SharedStringTable} table
 * @returns {string}
 */
function buildSharedStringsXml(table) {
  const strings = table.getStrings();
  const siXml = strings.map(entry => {
    if (entry.type === 'plain') {
      const needsPreserve = /^\s|\s$/.test(entry.value);
      const tAttr = needsPreserve ? { 'xml:space': 'preserve' } : {};
      return el('si', {}, el('t', tAttr, escapeXml(entry.value)));
    }
    // Rich text
    const runs = entry.runs.map(run => {
      const fontParts = run.font
        ? _fontToRPr(run.font)
        : '';
      const tAttr = /^\s|\s$/.test(run.text) ? { 'xml:space': 'preserve' } : {};
      return el('r', {}, fontParts + el('t', tAttr, escapeXml(run.text)));
    }).join('');
    return el('si', {}, runs);
  }).join('');

  return xmlDoc('sst', {
    ...NS,
    count: table.totalCount,
    uniqueCount: table.uniqueCount,
  }, siXml);
}

function _fontToRPr(font) {
  const parts = [];
  if (font.bold)          parts.push(el('b', {}));
  if (font.italic)        parts.push(el('i', {}));
  if (font.strikethrough) parts.push(el('strike', {}));
  if (font.underline && font.underline !== 'none') parts.push(el('u', font.underline !== 'single' ? { val: font.underline } : {}));
  if (font.vertAlign && font.vertAlign !== 'baseline') parts.push(el('vertAlign', { val: font.vertAlign }));
  if (font.size) parts.push(el('sz', { val: font.size }));
  if (font.color) parts.push(el('color', { rgb: font.color.replace('#','').toUpperCase().padStart(8, 'FF') }));
  if (font.name) parts.push(el('name', { val: font.name }));
  return parts.length ? el('rPr', {}, parts.join('')) : '';
}

module.exports = { buildSharedStringsXml };
