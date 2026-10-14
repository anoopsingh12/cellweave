/**
 * @fileoverview CellWeave – XmlBuilder (Single Responsibility)
 * Lightweight, zero-dependency XML building utility.
 * Replaces heavy DOM libraries — pure string concatenation with proper escaping.
 */

'use strict';

const { escapeXml } = require('../utils/addressUtils');

/**
 * Build an XML attribute string from a plain object.
 * @param {Object} attrs
 * @returns {string}
 */
function attrs(obj) {
  return Object.entries(obj)
    .filter(([, v]) => v !== undefined && v !== null)
    .map(([k, v]) => ` ${k}="${escapeXml(String(v))}"`)
    .join('');
}

/**
 * Wrap content in an XML element.
 * @param {string} tag
 * @param {Object} [attrObj]
 * @param {string} [content]
 * @returns {string}
 */
function el(tag, attrObj, content) {
  const a = attrObj ? attrs(attrObj) : '';
  if (content === undefined || content === null || content === '') {
    return `<${tag}${a}/>`;
  }
  return `<${tag}${a}>${content}</${tag}>`;
}

/**
 * Standard OOXML prologue + root element open.
 * @param {string} rootTag
 * @param {Object} rootAttrs
 * @param {string} content
 * @returns {string}
 */
function xmlDoc(rootTag, rootAttrs, content) {
  const a = rootAttrs ? attrs(rootAttrs) : '';
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<${rootTag}${a}>${content}</${rootTag}>`;
}

module.exports = { el, attrs, xmlDoc, escapeXml };
