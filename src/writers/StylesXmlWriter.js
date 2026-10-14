/**
 * @fileoverview CellWeave – StylesXmlWriter (Single Responsibility)
 * Generates the xl/styles.xml part from a StyleRegistry.
 */

'use strict';

const { el, xmlDoc } = require('../utils/xmlBuilder');
const { hexToArgb } = require('../utils/addressUtils');

const NS = {
  xmlns: 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
  'xmlns:mc': 'http://schemas.openxmlformats.org/markup-compatibility/2006',
};

/**
 * Build a color attribute set from hex or theme.
 * @param {string|Object} color
 * @returns {Object}
 */
function colorAttrs(color) {
  if (!color) return {};
  if (typeof color === 'string') return { rgb: hexToArgb(color) };
  if (color.theme !== undefined) return { theme: color.theme, tint: color.tint };
  if (color.indexed !== undefined) return { indexed: color.indexed };
  return {};
}

/**
 * Serialize a border side.
 * @param {import('../interfaces').IBorderSide|undefined} side
 * @param {string} tag
 * @returns {string}
 */
function borderSide(side, tag) {
  if (!side || !side.style) return el(tag, {});
  const colorEl = side.color ? el('color', colorAttrs(side.color)) : '';
  return el(tag, { style: side.style }, colorEl);
}

/**
 * @param {import('../core/StyleRegistry').StyleRegistry} registry
 * @returns {string} XML string for xl/styles.xml
 */
function buildStylesXml(registry) {
  // ── numFmts ───────────────────────────────────────────────────────────────
  const customFmts = registry.getNumFmts();
  const numFmtsXml = customFmts.length
    ? el('numFmts', { count: customFmts.length },
        customFmts.map(f => el('numFmt', { numFmtId: f.id, formatCode: f.code })).join(''))
    : '';

  // ── fonts ─────────────────────────────────────────────────────────────────
  const fontList = registry.getFonts();
  const fontsXml = el('fonts', { count: fontList.length },
    fontList.map(f => {
      const parts = [];
      if (f.bold)          parts.push(el('b', {}));
      if (f.italic)        parts.push(el('i', {}));
      if (f.strikethrough) parts.push(el('strike', {}));
      if (f.shadow)        parts.push(el('shadow', {}));
      if (f.condense)      parts.push(el('condense', {}));
      if (f.extend)        parts.push(el('extend', {}));
      if (f.underline && f.underline !== 'none') {
        parts.push(el('u', f.underline === 'single' ? {} : { val: f.underline }));
      }
      if (f.vertAlign && f.vertAlign !== 'baseline') {
        parts.push(el('vertAlign', { val: f.vertAlign }));
      }
      if (f.size) parts.push(el('sz', { val: f.size }));
      if (f.color) parts.push(el('color', colorAttrs(f.color)));
      if (f.name) parts.push(el('name', { val: f.name }));
      if (f.family) parts.push(el('family', { val: _familyVal(f.family) }));
      if (f.charset) parts.push(el('charset', { val: f.charset }));
      if (f.scheme && f.scheme !== 'none') parts.push(el('scheme', { val: f.scheme }));
      return el('font', {}, parts.join(''));
    }).join('')
  );

  // ── fills ─────────────────────────────────────────────────────────────────
  const fillList = registry.getFills();
  const fillsXml = el('fills', { count: fillList.length },
    fillList.map(f => {
      if (f.gradient) {
        const stopXml = (f.gradient.stops || []).map(s =>
          el('gradientStop', { position: s.position }, el('color', colorAttrs(s.color)))
        ).join('');
        return el('fill', {}, el('gradientFill', {
          type: f.gradient.type || 'linear',
          degree: f.gradient.degree || 0,
        }, stopXml));
      }
      const patAttr = { patternType: f.patternType || 'none' };
      const patContent = [
        f.fgColor ? el('fgColor', colorAttrs(f.fgColor)) : '',
        f.bgColor ? el('bgColor', colorAttrs(f.bgColor)) : '',
      ].join('');
      return el('fill', {}, el('patternFill', patAttr, patContent || undefined));
    }).join('')
  );

  // ── borders ───────────────────────────────────────────────────────────────
  const borderList = registry.getBorders();
  const bordersXml = el('borders', { count: borderList.length },
    borderList.map(b => {
      const dAttr = {};
      if (b.diagonalUp) dAttr.diagonalUp = 1;
      if (b.diagonalDown) dAttr.diagonalDown = 1;
      return el('border', dAttr,
        borderSide(b.left, 'left') +
        borderSide(b.right, 'right') +
        borderSide(b.top, 'top') +
        borderSide(b.bottom, 'bottom') +
        borderSide(b.diagonal, 'diagonal')
      );
    }).join('')
  );

  // ── cellStyleXfs (required placeholder) ──────────────────────────────────
  const cellStyleXfsXml = el('cellStyleXfs', { count: 1 },
    el('xf', { numFmtId: 0, fontId: 0, fillId: 0, borderId: 0 })
  );

  // ── cellXfs ───────────────────────────────────────────────────────────────
  const xfList = registry.getCellFormats();
  const cellXfsXml = el('cellXfs', { count: xfList.length },
    xfList.map(xf => {
      const a = {
        numFmtId:  xf.numFmtId,
        fontId:    xf.fontIdx,
        fillId:    xf.fillIdx,
        borderId:  xf.borderIdx,
        xfId:      0,
        applyFont:      xf.fontIdx   > 0 ? 1 : undefined,
        applyFill:      xf.fillIdx   > 1 ? 1 : undefined,  // 0,1 are defaults
        applyBorder:    xf.borderIdx > 0 ? 1 : undefined,
        applyNumberFormat: xf.numFmtId > 0 ? 1 : undefined,
        applyAlignment:    xf.alignment ? 1 : undefined,
        applyProtection:   (xf.locked === false || xf.hidden) ? 1 : undefined,
      };
      let content = '';
      if (xf.alignment) {
        const al = xf.alignment;
        content += el('alignment', {
          horizontal:        al.horizontal,
          vertical:          al.vertical,
          textRotation:      al.rotation,
          wrapText:          al.wrapText ? 1 : undefined,
          shrinkToFit:       al.shrinkToFit ? 1 : undefined,
          indent:            al.indent,
          justifyLastLine:   al.justifyLastLine ? 1 : undefined,
          readingOrder:      al.readingOrder,
        });
      }
      if (xf.locked === false || xf.hidden) {
        content += el('protection', { locked: xf.locked ? 1 : 0, hidden: xf.hidden ? 1 : 0 });
      }
      return el('xf', a, content || undefined);
    }).join('')
  );

  // ── cellStyles ────────────────────────────────────────────────────────────
  const cellStylesXml = el('cellStyles', { count: 1 },
    el('cellStyle', { name: 'Normal', xfId: 0, builtinId: 0 })
  );

  const body =
    numFmtsXml +
    fontsXml +
    fillsXml +
    bordersXml +
    cellStyleXfsXml +
    cellXfsXml +
    cellStylesXml;

  return xmlDoc('styleSheet', NS, body);
}

function _familyVal(family) {
  const map = { Automatic: 0, Roman: 1, Swiss: 2, Modern: 3, Script: 4, Decorative: 5 };
  return map[family] !== undefined ? map[family] : family;
}

module.exports = { buildStylesXml };
