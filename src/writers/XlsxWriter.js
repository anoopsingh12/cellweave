/**
 * @fileoverview CellWeave – XlsxWriter (Open-Closed / Dependency Inversion)
 * Orchestrates all writers to produce a valid .xlsx binary (ZIP-based).
 * Uses fflate for zero-native-dep zip creation.
 *
 * SOLID:
 *  - Open-Closed: new features injected via writers, not modifying this class.
 *  - Dependency Inversion: depends on writer abstractions, not concrete XML strings.
 *  - Single Responsibility: only assembles the ZIP; delegates XML generation.
 */

'use strict';

const { strToU8, zipSync } = require('fflate');

const { buildWorksheetXml }   = require('./WorksheetXmlWriter');
const { buildStylesXml }       = require('./StylesXmlWriter');
const { buildSharedStringsXml }= require('./SharedStringsXmlWriter');
const { buildChartXml, buildDrawingXml } = require('./ChartXmlWriter');
const { buildTableXml }        = require('./TableXmlWriter');
const {
  buildWorkbookXml,
  buildWorkbookRelsXml,
  buildRootRelsXml,
  buildContentTypesXml,
  buildCorePropsXml,
  buildAppPropsXml,
  buildThemeXml,
} = require('./WorkbookXmlWriter');

const { el, xmlDoc } = require('../utils/xmlBuilder');
const { encodeRange } = require('../utils/addressUtils');

/**
 * @class XlsxWriter
 * Converts a Workbook model into an .xlsx Buffer using fflate.
 */
class XlsxWriter {
  /**
   * @param {import('../models/Workbook').Workbook} workbook
   */
  constructor(workbook) {
    this._wb = workbook;
  }

  /**
   * Serialise the workbook to a Uint8Array (.xlsx bytes).
   * @returns {Uint8Array}
   */
  build() {
    const wb  = this._wb;
    const zip = {};

    // ── Pre-pass: register all cell styles and strings ─────────────────────
    // (StyleRegistry & SST are built lazily during worksheet XML generation)

    // ── Worksheet XMLs ─────────────────────────────────────────────────────
    const sheetDrawingMap  = new Map();   // sheetIdx → drawingId
    const sheetTableMap    = new Map();   // sheetIdx → [tableId, ...]
    let   drawingCounter   = 1;
    let   chartCounter     = 1;
    let   tableCounter     = 1;
    let   imageCounter     = 1;

    wb.sheets.forEach((ws, sheetIdx) => {
      // Worksheet XML
      const wsXml = buildWorksheetXml(ws, wb.styles, wb.sharedStrings, wb.properties.date1904);
      zip[`xl/worksheets/sheet${sheetIdx + 1}.xml`] = strToU8(wsXml);

      // Per-sheet relationships
      const wsRels = [];
      const drawItems = [];

      // ── Charts ────────────────────────────────────────────────────────────
      if (ws.charts.length > 0) {
        const drawId = drawingCounter++;
        sheetDrawingMap.set(sheetIdx, drawId);

        ws.charts.forEach(chart => {
          const cId = chartCounter++;
          const chartXml = buildChartXml(chart, cId, ws.name);
          zip[`xl/charts/chart${cId}.xml`] = strToU8(chartXml);

          // Chart rels
          zip[`xl/charts/_rels/chart${cId}.xml.rels`] = strToU8(
            xmlDoc('Relationships', { xmlns: 'http://schemas.openxmlformats.org/package/2006/relationships' }, '')
          );

          const relId = `rIdChart${cId}`;
          wsRels.push({ id: relId, type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart', target: `../charts/chart${cId}.xml` });
          const pos = chart.position || { row: 1, col: 1, rowSpan: 15, colSpan: 8 };
          drawItems.push({ type: 'chart', relId, id: cId, pos });
        });

        // Drawing XML for this sheet
        const drawXml = buildDrawingXml(drawItems);
        zip[`xl/drawings/drawing${drawId}.xml`] = strToU8(drawXml);
        wsRels.push({
          id:     `rIdDraw${drawId}`,
          type:   'http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing',
          target: `../drawings/drawing${drawId}.xml`,
        });
      }

      // ── Images ────────────────────────────────────────────────────────────
      if (ws.images.length > 0) {
        let drawId = sheetDrawingMap.get(sheetIdx);
        if (!drawId) {
          drawId = drawingCounter++;
          sheetDrawingMap.set(sheetIdx, drawId);
        }

        ws.images.forEach(img => {
          const imgId = imageCounter++;
          const ext = img.type || 'png';
          const imgPath = `xl/media/image${imgId}.${ext}`;
          if (typeof img.data === 'string') {
            // base64
            const binary = Buffer.from(img.data, 'base64');
            zip[imgPath] = new Uint8Array(binary);
          } else if (img.data instanceof Uint8Array || Buffer.isBuffer(img.data)) {
            zip[imgPath] = new Uint8Array(img.data);
          }
          const relId = `rIdImg${imgId}`;
          wsRels.push({ id: relId, type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image', target: `../media/image${imgId}.${ext}` });
          const pos = img.position || { row: 0, col: 0, rowSpan: 10, colSpan: 5 };
          drawItems.push({ type: 'image', relId, id: imgId, pos });
        });
      }

      // ── Tables ────────────────────────────────────────────────────────────
      if (ws.tables.length > 0) {
        const tIds = [];
        ws.tables.forEach(tbl => {
          const tId = tableCounter++;
          const tblXml = buildTableXml(tbl, tId);
          zip[`xl/tables/table${tId}.xml`] = strToU8(tblXml);
          tIds.push(tId);
          wsRels.push({
            id:     `rIdTbl${tId}`,
            type:   'http://schemas.openxmlformats.org/officeDocument/2006/relationships/table',
            target: `../tables/table${tId}.xml`,
          });
        });
        sheetTableMap.set(sheetIdx, tIds);
      }

      // ── Hyperlinks ────────────────────────────────────────────────────────
      let hyperlinkCounter = 1;
      for (const cell of ws.cells()) {
        if (cell.hyperlink) {
          wsRels.push({
            id:         `rIdHL${hyperlinkCounter++}`,
            type:       'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink',
            target:     cell.hyperlink,
            targetMode: 'External',
          });
        }
      }

      // Write worksheet rels if any
      if (wsRels.length > 0) {
        const relsXml = this._buildRelsXml(wsRels);
        zip[`xl/worksheets/_rels/sheet${sheetIdx + 1}.xml.rels`] = strToU8(relsXml);
      }
    });

    // ── Styles ─────────────────────────────────────────────────────────────
    zip['xl/styles.xml'] = strToU8(buildStylesXml(wb.styles));

    // ── Shared Strings ─────────────────────────────────────────────────────
    const hasSharedStrings = wb.sharedStrings.uniqueCount > 0;
    if (hasSharedStrings) {
      zip['xl/sharedStrings.xml'] = strToU8(buildSharedStringsXml(wb.sharedStrings));
    }

    // ── Theme ──────────────────────────────────────────────────────────────
    zip['xl/theme/theme1.xml'] = strToU8(buildThemeXml());

    // ── Workbook XML ───────────────────────────────────────────────────────
    zip['xl/workbook.xml'] = strToU8(buildWorkbookXml(wb));

    // ── Workbook Rels ──────────────────────────────────────────────────────
    zip['xl/_rels/workbook.xml.rels'] = strToU8(buildWorkbookRelsXml(wb, hasSharedStrings));

    // ── Root Rels ──────────────────────────────────────────────────────────
    zip['_rels/.rels'] = strToU8(buildRootRelsXml());

    // ── Content Types ──────────────────────────────────────────────────────
    zip['[Content_Types].xml'] = strToU8(buildContentTypesXml(wb, hasSharedStrings, sheetDrawingMap, sheetTableMap));

    // ── Core / App props ───────────────────────────────────────────────────
    zip['docProps/core.xml'] = strToU8(buildCorePropsXml(wb.properties));
    zip['docProps/app.xml']  = strToU8(buildAppPropsXml(wb.properties, wb));

    // ── Compress and return ────────────────────────────────────────────────
    return zipSync(zip, { level: 6 });
  }

  _buildRelsXml(rels) {
    const REL_NS = 'http://schemas.openxmlformats.org/package/2006/relationships';
    const relsXml = rels.map(r => {
      const a = { Id: r.id, Type: r.type, Target: r.target };
      if (r.targetMode) a.TargetMode = r.targetMode;
      return el('Relationship', a);
    }).join('');
    return xmlDoc('Relationships', { xmlns: REL_NS }, relsXml);
  }
}

module.exports = { XlsxWriter };
