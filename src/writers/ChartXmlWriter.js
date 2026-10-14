/**
 * @fileoverview CellWeave – ChartXmlWriter (Single Responsibility)
 * Generates xl/charts/chartN.xml and xl/drawings/drawingN.xml for chart embeds.
 */

'use strict';

const { el, escapeXml } = require('../utils/xmlBuilder');
const { encodeRange } = require('../utils/addressUtils');

const DML  = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const C_NS = 'http://schemas.openxmlformats.org/drawingml/2006/chart';
const R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const XDR   = 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing';

/**
 * Map CellWeave chart type → OOXML chart element name.
 */
const CHART_TAG = {
  bar:     'barChart',
  column:  'barChart',   // barDir="col"
  line:    'lineChart',
  pie:     'pieChart',
  doughnut:'doughnutChart',
  scatter: 'scatterChart',
  area:    'areaChart',
  radar:   'radarChart',
  bubble:  'bubbleChart',
};

/**
 * Build the XML for a single chart.
 * @param {import('../interfaces').IChartConfig} chart
 * @param {number} chartId
 * @param {string} sheetName - for data ranges
 * @returns {string}
 */
function buildChartXml(chart, chartId, sheetName) {
  const tag  = CHART_TAG[chart.type] || 'lineChart';
  const isBar = chart.type === 'bar' || chart.type === 'column';

  let plotAreaContent = '';

  // ── Series ────────────────────────────────────────────────────────────────
  const seriesXml = (chart.series || []).map((s, idx) => {
    let serXml = el('c:idx', { val: idx }) + el('c:order', { val: idx });

    // Series name
    if (s.name || s.nameRange) {
      const tx = s.nameRange
        ? el('c:strRef', {}, el('c:f', {}, `'${sheetName}'!${encodeRange(s.nameRange.startRow, s.nameRange.startCol, s.nameRange.endRow, s.nameRange.endCol)}`))
        : el('c:v', {}, escapeXml(s.name || ''));
      serXml += el('c:tx', {}, s.nameRange ? tx : el('c:strRef', {}, el('c:v', {}, tx)));
    }

    // Categories
    if (s.categoriesRange) {
      const catRef = `'${sheetName}'!${encodeRange(s.categoriesRange.startRow, s.categoriesRange.startCol, s.categoriesRange.endRow, s.categoriesRange.endCol)}`;
      serXml += el('c:cat', {}, el('c:numRef', {}, el('c:f', {}, catRef)));
    }

    // Values
    if (s.valuesRange) {
      const valRef = `'${sheetName}'!${encodeRange(s.valuesRange.startRow, s.valuesRange.startCol, s.valuesRange.endRow, s.valuesRange.endCol)}`;
      serXml += el('c:val', {}, el('c:numRef', {}, el('c:f', {}, valRef)));
    }

    return el('c:ser', {}, serXml);
  }).join('');

  // ── Chart-specific attributes ─────────────────────────────────────────────
  let chartSpecAttr = {};
  if (isBar)  chartSpecAttr = { barDir: chart.type === 'column' ? 'col' : 'bar', grouping: 'clustered' };
  if (chart.type === 'line') chartSpecAttr = { grouping: 'standard' };
  if (chart.type === 'doughnut') chartSpecAttr = { holeSize: 50 };

  const dataLabels = chart.showDataLabels
    ? el('c:dLbls', {}, el('c:showLegendKey', { val: 0 }) + el('c:showVal', { val: 1 }) + el('c:showCatName', { val: 0 }) + el('c:showSerName', { val: 0 }) + el('c:showPercent', { val: 0 }) + el('c:showBubbleSize', { val: 0 }))
    : el('c:dLbls', {}, el('c:showLegendKey', { val: 0 }) + el('c:showVal', { val: 0 }) + el('c:showCatName', { val: 0 }) + el('c:showSerName', { val: 0 }) + el('c:showPercent', { val: 0 }) + el('c:showBubbleSize', { val: 0 }));

  plotAreaContent = el(`c:${tag}`, chartSpecAttr, seriesXml + dataLabels);

  // Axes (not for pie/doughnut)
  if (!['pie', 'doughnut'].includes(chart.type)) {
    plotAreaContent += el('c:catAx', {},
      el('c:axId', { val: 1 }) + el('c:scaling', {}, el('c:orientation', { val: 'minMax' })) +
      el('c:delete', { val: 0 }) + el('c:axPos', { val: 'b' }) + el('c:crossAx', { val: 2 })
    );
    plotAreaContent += el('c:valAx', {},
      el('c:axId', { val: 2 }) + el('c:scaling', {}, el('c:orientation', { val: 'minMax' })) +
      el('c:delete', { val: 0 }) + el('c:axPos', { val: 'l' }) + el('c:crossAx', { val: 1 })
    );
  }

  const titleXml = chart.title
    ? el('c:title', {},
        el('c:tx', {},
          el('c:rich', {},
            el('a:bodyPr', { xmlns_a: DML }) +
            el('a:lstStyle', { xmlns_a: DML }) +
            el('a:p', { xmlns_a: DML }, el('a:r', { xmlns_a: DML }, el('a:t', { xmlns_a: DML }, escapeXml(chart.title))))
          )
        ) + el('c:overlay', { val: 0 })
      )
    : '';

  const legend = (chart.legend !== false)
    ? el('c:legend', {}, el('c:legendPos', { val: 'b' }) + el('c:overlay', { val: 0 }))
    : '';

  const chartXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="${C_NS}" xmlns:a="${DML}" xmlns:r="${R_NS}">
  <c:date1904 val="0"/>
  <c:lang val="en-US"/>
  <c:roundedCorners val="0"/>
  <c:chart>
    ${titleXml}
    <c:autoTitleDeleted val="${chart.title ? 0 : 1}"/>
    <c:plotArea>${plotAreaContent}</c:plotArea>
    ${legend}
    <c:plotVisOnly val="1"/>
    <c:dispBlanksAs val="gap"/>
    <c:showDLblsOverMax val="0"/>
  </c:chart>
</c:chartSpace>`;

  return chartXml;
}

/**
 * Build xl/drawings/drawingN.xml with anchors for charts and/or images.
 * @param {Array<{type:'chart'|'image', relId:string, pos:{row:number,col:number,rowSpan:number,colSpan:number}}>} items
 * @returns {string}
 */
function buildDrawingXml(items) {
  const anchors = items.map(item => {
    const { pos } = item;
    const from = `<xdr:from><xdr:col>${pos.col}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>${pos.row}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>`;
    const to   = `<xdr:to><xdr:col>${pos.col + (pos.colSpan || 8)}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>${pos.row + (pos.rowSpan || 15)}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>`;

    let graphic;
    if (item.type === 'chart') {
      graphic = `<xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="${item.id}" name="Chart ${item.id}"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="${item.relId}"/></a:graphicData></a:graphic></xdr:graphicFrame>`;
    } else {
      graphic = `<xdr:pic><xdr:nvPicPr><xdr:cNvPr id="${item.id}" name="Image ${item.id}"/><xdr:cNvPicPr/></xdr:nvPicPr><xdr:blipFill><a:blip r:embed="${item.relId}"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill><xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1000000" cy="1000000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr></xdr:pic>`;
    }

    return `<xdr:twoCellAnchor editAs="oneCell">${from}${to}${graphic}<xdr:clientData/></xdr:twoCellAnchor>`;
  }).join('');

  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="${XDR}" xmlns:a="${DML}" xmlns:r="${R_NS}">
${anchors}
</xdr:wsDr>`;
}

module.exports = { buildChartXml, buildDrawingXml };
