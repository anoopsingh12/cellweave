/**
 * @fileoverview CellWeave – WorkbookXmlWriter (Single Responsibility)
 * Generates workbook-level XML parts: xl/workbook.xml, _rels, [Content_Types].xml,
 * xl/app.xml, xl/core.xml, xl/theme/theme1.xml.
 */

'use strict';

const { el, xmlDoc, escapeXml } = require('../utils/xmlBuilder');

const WB_NS     = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const R_NS      = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const REL_NS    = 'http://schemas.openxmlformats.org/package/2006/relationships';
const CORE_NS   = 'http://schemas.openxmlformats.org/package/2006/metadata/core-properties';
const DC_NS     = 'http://purl.org/dc/elements/1.1/';
const DCTERMS   = 'http://purl.org/dc/terms/';
const XSI       = 'http://www.w3.org/2001/XMLSchema-instance';
const APP_NS    = 'http://schemas.openxmlformats.org/officeDocument/2006/extended-properties';
const VT_NS     = 'http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes';

// Relationship type prefixes
const RT = {
  SHEET:        'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet',
  STYLES:       'http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles',
  SHARED_STR:   'http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings',
  THEME:        'http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme',
  CHART:        'http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart',
  DRAWING:      'http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing',
  IMAGE:        'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
  WORKBOOK:     'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument',
  CORE:         'http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties',
  APP:          'http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties',
  TABLE:        'http://schemas.openxmlformats.org/officeDocument/2006/relationships/table',
};

/**
 * Build xl/workbook.xml
 * @param {import('../models/Workbook').Workbook} wb
 * @returns {string}
 */
function buildWorkbookXml(wb) {
  const fileVersion = el('fileVersion', { appName: 'xl', lastEdited: 5, lowestEdited: 5 });
  const wbPr = el('workbookPr', {
    date1904: wb.properties.date1904 ? 1 : undefined,
  });
  const bookViews = el('bookViews', {},
    el('workbookView', {
      xWindow:          -120,
      yWindow:          -120,
      windowWidth:      19440,
      windowHeight:     10260,
      activeTab:        wb._activeSheetIndex || 0,
    })
  );

  const sheetsXml = el('sheets', {},
    wb.sheets.map((ws, idx) =>
      el('sheet', {
        name:    escapeXml(ws.name),
        sheetId: ws.sheetId,
        'r:id':  `rId${idx + 1}`,
        state:   ws.sheetView && ws.sheetView.hidden ? 'hidden' : undefined,
      })
    ).join('')
  );

  let definedNamesXml = '';
  if (wb.definedNames.length > 0) {
    const names = wb.definedNames.map(dn => {
      const a = { name: dn.name, comment: dn.comment };
      if (dn.scope) {
        const sheetIdx = wb.sheets.findIndex(s => s.name === dn.scope);
        if (sheetIdx >= 0) a.localSheetId = sheetIdx;
      }
      return el('definedName', a, escapeXml(dn.refers));
    }).join('');
    definedNamesXml = el('definedNames', {}, names);
  }

  const calcPr = el('calcPr', { calcId: 0, concurrentCalc: 0 });

  return xmlDoc('workbook', { xmlns: WB_NS, 'xmlns:r': R_NS },
    fileVersion + wbPr + bookViews + sheetsXml + definedNamesXml + calcPr
  );
}

/**
 * Build xl/_rels/workbook.xml.rels
 * @param {import('../models/Workbook').Workbook} wb
 * @param {boolean} hasSharedStrings
 * @returns {string}
 */
function buildWorkbookRelsXml(wb, hasSharedStrings) {
  const rels = [];
  wb.sheets.forEach((ws, idx) => {
    rels.push(el('Relationship', {
      Id:     `rId${idx + 1}`,
      Type:   RT.SHEET,
      Target: `worksheets/sheet${idx + 1}.xml`,
    }));
  });

  const nextId = wb.sheets.length + 1;
  rels.push(el('Relationship', { Id: `rId${nextId}`,     Type: RT.STYLES,    Target: 'styles.xml' }));
  rels.push(el('Relationship', { Id: `rId${nextId + 1}`, Type: RT.THEME,     Target: 'theme/theme1.xml' }));
  if (hasSharedStrings) {
    rels.push(el('Relationship', { Id: `rId${nextId + 2}`, Type: RT.SHARED_STR, Target: 'sharedStrings.xml' }));
  }

  return xmlDoc('Relationships', { xmlns: REL_NS }, rels.join(''));
}

/**
 * Build _rels/.rels (root)
 * @returns {string}
 */
function buildRootRelsXml() {
  return xmlDoc('Relationships', { xmlns: REL_NS },
    el('Relationship', { Id: 'rId1', Type: RT.WORKBOOK, Target: 'xl/workbook.xml' }) +
    el('Relationship', { Id: 'rId2', Type: RT.CORE,     Target: 'docProps/core.xml' }) +
    el('Relationship', { Id: 'rId3', Type: RT.APP,      Target: 'docProps/app.xml' })
  );
}

/**
 * Build [Content_Types].xml
 * @param {import('../models/Workbook').Workbook} wb
 * @param {boolean} hasSharedStrings
 * @param {Map<number, number>} sheetDrawingMap - sheetIndex → drawingId
 * @param {Map<number, number[]>} sheetTableMap - sheetIndex → [tableIds]
 * @returns {string}
 */
function buildContentTypesXml(wb, hasSharedStrings, sheetDrawingMap = new Map(), sheetTableMap = new Map()) {
  const CT = 'http://schemas.openxmlformats.org/package/2006/content-types';
  const MIME = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

  const defaults = [
    el('Default', { Extension: 'rels', ContentType: 'application/vnd.openxmlformats-package.relationships+xml' }),
    el('Default', { Extension: 'xml',  ContentType: 'application/xml' }),
  ];

  const overrides = [
    el('Override', { PartName: '/xl/workbook.xml',           ContentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml' }),
    el('Override', { PartName: '/xl/styles.xml',             ContentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml' }),
    el('Override', { PartName: '/xl/theme/theme1.xml',       ContentType: 'application/vnd.openxmlformats-officedocument.theme+xml' }),
    el('Override', { PartName: '/docProps/core.xml',         ContentType: 'application/vnd.openxmlformats-package.core-properties+xml' }),
    el('Override', { PartName: '/docProps/app.xml',          ContentType: 'application/vnd.openxmlformats-officedocument.extended-properties+xml' }),
  ];

  if (hasSharedStrings) {
    overrides.push(el('Override', { PartName: '/xl/sharedStrings.xml', ContentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml' }));
  }

  wb.sheets.forEach((ws, idx) => {
    overrides.push(el('Override', {
      PartName:    `/xl/worksheets/sheet${idx + 1}.xml`,
      ContentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml',
    }));

    if (sheetDrawingMap.has(idx)) {
      const drawId = sheetDrawingMap.get(idx);
      overrides.push(el('Override', {
        PartName:    `/xl/drawings/drawing${drawId}.xml`,
        ContentType: 'application/vnd.openxmlformats-officedocument.drawing+xml',
      }));
    }

    if (sheetTableMap.has(idx)) {
      for (const tid of sheetTableMap.get(idx)) {
        overrides.push(el('Override', {
          PartName:    `/xl/tables/table${tid}.xml`,
          ContentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml',
        }));
      }
    }
  });

  return xmlDoc('Types', { xmlns: CT }, defaults.join('') + overrides.join(''));
}

/**
 * Build docProps/core.xml
 * @param {import('../interfaces').IWorkbookProperties} props
 * @returns {string}
 */
function buildCorePropsXml(props) {
  const fmt = (d) => d ? d.toISOString().replace(/\.\d+Z$/, 'Z') : new Date().toISOString().replace(/\.\d+Z$/, 'Z');
  return xmlDoc('cp:coreProperties', {
    'xmlns:cp': CORE_NS,
    'xmlns:dc': DC_NS,
    'xmlns:dcterms': DCTERMS,
    'xmlns:xsi': XSI,
  },
    (props.title       ? `<dc:title>${escapeXml(props.title)}</dc:title>`                                              : '') +
    (props.subject     ? `<dc:subject>${escapeXml(props.subject)}</dc:subject>`                                        : '') +
    (props.author      ? `<dc:creator>${escapeXml(props.author)}</dc:creator>`                                         : '') +
    (props.keywords    ? `<cp:keywords>${escapeXml(props.keywords)}</cp:keywords>`                                     : '') +
    (props.description ? `<dc:description>${escapeXml(props.description)}</dc:description>`                            : '') +
    `<cp:lastModifiedBy>${escapeXml(props.lastModifiedBy || 'CellWeave')}</cp:lastModifiedBy>`                        +
    `<dcterms:created xsi:type="dcterms:W3CDTF">${fmt(props.created)}</dcterms:created>`                             +
    `<dcterms:modified xsi:type="dcterms:W3CDTF">${fmt(props.modified || new Date())}</dcterms:modified>`            +
    (props.category    ? `<cp:category>${escapeXml(props.category)}</cp:category>`                                     : '')
  );
}

/**
 * Build docProps/app.xml
 * @param {import('../interfaces').IWorkbookProperties} props
 * @param {import('../models/Workbook').Workbook} wb
 * @returns {string}
 */
function buildAppPropsXml(props, wb) {
  const sheetsVt = wb.sheets.map(s => `<vt:lpstr>${escapeXml(s.name)}</vt:lpstr>`).join('');
  return xmlDoc('Properties', { xmlns: APP_NS, 'xmlns:vt': VT_NS },
    `<Application>CellWeave</Application>` +
    `<DocSecurity>0</DocSecurity>` +
    `<ScaleCrop>false</ScaleCrop>` +
    (props.company ? `<Company>${escapeXml(props.company)}</Company>` : '') +
    `<LinksUpToDate>false</LinksUpToDate>` +
    `<SharedDoc>false</SharedDoc>` +
    `<HyperlinksChanged>false</HyperlinksChanged>` +
    `<AppVersion>1.0.0</AppVersion>` +
    `<HeadingPairs><vt:vector size="2" baseType="variant"><vt:variant><vt:lpstr>Worksheets</vt:lpstr></vt:variant><vt:variant><vt:i4>${wb.sheetCount}</vt:i4></vt:variant></vt:vector></HeadingPairs>` +
    `<TitlesOfParts><vt:vector size="${wb.sheetCount}" baseType="lpstr">${sheetsVt}</vt:vector></TitlesOfParts>`
  );
}

/**
 * Minimal but valid OOXML theme XML.
 * Uses the standard OOXML DrawingML namespace (public open standard: ECMA-376).
 * @returns {string}
 */
function buildThemeXml() {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="CellWeave Theme">
  <a:themeElements>
    <a:clrScheme name="CellWeave">
      <a:dk1><a:sysClr lastClr="000000" val="windowText"/></a:dk1>
      <a:lt1><a:sysClr lastClr="FFFFFF" val="window"/></a:lt1>
      <a:dk2><a:srgbClr val="44546A"/></a:dk2>
      <a:lt2><a:srgbClr val="E7E6E6"/></a:lt2>
      <a:accent1><a:srgbClr val="4472C4"/></a:accent1>
      <a:accent2><a:srgbClr val="ED7D31"/></a:accent2>
      <a:accent3><a:srgbClr val="A9D18E"/></a:accent3>
      <a:accent4><a:srgbClr val="FFC000"/></a:accent4>
      <a:accent5><a:srgbClr val="5B9BD5"/></a:accent5>
      <a:accent6><a:srgbClr val="70AD47"/></a:accent6>
      <a:hlink><a:srgbClr val="0563C1"/></a:hlink>
      <a:folHlink><a:srgbClr val="954F72"/></a:folHlink>
    </a:clrScheme>
    <a:fontScheme name="CellWeave">
      <a:majorFont><a:latin typeface="Calibri Light"/></a:majorFont>
      <a:minorFont><a:latin typeface="Calibri"/></a:minorFont>
    </a:fontScheme>
    <a:fmtScheme name="CellWeave">
      <a:fillStyleLst>
        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
        <a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:lumMod val="110000"/><a:satMod val="105000"/><a:tint val="67000"/></a:schemeClr></a:gs><a:gs pos="50000"><a:schemeClr val="phClr"><a:lumMod val="105000"/><a:satMod val="103000"/><a:tint val="73000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:lumMod val="105000"/><a:satMod val="109000"/><a:tint val="81000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill>
        <a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:satMod val="103000"/><a:lumMod val="102000"/><a:tint val="94000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:lumMod val="99000"/><a:satMod val="120000"/><a:shade val="78000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill>
      </a:fillStyleLst>
      <a:lnStyleLst>
        <a:ln w="6350" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln>
        <a:ln w="12700" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln>
        <a:ln w="19050" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln>
      </a:lnStyleLst>
      <a:effectStyleLst>
        <a:effectStyle><a:effectLst/></a:effectStyle>
        <a:effectStyle><a:effectLst/></a:effectStyle>
        <a:effectStyle><a:effectLst><a:outerShdw blurRad="57150" dist="19050" dir="5400000" algn="ctr" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="63000"/></a:srgbClr></a:outerShdw></a:effectLst></a:effectStyle>
      </a:effectStyleLst>
      <a:bgFillStyleLst>
        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
        <a:solidFill><a:schemeClr val="phClr"><a:tint val="95000"/><a:satMod val="170000"/></a:schemeClr></a:solidFill>
        <a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="93000"/><a:satMod val="150000"/><a:shade val="98000"/><a:lumMod val="102000"/></a:schemeClr></a:gs><a:gs pos="50000"><a:schemeClr val="phClr"><a:tint val="98000"/><a:lumMod val="102000"/><a:satMod val="130000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="90000"/><a:lumMod val="103000"/><a:satMod val="120000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill>
      </a:bgFillStyleLst>
    </a:fmtScheme>
  </a:themeElements>
</a:theme>`;
}

module.exports = {
  buildWorkbookXml,
  buildWorkbookRelsXml,
  buildRootRelsXml,
  buildContentTypesXml,
  buildCorePropsXml,
  buildAppPropsXml,
  buildThemeXml,
  RT,
};
