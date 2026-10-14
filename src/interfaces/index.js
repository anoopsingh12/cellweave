/**
 * @fileoverview CellWeave – Interface Definitions (ISP: Interface Segregation Principle)
 * All contracts for the library are defined here as JSDoc typedefs.
 */

'use strict';

/**
 * @typedef {Object} ICellAddress
 * @property {number} row  - 0-based row index
 * @property {number} col  - 0-based column index
 */

/**
 * @typedef {Object} ICellRange
 * @property {number} startRow
 * @property {number} startCol
 * @property {number} endRow
 * @property {number} endCol
 */

/**
 * @typedef {'general'|'number'|'currency'|'accounting'|'date'|'time'|'datetime'|'percentage'|'fraction'|'scientific'|'text'|'special'|'custom'} NumberFormatType
 */

/**
 * @typedef {Object} INumberFormat
 * @property {NumberFormatType} type
 * @property {string} [formatCode]     - Raw Excel format code e.g. "#,##0.00"
 * @property {number} [decimalPlaces]
 * @property {boolean} [useThousandSeparator]
 * @property {string} [currencySymbol]
 * @property {string} [locale]
 */

/**
 * @typedef {'left'|'center'|'right'|'justify'|'fill'|'centerAcrossSelection'|'distributed'|'general'} HorizontalAlignment
 * @typedef {'top'|'center'|'bottom'|'justify'|'distributed'} VerticalAlignment
 */

/**
 * @typedef {Object} IAlignment
 * @property {HorizontalAlignment} [horizontal]
 * @property {VerticalAlignment}   [vertical]
 * @property {number}              [rotation]      - -90 to 90 degrees
 * @property {boolean}             [wrapText]
 * @property {boolean}             [shrinkToFit]
 * @property {number}              [indent]
 * @property {boolean}             [justifyLastLine]
 * @property {boolean}             [readingOrder]  - 0=Context, 1=LTR, 2=RTL
 */

/**
 * @typedef {'thin'|'medium'|'thick'|'hair'|'dotted'|'dashed'|'dashDot'|'dashDotDot'|'mediumDashed'|'mediumDashDot'|'mediumDashDotDot'|'slantDashDot'|'double'} BorderStyle
 */

/**
 * @typedef {Object} IBorderSide
 * @property {BorderStyle} [style]
 * @property {string}      [color]  - hex color e.g. "#000000"
 */

/**
 * @typedef {Object} IBorder
 * @property {IBorderSide} [top]
 * @property {IBorderSide} [bottom]
 * @property {IBorderSide} [left]
 * @property {IBorderSide} [right]
 * @property {IBorderSide} [diagonal]
 * @property {boolean}     [diagonalUp]
 * @property {boolean}     [diagonalDown]
 */

/**
 * @typedef {'none'|'solid'|'mediumGray'|'darkGray'|'lightGray'|'darkHorizontal'|'darkVertical'|'darkDown'|'darkUp'|'darkGrid'|'darkTrellis'|'lightHorizontal'|'lightVertical'|'lightDown'|'lightUp'|'lightGrid'|'lightTrellis'|'gray125'|'gray0625'} FillPattern
 */

/**
 * @typedef {Object} IFill
 * @property {FillPattern} [patternType]
 * @property {string}      [fgColor]    - foreground hex color
 * @property {string}      [bgColor]    - background hex color
 * @property {Object}      [gradient]   - gradient fill config
 * @property {'linear'|'path'} [gradient.type]
 * @property {number}      [gradient.degree]
 * @property {Array<{position:number,color:string}>} [gradient.stops]
 */

/**
 * @typedef {'normal'|'bold'} FontWeight
 * @typedef {'normal'|'italic'} FontStyle
 * @typedef {'none'|'single'|'double'|'singleAccounting'|'doubleAccounting'} FontUnderline
 * @typedef {'baseline'|'superscript'|'subscript'} FontVerticalAlign
 */

/**
 * @typedef {Object} IFont
 * @property {string}           [name]          - e.g. "Calibri"
 * @property {number}           [size]          - points
 * @property {string}           [color]         - hex
 * @property {FontWeight}       [bold]          - shorthand: pass true
 * @property {FontStyle}        [italic]        - shorthand: pass true
 * @property {FontUnderline}    [underline]
 * @property {boolean}          [strikethrough]
 * @property {FontVerticalAlign}[vertAlign]
 * @property {boolean}          [shadow]
 * @property {boolean}          [condense]
 * @property {boolean}          [extend]
 * @property {string}           [charset]
 * @property {string}           [family]        - 'Automatic'|'Roman'|'Swiss'|'Modern'|'Script'|'Decorative'
 * @property {string}           [scheme]        - 'major'|'minor'|'none'
 */

/**
 * @typedef {Object} ICellStyle
 * @property {IFont}         [font]
 * @property {IFill}         [fill]
 * @property {IBorder}       [border]
 * @property {IAlignment}    [alignment]
 * @property {INumberFormat} [numberFormat]
 * @property {boolean}       [locked]
 * @property {boolean}       [hidden]
 */

/**
 * @typedef {'b'|'n'|'s'|'str'|'inlineStr'|'e'|'d'|'z'} CellType
 * b=boolean, n=number, s=shared string, str=formula string, inlineStr, e=error, d=date, z=empty
 */

/**
 * @typedef {Object} ICellData
 * @property {*}            value
 * @property {CellType}     [type]
 * @property {string}       [formula]
 * @property {ICellStyle}   [style]
 * @property {string}       [comment]
 * @property {string}       [hyperlink]
 * @property {boolean}      [isArrayFormula]
 * @property {ICellRange}   [arrayFormulaRange]
 */

/**
 * @typedef {Object} IRowProperties
 * @property {number}  [height]      - points
 * @property {boolean} [hidden]
 * @property {boolean} [collapsed]
 * @property {number}  [outlineLevel]
 * @property {boolean} [customHeight]
 * @property {ICellStyle} [style]
 */

/**
 * @typedef {Object} IColumnProperties
 * @property {number}  [width]       - character units
 * @property {boolean} [hidden]
 * @property {boolean} [collapsed]
 * @property {number}  [outlineLevel]
 * @property {boolean} [bestFit]
 * @property {ICellStyle} [style]
 */

/**
 * @typedef {'none'|'ascending'|'descending'} SortOrder
 */

/**
 * @typedef {Object} ISortState
 * @property {ICellRange} range
 * @property {Array<{colIndex:number, order:SortOrder, caseSensitive:boolean}>} conditions
 */

/**
 * @typedef {Object} IAutoFilter
 * @property {ICellRange} range
 * @property {Array<{colIndex:number, filters:Array<string>}>} [columnFilters]
 */

/**
 * @typedef {Object} IConditionalFormattingRule
 * @property {'cellIs'|'expression'|'colorScale'|'dataBar'|'iconSet'|'top10'|'uniqueValues'|'duplicateValues'|'containsText'|'notContainsText'|'beginsWith'|'endsWith'|'containsBlanks'|'notContainsBlanks'|'containsErrors'|'notContainsErrors'|'timePeriod'|'aboveAverage'} type
 * @property {string}  [operator]
 * @property {Array<string>} [formula]
 * @property {ICellStyle} [style]
 * @property {Object}  [colorScale]
 * @property {Object}  [dataBar]
 * @property {Object}  [iconSet]
 * @property {number}  [priority]
 * @property {boolean} [stopIfTrue]
 * @property {string}  [text]
 */

/**
 * @typedef {Object} IConditionalFormatting
 * @property {ICellRange} range
 * @property {IConditionalFormattingRule[]} rules
 */

/**
 * @typedef {Object} IDataValidation
 * @property {ICellRange} range
 * @property {'whole'|'decimal'|'list'|'date'|'time'|'textLength'|'custom'} type
 * @property {'between'|'notBetween'|'equal'|'notEqual'|'greaterThan'|'lessThan'|'greaterThanOrEqual'|'lessThanOrEqual'} operator
 * @property {string}  [formula1]
 * @property {string}  [formula2]
 * @property {boolean} [allowBlank]
 * @property {boolean} [showDropDown]
 * @property {boolean} [showInputMessage]
 * @property {string}  [promptTitle]
 * @property {string}  [prompt]
 * @property {boolean} [showErrorAlert]
 * @property {'stop'|'warning'|'information'} [errorStyle]
 * @property {string}  [errorTitle]
 * @property {string}  [error]
 */

/**
 * @typedef {'bar'|'line'|'pie'|'doughnut'|'scatter'|'area'|'radar'|'bubble'|'stock'|'surface'|'column'|'combo'} ChartType
 */

/**
 * @typedef {Object} IChartSeries
 * @property {string}      name
 * @property {ICellRange}  [nameRange]
 * @property {ICellRange}  valuesRange
 * @property {ICellRange}  [categoriesRange]
 * @property {string}      [color]
 * @property {Object}      [marker]
 */

/**
 * @typedef {Object} IChartConfig
 * @property {ChartType}      type
 * @property {string}         title
 * @property {IChartSeries[]} series
 * @property {ICellRange}     [sourceRange]
 * @property {Object}         [position]    - {row, col, rowSpan, colSpan}
 * @property {Object}         [xAxis]
 * @property {Object}         [yAxis]
 * @property {Object}         [legend]
 * @property {boolean}        [showDataLabels]
 * @property {Object}         [style]
 */

/**
 * @typedef {Object} IPivotTable
 * @property {string}         name
 * @property {ICellRange}     sourceRange
 * @property {ICellAddress}   targetCell
 * @property {string[]}       rowFields
 * @property {string[]}       colFields
 * @property {string[]}       dataFields
 * @property {string[]}       [filterFields]
 */

/**
 * @typedef {Object} IImage
 * @property {string|Buffer} data     - base64 string or Buffer
 * @property {'jpeg'|'png'|'gif'|'bmp'|'tiff'|'wmf'|'emf'} type
 * @property {Object}        position  - {row, col, rowSpan, colSpan} or {x, y, width, height}
 * @property {number}        [rotation]
 * @property {number}        [opacity] - 0-1
 * @property {boolean}       [lockAspectRatio]
 */

/**
 * @typedef {Object} IPageSetup
 * @property {'portrait'|'landscape'} [orientation]
 * @property {'letter'|'legal'|'A4'|'A3'|'A5'|'B4'|'B5'|string} [paperSize]
 * @property {number}  [scale]            - 10-400 percent
 * @property {number}  [fitToPage]
 * @property {number}  [fitToWidth]
 * @property {number}  [fitToHeight]
 * @property {boolean} [blackAndWhite]
 * @property {boolean} [draft]
 * @property {'displayed'|'blank'|'NA'} [cellComments]
 * @property {string}  [errors]
 * @property {number}  [firstPageNumber]
 * @property {string}  [pageOrder]        - 'downThenOver'|'overThenDown'
 */

/**
 * @typedef {Object} IPageMargins
 * @property {number} [top]
 * @property {number} [bottom]
 * @property {number} [left]
 * @property {number} [right]
 * @property {number} [header]
 * @property {number} [footer]
 */

/**
 * @typedef {Object} IHeaderFooter
 * @property {string} [oddHeader]
 * @property {string} [oddFooter]
 * @property {string} [evenHeader]
 * @property {string} [evenFooter]
 * @property {string} [firstHeader]
 * @property {string} [firstFooter]
 * @property {boolean} [differentOddEven]
 * @property {boolean} [differentFirst]
 * @property {boolean} [scaleWithDoc]
 * @property {boolean} [alignWithMargins]
 */

/**
 * @typedef {Object} IFreezePanes
 * @property {number} [rows]   - number of rows to freeze
 * @property {number} [cols]   - number of cols to freeze
 * @property {number} [xSplit] - horizontal split point (col index)
 * @property {number} [ySplit] - vertical split point (row index)
 */

/**
 * @typedef {Object} ISheetView
 * @property {boolean}      [showGridLines]
 * @property {boolean}      [showRowColHeaders]
 * @property {boolean}      [showRuler]
 * @property {boolean}      [showFormulas]
 * @property {boolean}      [showWhiteSpace]
 * @property {boolean}      [showOutlineSymbols]
 * @property {boolean}      [showZeros]
 * @property {number}       [zoomScale]      - 10-400
 * @property {number}       [zoomScaleNormal]
 * @property {number}       [zoomScalePageLayout]
 * @property {'normal'|'pageBreakPreview'|'pageLayout'} [viewType]
 * @property {IFreezePanes} [freezePanes]
 * @property {ICellAddress} [topLeftCell]
 * @property {string}       [colorId]
 * @property {boolean}      [rightToLeft]
 * @property {boolean}      [tabSelected]
 */

/**
 * @typedef {Object} IProtection
 * @property {boolean} [sheet]
 * @property {string}  [password]
 * @property {boolean} [selectLockedCells]
 * @property {boolean} [selectUnlockedCells]
 * @property {boolean} [formatCells]
 * @property {boolean} [formatColumns]
 * @property {boolean} [formatRows]
 * @property {boolean} [insertColumns]
 * @property {boolean} [insertRows]
 * @property {boolean} [insertHyperlinks]
 * @property {boolean} [deleteColumns]
 * @property {boolean} [deleteRows]
 * @property {boolean} [sort]
 * @property {boolean} [autoFilter]
 * @property {boolean} [pivotTables]
 * @property {boolean} [objects]
 * @property {boolean} [scenarios]
 */

/**
 * @typedef {Object} IDefinedName
 * @property {string} name
 * @property {string} refers   - formula / cell reference
 * @property {string} [scope]  - sheet name or undefined for workbook scope
 * @property {string} [comment]
 */

/**
 * @typedef {Object} ITable
 * @property {string}     name
 * @property {ICellRange} range
 * @property {string}     [displayName]
 * @property {string}     [tableStyle]    - 'TableStyleLight1' etc.
 * @property {boolean}    [showFirstColumn]
 * @property {boolean}    [showLastColumn]
 * @property {boolean}    [showRowStripes]
 * @property {boolean}    [showColumnStripes]
 * @property {boolean}    [headerRow]
 * @property {boolean}    [totalsRow]
 * @property {Array<{name:string, formula:string, totalsRowFunction:string, style:ICellStyle}>} [columns]
 */

/**
 * @typedef {Object} ISparkline
 * @property {'line'|'column'|'winLoss'} type
 * @property {ICellRange} dataRange
 * @property {ICellRange} locationRange
 * @property {Object}     [style]
 * @property {boolean}    [showHighPoint]
 * @property {boolean}    [showLowPoint]
 * @property {boolean}    [showFirstPoint]
 * @property {boolean}    [showLastPoint]
 * @property {boolean}    [showNegativePoint]
 * @property {boolean}    [showMarkers]
 */

/**
 * @typedef {Object} IWorkbookProperties
 * @property {string}  [title]
 * @property {string}  [subject]
 * @property {string}  [author]
 * @property {string}  [manager]
 * @property {string}  [company]
 * @property {string}  [category]
 * @property {string}  [keywords]
 * @property {string}  [description]
 * @property {string}  [language]
 * @property {Date}    [created]
 * @property {Date}    [modified]
 * @property {string}  [lastModifiedBy]
 * @property {boolean} [date1904]
 */

/**
 * @typedef {Object} ISheetGrouping
 * @property {'row'|'col'} type
 * @property {number}      startIndex
 * @property {number}      endIndex
 * @property {boolean}     [collapsed]
 */

module.exports = {};
