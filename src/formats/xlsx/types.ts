export type XlsxInput = ArrayBuffer | Uint8Array | Blob;

export type XlsxErrorMode = 'strict' | 'tolerant';
export type XlsxDisplayTextMode = 'none' | 'supported';
export type XlsxImageMode = 'base64' | 'blob' | 'both' | 'none';
export type XlsxPivotCacheMode = 'metadata' | 'none' | 'records';

export type XlsxDiagnosticCode =
  | 'invalid-package'
  | 'invalid-document-structure'
  | 'invalid-document-value'
  | 'invalid-cell-reference'
  | 'invalid-formula'
  | 'invalid-selection'
  | 'invalid-relationship-target'
  | 'missing-required-part'
  | 'resource-limit-exceeded'
  | 'security-rejected-content'
  | 'unsupported-feature'
  | 'xml-parse-failed'
  | 'xml-read-failed';

export interface XlsxDiagnostic {
  actual?: number;
  cell?: string;
  code: XlsxDiagnosticCode;
  limit?: number;
  limitName?: keyof XlsxResourceLimits;
  message: string;
  part?: string;
  range?: string;
  relationshipType?: string;
  severity: 'error' | 'warning';
  sheet?: string;
}

export interface XlsxSelection {
  ranges?: Readonly<Record<string, readonly string[]>>;
  sheetNames?: readonly string[];
}

export interface XlsxResourceLimits {
  maxCharts?: number;
  maxColumnsPerWorksheet?: number;
  maxComments?: number;
  maxConditionalFormattingRules?: number;
  maxDefinedNames?: number;
  maxDrawings?: number;
  maxEntries?: number;
  maxFormulaCharacters?: number;
  maxFormulaGroups?: number;
  maxHyperlinks?: number;
  maxInputBytes?: number;
  maxMediaBytes?: number;
  maxMergedRanges?: number;
  maxPartBytes?: number;
  maxPivotRecords?: number;
  maxRangeAreas?: number;
  maxRelationships?: number;
  maxReturnedCells?: number;
  maxRichTextRuns?: number;
  maxRowsPerWorksheet?: number;
  maxScannedCells?: number;
  maxSharedStrings?: number;
  maxStyles?: number;
  maxTables?: number;
  maxTextCharacters?: number;
  maxTotalFormulaCharacters?: number;
  maxTotalUncompressedBytes?: number;
  maxTotalXmlNodes?: number;
  maxValidationRules?: number;
  maxWorksheets?: number;
  maxXmlBytes?: number;
  maxXmlDepth?: number;
  maxXmlNodes?: number;
}

export interface XlsxParseOptions {
  displayTextMode?: XlsxDisplayTextMode;
  errorMode?: XlsxErrorMode;
  imageMode?: XlsxImageMode;
  limits?: XlsxResourceLimits;
  pivotCacheMode?: XlsxPivotCacheMode;
  selection?: XlsxSelection;
}

export interface XlsxParseResult {
  diagnostics: XlsxDiagnostic[];
  document: XlsxDocument;
}

export interface XlsxRange {
  end: { column: number; row: number };
  reference: string;
  start: { column: number; row: number };
}

export interface XlsxRichTextRun {
  text: string;
}

export type XlsxCellValue =
  | { kind: 'text'; runs?: XlsxRichTextRun[]; text: string }
  | { kind: 'number'; value: number }
  | { kind: 'boolean'; value: boolean }
  | { code: string; kind: 'error' }
  | {
      kind: 'date';
      normalized: string | null;
      precision: 'date' | 'date-time' | 'duration' | 'time';
      source:
        | { dateSystem: '1900' | '1904'; kind: 'serial'; value: number }
        | { kind: 'iso'; value: string };
    };

export interface XlsxFormula {
  expression: string;
  kind: 'array' | 'data-table' | 'dynamic-array' | 'normal';
  range?: XlsxRange;
}

export interface XlsxCellBase {
  address: string;
  column: number;
  displayText?: string;
  style?: number;
}

export type XlsxCell = XlsxCellBase &
  (
    | { content: { kind: 'blank' } }
    | { content: { kind: 'value'; value: XlsxCellValue } }
    | {
        content: {
          cached: XlsxCellValue | { kind: 'missing' };
          formula: XlsxFormula;
          kind: 'formula';
        };
      }
  );

export interface XlsxRow {
  cells: XlsxCell[];
  collapsed?: boolean;
  height?: number;
  hidden?: boolean;
  index: number;
  outlineLevel?: number;
  style?: number;
}

export interface XlsxColumnRange {
  collapsed?: boolean;
  end: number;
  hidden?: boolean;
  outlineLevel?: number;
  start: number;
  style?: number;
  width?: number;
}

export type XlsxPanePosition =
  'bottom-left' | 'bottom-right' | 'top-left' | 'top-right';

export interface XlsxWorksheetPane {
  activePane: XlsxPanePosition;
  state: 'frozen' | 'frozen-split' | 'split';
  topLeftCell?: string;
  xSplit: number;
  ySplit: number;
}

export interface XlsxWorksheetViewSelection {
  activeCell?: string;
  activeCellId?: number;
  pane: XlsxPanePosition;
  ranges: XlsxRange[];
}

export interface XlsxWorksheetView {
  kind: 'normal' | 'page-break-preview' | 'page-layout';
  pane?: XlsxWorksheetPane;
  rightToLeft: boolean;
  selections: XlsxWorksheetViewSelection[];
  showGridLines: boolean;
  showRowColumnHeaders: boolean;
  tabSelected: boolean;
  topLeftCell?: string;
  workbookViewId: number;
  zoomScale: number;
  zoomScaleNormal?: number;
  zoomScalePageLayout?: number;
  zoomScaleSheetLayout?: number;
}

export interface XlsxWorksheetFormat {
  baseColumnWidth: number;
  customHeight: boolean;
  defaultColumnWidth?: number;
  defaultRowHeight: number;
  outlineColumnLevel: number;
  outlineRowLevel: number;
  thickBottom: boolean;
  thickTop: boolean;
  zeroHeight: boolean;
}

export interface XlsxWorksheetOutline {
  applyStyles: boolean;
  showOutlineSymbols: boolean;
  summaryBelow: boolean;
  summaryRight: boolean;
}

export type XlsxColor =
  | { argb: string; kind: 'rgb'; tint?: number }
  | { index: number; kind: 'theme'; tint?: number }
  | { index: number; kind: 'indexed'; tint?: number }
  | { kind: 'automatic' };

export interface XlsxFont {
  bold?: boolean;
  charset?: number;
  color?: XlsxColor;
  condense?: boolean;
  extend?: boolean;
  family?: number;
  italic?: boolean;
  name?: string;
  outline?: boolean;
  scheme?: 'major' | 'minor';
  shadow?: boolean;
  size?: number;
  strike?: boolean;
  underline?: 'double' | 'double-accounting' | 'single' | 'single-accounting';
  verticalAlignment?: 'subscript' | 'superscript';
}

export type XlsxPatternType =
  | 'darkDown'
  | 'darkGray'
  | 'darkGrid'
  | 'darkHorizontal'
  | 'darkTrellis'
  | 'darkUp'
  | 'darkVertical'
  | 'gray0625'
  | 'gray125'
  | 'lightDown'
  | 'lightGray'
  | 'lightGrid'
  | 'lightHorizontal'
  | 'lightTrellis'
  | 'lightUp'
  | 'lightVertical'
  | 'mediumGray'
  | 'none'
  | 'solid';

export interface XlsxGradientStop {
  color: XlsxColor;
  position: number;
}

export type XlsxFill =
  | {
      backgroundColor?: XlsxColor;
      foregroundColor?: XlsxColor;
      kind: 'pattern';
      pattern: XlsxPatternType;
    }
  | {
      angle?: number;
      bottom?: number;
      kind: 'gradient';
      left?: number;
      right?: number;
      stops: XlsxGradientStop[];
      top?: number;
      type: 'linear' | 'path';
    };

export type XlsxBorderStyle =
  | 'dashDot'
  | 'dashDotDot'
  | 'dashed'
  | 'dotted'
  | 'double'
  | 'hair'
  | 'medium'
  | 'mediumDashDot'
  | 'mediumDashDotDot'
  | 'mediumDashed'
  | 'slantDashDot'
  | 'thick'
  | 'thin';

export interface XlsxBorderSide {
  color?: XlsxColor;
  style?: XlsxBorderStyle;
}

export interface XlsxBorder {
  bottom?: XlsxBorderSide;
  diagonal?: XlsxBorderSide;
  diagonalDown?: boolean;
  diagonalUp?: boolean;
  end?: XlsxBorderSide;
  horizontal?: XlsxBorderSide;
  left?: XlsxBorderSide;
  outline?: boolean;
  right?: XlsxBorderSide;
  start?: XlsxBorderSide;
  top?: XlsxBorderSide;
  vertical?: XlsxBorderSide;
}

export interface XlsxAlignment {
  horizontal?:
    | 'center'
    | 'centerContinuous'
    | 'distributed'
    | 'fill'
    | 'justify'
    | 'left'
    | 'right';
  indent?: number;
  justifyLastLine?: boolean;
  readingOrder?: 'left-to-right' | 'right-to-left';
  relativeIndent?: number;
  shrinkToFit?: boolean;
  textRotation?: number;
  vertical?: 'center' | 'distributed' | 'justify' | 'top';
  wrapText?: boolean;
}

export interface XlsxProtection {
  hidden?: boolean;
  locked?: boolean;
}

export interface XlsxStyle {
  alignment?: XlsxAlignment;
  border?: XlsxBorder;
  fill?: XlsxFill;
  font?: XlsxFont;
  numberFormat?: string;
  protection?: XlsxProtection;
}

export interface XlsxNamedStyle {
  builtinId?: number;
  customBuiltin?: boolean;
  hidden?: boolean;
  name: string;
  outlineLevel?: number;
  style: XlsxStyle;
}

export interface XlsxDefinedName {
  comment?: string;
  customMenu?: string;
  description?: string;
  expression: string;
  function?: boolean;
  functionGroupId?: number;
  help?: string;
  hidden: boolean;
  name: string;
  publishToServer?: boolean;
  sheetIndex?: number;
  shortcutKey?: string;
  statusBar?: string;
  vbProcedure?: boolean;
  workbookParameter?: boolean;
  xlm?: boolean;
}

export interface XlsxStrongProtectionHash {
  algorithmName: string;
  hashValue: string;
  saltValue: string;
  spinCount: number;
}

export interface XlsxProtectionCredential {
  legacyHash?: string;
  strongHash?: XlsxStrongProtectionHash;
}

export interface XlsxWorkbookProtection {
  lockRevisions: boolean;
  lockStructure: boolean;
  lockWindows: boolean;
  revisionsCredential?: XlsxProtectionCredential;
  workbookCredential?: XlsxProtectionCredential;
}

export interface XlsxWorkbookProperties {
  calculation: {
    forceFullCalculation: boolean;
    fullCalculationOnLoad: boolean;
    mode: 'automatic' | 'automatic-except-tables' | 'manual';
  };
  commentPersons?: XlsxCommentPerson[];
  dateSystem: '1900' | '1904';
  definedNames: XlsxDefinedName[];
  protection?: XlsxWorkbookProtection;
  views: XlsxWorkbookView[];
}

export interface XlsxWorkbookView {
  activeSheetIndex: number;
  autoFilterDateGrouping: boolean;
  firstVisibleSheetIndex: number;
  minimized: boolean;
  showHorizontalScroll: boolean;
  showSheetTabs: boolean;
  showVerticalScroll: boolean;
  tabRatio: number;
  visibility: 'hidden' | 'very-hidden' | 'visible';
  windowHeight?: number;
  windowWidth?: number;
  xWindow?: number;
  yWindow?: number;
}

export type XlsxTableTotalsFunction =
  | 'average'
  | 'count'
  | 'count-numbers'
  | 'custom'
  | 'maximum'
  | 'minimum'
  | 'none'
  | 'standard-deviation'
  | 'sum'
  | 'variance';

export interface XlsxTableFormula {
  array: boolean;
  expression: string;
}

export interface XlsxTableColumn {
  calculatedFormula?: XlsxTableFormula;
  dataCellStyle?: string;
  dataDifferentialStyle?: number;
  headerCellStyle?: string;
  headerDifferentialStyle?: number;
  id: number;
  name: string;
  queryTableFieldId?: number;
  totalsCellStyle?: string;
  totalsDifferentialStyle?: number;
  totalsFormula?: XlsxTableFormula;
  totalsLabel?: string;
  totalsFunction: XlsxTableTotalsFunction;
  uniqueName?: string;
}

export interface XlsxTableStyleInfo {
  name?: string;
  showColumnStripes: boolean;
  showFirstColumn: boolean;
  showLastColumn: boolean;
  showRowStripes: boolean;
}

export interface XlsxDateGroupFilter {
  day?: number;
  grouping: 'day' | 'hour' | 'minute' | 'month' | 'second' | 'year';
  hour?: number;
  minute?: number;
  month?: number;
  second?: number;
  year: number;
}

export type XlsxCalendarType =
  | 'gregorian'
  | 'gregorianArabic'
  | 'gregorianMeFrench'
  | 'gregorianUs'
  | 'gregorianXlitEnglish'
  | 'gregorianXlitFrench'
  | 'hebrew'
  | 'hijri'
  | 'japan'
  | 'korea'
  | 'none'
  | 'saka'
  | 'taiwan'
  | 'thai';

export type XlsxDynamicFilterType =
  | 'aboveAverage'
  | 'belowAverage'
  | 'lastMonth'
  | 'lastQuarter'
  | 'lastWeek'
  | 'lastYear'
  | 'M1'
  | 'M2'
  | 'M3'
  | 'M4'
  | 'M5'
  | 'M6'
  | 'M7'
  | 'M8'
  | 'M9'
  | 'M10'
  | 'M11'
  | 'M12'
  | 'nextMonth'
  | 'nextQuarter'
  | 'nextWeek'
  | 'nextYear'
  | 'null'
  | 'Q1'
  | 'Q2'
  | 'Q3'
  | 'Q4'
  | 'thisMonth'
  | 'thisQuarter'
  | 'thisWeek'
  | 'thisYear'
  | 'today'
  | 'tomorrow'
  | 'yearToDate'
  | 'yesterday';

export type XlsxIconSet =
  | '3Arrows'
  | '3ArrowsGray'
  | '3Flags'
  | '3Signs'
  | '3Symbols'
  | '3Symbols2'
  | '3TrafficLights1'
  | '3TrafficLights2'
  | '4Arrows'
  | '4ArrowsGray'
  | '4Rating'
  | '4RedToBlack'
  | '4TrafficLights'
  | '5Arrows'
  | '5ArrowsGray'
  | '5Quarters'
  | '5Rating';

export type XlsxFilterRule =
  | {
      blank: boolean;
      calendarType?: XlsxCalendarType;
      dates: XlsxDateGroupFilter[];
      kind: 'values';
      values: string[];
    }
  | {
      and: boolean;
      conditions: Array<{
        operator:
          | 'equal'
          | 'greater-than'
          | 'greater-than-or-equal'
          | 'less-than'
          | 'less-than-or-equal'
          | 'not-equal';
        value: string;
      }>;
      kind: 'custom';
    }
  | {
      kind: 'dynamic';
      maxValue?: number;
      type: XlsxDynamicFilterType;
      value?: number;
    }
  | {
      filterValue?: number;
      kind: 'top';
      percent: boolean;
      top: boolean;
      value: number;
    }
  | {
      cellColor: boolean;
      differentialStyle?: number;
      kind: 'color';
    }
  | {
      iconId: number;
      iconSet: XlsxIconSet;
      kind: 'icon';
    }
  | { kind: 'none' };

export interface XlsxFilterColumn {
  columnId: number;
  hiddenButton: boolean;
  rule: XlsxFilterRule;
  showButton: boolean;
}

export interface XlsxSortCondition {
  customList?: string;
  descending: boolean;
  differentialStyle?: number;
  iconId?: number;
  iconSet?: XlsxIconSet;
  range: XlsxRange;
  sortBy: 'cell-color' | 'font-color' | 'icon' | 'value';
}

export interface XlsxSortState {
  caseSensitive: boolean;
  columnSort: boolean;
  conditions: XlsxSortCondition[];
  range: XlsxRange;
  sortMethod: 'none' | 'pin-yin' | 'stroke';
}

export interface XlsxAutoFilter {
  columns: XlsxFilterColumn[];
  range: XlsxRange;
  selectionRelation: 'full-sheet' | 'intersects-selection';
  sort?: XlsxSortState;
}

export type XlsxDataValidationType =
  | 'custom'
  | 'date'
  | 'decimal'
  | 'list'
  | 'none'
  | 'text-length'
  | 'time'
  | 'whole';

export type XlsxDataValidationOperator =
  | 'between'
  | 'equal'
  | 'greater-than'
  | 'greater-than-or-equal'
  | 'less-than'
  | 'less-than-or-equal'
  | 'not-between'
  | 'not-equal';

export type XlsxDataValidationImeMode =
  | 'disabled'
  | 'full-alpha'
  | 'full-hangul'
  | 'full-katakana'
  | 'half-alpha'
  | 'half-hangul'
  | 'half-katakana'
  | 'hiragana'
  | 'no-control'
  | 'off'
  | 'on';

export interface XlsxDataValidation {
  allowBlank: boolean;
  error?: string;
  errorStyle: 'information' | 'stop' | 'warning';
  errorTitle?: string;
  formula1?: string;
  formula2?: string;
  imeMode: XlsxDataValidationImeMode;
  operator: XlsxDataValidationOperator;
  prompt?: string;
  promptTitle?: string;
  ranges: XlsxRange[];
  selectionRelation: 'full-sheet' | 'intersects-selection';
  showDropDown: boolean;
  showErrorMessage: boolean;
  showInputMessage: boolean;
  type: XlsxDataValidationType;
}

export interface XlsxDataValidationSettings {
  disablePrompts: boolean;
  xWindow?: number;
  yWindow?: number;
}

export type XlsxConditionalFormattingRuleType =
  | 'above-average'
  | 'begins-with'
  | 'cell-is'
  | 'color-scale'
  | 'contains-blanks'
  | 'contains-errors'
  | 'contains-text'
  | 'data-bar'
  | 'duplicate-values'
  | 'ends-with'
  | 'expression'
  | 'icon-set'
  | 'not-contains-blanks'
  | 'not-contains-errors'
  | 'not-contains-text'
  | 'time-period'
  | 'top'
  | 'unique-values';

export type XlsxConditionalFormattingOperator =
  | 'between'
  | 'equal'
  | 'greater-than'
  | 'greater-than-or-equal'
  | 'less-than'
  | 'less-than-or-equal'
  | 'not-between'
  | 'not-equal';

export type XlsxConditionalValueObject =
  | { greaterThanOrEqual: boolean; kind: 'maximum' | 'minimum' }
  | {
      greaterThanOrEqual: boolean;
      kind: 'number' | 'percent' | 'percentile';
      value: number;
    }
  | {
      expression: string;
      greaterThanOrEqual: boolean;
      kind: 'formula';
    };

export interface XlsxConditionalColorScale {
  stops: Array<{
    color: XlsxColor;
    threshold: XlsxConditionalValueObject;
  }>;
}

export interface XlsxConditionalDataBar {
  color: XlsxColor;
  maximumLength: number;
  minimumLength: number;
  showValue: boolean;
  thresholds: [XlsxConditionalValueObject, XlsxConditionalValueObject];
}

export interface XlsxConditionalIconSet {
  iconSet: XlsxIconSet;
  percent: boolean;
  reverse: boolean;
  showValue: boolean;
  thresholds: XlsxConditionalValueObject[];
}

export interface XlsxConditionalFormattingRule {
  aboveAverage?: boolean;
  bottom?: boolean;
  colorScale?: XlsxConditionalColorScale;
  dataBar?: XlsxConditionalDataBar;
  differentialStyle?: number;
  equalAverage?: boolean;
  formulas: string[];
  iconSet?: XlsxConditionalIconSet;
  operator?: XlsxConditionalFormattingOperator;
  percent?: boolean;
  priority: number;
  rank?: number;
  standardDeviations?: number;
  stopIfTrue: boolean;
  text?: string;
  timePeriod?:
    | 'last-7-days'
    | 'last-month'
    | 'last-week'
    | 'next-month'
    | 'next-week'
    | 'this-month'
    | 'this-week'
    | 'today'
    | 'tomorrow'
    | 'yesterday';
  type: XlsxConditionalFormattingRuleType;
}

export interface XlsxConditionalFormatting {
  pivot: boolean;
  ranges: XlsxRange[];
  rules: XlsxConditionalFormattingRule[];
  selectionRelation: 'full-sheet' | 'intersects-selection';
}

export interface XlsxProtectedRange {
  credential?: XlsxProtectionCredential;
  name: string;
  ranges: XlsxRange[];
  securityDescriptor?: string;
  selectionRelation: 'full-sheet' | 'intersects-selection';
}

export interface XlsxWorksheetProtection {
  credential?: XlsxProtectionCredential;
  protectAutoFilter: boolean;
  protectDeleteColumns: boolean;
  protectDeleteRows: boolean;
  protectFormatCells: boolean;
  protectFormatColumns: boolean;
  protectFormatRows: boolean;
  protectInsertColumns: boolean;
  protectInsertHyperlinks: boolean;
  protectInsertRows: boolean;
  protectObjects: boolean;
  protectPivotTables: boolean;
  protectScenarios: boolean;
  protectSelectLockedCells: boolean;
  protectSelectUnlockedCells: boolean;
  protectSheet: boolean;
  protectSort: boolean;
}

export interface XlsxUniversalMeasure {
  unit: 'cm' | 'in' | 'mm' | 'pc' | 'pi' | 'pt';
  value: number;
}

export interface XlsxPageMargins {
  bottom: number;
  footer: number;
  header: number;
  left: number;
  right: number;
  top: number;
}

export interface XlsxPrintOptions {
  gridLines: boolean;
  gridLinesSet: boolean;
  headings: boolean;
  horizontalCentered: boolean;
  verticalCentered: boolean;
}

export interface XlsxPageSetupProperties {
  autoPageBreaks: boolean;
  fitToPage: boolean;
}

export interface XlsxPageSetup {
  blackAndWhite: boolean;
  cellComments: 'as-displayed' | 'at-end' | 'none';
  copies?: number;
  draft: boolean;
  errors: 'blank' | 'dash' | 'displayed' | 'not-available';
  firstPageNumber?: number;
  fitToHeight?: number;
  fitToWidth?: number;
  horizontalDpi?: number;
  orientation: 'default' | 'landscape' | 'portrait';
  pageOrder: 'down-then-over' | 'over-then-down';
  paperHeight?: XlsxUniversalMeasure;
  paperSize?: number;
  paperWidth?: XlsxUniversalMeasure;
  scale?: number;
  useFirstPageNumber: boolean;
  usePrinterDefaults: boolean;
  verticalDpi?: number;
}

export interface XlsxHeaderFooter {
  alignWithMargins: boolean;
  differentFirst: boolean;
  differentOddEven: boolean;
  evenFooter?: string;
  evenHeader?: string;
  firstFooter?: string;
  firstHeader?: string;
  oddFooter?: string;
  oddHeader?: string;
  scaleWithDocument: boolean;
}

export interface XlsxPageBreak {
  end: number;
  manual: boolean;
  pivot: boolean;
  position: number;
  start: number;
}

export interface XlsxWorksheetPrintSettings {
  columnBreaks?: XlsxPageBreak[];
  headerFooter?: XlsxHeaderFooter;
  margins?: XlsxPageMargins;
  options?: XlsxPrintOptions;
  pageSetup?: XlsxPageSetup;
  properties?: XlsxPageSetupProperties;
  rowBreaks?: XlsxPageBreak[];
}

export interface XlsxTable {
  autoFilter?: XlsxAutoFilter;
  columns: XlsxTableColumn[];
  comment?: string;
  connectionId?: number;
  dataCellStyle?: string;
  dataDifferentialStyle?: number;
  displayName: string;
  headerCellStyle?: string;
  headerDifferentialStyle?: number;
  headerRow: boolean;
  headerRowBorderDifferentialStyle?: number;
  id: number;
  insertRow: boolean;
  insertRowShift: boolean;
  name: string;
  published: boolean;
  range: XlsxRange;
  selectionRelation: 'full-sheet' | 'intersects-selection';
  style?: XlsxTableStyleInfo;
  tableBorderDifferentialStyle?: number;
  tableType: 'query-table' | 'worksheet' | 'xml';
  totalsCellStyle?: string;
  totalsDifferentialStyle?: number;
  totalsRow: boolean;
  totalsRowBorderDifferentialStyle?: number;
  totalsRowShown: boolean;
}

export interface XlsxDrawingMarker {
  column: number;
  columnOffset: number;
  row: number;
  rowOffset: number;
}

export interface XlsxDrawingExtent {
  height: number;
  width: number;
}

export interface XlsxDrawingTransform {
  flipHorizontal: boolean;
  flipVertical: boolean;
  rotation: number;
}

export interface XlsxDrawingObjectTransform extends XlsxDrawingTransform {
  height: number;
  width: number;
  x: number;
  y: number;
}

export type XlsxDrawingColor =
  | { kind: 'rgb'; value: string }
  | { kind: 'scheme'; value: string }
  | { kind: 'system'; lastColor?: string; value: string };

export type XlsxDrawingFill =
  { kind: 'none' } | { color: XlsxDrawingColor; kind: 'solid' };

export interface XlsxDrawingLine {
  dash?: string;
  fill?: XlsxDrawingFill;
  width?: number;
}

export type XlsxDrawingGeometry =
  { kind: 'custom' } | { kind: 'preset'; preset: string };

interface XlsxDrawingShapeBase {
  description?: string;
  fill?: XlsxDrawingFill;
  geometry: XlsxDrawingGeometry;
  hidden: boolean;
  id: number;
  line?: XlsxDrawingLine;
  name: string;
  text?: string;
  transform: XlsxDrawingObjectTransform;
}

export interface XlsxDrawingShape extends XlsxDrawingShapeBase {
  kind: 'shape';
}

export interface XlsxDrawingConnector extends XlsxDrawingShapeBase {
  endConnection?: { shapeId: number; site: number };
  kind: 'connector';
  startConnection?: { shapeId: number; site: number };
}

export interface XlsxDrawingGroup {
  children: XlsxDrawingObject[];
  description?: string;
  hidden: boolean;
  id: number;
  kind: 'group';
  name: string;
  transform: XlsxDrawingObjectTransform & {
    childHeight: number;
    childWidth: number;
    childX: number;
    childY: number;
  };
}

export interface XlsxImageCrop {
  bottom: number;
  left: number;
  right: number;
  top: number;
}

export interface XlsxEmbeddedImage {
  base64?: string;
  blobUrl?: string;
  byteLength?: number;
  contentType: string;
  crop: XlsxImageCrop;
  description?: string;
  hidden: boolean;
  id: number;
  kind: 'image';
  name: string;
  part: string;
  transform: XlsxDrawingTransform;
}

export type XlsxChartType =
  | 'area'
  | 'area-3d'
  | 'bar'
  | 'bar-3d'
  | 'bubble'
  | 'doughnut'
  | 'line'
  | 'line-3d'
  | 'of-pie'
  | 'pie'
  | 'pie-3d'
  | 'radar'
  | 'scatter'
  | 'stock'
  | 'surface'
  | 'surface-3d';

export interface XlsxChartNumericPoint {
  index: number;
  value: number;
}

export interface XlsxChartStringPoint {
  index: number;
  value: string;
}

export type XlsxChartDataSource =
  | {
      formatCode?: string;
      formula?: string;
      kind: 'number';
      pointCount?: number;
      points: XlsxChartNumericPoint[];
    }
  | {
      formula?: string;
      kind: 'string';
      pointCount?: number;
      points: XlsxChartStringPoint[];
    }
  | {
      formula?: string;
      kind: 'multi-level-string';
      levels: XlsxChartStringPoint[][];
      pointCount?: number;
    };

export interface XlsxChartText {
  formula?: string;
  text: string;
}

export interface XlsxChartMarker {
  size?: number;
  symbol?: string;
}

export interface XlsxChartDataLabels {
  position?: string;
  separator?: string;
  showBubbleSize: boolean;
  showCategoryName: boolean;
  showLegendKey: boolean;
  showPercent: boolean;
  showSeriesName: boolean;
  showValue: boolean;
}

export interface XlsxChartSeries {
  bubbleSizes?: XlsxChartDataSource;
  categories?: XlsxChartDataSource;
  color?: XlsxDrawingColor;
  dataLabels?: XlsxChartDataLabels;
  index: number;
  marker?: XlsxChartMarker;
  name?: XlsxChartText;
  order: number;
  smooth?: boolean;
  values?: XlsxChartDataSource;
  xValues?: XlsxChartDataSource;
  yValues?: XlsxChartDataSource;
}

export interface XlsxChartPlot {
  axisIds: number[];
  barDirection?: 'bar' | 'column';
  bubbleScale?: number;
  dataLabels?: XlsxChartDataLabels;
  firstSliceAngle?: number;
  gapDepth?: number;
  gapWidth?: number;
  grouping?: string;
  holeSize?: number;
  overlap?: number;
  radarStyle?: string;
  scatterStyle?: string;
  series: XlsxChartSeries[];
  type: XlsxChartType;
  varyColors: boolean;
}

export interface XlsxChartAxis {
  crossAxis?: number;
  crosses?: string;
  crossesAt?: number;
  deleted: boolean;
  id: number;
  kind: 'category' | 'date' | 'series' | 'value';
  logBase?: number;
  majorGridlines: boolean;
  majorUnit?: number;
  maximum?: number;
  minimum?: number;
  minorGridlines: boolean;
  minorUnit?: number;
  numberFormat?: { code: string; sourceLinked: boolean };
  orientation: 'max-min' | 'min-max';
  position?: 'bottom' | 'left' | 'right' | 'top';
  title?: XlsxChartText;
}

export interface XlsxChartLegendEntry {
  deleted: boolean;
  index: number;
}

export interface XlsxChartLegend {
  entries: XlsxChartLegendEntry[];
  overlay: boolean;
  position?: 'bottom' | 'left' | 'right' | 'top' | 'top-right';
}

export interface XlsxChart {
  axes: XlsxChartAxis[];
  autoTitleDeleted: boolean;
  description?: string;
  displayBlanksAs: 'gap' | 'span' | 'zero';
  hidden: boolean;
  id: number;
  kind: 'chart';
  legend?: XlsxChartLegend;
  name: string;
  part: string;
  plots: XlsxChartPlot[];
  plotVisibleOnly: boolean;
  roundedCorners: boolean;
  showDataLabelsOverMaximum: boolean;
  style?: number;
  title?: XlsxChartText;
  transform: XlsxDrawingObjectTransform;
}

export type XlsxDrawingObject =
  | XlsxChart
  | XlsxDrawingConnector
  | XlsxDrawingGroup
  | XlsxDrawingShape
  | XlsxEmbeddedImage;

export interface XlsxDrawing {
  editAs?: 'absolute' | 'one-cell' | 'two-cell';
  extent: XlsxDrawingExtent;
  from?: XlsxDrawingMarker;
  kind: 'absolute' | 'one-cell' | 'two-cell';
  object: XlsxDrawingObject;
  position?: { x: number; y: number };
  selectionRelation: 'full-sheet' | 'intersects-selection' | 'worksheet-global';
  to?: XlsxDrawingMarker;
}

export type XlsxHyperlinkTarget =
  | { kind: 'internal'; location: string }
  | { kind: 'external'; location?: string; url: string };

export interface XlsxHyperlink {
  display?: string;
  range: XlsxRange;
  selectionRelation: 'full-sheet' | 'intersects-selection';
  target: XlsxHyperlinkTarget;
  tooltip?: string;
}

export interface XlsxCommentPerson {
  displayName: string;
  id: string;
  providerId?: string;
  userId?: string;
}

interface XlsxCommentBase {
  reference: string;
  selectionRelation: 'full-sheet' | 'intersects-selection';
  text: string;
}

export interface XlsxLegacyComment extends XlsxCommentBase {
  author: string;
  kind: 'note';
  visible: boolean;
}

export interface XlsxThreadedComment extends XlsxCommentBase {
  id: string;
  kind: 'threaded';
  parentId?: string;
  personId: string;
  timestamp: string;
}

export type XlsxComment = XlsxLegacyComment | XlsxThreadedComment;

export interface XlsxSparkline {
  dataFormula: string;
  location: string;
  selectionRelation: 'full-sheet' | 'intersects-selection';
}

export interface XlsxSparklineColors {
  axis?: XlsxColor;
  first?: XlsxColor;
  high?: XlsxColor;
  last?: XlsxColor;
  low?: XlsxColor;
  markers?: XlsxColor;
  negative?: XlsxColor;
  series?: XlsxColor;
}

export interface XlsxSparklineGroup {
  colors: XlsxSparklineColors;
  dateAxis: boolean;
  displayEmptyCellsAs: 'gap' | 'span' | 'zero';
  displayHidden: boolean;
  displayXAxis: boolean;
  first: boolean;
  high: boolean;
  last: boolean;
  lineWeight?: number;
  low: boolean;
  manualMaximum?: number;
  manualMinimum?: number;
  markers: boolean;
  maximumAxisType: 'custom' | 'group' | 'individual';
  minimumAxisType: 'custom' | 'group' | 'individual';
  negative: boolean;
  rightToLeft: boolean;
  sparklines: XlsxSparkline[];
  type: 'column' | 'line' | 'stacked';
}

export interface XlsxSheetBase {
  index: number;
  name: string;
  payload: 'full-sheet' | 'not-selected' | 'selected-ranges';
  state: 'hidden' | 'very-hidden' | 'visible';
}

export interface XlsxWorksheet extends XlsxSheetBase {
  autoFilter?: XlsxAutoFilter;
  columns: XlsxColumnRange[];
  comments: XlsxComment[];
  conditionalFormattings: XlsxConditionalFormatting[];
  declaredDimension?: XlsxRange;
  dataValidationSettings?: XlsxDataValidationSettings;
  dataValidations: XlsxDataValidation[];
  drawings: XlsxDrawing[];
  hyperlinks: XlsxHyperlink[];
  kind: 'worksheet';
  mergedRanges: XlsxRange[];
  outline?: XlsxWorksheetOutline;
  print?: XlsxWorksheetPrintSettings;
  protectedRanges: XlsxProtectedRange[];
  protection?: XlsxWorksheetProtection;
  rows: XlsxRow[];
  sheetFormat?: XlsxWorksheetFormat;
  sparklineGroups?: XlsxSparklineGroup[];
  tabColor?: XlsxColor;
  tables: XlsxTable[];
  views: XlsxWorksheetView[];
}

export interface XlsxChartSheet extends XlsxSheetBase {
  kind: 'chart-sheet';
}

export type XlsxSheet = XlsxChartSheet | XlsxWorksheet;

export interface XlsxDocument {
  differentialStyles: XlsxStyle[];
  namedStyles: XlsxNamedStyle[];
  sheets: XlsxSheet[];
  styles: XlsxStyle[];
  workbook: XlsxWorkbookProperties;
}
