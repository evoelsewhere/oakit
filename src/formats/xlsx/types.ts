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
  height?: number;
  hidden?: boolean;
  index: number;
  outlineLevel?: number;
}

export interface XlsxColumnRange {
  end: number;
  hidden?: boolean;
  outlineLevel?: number;
  start: number;
  style?: number;
  width?: number;
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
  border?: XlsxBorder;
  fill?: XlsxFill;
  font?: XlsxFont;
  numberFormat?: string;
}

export interface XlsxDefinedName {
  expression: string;
  hidden: boolean;
  name: string;
  sheetIndex?: number;
}

export interface XlsxWorkbookProperties {
  calculation: {
    forceFullCalculation: boolean;
    fullCalculationOnLoad: boolean;
    mode: 'automatic' | 'automatic-except-tables' | 'manual';
  };
  dateSystem: '1900' | '1904';
  definedNames: XlsxDefinedName[];
}

export interface XlsxTable {
  name: string;
  range: XlsxRange;
}

export interface XlsxDrawing {
  kind: 'absolute' | 'one-cell' | 'two-cell';
}

export interface XlsxSheetBase {
  index: number;
  name: string;
  payload: 'full-sheet' | 'not-selected' | 'selected-ranges';
  state: 'hidden' | 'very-hidden' | 'visible';
}

export interface XlsxWorksheet extends XlsxSheetBase {
  columns: XlsxColumnRange[];
  drawings: XlsxDrawing[];
  kind: 'worksheet';
  mergedRanges: XlsxRange[];
  rows: XlsxRow[];
  tables: XlsxTable[];
}

export interface XlsxChartSheet extends XlsxSheetBase {
  kind: 'chart-sheet';
}

export type XlsxSheet = XlsxChartSheet | XlsxWorksheet;

export interface XlsxDocument {
  sheets: XlsxSheet[];
  styles: XlsxStyle[];
  workbook: XlsxWorkbookProperties;
}
