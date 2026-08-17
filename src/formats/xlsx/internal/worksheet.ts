import { XlsxParseError } from '../errors';
import type {
  XlsxCell,
  XlsxCellValue,
  XlsxColumnRange,
  XlsxFormula,
  XlsxRange,
  XlsxRichTextRun,
  XlsxRow,
  XlsxWorksheetView,
} from '../types';
import {
  parseXlsxCellReference,
  parseXlsxRangeReference,
  xlsxColumnName,
} from './cell-reference';
import {
  parseXlsxScalarCellValue,
  type XlsxScalarCellType,
} from './cell-value';
import { normalizeXlsxSerialDate } from './date-system';
import { translateXlsxSharedFormula } from './formula';
import { xlsxNumberFormatDatePrecision } from './number-format';
import { XlsxPartReader } from './part-reader';
import {
  type ResolvedXlsxResourceLimits,
  XlsxResourceLimitError,
} from './resource-limits';
import {
  type XlsxResolvedSheetSelection,
  xlsxSelectionIncludesCell,
  xlsxSelectionIncludesRow,
} from './selection';
import type { XlsxSharedStringTable } from './shared-strings';
import { EMPTY_XLSX_STYLE_TABLE, type XlsxStyleTable } from './styles';
import type { XlsxXmlElement, XlsxXmlEventSink } from './streaming-xml';
import {
  type XlsxWorkbookDiscovery,
  XLSX_SPREADSHEET_NAMESPACES,
} from './workbook-discovery';
import {
  normalizeXlsxColumnRanges,
  type XlsxAuthoredColumnRange,
  xlsxMergedRangesOverlap,
} from './worksheet-layout';
import {
  parseXlsxWorksheetPane,
  parseXlsxWorksheetView,
  parseXlsxWorksheetViewSelection,
  validateXlsxWorksheetView,
} from './worksheet-view';

const XML_NAMESPACE = 'http://www.w3.org/XML/1998/namespace';
const UNSIGNED_INTEGER_PATTERN = /^(?:0|[1-9]\d*)$/u;
const NONNEGATIVE_DECIMAL_PATTERN = /^(?:\d+(?:\.\d*)?|\.\d+)$/u;

type CellType = XlsxScalarCellType | 'inlineStr';
type TextCapture = 'formula' | 'inline' | 'value' | null;

interface PendingFormula {
  reference?: string;
  sharedIndex?: string;
  text: string;
  type: 'array' | 'dataTable' | 'normal' | 'shared';
}

interface PendingCell {
  address: string;
  column: number;
  formula?: PendingFormula;
  hasInlineString: boolean;
  hasValue: boolean;
  inlineMode: 'plain' | 'rich' | 'unset';
  inlineRuns: XlsxRichTextRun[];
  inlineText: string;
  numberFormat?: string;
  selected: boolean;
  style?: number;
  type: CellType;
  valueText: string;
}

interface PendingInlineRun {
  hasText: boolean;
  text: string;
}

export interface XlsxWorksheetBudget {
  formulaCharacters: number;
  formulaGroups: number;
  rangeAreas: number;
  returnedCells: number;
  richTextRuns: number;
  scannedCells: number;
  textCharacters: number;
}

interface SharedFormulaMaster {
  column: number;
  expression: string;
  range: XlsxRange;
  row: number;
}

export interface XlsxWorksheetPayload {
  columns: XlsxColumnRange[];
  mergedRanges: XlsxRange[];
  rows: XlsxRow[];
  views: XlsxWorksheetView[];
}

export interface XlsxWorksheetSemantics {
  dateSystem: '1900' | '1904';
  styles: XlsxStyleTable;
}

const DEFAULT_WORKSHEET_SEMANTICS: XlsxWorksheetSemantics = Object.freeze({
  dateSystem: '1900',
  styles: EMPTY_XLSX_STYLE_TABLE,
});

function structureFailure(
  part: string,
  cell: string | undefined,
  message: string,
): never {
  throw new XlsxParseError({
    ...(cell === undefined ? {} : { cell }),
    code: 'invalid-document-structure',
    message,
    part,
    severity: 'error',
  });
}

function valueFailure(
  part: string,
  cell: string | undefined,
  message: string,
): never {
  throw new XlsxParseError({
    ...(cell === undefined ? {} : { cell }),
    code: 'invalid-document-value',
    message,
    part,
    severity: 'error',
  });
}

function formulaFailure(part: string, cell: string, message: string): never {
  throw new XlsxParseError({
    cell,
    code: 'invalid-formula',
    message,
    part,
    severity: 'error',
  });
}

function attribute(
  element: XlsxXmlElement,
  localName: string,
): string | undefined {
  return element.attributes.get(`{}${localName}`);
}

function unsignedInteger(
  value: string | undefined,
  part: string,
  cell: string | undefined,
  message: string,
): number | undefined {
  if (value === undefined) return undefined;
  if (!UNSIGNED_INTEGER_PATTERN.test(value)) valueFailure(part, cell, message);
  const parsed = Number(value);
  if (!Number.isSafeInteger(parsed) || parsed > 0xffff_ffff) {
    valueFailure(part, cell, message);
  }
  return parsed;
}

function optionalBoolean(
  value: string | undefined,
  part: string,
  message: string,
): boolean | undefined {
  if (value === undefined) return undefined;
  if (value === '0' || value === 'false') return false;
  if (value === '1' || value === 'true') return true;
  valueFailure(part, undefined, message);
}

function optionalHeight(
  value: string | undefined,
  part: string,
): number | undefined {
  if (value === undefined) return undefined;
  if (!NONNEGATIVE_DECIMAL_PATTERN.test(value)) {
    valueFailure(part, undefined, 'Worksheet row height is invalid');
  }
  const parsed = Number(value);
  if (!Number.isFinite(parsed)) {
    valueFailure(
      part,
      undefined,
      'Worksheet row height is outside the finite range',
    );
  }
  return parsed;
}

function optionalWidth(
  value: string | undefined,
  part: string,
): number | undefined {
  if (value === undefined) return undefined;
  if (!NONNEGATIVE_DECIMAL_PATTERN.test(value)) {
    valueFailure(part, undefined, 'Worksheet column width is invalid');
  }
  const parsed = Number(value);
  if (!Number.isFinite(parsed) || parsed > 255) {
    valueFailure(part, undefined, 'Worksheet column width is invalid');
  }
  return parsed;
}

function resolvedStyle(
  rawStyle: number | undefined,
  styles: XlsxStyleTable,
  part: string,
  location: string | undefined,
  message: string,
) {
  if (rawStyle === undefined) return undefined;
  const style = styles.cellXfs[rawStyle];
  if (!style) valueFailure(part, location, message);
  return style;
}

function cellType(
  value: string | undefined,
  part: string,
  cell: string,
): CellType {
  if (value === undefined) return 'n';
  if (
    value === 'b' ||
    value === 'd' ||
    value === 'e' ||
    value === 'inlineStr' ||
    value === 'n' ||
    value === 's' ||
    value === 'str'
  ) {
    return value;
  }
  valueFailure(part, cell, 'Worksheet cell type is invalid');
}

function validateXmlSpace(
  element: XlsxXmlElement,
  part: string,
  cell: string,
): void {
  const value = element.attributes.get(`{${XML_NAMESPACE}}space`);
  if (value !== undefined && value !== 'default' && value !== 'preserve') {
    valueFailure(part, cell, 'Inline-string xml:space value is invalid');
  }
}

function consume(
  budget: XlsxWorksheetBudget,
  key: keyof XlsxWorksheetBudget,
  amount: number,
  limitName:
    | 'maxFormulaGroups'
    | 'maxRangeAreas'
    | 'maxReturnedCells'
    | 'maxRichTextRuns'
    | 'maxScannedCells'
    | 'maxTextCharacters',
  limits: ResolvedXlsxResourceLimits,
  part: string,
): void {
  const actual = budget[key] + amount;
  if (!Number.isSafeInteger(actual) || actual > limits[limitName]) {
    throw new XlsxResourceLimitError(
      limitName,
      actual,
      limits[limitName],
      part,
    );
  }
  budget[key] = actual;
}

export function createXlsxWorksheetBudget(
  sharedStrings: XlsxSharedStringTable,
  initial: Partial<
    Pick<XlsxWorksheetBudget, 'formulaCharacters' | 'textCharacters'>
  > = {},
): XlsxWorksheetBudget {
  let richTextRuns = 0;
  let textCharacters = 0;
  for (const value of sharedStrings.values) {
    textCharacters += value.text.length;
    richTextRuns += value.runs?.length ?? 0;
    for (const run of value.phoneticRuns ?? []) {
      textCharacters += run.text.length;
      richTextRuns += 1;
    }
  }
  return {
    formulaCharacters: initial.formulaCharacters ?? 0,
    formulaGroups: 0,
    rangeAreas: 0,
    returnedCells: 0,
    richTextRuns,
    scannedCells: 0,
    textCharacters: textCharacters + (initial.textCharacters ?? 0),
  };
}

class WorksheetSink implements XlsxXmlEventSink {
  private readonly authoredColumns: XlsxAuthoredColumnRange[] = [];
  private capture: TextCapture = null;
  private columnsSeen = false;
  private currentCell: PendingCell | undefined;
  private currentInlineRun: PendingInlineRun | undefined;
  private currentRow: XlsxRow | undefined;
  private currentRowSelected!: boolean;
  private currentView: XlsxWorksheetView | undefined;
  private ignoredDepth = 0;
  private lastCellColumn = 0;
  private lastRow = 0;
  private mergeCellsExpected: number | undefined;
  private mergeCellsSeen = false;
  private readonly mergedRanges: XlsxRange[] = [];
  private readonly selectedColumnPrefix: Uint32Array;
  private sheetDataSeen = false;
  private sheetViewsSeen = false;
  private readonly stack: XlsxXmlElement[] = [];
  private readonly rows: XlsxRow[] = [];
  private readonly viewIds = new Set<number>();
  private readonly views: XlsxWorksheetView[] = [];
  private readonly sharedFormulaMasters = new Map<
    number,
    SharedFormulaMaster
  >();

  constructor(
    private readonly part: string,
    private readonly namespace: string,
    private readonly sharedStrings: XlsxSharedStringTable,
    private readonly budget: XlsxWorksheetBudget,
    private readonly limits: ResolvedXlsxResourceLimits,
    private readonly selection: XlsxResolvedSheetSelection,
    private readonly semantics: XlsxWorksheetSemantics,
  ) {
    this.selectedColumnPrefix = new Uint32Array(
      this.limits.maxColumnsPerWorksheet + 1,
    );
    if (selection.kind === 'selected-ranges') {
      const differences = new Int32Array(
        this.limits.maxColumnsPerWorksheet + 2,
      );
      for (const range of selection.ranges) {
        const start = Math.min(
          range.start.column,
          this.limits.maxColumnsPerWorksheet,
        );
        const end = Math.min(
          range.end.column,
          this.limits.maxColumnsPerWorksheet,
        );
        differences[start]! += 1;
        differences[end + 1]! -= 1;
      }
      let active = 0;
      let selected = 0;
      Array.from(
        { length: this.limits.maxColumnsPerWorksheet },
        (_, index) => index + 1,
      ).forEach((column) => {
        active += differences[column]!;
        if (active > 0) selected += 1;
        this.selectedColumnPrefix[column] = selected;
      });
    }
  }

  openElement(element: XlsxXmlElement): void {
    if (this.ignoredDepth > 0) {
      this.ignoredDepth += 1;
      this.stack.push(element);
      return;
    }
    if (element.namespace !== this.namespace) {
      structureFailure(
        this.part,
        this.currentCell?.address,
        'Worksheet element has an unsupported namespace',
      );
    }
    const parent = this.stack.at(-1);
    if (!parent) {
      if (element.localName !== 'worksheet') {
        structureFailure(this.part, undefined, 'Worksheet root is missing');
      }
      this.stack.push(element);
      return;
    }

    this.openChild(parent.localName, element);
    this.stack.push(element);
  }

  closeElement(element: XlsxXmlElement): void {
    if (this.ignoredDepth > 0) {
      this.ignoredDepth -= 1;
      this.stack.pop();
      return;
    }
    this.capture = null;
    if (element.localName === 'r') this.closeInlineRun();
    if (element.localName === 'c') this.closeCell();
    if (element.localName === 'row') this.closeRow();
    if (element.localName === 'sheetView') this.closeView();
    this.stack.pop();
  }

  text(value: string): void {
    if (this.ignoredDepth > 0) return;
    if (this.capture === 'value') {
      this.currentCell!.valueText += value;
      return;
    }
    if (this.capture === 'formula') {
      this.currentCell!.formula!.text += value;
      return;
    }
    if (this.capture === 'inline') {
      if (this.currentInlineRun) this.currentInlineRun.text += value;
      else this.currentCell!.inlineText += value;
      return;
    }
    if (value.trim().length > 0) {
      structureFailure(
        this.part,
        this.currentCell?.address,
        'Worksheet text is outside a value or inline-string text element',
      );
    }
  }

  result(): XlsxWorksheetPayload {
    if (!this.sheetDataSeen) {
      structureFailure(this.part, undefined, 'Worksheet sheetData is missing');
    }
    if (this.sheetViewsSeen && this.views.length === 0) {
      structureFailure(
        this.part,
        undefined,
        'Worksheet sheetViews collection is empty',
      );
    }
    if (
      this.mergeCellsExpected !== undefined &&
      this.mergeCellsExpected !== this.mergedRanges.length
    ) {
      structureFailure(
        this.part,
        undefined,
        'Worksheet merged-range count does not match',
      );
    }
    if (xlsxMergedRangesOverlap(this.mergedRanges)) {
      valueFailure(this.part, undefined, 'Worksheet merged ranges overlap');
    }
    return {
      columns: normalizeXlsxColumnRanges(this.authoredColumns).filter((range) =>
        this.columnRangeSelected(range),
      ),
      mergedRanges: this.mergedRanges.filter((range) =>
        this.mergedRangeSelected(range),
      ),
      rows: this.rows,
      views: this.views,
    };
  }

  private columnRangeSelected(range: XlsxColumnRange): boolean {
    if (this.selection.kind !== 'selected-ranges') {
      return this.selection.kind === 'full-sheet';
    }
    return (
      this.selectedColumnPrefix[range.end]! -
        this.selectedColumnPrefix[range.start - 1]! >
      0
    );
  }

  private mergedRangeSelected(range: XlsxRange): boolean {
    if (this.selection.kind !== 'selected-ranges') {
      return this.selection.kind === 'full-sheet';
    }
    for (const selected of this.selection.ranges) {
      consume(
        this.budget,
        'scannedCells',
        1,
        'maxScannedCells',
        this.limits,
        this.part,
      );
      if (
        selected.start.row <= range.end.row &&
        selected.end.row >= range.start.row &&
        selected.start.column <= range.end.column &&
        selected.end.column >= range.start.column
      ) {
        return true;
      }
    }
    return false;
  }

  private beginIgnore(): void {
    this.ignoredDepth = 1;
  }

  private openChild(parent: string, element: XlsxXmlElement): void {
    if (parent === 'worksheet') {
      if (element.localName === 'cols') {
        if (this.columnsSeen) {
          structureFailure(
            this.part,
            undefined,
            'Worksheet contains duplicate cols elements',
          );
        }
        this.columnsSeen = true;
        return;
      }
      if (element.localName === 'sheetData') {
        if (this.sheetDataSeen) {
          structureFailure(
            this.part,
            undefined,
            'Worksheet contains duplicate sheetData elements',
          );
        }
        this.sheetDataSeen = true;
        return;
      }
      if (element.localName === 'sheetViews') {
        if (this.sheetViewsSeen) {
          structureFailure(
            this.part,
            undefined,
            'Worksheet contains duplicate sheetViews elements',
          );
        }
        this.sheetViewsSeen = true;
        return;
      }
      if (element.localName === 'mergeCells') {
        if (this.mergeCellsSeen) {
          structureFailure(
            this.part,
            undefined,
            'Worksheet contains duplicate mergeCells elements',
          );
        }
        this.mergeCellsSeen = true;
        this.mergeCellsExpected = unsignedInteger(
          attribute(element, 'count'),
          this.part,
          undefined,
          'Worksheet merged-range count is invalid',
        );
        if (
          this.mergeCellsExpected !== undefined &&
          this.mergeCellsExpected > this.limits.maxMergedRanges
        ) {
          throw new XlsxResourceLimitError(
            'maxMergedRanges',
            this.mergeCellsExpected,
            this.limits.maxMergedRanges,
            this.part,
          );
        }
        return;
      }
      this.beginIgnore();
      return;
    }
    if (parent === 'cols' && element.localName === 'col') {
      this.openColumn(element);
      return;
    }
    if (parent === 'sheetViews' && element.localName === 'sheetView') {
      this.openView(element);
      return;
    }
    if (parent === 'sheetView') {
      if (element.localName === 'pane') {
        this.openPane(element);
        return;
      }
      if (element.localName === 'selection') {
        this.openViewSelection(element);
        return;
      }
      if (
        element.localName === 'extLst' ||
        element.localName === 'pivotSelection'
      ) {
        this.beginIgnore();
        return;
      }
    }
    if (parent === 'mergeCells' && element.localName === 'mergeCell') {
      this.openMergedRange(element);
      return;
    }
    if (parent === 'sheetData' && element.localName === 'row') {
      this.openRow(element);
      return;
    }
    if (parent === 'row') {
      if (element.localName === 'c') {
        this.openCell(element);
        return;
      }
      if (element.localName === 'extLst') {
        this.beginIgnore();
        return;
      }
    }
    if (parent === 'c') {
      if (element.localName === 'f') {
        this.openFormula(element);
        return;
      }
      if (element.localName === 'v') {
        this.openValue();
        return;
      }
      if (element.localName === 'is') {
        this.openInlineString();
        return;
      }
      if (element.localName === 'extLst') {
        this.beginIgnore();
        return;
      }
    }
    if (parent === 'is') {
      if (element.localName === 't') {
        this.openInlinePlainText(element);
        return;
      }
      if (element.localName === 'r') {
        this.openInlineRun();
        return;
      }
      if (element.localName === 'rPh' || element.localName === 'phoneticPr') {
        this.beginIgnore();
        return;
      }
    }
    if (parent === 'r') {
      if (element.localName === 'rPr') {
        this.beginIgnore();
        return;
      }
      if (element.localName === 't') {
        this.openInlineRunText(element);
        return;
      }
    }
    structureFailure(
      this.part,
      this.currentCell?.address,
      'Worksheet element nesting is invalid',
    );
  }

  private openColumn(element: XlsxXmlElement): void {
    const start = unsignedInteger(
      attribute(element, 'min'),
      this.part,
      undefined,
      'Worksheet column start is invalid',
    );
    const end = unsignedInteger(
      attribute(element, 'max'),
      this.part,
      undefined,
      'Worksheet column end is invalid',
    );
    if (start === undefined || start === 0) {
      valueFailure(this.part, undefined, 'Worksheet column start is invalid');
    }
    if (end === undefined || end < start) {
      valueFailure(this.part, undefined, 'Worksheet column end is invalid');
    }
    if (end > this.limits.maxColumnsPerWorksheet) {
      throw new XlsxResourceLimitError(
        'maxColumnsPerWorksheet',
        end,
        this.limits.maxColumnsPerWorksheet,
        this.part,
      );
    }
    const actualRanges = this.authoredColumns.length + 1;
    if (actualRanges > this.limits.maxColumnsPerWorksheet) {
      throw new XlsxResourceLimitError(
        'maxColumnsPerWorksheet',
        actualRanges,
        this.limits.maxColumnsPerWorksheet,
        this.part,
      );
    }
    const outlineLevel = unsignedInteger(
      attribute(element, 'outlineLevel'),
      this.part,
      undefined,
      'Worksheet column outline level is invalid',
    );
    if (outlineLevel !== undefined && outlineLevel > 7) {
      valueFailure(
        this.part,
        undefined,
        'Worksheet column outline level is invalid',
      );
    }
    const collapsed = optionalBoolean(
      attribute(element, 'collapsed'),
      this.part,
      'Worksheet column collapsed flag is invalid',
    );
    const hidden = optionalBoolean(
      attribute(element, 'hidden'),
      this.part,
      'Worksheet column hidden flag is invalid',
    );
    optionalBoolean(
      attribute(element, 'bestFit'),
      this.part,
      'Worksheet column bestFit flag is invalid',
    );
    optionalBoolean(
      attribute(element, 'customWidth'),
      this.part,
      'Worksheet column customWidth flag is invalid',
    );
    optionalBoolean(
      attribute(element, 'phonetic'),
      this.part,
      'Worksheet column phonetic flag is invalid',
    );
    const width = optionalWidth(attribute(element, 'width'), this.part);
    const rawStyle = unsignedInteger(
      attribute(element, 'style'),
      this.part,
      undefined,
      'Worksheet column style index is invalid',
    );
    const style = resolvedStyle(
      rawStyle,
      this.semantics.styles,
      this.part,
      undefined,
      'Worksheet column style reference is invalid',
    );
    this.authoredColumns.push({
      ...(collapsed === undefined ? {} : { collapsed }),
      end,
      ...(hidden === undefined ? {} : { hidden }),
      order: this.authoredColumns.length,
      ...(outlineLevel === undefined ? {} : { outlineLevel }),
      start,
      ...(style === undefined ? {} : { style: style.normalizedStyle }),
      ...(width === undefined ? {} : { width }),
    });
  }

  private openView(element: XlsxXmlElement): void {
    const view = parseXlsxWorksheetView(element, this.part);
    if (this.viewIds.has(view.workbookViewId)) {
      valueFailure(
        this.part,
        undefined,
        'Worksheet contains duplicate workbook view references',
      );
    }
    this.viewIds.add(view.workbookViewId);
    this.views.push(view);
    this.currentView = view;
  }

  private openPane(element: XlsxXmlElement): void {
    if (this.currentView!.pane) {
      structureFailure(
        this.part,
        undefined,
        'Worksheet view contains duplicate pane elements',
      );
    }
    this.currentView!.pane = parseXlsxWorksheetPane(
      element,
      this.part,
      this.limits,
    );
  }

  private openViewSelection(element: XlsxXmlElement): void {
    const parsed = parseXlsxWorksheetViewSelection(element, this.part);
    consume(
      this.budget,
      'rangeAreas',
      parsed.rangeAreaCount,
      'maxRangeAreas',
      this.limits,
      this.part,
    );
    this.currentView!.selections.push(parsed.selection);
  }

  private closeView(): void {
    validateXlsxWorksheetView(this.currentView!, this.part);
    this.currentView = undefined;
  }

  private openMergedRange(element: XlsxXmlElement): void {
    const source = attribute(element, 'ref');
    const range = parseXlsxRangeReference(source);
    if (!range || source?.includes('$')) {
      valueFailure(
        this.part,
        undefined,
        'Worksheet merged-range reference is invalid',
      );
    }
    if (
      range.start.row === range.end.row &&
      range.start.column === range.end.column
    ) {
      valueFailure(
        this.part,
        undefined,
        'Worksheet merged range must contain multiple cells',
      );
    }
    const actual = this.mergedRanges.length + 1;
    if (actual > this.limits.maxMergedRanges) {
      throw new XlsxResourceLimitError(
        'maxMergedRanges',
        actual,
        this.limits.maxMergedRanges,
        this.part,
      );
    }
    this.mergedRanges.push(range);
  }

  private openRow(element: XlsxXmlElement): void {
    const reference = unsignedInteger(
      attribute(element, 'r'),
      this.part,
      undefined,
      'Worksheet row reference is invalid',
    );
    const index = reference ?? this.lastRow + 1;
    if (index <= this.lastRow) {
      valueFailure(this.part, undefined, 'Worksheet rows are out of order');
    }
    if (index > this.limits.maxRowsPerWorksheet) {
      throw new XlsxResourceLimitError(
        'maxRowsPerWorksheet',
        index,
        this.limits.maxRowsPerWorksheet,
        this.part,
      );
    }
    const outlineLevel = unsignedInteger(
      attribute(element, 'outlineLevel'),
      this.part,
      undefined,
      'Worksheet row outline level is invalid',
    );
    if (outlineLevel !== undefined && outlineLevel > 7) {
      valueFailure(
        this.part,
        undefined,
        'Worksheet row outline level is invalid',
      );
    }
    const height = optionalHeight(attribute(element, 'ht'), this.part);
    const collapsed = optionalBoolean(
      attribute(element, 'collapsed'),
      this.part,
      'Worksheet row collapsed flag is invalid',
    );
    const customFormat = optionalBoolean(
      attribute(element, 'customFormat'),
      this.part,
      'Worksheet row customFormat flag is invalid',
    );
    optionalBoolean(
      attribute(element, 'customHeight'),
      this.part,
      'Worksheet row customHeight flag is invalid',
    );
    const hidden = optionalBoolean(
      attribute(element, 'hidden'),
      this.part,
      'Worksheet row hidden flag is invalid',
    );
    const rawStyle = unsignedInteger(
      attribute(element, 's'),
      this.part,
      undefined,
      'Worksheet row style index is invalid',
    );
    if (customFormat === true && rawStyle === undefined) {
      valueFailure(
        this.part,
        undefined,
        'Worksheet custom-formatted row style is missing',
      );
    }
    const style = resolvedStyle(
      rawStyle,
      this.semantics.styles,
      this.part,
      undefined,
      'Worksheet row style reference is invalid',
    );
    this.currentRow = {
      cells: [],
      ...(collapsed === undefined ? {} : { collapsed }),
      ...(height === undefined ? {} : { height }),
      ...(hidden === undefined ? {} : { hidden }),
      index,
      ...(outlineLevel === undefined ? {} : { outlineLevel }),
      ...(style === undefined ? {} : { style: style.normalizedStyle }),
    };
    this.currentRowSelected = xlsxSelectionIncludesRow(this.selection, index);
    this.lastCellColumn = 0;
    this.lastRow = index;
  }

  private closeRow(): void {
    if (this.currentRowSelected) this.rows.push(this.currentRow!);
    this.currentRow = undefined;
  }

  private openCell(element: XlsxXmlElement): void {
    const sourceReference = attribute(element, 'r');
    let address: string;
    let column: number;
    if (sourceReference === undefined) {
      column = this.lastCellColumn + 1;
      if (column > this.limits.maxColumnsPerWorksheet) {
        throw new XlsxResourceLimitError(
          'maxColumnsPerWorksheet',
          column,
          this.limits.maxColumnsPerWorksheet,
          this.part,
        );
      }
      address = `${xlsxColumnName(column)!}${this.currentRow!.index}`;
    } else {
      const parsed = parseXlsxCellReference(sourceReference);
      if (!parsed || parsed.absoluteColumn || parsed.absoluteRow) {
        valueFailure(
          this.part,
          undefined,
          'Worksheet cell reference is invalid',
        );
      }
      address = parsed.address;
      column = parsed.column;
      if (parsed.row !== this.currentRow!.index) {
        valueFailure(
          this.part,
          address,
          'Worksheet cell reference does not belong to its row',
        );
      }
      if (column > this.limits.maxColumnsPerWorksheet) {
        throw new XlsxResourceLimitError(
          'maxColumnsPerWorksheet',
          column,
          this.limits.maxColumnsPerWorksheet,
          this.part,
        );
      }
    }
    if (column <= this.lastCellColumn) {
      valueFailure(this.part, address, 'Worksheet cells are out of order');
    }
    const rawStyle = unsignedInteger(
      attribute(element, 's'),
      this.part,
      address,
      'Worksheet cell style index is invalid',
    );
    const styleXf = resolvedStyle(
      rawStyle,
      this.semantics.styles,
      this.part,
      address,
      'Worksheet cell style reference is invalid',
    );
    consume(
      this.budget,
      'scannedCells',
      1,
      'maxScannedCells',
      this.limits,
      this.part,
    );
    const selected = xlsxSelectionIncludesCell(
      this.selection,
      this.currentRow!.index,
      column,
    );
    if (selected) {
      consume(
        this.budget,
        'returnedCells',
        1,
        'maxReturnedCells',
        this.limits,
        this.part,
      );
    }
    this.currentCell = {
      address,
      column,
      hasInlineString: false,
      hasValue: false,
      inlineMode: 'unset',
      inlineRuns: [],
      inlineText: '',
      ...(styleXf?.numberFormat === undefined
        ? {}
        : { numberFormat: styleXf.numberFormat }),
      selected,
      ...(styleXf === undefined ? {} : { style: styleXf.normalizedStyle }),
      type: cellType(attribute(element, 't'), this.part, address),
      valueText: '',
    };
    this.lastCellColumn = column;
  }

  private openValue(): void {
    const cell = this.currentCell!;
    if (cell.hasValue || cell.hasInlineString || cell.type === 'inlineStr') {
      structureFailure(
        this.part,
        cell.address,
        'Worksheet cell value structure is invalid',
      );
    }
    cell.hasValue = true;
    this.capture = 'value';
  }

  private openFormula(element: XlsxXmlElement): void {
    const cell = this.currentCell!;
    if (
      cell.formula !== undefined ||
      cell.hasValue ||
      cell.hasInlineString ||
      cell.type === 'inlineStr' ||
      cell.type === 's'
    ) {
      structureFailure(
        this.part,
        cell.address,
        'Worksheet formula structure is invalid',
      );
    }
    const sourceType = attribute(element, 't');
    const type = sourceType ?? 'normal';
    if (
      type !== 'normal' &&
      type !== 'shared' &&
      type !== 'array' &&
      type !== 'dataTable'
    ) {
      formulaFailure(this.part, cell.address, 'Formula type is invalid');
    }
    const reference = attribute(element, 'ref');
    const sharedIndex = attribute(element, 'si');
    cell.formula = {
      ...(reference === undefined ? {} : { reference }),
      ...(sharedIndex === undefined ? {} : { sharedIndex }),
      text: '',
      type,
    };
    this.capture = 'formula';
  }

  private openInlineString(): void {
    const cell = this.currentCell!;
    if (cell.hasValue || cell.hasInlineString || cell.type !== 'inlineStr') {
      structureFailure(
        this.part,
        cell.address,
        'Worksheet inline-string structure is invalid',
      );
    }
    cell.hasInlineString = true;
  }

  private openInlinePlainText(element: XlsxXmlElement): void {
    const cell = this.currentCell!;
    if (cell.inlineMode !== 'unset') {
      structureFailure(
        this.part,
        cell.address,
        'Worksheet inline-string plain text is out of order',
      );
    }
    cell.inlineMode = 'plain';
    validateXmlSpace(element, this.part, cell.address);
    this.capture = 'inline';
  }

  private openInlineRun(): void {
    const cell = this.currentCell!;
    if (cell.inlineMode === 'plain') {
      structureFailure(
        this.part,
        cell.address,
        'Worksheet inline-string rich text is out of order',
      );
    }
    cell.inlineMode = 'rich';
    consume(
      this.budget,
      'richTextRuns',
      1,
      'maxRichTextRuns',
      this.limits,
      this.part,
    );
    this.currentInlineRun = { hasText: false, text: '' };
  }

  private openInlineRunText(element: XlsxXmlElement): void {
    const cell = this.currentCell!;
    if (this.currentInlineRun!.hasText) {
      structureFailure(
        this.part,
        cell.address,
        'Worksheet inline-string run has duplicate text',
      );
    }
    this.currentInlineRun!.hasText = true;
    validateXmlSpace(element, this.part, cell.address);
    this.capture = 'inline';
  }

  private closeInlineRun(): void {
    const cell = this.currentCell!;
    if (!this.currentInlineRun!.hasText) {
      structureFailure(
        this.part,
        cell.address,
        'Worksheet inline-string run text is missing',
      );
    }
    cell.inlineRuns.push({ text: this.currentInlineRun!.text });
    cell.inlineText += this.currentInlineRun!.text;
    this.currentInlineRun = undefined;
  }

  private closeCell(): void {
    const cell = this.currentCell!;
    let content: XlsxCell['content'];
    if (cell.formula !== undefined) {
      const formula = this.resolveFormula(cell);
      const cached = cell.hasValue
        ? this.applyNumberFormat(
            parseXlsxScalarCellValue(
              cell.type as XlsxScalarCellType,
              cell.valueText,
              this.sharedStrings,
              this.part,
              cell.address,
            ),
            cell,
          )
        : ({ kind: 'missing' } as const);
      if (cell.selected && cached.kind !== 'missing') {
        this.consumeReturnedText(cached);
      }
      content = { cached, formula, kind: 'formula' };
    } else if (!cell.hasValue && !cell.hasInlineString) {
      content = { kind: 'blank' };
    } else if (cell.hasInlineString) {
      const value: XlsxCellValue = {
        kind: 'text',
        ...(cell.inlineMode === 'rich' ? { runs: cell.inlineRuns } : {}),
        text: cell.inlineText,
      };
      if (cell.selected) this.consumeTextCharacters(value.text);
      content = { kind: 'value', value };
    } else {
      const value = this.applyNumberFormat(
        parseXlsxScalarCellValue(
          cell.type as XlsxScalarCellType,
          cell.valueText,
          this.sharedStrings,
          this.part,
          cell.address,
        ),
        cell,
      );
      if (cell.selected) this.consumeReturnedText(value);
      content = { kind: 'value', value };
    }
    if (!cell.selected) {
      this.currentCell = undefined;
      return;
    }
    const base = {
      address: cell.address,
      column: cell.column,
      ...(cell.style === undefined ? {} : { style: cell.style }),
    };
    if (content.kind === 'blank') {
      this.currentRow!.cells.push({ ...base, content: { kind: 'blank' } });
    } else if (content.kind === 'formula') {
      this.currentRow!.cells.push({ ...base, content });
    } else {
      this.currentRow!.cells.push({
        ...base,
        content: { kind: 'value', value: content.value },
      });
    }
    this.currentCell = undefined;
  }

  private applyNumberFormat(
    value: XlsxCellValue,
    cell: PendingCell,
  ): XlsxCellValue {
    if (value.kind !== 'number' || cell.numberFormat === undefined)
      return value;
    const precision = xlsxNumberFormatDatePrecision(
      cell.numberFormat,
      value.value,
    );
    if (precision === undefined) return value;
    return {
      kind: 'date',
      normalized: normalizeXlsxSerialDate(
        value.value,
        this.semantics.dateSystem,
        precision,
      ),
      precision,
      source: {
        dateSystem: this.semantics.dateSystem,
        kind: 'serial',
        value: value.value,
      },
    };
  }

  private resolveFormula(cell: PendingCell): XlsxFormula {
    const pending = cell.formula!;
    if (pending.text.startsWith('=')) {
      formulaFailure(
        this.part,
        cell.address,
        'Formula expression must not include a leading equals sign',
      );
    }
    if (pending.type === 'normal') {
      if (
        pending.text.length === 0 ||
        pending.reference !== undefined ||
        pending.sharedIndex !== undefined
      ) {
        formulaFailure(this.part, cell.address, 'Normal formula is invalid');
      }
      this.consumeFormulaCharacters(pending.text);
      return { expression: pending.text, kind: 'normal' };
    }
    if (pending.type === 'shared') return this.resolveSharedFormula(cell);

    if (pending.sharedIndex !== undefined) {
      formulaFailure(
        this.part,
        cell.address,
        'Grouped formula shared index is invalid',
      );
    }
    const range = this.formulaRange(pending.reference, cell);
    if (
      range.start.row !== this.currentRow!.index ||
      range.start.column !== cell.column
    ) {
      formulaFailure(
        this.part,
        cell.address,
        'Grouped formula must start at its owning cell',
      );
    }
    if (pending.type === 'array' && pending.text.length === 0) {
      formulaFailure(this.part, cell.address, 'Array formula is empty');
    }
    this.consumeFormulaGroup();
    this.consumeFormulaCharacters(pending.text);
    return {
      expression: pending.text,
      kind: pending.type === 'array' ? 'array' : 'data-table',
      range,
    };
  }

  private resolveSharedFormula(cell: PendingCell): XlsxFormula {
    const pending = cell.formula!;
    const sharedIndex = this.formulaSharedIndex(pending.sharedIndex, cell);
    if (pending.reference !== undefined) {
      if (
        pending.text.length === 0 ||
        this.sharedFormulaMasters.has(sharedIndex)
      ) {
        formulaFailure(
          this.part,
          cell.address,
          'Shared formula master is invalid',
        );
      }
      const range = this.formulaRange(pending.reference, cell);
      if (
        range.start.row !== this.currentRow!.index ||
        range.start.column !== cell.column
      ) {
        formulaFailure(
          this.part,
          cell.address,
          'Shared formula master must start at its owning cell',
        );
      }
      this.consumeFormulaGroup();
      this.sharedFormulaMasters.set(sharedIndex, {
        column: cell.column,
        expression: pending.text,
        range,
        row: this.currentRow!.index,
      });
      this.consumeFormulaCharacters(pending.text);
      return { expression: pending.text, kind: 'normal' };
    }
    if (pending.text.length !== 0) {
      formulaFailure(
        this.part,
        cell.address,
        'Shared formula dependent contains an expression',
      );
    }
    const master = this.sharedFormulaMasters.get(sharedIndex);
    if (
      master === undefined ||
      !this.rangeContains(master.range, this.currentRow!.index, cell.column)
    ) {
      formulaFailure(
        this.part,
        cell.address,
        'Shared formula master is missing or does not own the cell',
      );
    }
    const expression = translateXlsxSharedFormula(
      master.expression,
      { column: master.column, row: master.row },
      { column: cell.column, row: this.currentRow!.index },
    );
    if (expression === undefined) {
      formulaFailure(
        this.part,
        cell.address,
        'Shared formula translation is outside the worksheet grid',
      );
    }
    this.consumeFormulaCharacters(expression);
    return { expression, kind: 'normal' };
  }

  private formulaRange(
    value: string | undefined,
    cell: PendingCell,
  ): XlsxRange {
    const range = parseXlsxRangeReference(value);
    if (!range) {
      formulaFailure(this.part, cell.address, 'Formula range is invalid');
    }
    return range;
  }

  private formulaSharedIndex(
    value: string | undefined,
    cell: PendingCell,
  ): number {
    if (value === undefined || !UNSIGNED_INTEGER_PATTERN.test(value)) {
      formulaFailure(
        this.part,
        cell.address,
        'Shared formula index is invalid',
      );
    }
    const parsed = Number(value);
    if (!Number.isSafeInteger(parsed) || parsed > 0xffff_ffff) {
      formulaFailure(
        this.part,
        cell.address,
        'Shared formula index is invalid',
      );
    }
    return parsed;
  }

  private rangeContains(
    range: XlsxRange,
    row: number,
    column: number,
  ): boolean {
    return row <= range.end.row && column <= range.end.column;
  }

  private consumeFormulaGroup(): void {
    consume(
      this.budget,
      'formulaGroups',
      1,
      'maxFormulaGroups',
      this.limits,
      this.part,
    );
  }

  private consumeFormulaCharacters(expression: string): void {
    if (expression.length > this.limits.maxFormulaCharacters) {
      throw new XlsxResourceLimitError(
        'maxFormulaCharacters',
        expression.length,
        this.limits.maxFormulaCharacters,
        this.part,
      );
    }
    const actual = this.budget.formulaCharacters + expression.length;
    if (
      !Number.isSafeInteger(actual) ||
      actual > this.limits.maxTotalFormulaCharacters
    ) {
      throw new XlsxResourceLimitError(
        'maxTotalFormulaCharacters',
        actual,
        this.limits.maxTotalFormulaCharacters,
        this.part,
      );
    }
    this.budget.formulaCharacters = actual;
  }

  private consumeTextCharacters(value: string): void {
    consume(
      this.budget,
      'textCharacters',
      value.length,
      'maxTextCharacters',
      this.limits,
      this.part,
    );
  }

  private consumeReturnedText(value: XlsxCellValue): void {
    if (value.kind !== 'text') return;
    this.consumeTextCharacters(value.text);
    if (value.runs) {
      consume(
        this.budget,
        'richTextRuns',
        value.runs.length,
        'maxRichTextRuns',
        this.limits,
        this.part,
      );
    }
  }
}

export async function parseXlsxWorksheetPart(
  part: string,
  dialect: XlsxWorkbookDiscovery['dialect'],
  reader: XlsxPartReader,
  limits: ResolvedXlsxResourceLimits,
  sharedStrings: XlsxSharedStringTable,
  budget: XlsxWorksheetBudget,
  selection: XlsxResolvedSheetSelection = { kind: 'full-sheet' },
  semantics: XlsxWorksheetSemantics = DEFAULT_WORKSHEET_SEMANTICS,
): Promise<XlsxWorksheetPayload> {
  const sink = new WorksheetSink(
    part,
    XLSX_SPREADSHEET_NAMESPACES[dialect],
    sharedStrings,
    budget,
    limits,
    selection,
    semantics,
  );
  await reader.streamXml(part, sink, { required: true });
  return sink.result();
}
