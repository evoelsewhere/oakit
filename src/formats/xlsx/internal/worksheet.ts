import { XlsxParseError } from '../errors';
import type {
  XlsxCell,
  XlsxCellValue,
  XlsxRichTextRun,
  XlsxRow,
} from '../types';
import { parseXlsxCellReference, xlsxColumnName } from './cell-reference';
import {
  parseXlsxScalarCellValue,
  type XlsxScalarCellType,
} from './cell-value';
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
import type { XlsxXmlElement, XlsxXmlEventSink } from './streaming-xml';
import {
  type XlsxWorkbookDiscovery,
  XLSX_SPREADSHEET_NAMESPACES,
} from './workbook-discovery';

const XML_NAMESPACE = 'http://www.w3.org/XML/1998/namespace';
const UNSIGNED_INTEGER_PATTERN = /^(?:0|[1-9]\d*)$/u;
const NONNEGATIVE_DECIMAL_PATTERN = /^(?:\d+(?:\.\d*)?|\.\d+)$/u;

type CellType = XlsxScalarCellType | 'inlineStr';
type TextCapture = 'inline' | 'value' | null;

interface PendingCell {
  address: string;
  column: number;
  hasInlineString: boolean;
  hasValue: boolean;
  inlineMode: 'plain' | 'rich' | 'unset';
  inlineRuns: XlsxRichTextRun[];
  inlineText: string;
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
  returnedCells: number;
  richTextRuns: number;
  scannedCells: number;
  textCharacters: number;
}

export interface XlsxWorksheetPayload {
  rows: XlsxRow[];
}

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

function unsupportedFormula(part: string, cell: string): never {
  throw new XlsxParseError({
    cell,
    code: 'unsupported-feature',
    message: 'XLSX formula cells are not supported yet',
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
  return { returnedCells: 0, richTextRuns, scannedCells: 0, textCharacters };
}

class WorksheetSink implements XlsxXmlEventSink {
  private capture: TextCapture = null;
  private currentCell: PendingCell | undefined;
  private currentInlineRun: PendingInlineRun | undefined;
  private currentRow: XlsxRow | undefined;
  private currentRowSelected!: boolean;
  private ignoredDepth = 0;
  private lastCellColumn = 0;
  private lastRow = 0;
  private sheetDataSeen = false;
  private readonly stack: XlsxXmlElement[] = [];
  private readonly rows: XlsxRow[] = [];

  constructor(
    private readonly part: string,
    private readonly namespace: string,
    private readonly sharedStrings: XlsxSharedStringTable,
    private readonly budget: XlsxWorksheetBudget,
    private readonly limits: ResolvedXlsxResourceLimits,
    private readonly selection: XlsxResolvedSheetSelection,
  ) {}

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
    this.stack.pop();
  }

  text(value: string): void {
    if (this.ignoredDepth > 0) return;
    if (this.capture === 'value') {
      this.currentCell!.valueText += value;
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
    return { rows: this.rows };
  }

  private beginIgnore(): void {
    this.ignoredDepth = 1;
  }

  private openChild(parent: string, element: XlsxXmlElement): void {
    if (parent === 'worksheet') {
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
      this.beginIgnore();
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
        unsupportedFormula(this.part, this.currentCell!.address);
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
    const hidden = optionalBoolean(
      attribute(element, 'hidden'),
      this.part,
      'Worksheet row hidden flag is invalid',
    );
    this.currentRow = {
      cells: [],
      ...(height === undefined ? {} : { height }),
      ...(hidden === undefined ? {} : { hidden }),
      index,
      ...(outlineLevel === undefined ? {} : { outlineLevel }),
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
    const style = unsignedInteger(
      attribute(element, 's'),
      this.part,
      address,
      'Worksheet cell style index is invalid',
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
      selected,
      ...(style === undefined ? {} : { style }),
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
    if (!cell.hasValue && !cell.hasInlineString) {
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
      const value = parseXlsxScalarCellValue(
        cell.type as XlsxScalarCellType,
        cell.valueText,
        this.sharedStrings,
        this.part,
        cell.address,
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
    } else {
      this.currentRow!.cells.push({
        ...base,
        content: { kind: 'value', value: content.value },
      });
    }
    this.currentCell = undefined;
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
): Promise<XlsxWorksheetPayload> {
  const sink = new WorksheetSink(
    part,
    XLSX_SPREADSHEET_NAMESPACES[dialect],
    sharedStrings,
    budget,
    limits,
    selection,
  );
  await reader.streamXml(part, sink, { required: true });
  return sink.result();
}
