import type {
  PptxSceneTableBorder,
  PptxSceneTableCell,
  PptxSceneTableElement,
  PptxSceneTextNode,
} from '../scene-types';
import type { PptxDocument, Table } from '../types';

type ParsedTableBorder =
  Table['data'][number][number]['borders'][keyof Table['data'][number][number]['borders']];

export interface PptxTableVerificationDependencies {
  expectedPointValue(value: number): number;
  plainText(html: string): string;
  textNodeValue(node: PptxSceneTextNode): string;
  verifyTransform(
    generated: Table,
    expected: PptxSceneTableElement,
    location: string,
  ): void;
}

function verifyTableBorder(
  generated: ParsedTableBorder,
  expected: PptxSceneTableBorder | undefined,
  location: string,
): void {
  if (expected === undefined) {
    if (generated !== undefined) {
      throw new Error(
        `Generated PowerPoint table border mismatch at ${location}`,
      );
    }
    return;
  }
  if (
    generated?.borderColor !== expected.color ||
    generated.borderWidth !== expected.width ||
    generated.borderType !== (expected.style ?? 'solid')
  ) {
    throw new Error(
      `Generated PowerPoint table border mismatch at ${location}`,
    );
  }
}

function verifyTableCell(
  generated: Table['data'][number][number] | undefined,
  expected: PptxSceneTableCell,
  location: string,
  dependencies: PptxTableVerificationDependencies,
): void {
  if (generated === undefined) {
    throw new Error(`Generated PowerPoint table cell missing at ${location}`);
  }
  if (
    generated.fillColor !== expected.fillColor ||
    generated.colSpan !== expected.colSpan ||
    generated.rowSpan !== expected.rowSpan ||
    generated.hMerge !== (expected.hMerge ? 1 : undefined) ||
    generated.vMerge !== (expected.vMerge ? 1 : undefined)
  ) {
    throw new Error(`Generated PowerPoint table cell mismatch at ${location}`);
  }
  if (!expected.hMerge && !expected.vMerge) {
    const expectedText = expected.text.paragraphs
      .map((paragraph) =>
        paragraph.children
          .map((node) => dependencies.textNodeValue(node))
          .join(''),
      )
      .join('\n');
    if (dependencies.plainText(generated.text) !== expectedText) {
      throw new Error(
        `Generated PowerPoint table text mismatch at ${location}`,
      );
    }
  }
  for (const key of ['bottom', 'left', 'right', 'top'] as const) {
    verifyTableBorder(
      generated.borders[key],
      expected.borders?.[key],
      `${location}, ${key} border`,
    );
  }
}

export function verifyPowerPointTableElement(
  generated: PptxDocument['slides'][number]['elements'][number] | undefined,
  expected: PptxSceneTableElement,
  slideIndex: number,
  elementIndex: number,
  dependencies: PptxTableVerificationDependencies,
): void {
  const location = `slide ${slideIndex + 1}, element ${elementIndex + 1}`;
  if (generated?.type !== 'table') {
    throw new Error(`Generated PowerPoint table missing at ${location}`);
  }
  dependencies.verifyTransform(generated, expected, location);
  const expectedColumns = expected.columns.map((value) =>
    dependencies.expectedPointValue(value),
  );
  const expectedRows = expected.rows.map((row) =>
    dependencies.expectedPointValue(row.height),
  );
  if (
    JSON.stringify(generated.colWidths) !== JSON.stringify(expectedColumns) ||
    JSON.stringify(generated.rowHeights) !== JSON.stringify(expectedRows) ||
    generated.data.length !== expected.rows.length
  ) {
    throw new Error(`Generated PowerPoint table grid mismatch at ${location}`);
  }
  expected.rows.forEach((row, rowIndex) => {
    const generatedRow = generated.data[rowIndex];
    if (generatedRow?.length !== row.cells.length) {
      throw new Error(
        `Generated PowerPoint table row mismatch at ${location}, row ${rowIndex + 1}`,
      );
    }
    row.cells.forEach((cell, cellIndex) =>
      verifyTableCell(
        generatedRow[cellIndex],
        cell,
        `${location}, row ${rowIndex + 1}, cell ${cellIndex + 1}`,
        dependencies,
      ),
    );
  });
}
