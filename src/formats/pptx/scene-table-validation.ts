import type {
  PptxSceneValidationCode,
  PptxSceneValidationIssue,
} from './scene-types';
import { validatePptxSceneTableCell } from './scene-table-cell-validation';
import { validatePptxSceneTableMerges } from './scene-table-merge-validation';

export type PptxTableValidationObject = Record<string, unknown>;
export type PptxTableValidationProfile =
  'create-native-v1' | 'create-text-v1' | 'scene';

export interface PptxTableValidationDependencies {
  addIssue(
    issues: PptxSceneValidationIssue[],
    code: PptxSceneValidationCode,
    path: string,
    message: string,
  ): void;
  isCreationProfile(profile: PptxTableValidationProfile): boolean;
  isObject(value: unknown): value is PptxTableValidationObject;
  optionalBoolean(
    value: PptxTableValidationObject,
    key: string,
    path: string,
    issues: PptxSceneValidationIssue[],
  ): void;
  optionalColor(
    value: PptxTableValidationObject,
    key: string,
    path: string,
    issues: PptxSceneValidationIssue[],
  ): void;
  rejectUnknownKeys(
    value: PptxTableValidationObject,
    allowed: readonly string[],
    path: string,
    issues: PptxSceneValidationIssue[],
  ): void;
  requireArray(
    value: unknown,
    path: string,
    issues: PptxSceneValidationIssue[],
  ): unknown[] | undefined;
  requireFiniteNumber(
    value: unknown,
    path: string,
    issues: PptxSceneValidationIssue[],
    positive: boolean,
  ): void;
  requireObject(
    value: unknown,
    path: string,
    issues: PptxSceneValidationIssue[],
  ): PptxTableValidationObject | undefined;
  requireSerializableInteger(
    value: unknown,
    multiplier: number,
    path: string,
    issues: PptxSceneValidationIssue[],
    positive: boolean,
  ): void;
  validateTextBody(
    value: unknown,
    path: string,
    profile: PptxTableValidationProfile,
    keys: Set<string>,
    issues: PptxSceneValidationIssue[],
  ): void;
}

export function validatePptxSceneTable(
  value: PptxTableValidationObject,
  path: string,
  profile: PptxTableValidationProfile,
  keys: Set<string>,
  issues: PptxSceneValidationIssue[],
  emusPerPoint: number,
  dependencies: PptxTableValidationDependencies,
): void {
  const columns = dependencies.requireArray(
    value.columns,
    `${path}.columns`,
    issues,
  );
  if (columns?.length === 0) {
    dependencies.addIssue(
      issues,
      'invalid-scene-document',
      `${path}.columns`,
      'A table needs at least one column',
    );
  }
  columns?.forEach((width, index) => {
    const widthPath = `${path}.columns[${index}]`;
    dependencies.requireFiniteNumber(width, widthPath, issues, true);
    if (dependencies.isCreationProfile(profile)) {
      dependencies.requireSerializableInteger(
        width,
        emusPerPoint,
        widthPath,
        issues,
        true,
      );
    }
  });
  const rows = dependencies.requireArray(value.rows, `${path}.rows`, issues);
  if (rows?.length === 0) {
    dependencies.addIssue(
      issues,
      'invalid-scene-document',
      `${path}.rows`,
      'A table needs at least one row',
    );
  }
  rows?.forEach((rowValue, rowIndex) => {
    const rowPath = `${path}.rows[${rowIndex}]`;
    const row = dependencies.requireObject(rowValue, rowPath, issues);
    if (!row) return;
    dependencies.rejectUnknownKeys(row, ['cells', 'height'], rowPath, issues);
    dependencies.requireFiniteNumber(
      row.height,
      `${rowPath}.height`,
      issues,
      true,
    );
    if (dependencies.isCreationProfile(profile)) {
      dependencies.requireSerializableInteger(
        row.height,
        emusPerPoint,
        `${rowPath}.height`,
        issues,
        true,
      );
    }
    const cells = dependencies.requireArray(
      row.cells,
      `${rowPath}.cells`,
      issues,
    );
    if (cells && columns && cells.length !== columns.length) {
      dependencies.addIssue(
        issues,
        'invalid-scene-document',
        `${rowPath}.cells`,
        `Table row must contain exactly ${columns.length} grid cells`,
      );
    }
    cells?.forEach((cell, cellIndex) =>
      validatePptxSceneTableCell(
        cell,
        `${rowPath}.cells[${cellIndex}]`,
        profile,
        keys,
        issues,
        emusPerPoint,
        dependencies,
      ),
    );
  });
  if (rows && columns) {
    validatePptxSceneTableMerges(rows, columns, path, issues, dependencies);
  }
}

export function validatePptxSceneTableDimensions(
  element: PptxTableValidationObject,
  transform: PptxTableValidationObject,
  path: string,
  issues: PptxSceneValidationIssue[],
  emusPerPoint: number,
  dependencies: PptxTableValidationDependencies,
): void {
  if (
    element.type !== 'table' ||
    !Array.isArray(element.columns) ||
    !Array.isArray(element.rows)
  ) {
    return;
  }
  if (
    !element.columns.every((value): value is number =>
      Number.isFinite(value as number),
    ) ||
    !element.rows.every(
      (value): value is PptxTableValidationObject =>
        dependencies.isObject(value) && Number.isFinite(value.height as number),
    )
  ) {
    return;
  }
  const columnWidth = element.columns.reduce(
    (total, value) => total + value,
    0,
  );
  const rowHeight = element.rows.reduce(
    (total, value) => total + (value.height as number),
    0,
  );
  if (
    typeof transform.width === 'number' &&
    Math.round(columnWidth * emusPerPoint) !==
      Math.round(transform.width * emusPerPoint)
  ) {
    dependencies.addIssue(
      issues,
      'invalid-scene-document',
      `${path}.authored.transform.width`,
      'Table transform width must equal the sum of its column widths',
    );
  }
  if (
    typeof transform.height === 'number' &&
    Math.round(rowHeight * emusPerPoint) !==
      Math.round(transform.height * emusPerPoint)
  ) {
    dependencies.addIssue(
      issues,
      'invalid-scene-document',
      `${path}.authored.transform.height`,
      'Table transform height must equal the sum of its row heights',
    );
  }
}
