import type { PptxSceneValidationIssue } from './scene-types';
import type {
  PptxTableValidationDependencies,
  PptxTableValidationProfile,
} from './scene-table-validation';

const BORDER_KEYS = ['bottom', 'left', 'right', 'top'] as const;
const SPAN_KEYS = ['colSpan', 'rowSpan'] as const;

function validateTableBorder(
  value: unknown,
  path: string,
  profile: PptxTableValidationProfile,
  issues: PptxSceneValidationIssue[],
  emusPerPoint: number,
  dependencies: PptxTableValidationDependencies,
): void {
  const border = dependencies.requireObject(value, path, issues);
  if (!border) return;
  dependencies.rejectUnknownKeys(
    border,
    ['color', 'style', 'width'],
    path,
    issues,
  );
  dependencies.optionalColor(border, 'color', path, issues);
  if (border.color === undefined) {
    dependencies.addIssue(
      issues,
      'invalid-scene-document',
      `${path}.color`,
      'Expected a #RRGGBB color',
    );
  }
  if (
    border.style !== undefined &&
    border.style !== 'dashed' &&
    border.style !== 'dotted' &&
    border.style !== 'solid'
  ) {
    dependencies.addIssue(
      issues,
      'invalid-scene-document',
      `${path}.style`,
      'Unknown table border style',
    );
  }
  dependencies.requireFiniteNumber(border.width, `${path}.width`, issues, true);
  if (dependencies.isCreationProfile(profile)) {
    dependencies.requireSerializableInteger(
      border.width,
      emusPerPoint,
      `${path}.width`,
      issues,
      true,
    );
  }
}

export function validatePptxSceneTableCell(
  value: unknown,
  path: string,
  profile: PptxTableValidationProfile,
  keys: Set<string>,
  issues: PptxSceneValidationIssue[],
  emusPerPoint: number,
  dependencies: PptxTableValidationDependencies,
): void {
  const cell = dependencies.requireObject(value, path, issues);
  if (!cell) return;
  dependencies.rejectUnknownKeys(
    cell,
    ['borders', 'colSpan', 'fillColor', 'hMerge', 'rowSpan', 'text', 'vMerge'],
    path,
    issues,
  );
  dependencies.optionalColor(cell, 'fillColor', path, issues);
  dependencies.optionalBoolean(cell, 'hMerge', path, issues);
  dependencies.optionalBoolean(cell, 'vMerge', path, issues);
  for (const key of SPAN_KEYS) {
    if (
      cell[key] !== undefined &&
      (!Number.isSafeInteger(cell[key]) || Number(cell[key]) <= 1)
    ) {
      dependencies.addIssue(
        issues,
        'invalid-numeric-value',
        `${path}.${key}`,
        'Table span must be an integer greater than one',
      );
    }
  }
  if (cell.borders !== undefined) {
    const borders = dependencies.requireObject(
      cell.borders,
      `${path}.borders`,
      issues,
    );
    if (borders) {
      dependencies.rejectUnknownKeys(
        borders,
        BORDER_KEYS,
        `${path}.borders`,
        issues,
      );
      for (const key of BORDER_KEYS) {
        if (borders[key] !== undefined) {
          validateTableBorder(
            borders[key],
            `${path}.borders.${key}`,
            profile,
            issues,
            emusPerPoint,
            dependencies,
          );
        }
      }
    }
  }
  dependencies.validateTextBody(
    cell.text,
    `${path}.text`,
    profile,
    keys,
    issues,
  );
}
