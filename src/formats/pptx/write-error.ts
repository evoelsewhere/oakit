import type { PptxSceneValidationIssue } from './scene-types';
import type { PptxWriteErrorCode, PptxWriteErrorOptions } from './write-types';

function freezeIssues(
  issues: readonly PptxSceneValidationIssue[],
): readonly PptxSceneValidationIssue[] {
  return Object.freeze(issues.map((issue) => Object.freeze({ ...issue })));
}

export class PptxWriteError extends Error {
  readonly code: PptxWriteErrorCode;
  readonly issues: readonly PptxSceneValidationIssue[];

  constructor(
    code: PptxWriteErrorCode,
    message: string,
    options: PptxWriteErrorOptions = {},
  ) {
    super(message, options);
    this.name = 'PptxWriteError';
    this.code = code;
    this.issues = freezeIssues(options.issues ?? []);
  }
}
