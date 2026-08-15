import type { PptxDiagnostic } from './types';

export class PptxParseError extends Error {
  readonly diagnostic: PptxDiagnostic;

  constructor(diagnostic: PptxDiagnostic, options?: ErrorOptions) {
    super(diagnostic.message, options);
    this.name = 'PptxParseError';
    this.diagnostic = diagnostic;
  }
}
