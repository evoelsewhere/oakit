import type { PptxRenderErrorCode } from './render-types';

export class PptxRenderError extends Error {
  readonly code: PptxRenderErrorCode;

  constructor(
    code: PptxRenderErrorCode,
    message: string,
    options?: ErrorOptions,
  ) {
    super(message, options);
    this.name = 'PptxRenderError';
    this.code = code;
  }
}
