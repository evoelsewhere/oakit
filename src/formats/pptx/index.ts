import { parse } from './parser';
import type {
  PptxDocument,
  PptxInput,
  PptxParseOptions,
  PptxParseResult,
} from './types';

export { validatePptxScene } from './scene-validation';

export { PptxParseError } from './errors';

/** Parse a PowerPoint Open XML package into the current structured JSON model. */
export async function parsePptx(
  input: PptxInput,
  options: PptxParseOptions = {},
): Promise<PptxDocument> {
  return parse(input, options);
}

/** Parse a PowerPoint package and return recoverable diagnostics. */
export async function parsePptxWithDiagnostics(
  input: PptxInput,
  options: PptxParseOptions = {},
): Promise<PptxParseResult> {
  const diagnostics: PptxParseResult['diagnostics'] = [];
  const document = await parse(input, options, diagnostics);
  return { document, diagnostics };
}

export type * from './types';
export type * from './scene-types';
