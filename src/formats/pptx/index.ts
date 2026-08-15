import { parse } from './parser';
import type { PptxDocument, PptxInput, PptxParseOptions } from './types';

export { PptxParseError } from './errors';

/** Parse a PowerPoint Open XML package into the current structured JSON model. */
export async function parsePptx(
  input: PptxInput,
  options: PptxParseOptions = {},
): Promise<PptxDocument> {
  return parse(input, options);
}

export type * from './types';
