import { PptxWriteError } from '../write-error';

export function unsupportedPptxEdit(message: string, cause?: unknown): never {
  throw new PptxWriteError('unsupported-edit-operation', message, { cause });
}
