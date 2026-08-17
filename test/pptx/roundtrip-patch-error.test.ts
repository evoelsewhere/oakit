import { describe, expect, it } from 'vitest';

import { unsupportedPptxEdit } from '../../src/formats/pptx/roundtrip/patch-error';
import { PptxWriteError } from '../../src/formats/pptx/write-error';

describe('PowerPoint unsupported edit errors', () => {
  it('retains the typed code, message, and cause', () => {
    const cause = new Error('source');

    try {
      unsupportedPptxEdit('unsupported', cause);
    } catch (error) {
      expect(error).toBeInstanceOf(PptxWriteError);
      if (!(error instanceof PptxWriteError)) throw error;
      expect(error.cause).toBe(cause);
      expect(error.code).toBe('unsupported-edit-operation');
      expect(error.message).toBe('unsupported');
      expect(error.name).toBe('PptxWriteError');
    }
  });
});
