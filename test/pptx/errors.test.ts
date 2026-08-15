import { describe, expect, it } from 'vitest';

import { PptxParseError } from '../../src/formats/pptx/errors';
import type { PptxDiagnostic } from '../../src/formats/pptx/types';

describe('PPTX parse error contract', () => {
  it('preserves the exact diagnostic object and public Error fields', () => {
    const diagnostic: PptxDiagnostic = {
      code: 'missing-required-part',
      message: 'Required OOXML part is missing: ppt/presentation.xml',
      part: 'ppt/presentation.xml',
      severity: 'error',
    };

    const error = new PptxParseError(diagnostic);

    expect(error).toBeInstanceOf(Error);
    expect(error.name).toBe('PptxParseError');
    expect(error.message).toBe(diagnostic.message);
    expect(error.diagnostic).toBe(diagnostic);
    expect(error.cause).toBeUndefined();
  });

  it('preserves an underlying failure as the standard Error cause', () => {
    const cause = new Error('invalid ZIP directory');
    const diagnostic: PptxDiagnostic = {
      code: 'invalid-package',
      message: 'Unable to read the PowerPoint package',
      severity: 'error',
    };

    const error = new PptxParseError(diagnostic, { cause });

    expect(error.cause).toBe(cause);
    expect(error.diagnostic).toBe(diagnostic);
  });
});
