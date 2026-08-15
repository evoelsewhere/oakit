import { describe, expect, it } from 'vitest';

import { parsePptx, parsePptxWithDiagnostics } from '../../src';
import { createMinimalPptx } from './fixture';

describe('parsePptxWithDiagnostics', () => {
  it('returns a partial document with warnings for invalid optional XML', async () => {
    const input = await createMinimalPptx({
      'ppt/theme/theme1.xml': '<a:theme><a:themeElements></a:theme>',
    });

    const result = await parsePptxWithDiagnostics(input);

    expect(result.document.slides).toHaveLength(1);
    expect(result.diagnostics).toMatchObject([
      {
        code: 'xml-parse-failed',
        part: 'ppt/theme/theme1.xml',
        severity: 'warning',
      },
    ]);
  });

  it('reports missing required package parts', async () => {
    const input = await createMinimalPptx({ '[Content_Types].xml': null });

    const result = await parsePptxWithDiagnostics(input);

    expect(result.document.slides).toEqual([]);
    expect(result.diagnostics).toContainEqual(
      expect.objectContaining({
        code: 'missing-required-part',
        part: '[Content_Types].xml',
        severity: 'error',
      }),
    );
  });

  it('rejects invalid XML with a typed error in strict mode', async () => {
    const input = await createMinimalPptx({
      'ppt/theme/theme1.xml': '<a:theme><a:themeElements></a:theme>',
    });

    await expect(
      parsePptx(input, { errorMode: 'strict' }),
    ).rejects.toMatchObject({
      name: 'PptxParseError',
      diagnostic: {
        code: 'xml-parse-failed',
        part: 'ppt/theme/theme1.xml',
      },
    });
  });
});
