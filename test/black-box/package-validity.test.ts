import { describe, expect, it } from 'vitest';

import { parsePptx, parsePptxWithDiagnostics } from '../../src';
import { createIndependentPptx, PRESENTATION_NS } from './pptx-package';

describe('PPTX package validity at the public boundary', () => {
  it('rejects non-ZIP input with a typed package diagnostic', async () => {
    const input = new TextEncoder().encode('this is not an OPC package');

    await expect(parsePptx(input)).rejects.toMatchObject({
      name: 'PptxParseError',
      diagnostic: {
        code: 'invalid-package',
        severity: 'error',
      },
    });
  });

  it('reports a content-types part with the wrong root', async () => {
    const input = await createIndependentPptx({
      '[Content_Types].xml': '<NotTypes/>',
    });

    const tolerant = await parsePptxWithDiagnostics(input);

    expect(tolerant.document.slides).toHaveLength(0);
    expect(tolerant.diagnostics).toContainEqual(
      expect.objectContaining({
        code: 'invalid-document-structure',
        part: '[Content_Types].xml',
        severity: 'error',
      }),
    );
    await expect(
      parsePptx(input, { errorMode: 'strict' }),
    ).rejects.toMatchObject({
      diagnostic: {
        code: 'invalid-document-structure',
        part: '[Content_Types].xml',
      },
    });
  });

  it('reports a selected slide part with the wrong root', async () => {
    const input = await createIndependentPptx({
      'ppt/slides/slide1.xml': `<p:notSlide xmlns:p="${PRESENTATION_NS}"/>`,
    });

    const tolerant = await parsePptxWithDiagnostics(input);

    expect(tolerant.document.slides).toHaveLength(1);
    expect(tolerant.document.slides[0]?.elements).toHaveLength(0);
    expect(tolerant.diagnostics).toContainEqual(
      expect.objectContaining({
        code: 'invalid-document-structure',
        part: 'ppt/slides/slide1.xml',
        severity: 'error',
      }),
    );
    await expect(
      parsePptx(input, { errorMode: 'strict' }),
    ).rejects.toMatchObject({
      diagnostic: {
        code: 'invalid-document-structure',
        part: 'ppt/slides/slide1.xml',
      },
    });
  });
});
