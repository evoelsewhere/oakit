import { describe, expect, it } from 'vitest';

import { parsePptx, parsePptxWithDiagnostics, PptxParseError } from '../../src';
import { createIndependentPptx, PRESENTATION_NS } from './pptx-package';

describe('PPTX package validity at the public boundary', () => {
  it('rejects non-ZIP input with a typed package diagnostic', async () => {
    const input = new TextEncoder().encode('this is not an OPC package');

    try {
      await parsePptx(input);
      throw new Error('Expected non-ZIP input to reject');
    } catch (caught) {
      expect(caught).toBeInstanceOf(PptxParseError);
      if (!(caught instanceof PptxParseError)) throw caught;
      expect(caught.diagnostic).toMatchObject({
        code: 'invalid-package',
        severity: 'error',
      });
      expect(caught.diagnostic.message).toMatch(
        /^Failed to open OPC package: .+/,
      );
      expect(caught.cause).toBeInstanceOf(Error);
    }
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

  it('reports a missing presentation as a missing required part', async () => {
    const input = await createIndependentPptx({
      'ppt/presentation.xml': null,
    });

    const tolerant = await parsePptxWithDiagnostics(input);

    expect(tolerant.document.slides).toEqual([]);
    expect(tolerant.diagnostics).toContainEqual({
      code: 'missing-required-part',
      message: 'Required OOXML part is missing: ppt/presentation.xml',
      part: 'ppt/presentation.xml',
      severity: 'error',
    });
    await expect(
      parsePptx(input, { errorMode: 'strict' }),
    ).rejects.toMatchObject({
      diagnostic: {
        code: 'missing-required-part',
        message: 'Required OOXML part is missing: ppt/presentation.xml',
        part: 'ppt/presentation.xml',
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
