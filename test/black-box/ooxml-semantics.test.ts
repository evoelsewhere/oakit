import { describe, expect, it } from 'vitest';

import { parsePptx, parsePptxWithDiagnostics } from '../../src';
import {
  createIndependentPptx,
  independentTextSlide,
  OFFICE_REL_NS,
  PRESENTATION_NS,
} from './pptx-package';

function presentationWithSize(width: string, height: string): string {
  return `<p:presentation xmlns:p="${PRESENTATION_NS}" xmlns:r="${OFFICE_REL_NS}">
    <p:sldIdLst><p:sldId id="256" r:id="rIdSlide1"/></p:sldIdLst>
    <p:sldSz cx="${width}" cy="${height}"/>
  </p:presentation>`;
}

describe('PPTX semantic validity at the public boundary', () => {
  it.each([
    ['non-numeric', 'not-a-number', '5143500'],
    ['numeric suffix', '9144000px', '5143500'],
    ['leading zero', '09144000', '5143500'],
    ['surrounding whitespace', ' 9144000 ', '5143500'],
    ['decimal notation', '9144000.0', '5143500'],
    ['non-finite', '1e309', '5143500'],
    ['unsafe integer', '9007199254740992', '5143500'],
    ['zero', '0', '5143500'],
    ['negative', '-9144000', '5143500'],
  ])(
    'reports a %s presentation size without returning invalid numbers',
    async (_, width, height) => {
      const input = await createIndependentPptx({
        'ppt/presentation.xml': presentationWithSize(width, height),
      });

      const tolerant = await parsePptxWithDiagnostics(input);

      expect(tolerant.document.size).toEqual({ width: 0, height: 0 });
      expect(tolerant.diagnostics).toContainEqual(
        expect.objectContaining({
          code: 'invalid-document-value',
          part: 'ppt/presentation.xml',
          severity: 'error',
        }),
      );
      await expect(
        parsePptx(input, { errorMode: 'strict' }),
      ).rejects.toMatchObject({
        diagnostic: {
          code: 'invalid-document-value',
          part: 'ppt/presentation.xml',
        },
      });
    },
  );

  it('accepts canonical positive coordinates with an optional plus sign', async () => {
    const input = await createIndependentPptx({
      'ppt/presentation.xml': presentationWithSize('+9144000', '+5143500'),
    });

    const document = await parsePptx(input, { errorMode: 'strict' });

    expect(document.size).toEqual({ width: 720, height: 405 });
  });

  it('reports a valid XML part with the wrong presentation root', async () => {
    const input = await createIndependentPptx({
      'ppt/presentation.xml': `<p:notPresentation xmlns:p="${PRESENTATION_NS}"/>`,
    });

    const tolerant = await parsePptxWithDiagnostics(input);

    expect(tolerant.document.slides).toHaveLength(0);
    expect(tolerant.diagnostics).toContainEqual(
      expect.objectContaining({
        code: 'invalid-document-structure',
        message:
          'Required OOXML root p:presentation is missing from ppt/presentation.xml',
        part: 'ppt/presentation.xml',
        severity: 'error',
      }),
    );
    await expect(
      parsePptx(input, { errorMode: 'strict' }),
    ).rejects.toMatchObject({
      diagnostic: {
        code: 'invalid-document-structure',
        message:
          'Required OOXML root p:presentation is missing from ppt/presentation.xml',
        part: 'ppt/presentation.xml',
      },
    });
  });

  it('reports a missing required presentation size', async () => {
    const input = await createIndependentPptx({
      'ppt/presentation.xml': `<p:presentation xmlns:p="${PRESENTATION_NS}" xmlns:r="${OFFICE_REL_NS}">
        <p:sldIdLst><p:sldId id="256" r:id="rIdSlide1"/></p:sldIdLst>
      </p:presentation>`,
    });

    const tolerant = await parsePptxWithDiagnostics(input);

    expect(tolerant.document.size).toEqual({ width: 0, height: 0 });
    expect(tolerant.diagnostics).toContainEqual(
      expect.objectContaining({
        code: 'invalid-document-value',
        part: 'ppt/presentation.xml',
      }),
    );
  });

  it('never returns non-finite numbers from malformed shape coordinates', async () => {
    const input = await createIndependentPptx({
      'ppt/slides/slide1.xml': independentTextSlide('Bad coordinate').replace(
        'cx="914400"',
        'cx="not-a-number"',
      ),
    });

    const tolerant = await parsePptxWithDiagnostics(input);
    const element = tolerant.document.slides[0]?.elements[0];

    expect(element?.width).toBe(0);
    expect(tolerant.diagnostics).toContainEqual(
      expect.objectContaining({
        code: 'invalid-document-value',
        message:
          'Non-finite numeric values were replaced while parsing ppt/slides/slide1.xml',
        part: 'ppt/slides/slide1.xml',
      }),
    );
    await expect(
      parsePptx(input, { errorMode: 'strict' }),
    ).rejects.toMatchObject({
      diagnostic: {
        code: 'invalid-document-value',
        message:
          'Non-finite numeric values were replaced while parsing ppt/slides/slide1.xml',
        part: 'ppt/slides/slide1.xml',
      },
    });
  });
});
