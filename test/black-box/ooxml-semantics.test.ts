import { describe, expect, it } from 'vitest';

import { parsePptx, parsePptxWithDiagnostics } from '../../src';
import {
  createIndependentPptx,
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
    ['non-finite', '1e309', '5143500'],
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

  it('reports a valid XML part with the wrong presentation root', async () => {
    const input = await createIndependentPptx({
      'ppt/presentation.xml': `<p:notPresentation xmlns:p="${PRESENTATION_NS}"/>`,
    });

    const tolerant = await parsePptxWithDiagnostics(input);

    expect(tolerant.document.slides).toHaveLength(0);
    expect(tolerant.diagnostics).toContainEqual(
      expect.objectContaining({
        code: 'invalid-document-structure',
        part: 'ppt/presentation.xml',
        severity: 'error',
      }),
    );
    await expect(
      parsePptx(input, { errorMode: 'strict' }),
    ).rejects.toMatchObject({
      diagnostic: {
        code: 'invalid-document-structure',
        part: 'ppt/presentation.xml',
      },
    });
  });
});
