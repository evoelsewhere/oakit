import { describe, expect, it } from 'vitest';

import { parsePptx, parsePptxWithDiagnostics } from '../../src';
import {
  createIndependentPptx,
  independentTextSlide,
  OFFICE_REL_TYPE,
  PACKAGE_REL_NS,
} from './pptx-package';

describe('PPTX public API observable contract', () => {
  it('returns identical documents for Uint8Array, ArrayBuffer, and Blob inputs', async () => {
    const bytes = await createIndependentPptx();
    const arrayBuffer = bytes.buffer.slice(
      bytes.byteOffset,
      bytes.byteOffset + bytes.byteLength,
    ) as ArrayBuffer;

    const [fromBytes, fromBuffer, fromBlob] = await Promise.all([
      parsePptx(bytes),
      parsePptx(arrayBuffer),
      parsePptx(new Blob([arrayBuffer])),
    ]);

    expect(fromBuffer).toEqual(fromBytes);
    expect(fromBlob).toEqual(fromBytes);
  });

  it('does not mutate caller-owned input bytes', async () => {
    const input = await createIndependentPptx();
    const snapshot = input.slice();

    await parsePptx(input);

    expect(input).toEqual(snapshot);
  });

  it('keeps concurrent parses isolated and deterministic', async () => {
    const [firstInput, secondInput] = await Promise.all([
      createIndependentPptx({
        'ppt/slides/slide1.xml': independentTextSlide('First parse'),
      }),
      createIndependentPptx({
        'ppt/slides/slide1.xml': independentTextSlide('Second parse'),
      }),
    ]);

    const [first, second, repeated] = await Promise.all([
      parsePptx(firstInput),
      parsePptx(secondInput),
      parsePptx(secondInput),
    ]);

    expect(first.slides[0]?.elements).not.toEqual(second.slides[0]?.elements);
    expect(JSON.stringify(first)).toContain('First&nbsp;parse');
    expect(JSON.stringify(second)).toContain('Second&nbsp;parse');
    expect(second).toEqual(repeated);
  });

  it('reports a malformed optional part in tolerant mode and rejects it in strict mode', async () => {
    const input = await createIndependentPptx({
      'ppt/theme/theme1.xml': '<a:theme><a:broken></a:theme>',
    });

    const tolerant = await parsePptxWithDiagnostics(input);

    expect(tolerant.document.slides).toHaveLength(1);
    expect(tolerant.diagnostics).toEqual([
      expect.objectContaining({
        code: 'xml-parse-failed',
        part: 'ppt/theme/theme1.xml',
        severity: 'warning',
      }),
    ]);
    await expect(
      parsePptx(input, { errorMode: 'strict' }),
    ).rejects.toMatchObject({
      diagnostic: {
        code: 'xml-parse-failed',
        part: 'ppt/theme/theme1.xml',
      },
    });
  });

  it('surfaces an escaping relationship target without reading outside the package', async () => {
    const input = await createIndependentPptx({
      'ppt/_rels/presentation.xml.rels': `
        <Relationships xmlns="${PACKAGE_REL_NS}">
          <Relationship Id="rIdSlide1" Type="${OFFICE_REL_TYPE}slide" Target="../../outside.xml"/>
        </Relationships>`,
    });

    const tolerant = await parsePptxWithDiagnostics(input);

    expect(tolerant.document.slides).toHaveLength(0);
    expect(tolerant.diagnostics).toEqual([
      expect.objectContaining({
        code: 'invalid-relationship-target',
        part: 'ppt/presentation.xml',
        severity: 'warning',
      }),
    ]);
    await expect(
      parsePptx(input, { errorMode: 'strict' }),
    ).rejects.toMatchObject({
      diagnostic: {
        code: 'invalid-relationship-target',
        part: 'ppt/presentation.xml',
      },
    });
  });
});
