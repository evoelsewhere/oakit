import fc from 'fast-check';
import { describe, expect, it } from 'vitest';

import { parsePptx, PptxParseError } from '../../src';
import {
  createIndependentPptx,
  OFFICE_REL_NS,
  OFFICE_REL_TYPE,
  PACKAGE_REL_NS,
  PRESENTATION_NS,
} from './pptx-package';

const FUZZ_SEED = 0x02_0f_24;
const FUZZ_RUNS = 40;

function escapeXmlAttribute(value: string): string {
  return value
    .replaceAll('&', '&amp;')
    .replaceAll('"', '&quot;')
    .replaceAll('<', '&lt;');
}

function presentationWithWidth(width: string): string {
  return `<p:presentation xmlns:p="${PRESENTATION_NS}" xmlns:r="${OFFICE_REL_NS}">
    <p:sldIdLst><p:sldId id="256" r:id="rIdSlide1"/></p:sldIdLst>
    <p:sldSz cx="${escapeXmlAttribute(width)}" cy="5143500"/>
  </p:presentation>`;
}

function expectOnlyFiniteNumbers(value: unknown): void {
  if (typeof value === 'number') {
    expect(Number.isFinite(value)).toBe(true);
    return;
  }
  if (!value || typeof value !== 'object') return;
  for (const child of Object.values(value)) expectOnlyFiniteNumbers(child);
}

describe('PPTX seeded public-boundary properties', () => {
  it('rejects arbitrary non-positive-integer slide widths', async () => {
    const invalidWidth = fc
      .string({ minLength: 1, maxLength: 18 })
      .filter((value) => !/^\+?[1-9]\d*$/.test(value));

    await fc.assert(
      fc.asyncProperty(invalidWidth, async (width) => {
        const input = await createIndependentPptx({
          'ppt/presentation.xml': presentationWithWidth(width),
        });

        await expect(
          parsePptx(input, { errorMode: 'strict' }),
        ).rejects.toMatchObject({
          diagnostic: {
            code: 'invalid-document-value',
            part: 'ppt/presentation.xml',
          },
        });
      }),
      { numRuns: FUZZ_RUNS, seed: FUZZ_SEED },
    );
  });

  it('rejects generated relationship targets that escape the package root', async () => {
    await fc.assert(
      fc.asyncProperty(
        fc.integer({ min: 2, max: 12 }),
        fc.constantFrom('/', '\\'),
        async (depth, separator) => {
          const target = `${`..${separator}`.repeat(depth)}outside.xml`;
          const input = await createIndependentPptx({
            'ppt/_rels/presentation.xml.rels': `
              <Relationships xmlns="${PACKAGE_REL_NS}">
                <Relationship Id="rIdSlide1" Type="${OFFICE_REL_TYPE}slide" Target="${target}"/>
              </Relationships>`,
          });

          await expect(
            parsePptx(input, { errorMode: 'strict' }),
          ).rejects.toMatchObject({
            diagnostic: {
              code: 'invalid-relationship-target',
              part: 'ppt/presentation.xml',
            },
          });
        },
      ),
      { numRuns: FUZZ_RUNS, seed: FUZZ_SEED + 1 },
    );
  });

  it('rejects generated mismatched XML closing names', async () => {
    const localName = fc.stringMatching(/^[A-Za-z][A-Za-z0-9]{0,15}$/);

    await fc.assert(
      fc.asyncProperty(localName, async (name) => {
        const input = await createIndependentPptx({
          'ppt/presentation.xml': presentationWithWidth('9144000').replace(
            '</p:sldIdLst>',
            `</p:${name}>`,
          ),
        });

        await expect(
          parsePptx(input, { errorMode: 'strict' }),
        ).rejects.toMatchObject({
          diagnostic: {
            code: 'xml-parse-failed',
            part: 'ppt/presentation.xml',
          },
        });
      }),
      { numRuns: FUZZ_RUNS, seed: FUZZ_SEED + 2 },
    );
  });

  it('never returns non-finite numbers after random single-byte ZIP mutations', async () => {
    const baseline = await createIndependentPptx();

    await fc.assert(
      fc.asyncProperty(
        fc.integer({ min: 0, max: baseline.byteLength - 1 }),
        fc.integer({ min: 1, max: 255 }),
        async (index, mask) => {
          const mutated = baseline.slice();
          mutated[index] = mutated[index]! ^ mask;

          try {
            const result = await parsePptx(mutated, { errorMode: 'strict' });
            expectOnlyFiniteNumbers(result);
          } catch (error) {
            expect(error).toBeInstanceOf(PptxParseError);
            if (!(error instanceof PptxParseError)) throw error;
            expect(error.name).toBe('PptxParseError');
            expect(typeof error.diagnostic.code).toBe('string');
            expect(error.diagnostic.severity).toMatch(/^(?:error|warning)$/);
          }
        },
      ),
      { numRuns: FUZZ_RUNS, seed: FUZZ_SEED + 3 },
    );
  });
});
