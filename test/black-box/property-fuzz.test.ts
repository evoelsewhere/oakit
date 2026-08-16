import JSZip from 'jszip';
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

async function expectFiniteDocumentOrTypedError(
  input: Uint8Array,
): Promise<void> {
  try {
    expectOnlyFiniteNumbers(await parsePptx(input, { errorMode: 'strict' }));
  } catch (error) {
    expect(error).toBeInstanceOf(PptxParseError);
    if (!(error instanceof PptxParseError)) throw error;
    expect(error.name).toBe('PptxParseError');
    expect(typeof error.diagnostic.code).toBe('string');
    expect(error.diagnostic.severity).toMatch(/^(?:error|warning)$/);
  }
}

describe('PPTX seeded public-boundary properties', () => {
  it('rejects arbitrary non-positive-integer slide widths', async () => {
    const invalidWidth = fc
      .string({ minLength: 1, maxLength: 18 })
      .filter((value) => !/^\+?[1-9]\d*$/.test(value));

    await fc.assert(
      fc.asyncProperty(invalidWidth, async (width) => {
        const input = await createIndependentPptx(
          {
            'ppt/presentation.xml': presentationWithWidth(width),
          },
          { compression: 'STORE' },
        );

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
          const input = await createIndependentPptx(
            {
              'ppt/_rels/presentation.xml.rels': `
              <Relationships xmlns="${PACKAGE_REL_NS}">
                <Relationship Id="rIdSlide1" Type="${OFFICE_REL_TYPE}slide" Target="${target}"/>
              </Relationships>`,
            },
            { compression: 'STORE' },
          );

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

  it('normalizes generated safe relationship dot segments without changing the selected slide', async () => {
    await fc.assert(
      fc.asyncProperty(
        fc.integer({ min: 0, max: 8 }),
        fc.integer({ min: 0, max: 8 }),
        async (leadingDots, internalDots) => {
          const target = `${'./'.repeat(leadingDots)}slides/${'folder/../'.repeat(internalDots)}slide1.xml`;
          const input = await createIndependentPptx(
            {
              'ppt/_rels/presentation.xml.rels': `
              <Relationships xmlns="${PACKAGE_REL_NS}">
                <Relationship Id="rIdSlide1" Type="${OFFICE_REL_TYPE}slide" Target="${target}"/>
              </Relationships>`,
            },
            { compression: 'STORE' },
          );

          const document = await parsePptx(input, { errorMode: 'strict' });

          expect(document.slides).toHaveLength(1);
          expectOnlyFiniteNumbers(document);
        },
      ),
      { numRuns: FUZZ_RUNS, seed: FUZZ_SEED + 2 },
    );
  });

  it('never selects a generated external relationship as an internal slide', async () => {
    const externalTarget = fc
      .webUrl({ validSchemes: ['http', 'https'] })
      .map(escapeXmlAttribute);
    await fc.assert(
      fc.asyncProperty(
        externalTarget,
        fc.constantFrom('External', 'external', 'EXTERNAL'),
        async (target, targetMode) => {
          const input = await createIndependentPptx(
            {
              'ppt/_rels/presentation.xml.rels': `
              <Relationships xmlns="${PACKAGE_REL_NS}">
                <Relationship Id="rIdSlide1" Type="${OFFICE_REL_TYPE}slide" Target="${target}" TargetMode="${targetMode}"/>
              </Relationships>`,
            },
            { compression: 'STORE' },
          );

          const document = await parsePptx(input, { errorMode: 'strict' });

          expect(document.slides).toEqual([]);
        },
      ),
      { numRuns: FUZZ_RUNS, seed: FUZZ_SEED + 3 },
    );
  });

  it('rejects generated mismatched XML closing names', async () => {
    const localName = fc.stringMatching(/^[A-Za-z][A-Za-z0-9]{0,15}$/);

    await fc.assert(
      fc.asyncProperty(localName, async (name) => {
        const input = await createIndependentPptx(
          {
            'ppt/presentation.xml': presentationWithWidth('9144000').replace(
              '</p:sldIdLst>',
              `</p:${name}>`,
            ),
          },
          { compression: 'STORE' },
        );

        await expect(
          parsePptx(input, { errorMode: 'strict' }),
        ).rejects.toMatchObject({
          diagnostic: {
            code: 'xml-parse-failed',
            part: 'ppt/presentation.xml',
          },
        });
      }),
      { numRuns: FUZZ_RUNS, seed: FUZZ_SEED + 4 },
    );
  });

  it('rejects every generated ZIP truncation with a typed package error', async () => {
    const baseline = await createIndependentPptx();

    await fc.assert(
      fc.asyncProperty(
        fc.integer({ min: 0, max: baseline.byteLength - 1 }),
        async (end) => {
          await expect(parsePptx(baseline.slice(0, end))).rejects.toMatchObject(
            {
              name: 'PptxParseError',
              diagnostic: { code: 'invalid-package' },
            },
          );
        },
      ),
      { numRuns: FUZZ_RUNS, seed: FUZZ_SEED + 7 },
    );
  });

  it('returns finite data or a typed error after generated OPC part removal', async () => {
    const baseline = await createIndependentPptx();
    const archive = await JSZip.loadAsync(baseline);
    const partNames = Object.keys(archive.files).filter(
      (name) => !archive.files[name]?.dir,
    );

    await fc.assert(
      fc.asyncProperty(fc.constantFrom(...partNames), async (partName) => {
        const mutatedArchive = await JSZip.loadAsync(baseline);
        mutatedArchive.remove(partName);
        const mutated = await mutatedArchive.generateAsync({
          compression: 'DEFLATE',
          type: 'uint8array',
        });

        await expectFiniteDocumentOrTypedError(mutated);
      }),
      { numRuns: FUZZ_RUNS, seed: FUZZ_SEED + 8 },
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
          const current = mutated[index];
          if (current === undefined)
            throw new Error('Generated invalid ZIP index');
          mutated[index] = current ^ mask;

          await expectFiniteDocumentOrTypedError(mutated);
        },
      ),
      { numRuns: FUZZ_RUNS, seed: FUZZ_SEED + 9 },
    );
  });
});
