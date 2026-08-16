import fc from 'fast-check';
import { describe, expect, it } from 'vitest';

import { parsePptx } from '../../src';
import {
  createIndependentPptx,
  DRAWING_NS,
  independentTextSlide,
  OFFICE_REL_NS,
  PRESENTATION_NS,
} from './pptx-package';

const FUZZ_SEED = 0x58_4d_4c;
const FUZZ_RUNS = 40;

const READ_XML_PARTS = [
  '[Content_Types].xml',
  'ppt/presentation.xml',
  'ppt/_rels/presentation.xml.rels',
  'ppt/theme/theme1.xml',
  'ppt/slides/slide1.xml',
  'ppt/slides/_rels/slide1.xml.rels',
  'ppt/slideLayouts/slideLayout1.xml',
  'ppt/slideLayouts/_rels/slideLayout1.xml.rels',
  'ppt/slideMasters/slideMaster1.xml',
  'ppt/slideMasters/_rels/slideMaster1.xml.rels',
] as const;

function presentationWithPrefix(
  prefix: string,
  relationshipPrefix: string,
): string {
  return `<${prefix}:presentation xmlns:${prefix}="${PRESENTATION_NS}" xmlns:${relationshipPrefix}="${OFFICE_REL_NS}">
    <${prefix}:sldIdLst><${prefix}:sldId id="256" ${relationshipPrefix}:id="rIdSlide1"/></${prefix}:sldIdLst>
    <${prefix}:sldSz cx="9144000" cy="5143500"/>
  </${prefix}:presentation>`;
}

function slideWithPrefixes(
  presentationPrefix: string,
  drawingPrefix: string,
): string {
  return `<${presentationPrefix}:sld xmlns:${presentationPrefix}="${PRESENTATION_NS}" xmlns:${drawingPrefix}="${DRAWING_NS}">
    <${presentationPrefix}:cSld><${presentationPrefix}:spTree>
      <${presentationPrefix}:nvGrpSpPr><${presentationPrefix}:cNvPr id="1" name=""/><${presentationPrefix}:cNvGrpSpPr/><${presentationPrefix}:nvPr/></${presentationPrefix}:nvGrpSpPr>
      <${presentationPrefix}:grpSpPr/>
      <${presentationPrefix}:sp>
        <${presentationPrefix}:nvSpPr><${presentationPrefix}:cNvPr id="2" name="Text"/><${presentationPrefix}:cNvSpPr txBox="1"/><${presentationPrefix}:nvPr/></${presentationPrefix}:nvSpPr>
        <${presentationPrefix}:spPr>
          <${drawingPrefix}:xfrm><${drawingPrefix}:off x="0" y="0"/><${drawingPrefix}:ext cx="914400" cy="914400"/></${drawingPrefix}:xfrm>
          <${drawingPrefix}:prstGeom prst="rect"><${drawingPrefix}:avLst/></${drawingPrefix}:prstGeom>
        </${presentationPrefix}:spPr>
        <${presentationPrefix}:txBody>
          <${drawingPrefix}:bodyPr/><${drawingPrefix}:lstStyle/>
          <${drawingPrefix}:p><${drawingPrefix}:r><${drawingPrefix}:rPr lang="en-US"/><${drawingPrefix}:t>Namespace invariant</${drawingPrefix}:t></${drawingPrefix}:r></${drawingPrefix}:p>
        </${presentationPrefix}:txBody>
      </${presentationPrefix}:sp>
    </${presentationPrefix}:spTree></${presentationPrefix}:cSld>
  </${presentationPrefix}:sld>`;
}

const xmlPrefix = fc
  .stringMatching(/^[A-Za-z][A-Za-z0-9]{0,7}$/)
  .filter((value) => !/^(?:xml|xmlns)$/i.test(value));

describe('PowerPoint generated XML properties', () => {
  it('normalizes arbitrary namespace prefixes to the same public document', async () => {
    const baselineInput = await createIndependentPptx({
      'ppt/slides/slide1.xml': independentTextSlide('Namespace invariant'),
    });
    const baseline = await parsePptx(baselineInput, { errorMode: 'strict' });

    await fc.assert(
      fc.asyncProperty(
        fc
          .tuple(xmlPrefix, xmlPrefix, xmlPrefix)
          .filter((prefixes) => new Set(prefixes).size === prefixes.length),
        async ([presentationPrefix, drawingPrefix, relationshipPrefix]) => {
          const input = await createIndependentPptx({
            'ppt/presentation.xml': presentationWithPrefix(
              presentationPrefix,
              relationshipPrefix,
            ),
            'ppt/slides/slide1.xml': slideWithPrefixes(
              presentationPrefix,
              drawingPrefix,
            ),
          });

          await expect(
            parsePptx(input, { errorMode: 'strict' }),
          ).resolves.toEqual(baseline);
        },
      ),
      { numRuns: FUZZ_RUNS, seed: FUZZ_SEED },
    );
  });

  it('rejects generated XML control characters with a typed part diagnostic', async () => {
    const invalidControlCharacter = fc
      .integer({ min: 0, max: 31 })
      .filter((codePoint) => ![9, 10, 13].includes(codePoint))
      .map(String.fromCharCode);

    await fc.assert(
      fc.asyncProperty(invalidControlCharacter, async (character) => {
        const input = await createIndependentPptx({
          'ppt/slides/slide1.xml': independentTextSlide(
            `Before${character}After`,
          ),
        });

        await expect(
          parsePptx(input, { errorMode: 'strict' }),
        ).rejects.toMatchObject({
          name: 'PptxParseError',
          diagnostic: {
            code: 'xml-parse-failed',
            part: 'ppt/slides/slide1.xml',
          },
        });
      }),
      { numRuns: FUZZ_RUNS, seed: FUZZ_SEED + 1 },
    );
  });

  it('attributes generated malformed XML to every package part that is read', async () => {
    const localName = fc.stringMatching(/^[A-Za-z][A-Za-z0-9]{0,15}$/);

    await fc.assert(
      fc.asyncProperty(
        fc.constantFrom(...READ_XML_PARTS),
        localName,
        async (part, name) => {
          const input = await createIndependentPptx({
            [part]: `<root><${name}></root>`,
          });

          await expect(
            parsePptx(input, { errorMode: 'strict' }),
          ).rejects.toMatchObject({
            name: 'PptxParseError',
            diagnostic: {
              code: 'xml-parse-failed',
              part,
            },
          });
        },
      ),
      { numRuns: FUZZ_RUNS, seed: FUZZ_SEED + 2 },
    );
  });
});
