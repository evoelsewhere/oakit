import fc from 'fast-check';
import { describe, expect, it } from 'vitest';

import { parsePptx, parsePptxWithDiagnostics } from '../../src';
import { createIndependentPptx, PRESENTATION_NS } from './pptx-package';

const FUZZ_SEED = 0x47_45_4f;
const FUZZ_RUNS = 40;

function escapeXmlAttribute(value: string): string {
  return value
    .replaceAll('&', '&amp;')
    .replaceAll('"', '&quot;')
    .replaceAll('<', '&lt;');
}

function pointsFromEmus(value: number): number {
  return Number.parseFloat((value / 12_700).toFixed(4));
}

function slideWithGeometry(
  x: string,
  y: string,
  width: string,
  height: string,
): string {
  return `<p:sld xmlns:p="${PRESENTATION_NS}" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
    <p:cSld><p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>
      <p:sp>
        <p:nvSpPr><p:cNvPr id="2" name="Generated geometry"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
        <p:spPr>
          <a:xfrm><a:off x="${escapeXmlAttribute(x)}" y="${escapeXmlAttribute(y)}"/><a:ext cx="${escapeXmlAttribute(width)}" cy="${escapeXmlAttribute(height)}"/></a:xfrm>
          <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
          <a:solidFill><a:srgbClr val="4472C4"/></a:solidFill>
        </p:spPr>
      </p:sp>
    </p:spTree></p:cSld>
  </p:sld>`;
}

describe('PowerPoint generated geometry properties', () => {
  it('converts canonical coordinates exactly', async () => {
    await fc.assert(
      fc.asyncProperty(
        fc.record({
          x: fc.integer({ min: -10_000_000, max: 10_000_000 }),
          y: fc.integer({ min: -10_000_000, max: 10_000_000 }),
          width: fc.integer({ min: 0, max: 10_000_000 }),
          height: fc.integer({ min: 0, max: 10_000_000 }),
        }),
        async ({ x, y, width, height }) => {
          const input = await createIndependentPptx({
            'ppt/slides/slide1.xml': slideWithGeometry(
              String(x),
              String(y),
              String(width),
              String(height),
            ),
          });

          const document = await parsePptx(input, { errorMode: 'strict' });
          const element = document.slides[0]?.elements[0];

          expect(element).toMatchObject({
            left: pointsFromEmus(x),
            top: pointsFromEmus(y),
            width: pointsFromEmus(width),
            height: pointsFromEmus(height),
          });
        },
      ),
      { numRuns: FUZZ_RUNS, seed: FUZZ_SEED },
    );
  }, 15_000);

  it('reports malformed coordinates instead of returning ambiguous geometry', async () => {
    const malformedSignedInteger = fc.oneof(
      fc.constantFrom('', ' ', '+', '-', '1.0', '1e3', ' 1', '1 ', '01'),
      fc
        .array(fc.constantFrom(...'0123456789+-. eEabcXYZ'), {
          minLength: 1,
          maxLength: 12,
        })
        .map((characters) => characters.join(''))
        .filter((value) => {
          if (!/^[+-]?(?:0|[1-9]\d*)$/.test(value)) return true;
          return !Number.isSafeInteger(Number(value));
        }),
    );
    const malformedSize = fc.oneof(
      malformedSignedInteger,
      fc.integer({ min: -10_000_000, max: -1 }).map(String),
    );
    const generatedCase = fc.oneof(
      malformedSignedInteger.map((value) => ({
        x: value,
        y: '0',
        width: '914400',
        height: '914400',
      })),
      malformedSignedInteger.map((value) => ({
        x: '0',
        y: value,
        width: '914400',
        height: '914400',
      })),
      malformedSize.map((value) => ({
        x: '0',
        y: '0',
        width: value,
        height: '914400',
      })),
      malformedSize.map((value) => ({
        x: '0',
        y: '0',
        width: '914400',
        height: value,
      })),
    );

    await fc.assert(
      fc.asyncProperty(generatedCase, async ({ x, y, width, height }) => {
        const input = await createIndependentPptx({
          'ppt/slides/slide1.xml': slideWithGeometry(x, y, width, height),
        });

        const tolerant = await parsePptxWithDiagnostics(input);
        const element = tolerant.document.slides[0]?.elements[0];

        expect(element).toMatchObject({
          left: 0,
          top: 0,
          width: width === '914400' ? 72 : 0,
          height: height === '914400' ? 72 : 0,
        });
        expect(tolerant.diagnostics).toContainEqual(
          expect.objectContaining({
            code: 'invalid-document-value',
            part: 'ppt/slides/slide1.xml',
          }),
        );
        await expect(
          parsePptx(input, { errorMode: 'strict' }),
        ).rejects.toMatchObject({
          diagnostic: {
            code: 'invalid-document-value',
            part: 'ppt/slides/slide1.xml',
          },
        });
      }),
      { numRuns: FUZZ_RUNS, seed: FUZZ_SEED + 1 },
    );
  }, 15_000);
});
