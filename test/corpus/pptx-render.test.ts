import { readFile } from 'node:fs/promises';
import { resolve } from 'node:path';

import { describe, expect, it } from 'vitest';

import { parsePptx, renderPptxDocumentToSvg } from '../../src';
import { renderPptxDocumentToPng } from '../../src/formats/pptx/node';
import { findUnsafeSvgFeatures } from './svg-safety';

interface ResolvedCorpusEntry {
  expectedSlides: number;
  id: string;
  path: string;
  producer: string;
}

interface ResolvedCorpus {
  entries: ResolvedCorpusEntry[];
  version: number;
}

const PNG_SIGNATURE = [137, 80, 78, 71, 13, 10, 26, 10];
const UTF8_DECODER = new TextDecoder();

function pngDimensions(data: Uint8Array) {
  const view = new DataView(data.buffer, data.byteOffset, data.byteLength);
  return { height: view.getUint32(20), width: view.getUint32(16) };
}

const resolvedManifest = JSON.parse(
  await readFile(resolve('.cache/pptx-corpus/resolved.json'), 'utf8'),
) as ResolvedCorpus;

describe(`real-world PPTX render corpus v${resolvedManifest.version}`, () => {
  for (const entry of resolvedManifest.entries) {
    it(`${entry.producer}: ${entry.id}`, async () => {
      const bytes = new Uint8Array(await readFile(entry.path));
      const document = await parsePptx(bytes, {
        audioMode: 'none',
        errorMode: 'tolerant',
        imageMode: 'base64',
        limits: { maxInputBytes: 150 * 1024 * 1024 },
        videoMode: 'none',
      });
      const slideNumbers =
        entry.expectedSlides === 1 ? [1] : [1, entry.expectedSlides];
      const options = { scale: 0.25, slideNumbers };
      const svg = renderPptxDocumentToSvg(document, options);
      const png = renderPptxDocumentToPng(document, options);

      expect(svg.slides).toHaveLength(slideNumbers.length);
      expect(png.slides).toHaveLength(slideNumbers.length);
      for (let index = 0; index < slideNumbers.length; index += 1) {
        const svgSlide = svg.slides[index];
        const pngSlide = png.slides[index];
        expect(svgSlide?.slideNumber).toBe(slideNumbers[index]);
        expect(pngSlide?.slideNumber).toBe(slideNumbers[index]);
        expect(svgSlide?.data.byteLength).toBeGreaterThan(200);
        expect(pngSlide?.data.byteLength).toBeGreaterThan(60);
        expect(Array.from(pngSlide?.data.subarray(0, 8) ?? [])).toEqual(
          PNG_SIGNATURE,
        );
        expect(pngSlide ? pngDimensions(pngSlide.data) : null).toEqual({
          height: pngSlide?.height,
          width: pngSlide?.width,
        });
        const source = UTF8_DECODER.decode(svgSlide?.data);
        expect(source).toContain(
          `<title>PowerPoint slide ${slideNumbers[index]}</title>`,
        );
        expect(findUnsafeSvgFeatures(source)).toEqual([]);
      }
    });
  }
});
