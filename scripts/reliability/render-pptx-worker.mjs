import { Buffer } from 'node:buffer';
import { createHash } from 'node:crypto';
import { readFile, writeFile } from 'node:fs/promises';
import path from 'node:path';

import {
  parsePptx,
  parsePptxRoundTripJson,
  readPptxRoundTrip,
  renderPptxDocumentToSvg,
  serializePptxRoundTripJson,
  writePptxRoundTrip,
} from '../../dist/index.js';
import { renderPptxDocumentToPng } from '../../dist/pptx/node.js';

const [inputPath, outputDirectory] = process.argv.slice(2);
if (!inputPath || !outputDirectory) {
  throw new Error('Expected input PPTX and output directory arguments');
}

function sha256(bytes) {
  return createHash('sha256').update(bytes).digest('hex');
}

const source = new Uint8Array(await readFile(inputPath));
const runtime = await readPptxRoundTrip(source);
const portable = await serializePptxRoundTripJson(runtime);
const wireJson = JSON.stringify(portable);
const transported = await parsePptxRoundTripJson(JSON.parse(wireJson));
const restored = await writePptxRoundTrip(transported);
const document = await parsePptx(restored.data, {
  audioMode: 'none',
  errorMode: 'strict',
  imageMode: 'base64',
  videoMode: 'none',
});
const svg = renderPptxDocumentToSvg(document);
const png = renderPptxDocumentToPng(document);
const svgSlide = svg.slides[0];
const pngSlide = png.slides[0];
if (!svgSlide || !pngSlide) throw new Error('Expected one rendered slide');

await Promise.all([
  writeFile(path.join(outputDirectory, 'portable.json'), wireJson),
  writeFile(path.join(outputDirectory, 'restored.pptx'), restored.data),
  writeFile(path.join(outputDirectory, 'slide-1.svg'), svgSlide.data),
  writeFile(path.join(outputDirectory, 'slide-1.png'), pngSlide.data),
]);

process.stdout.write(
  JSON.stringify({
    fidelityLevel: restored.report.level,
    outputSha256: sha256(restored.data),
    path: process.env.PATH ?? null,
    png: {
      bytes: pngSlide.data.byteLength,
      height: pngSlide.height,
      mimeType: pngSlide.mimeType,
      width: pngSlide.width,
    },
    portableJsonBytes: Buffer.byteLength(wireJson),
    slideCount: document.slides.length,
    sourceSha256: sha256(source),
    svg: {
      bytes: svgSlide.data.byteLength,
      height: svgSlide.height,
      mimeType: svgSlide.mimeType,
      width: svgSlide.width,
    },
  }),
);
