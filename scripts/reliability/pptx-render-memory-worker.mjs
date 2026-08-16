import { Buffer } from 'node:buffer';
import { createHash } from 'node:crypto';
import { performance } from 'node:perf_hooks';

import { renderPptxDocumentToSvg } from '../../dist/index.js';
import { renderPptxDocumentToPng } from '../../dist/pptx/node.js';

const PNG_SIGNATURE = Buffer.from([137, 80, 78, 71, 13, 10, 26, 10]);
const OUTPUT_WIDTH = 360;
const OUTPUT_HEIGHT = 203;

function memorySample(stage) {
  const memory = process.memoryUsage();
  return {
    arrayBuffers: memory.arrayBuffers,
    external: memory.external,
    heapUsed: memory.heapUsed,
    rss: memory.rss,
    stage,
  };
}

function textElement(index) {
  return {
    borderColor: '#1f2937',
    borderStrokeDasharray: '',
    borderType: 'solid',
    borderWidth: 1,
    content: `Agent-ready slide ${index}`,
    fill: { type: 'color', value: '#ffffff' },
    height: 48,
    id: `text-${index}`,
    isFlipH: false,
    isFlipV: false,
    isVertical: false,
    left: 24,
    name: `Title ${index}`,
    order: 0,
    rotate: 0,
    top: 24,
    type: 'text',
    vAlign: 'top',
    width: 320,
    wrap: true,
  };
}

function chartElement(index) {
  return {
    chartType: 'barChart',
    colors: ['#4f46e5', '#0ea5e9', '#22c55e'],
    data: [
      {
        key: `Series ${index}`,
        values: [
          { x: 'A', y: index },
          { x: 'B', y: index + 1 },
          { x: 'C', y: index + 2 },
        ],
        xlabels: {},
      },
    ],
    height: 220,
    id: `chart-${index}`,
    left: 24,
    order: 1,
    top: 96,
    type: 'chart',
    width: 520,
  };
}

function document(slideCount) {
  return {
    size: { height: 405, width: 720 },
    slides: Array.from({ length: slideCount }, (_value, index) => ({
      elements: [textElement(index + 1), chartElement(index + 1)],
      fill: {
        type: 'color',
        value: index % 2 === 0 ? '#eef2ff' : '#ecfeff',
      },
      layoutElements: [],
      note: `Memory evidence note ${index + 1}`,
    })),
    themeColors: [],
    usedFonts: [],
  };
}

function timed(timings, samples, stage, operation) {
  const started = performance.now();
  const result = operation();
  timings[stage] = Math.round(performance.now() - started);
  samples.push(memorySample(stage));
  return result;
}

function digestSlides(slides) {
  const digest = createHash('sha256');
  for (const slide of slides) {
    digest.update(
      Buffer.from(
        slide.data.buffer,
        slide.data.byteOffset,
        slide.data.byteLength,
      ),
    );
  }
  return digest.digest('hex');
}

function assertSvg(slides, expectedCount) {
  if (slides.length !== expectedCount) {
    throw new Error(
      `Expected ${expectedCount} SVG slides, received ${slides.length}`,
    );
  }
  for (const [index, slide] of slides.entries()) {
    if (
      slide.slideNumber !== index + 1 ||
      slide.width !== OUTPUT_WIDTH ||
      slide.height !== OUTPUT_HEIGHT ||
      slide.format !== 'svg' ||
      slide.mimeType !== 'image/svg+xml'
    ) {
      throw new Error(`SVG metadata mismatch on slide ${index + 1}`);
    }
    const source = Buffer.from(
      slide.data.buffer,
      slide.data.byteOffset,
      slide.data.byteLength,
    ).toString('utf8');
    if (
      !source.startsWith('<?xml version="1.0"') ||
      source.includes('<script') ||
      /\b(?:href|src)="(?!data:image\/)/i.test(source)
    ) {
      throw new Error(`SVG safety mismatch on slide ${index + 1}`);
    }
  }
}

function pngDimensions(data) {
  const view = new DataView(data.buffer, data.byteOffset, data.byteLength);
  return { height: view.getUint32(20), width: view.getUint32(16) };
}

function assertPng(slides, expectedCount) {
  if (slides.length !== expectedCount) {
    throw new Error(
      `Expected ${expectedCount} PNG slides, received ${slides.length}`,
    );
  }
  for (const [index, slide] of slides.entries()) {
    const signature = Buffer.from(
      slide.data.buffer,
      slide.data.byteOffset,
      PNG_SIGNATURE.byteLength,
    );
    const dimensions = pngDimensions(slide.data);
    if (
      slide.slideNumber !== index + 1 ||
      slide.width !== OUTPUT_WIDTH ||
      slide.height !== OUTPUT_HEIGHT ||
      slide.format !== 'png' ||
      slide.mimeType !== 'image/png' ||
      !signature.equals(PNG_SIGNATURE) ||
      dimensions.width !== OUTPUT_WIDTH ||
      dimensions.height !== OUTPUT_HEIGHT
    ) {
      throw new Error(`PNG metadata mismatch on slide ${index + 1}`);
    }
  }
}

function totalBytes(slides) {
  return slides.reduce((total, slide) => total + slide.data.byteLength, 0);
}

const tier = process.argv[2];
const slideCount = Number(process.argv[3]);
if (!tier || !Number.isSafeInteger(slideCount) || slideCount <= 0) {
  throw new Error('Expected a tier and positive safe slide count');
}

const input = document(slideCount);
const inputBefore = JSON.stringify(input);
globalThis.gc?.();
const samples = [memorySample('document-created')];
const timings = {};
const options = {
  limits: {
    maxElementsPerSlide: 2,
    maxOutputPixels: OUTPUT_WIDTH * OUTPUT_HEIGHT,
    maxPngBytes: 2 * 1024 * 1024,
    maxSlides: slideCount,
    maxSvgBytes: 2 * 1024 * 1024,
  },
  scale: 0.5,
};

const svg = timed(timings, samples, 'render-svg', () =>
  renderPptxDocumentToSvg(input, options),
);
assertSvg(svg.slides, slideCount);
const png = timed(timings, samples, 'render-png', () =>
  renderPptxDocumentToPng(input, options),
);
assertPng(png.slides, slideCount);
if (JSON.stringify(input) !== inputBefore) {
  throw new Error('Renderer mutated the caller-owned document');
}
if (
  svg.slides.some(
    (slide, index) =>
      JSON.stringify(slide.warnings) !==
      JSON.stringify(png.slides[index]?.warnings),
  )
) {
  throw new Error('PNG output did not preserve SVG warnings');
}
samples.push(memorySample('outputs-verified'));

const peak = (field) => Math.max(...samples.map((sample) => sample[field]));
const totalDurationMs = Object.values(timings).reduce(
  (total, duration) => total + duration,
  0,
);

process.stdout.write(
  `${JSON.stringify({
    officeRuntimeRequired: false,
    pathEmpty: process.env.PATH === '',
    peakArrayBuffersBytes: peak('arrayBuffers'),
    peakExternalBytes: peak('external'),
    peakHeapUsedBytes: peak('heapUsed'),
    peakRssBytes: peak('rss'),
    pngBytes: totalBytes(png.slides),
    pngSha256: digestSlides(png.slides),
    samples,
    slideCount,
    svgBytes: totalBytes(svg.slides),
    svgSha256: digestSlides(svg.slides),
    tier,
    timingsMs: timings,
    totalDurationMs,
    warningCount: svg.slides.reduce(
      (total, slide) => total + slide.warnings.length,
      0,
    ),
  })}\n`,
);
