import assert from 'node:assert/strict';
import { createHash, randomUUID } from 'node:crypto';
import { mkdir, rm, writeFile } from 'node:fs/promises';
import path from 'node:path';
import { setTimeout as delay } from 'node:timers/promises';
import { TextDecoder } from 'node:util';

import JSZip from 'jszip';

import { parsePptx, renderPptxToSvg } from '../../dist/index.js';
import { renderPptxToPng } from '../../dist/pptx/node.js';
import { fetchSlidesManiaTemplate } from './fetch-slidesmania-template.mjs';
import { roundTripGoogleSlidesPresentation } from './google-slides-drive.mjs';
import {
  slidesManiaCorpus,
  slidesManiaCorpusProvenance,
} from './slidesmania-corpus.mjs';

const accessToken = process.env.GOOGLE_DRIVE_ACCESS_TOKEN;
const reportDirectory = path.resolve(
  'reports',
  'reliability',
  'pptx-google-slidesmania',
);
const decoder = new TextDecoder();

function sha256(bytes) {
  return createHash('sha256').update(bytes).digest('hex');
}

function artifactEvidence(bytes) {
  return { byteLength: bytes.byteLength, sha256: sha256(bytes) };
}

function isRetryableStatus(status) {
  return status === 408 || status === 429 || status >= 500;
}

async function retryingFetch(input, init) {
  let lastError;
  for (let attempt = 1; attempt <= 3; attempt += 1) {
    try {
      const response = await globalThis.fetch(input, init);
      if (!isRetryableStatus(response.status) || attempt === 3) return response;
      await response.body?.cancel();
    } catch (error) {
      lastError = error;
      if (attempt === 3) throw error;
    }
    await delay(750 * 2 ** (attempt - 1));
  }
  throw lastError;
}

function googleRoundTrip(source, index) {
  return roundTripGoogleSlidesPresentation(
    source,
    accessToken,
    `oakit-slidesmania-${String(index).padStart(2, '0')}-${randomUUID()}`,
    retryingFetch,
  );
}

function elementCount(document) {
  return document.slides.reduce(
    (total, slide) => total + slide.elements.length,
    0,
  );
}

function warningCounts(slides) {
  const counts = {};
  for (const slide of slides) {
    for (const warning of slide.warnings) {
      counts[warning.code] = (counts[warning.code] ?? 0) + 1;
    }
  }
  return counts;
}

function normalizeText(value) {
  return value
    .replace(/<[^>]*>/g, ' ')
    .replace(/&(?:nbsp|amp|lt|gt|quot|apos);/gi, ' ')
    .toLocaleLowerCase('en-US')
    .normalize('NFKC')
    .replace(/[^\p{Letter}\p{Number}]+/gu, ' ')
    .trim();
}

function visitTextContent(element, tokens) {
  if (typeof element.content === 'string') {
    for (const token of normalizeText(element.content).split(' ')) {
      if (token.length >= 3) tokens.add(token);
    }
  }
  if (Array.isArray(element.elements)) {
    for (const child of element.elements) visitTextContent(child, tokens);
  }
}

function documentTextTokens(document) {
  const tokens = new Set();
  for (const slide of document.slides) {
    for (const element of slide.elements) visitTextContent(element, tokens);
  }
  return tokens;
}

function tokenRetention(source, output) {
  if (source.size === 0) return 1;
  let retained = 0;
  for (const token of source) if (output.has(token)) retained += 1;
  return retained / source.size;
}

async function containsSlidesManiaAttribution(bytes) {
  const archive = await JSZip.loadAsync(bytes);
  for (const [name, entry] of Object.entries(archive.files)) {
    if (entry.dir || !name.endsWith('.xml')) continue;
    if (/SLIDESMANIA/i.test(await entry.async('string'))) return true;
  }
  return false;
}

function selectedSlideNumbers(count) {
  return [...new Set([1, Math.ceil(count / 2), count])];
}

async function renderEvidence(bytes, expectedSlideCount) {
  const svgResult = await renderPptxToSvg(bytes);
  assert.equal(svgResult.slides.length, expectedSlideCount);
  const svgDigest = createHash('sha256');
  let svgByteLength = 0;
  for (const slide of svgResult.slides) {
    svgDigest.update(slide.data);
    svgByteLength += slide.data.byteLength;
    const source = decoder.decode(slide.data);
    assert.equal(
      /(?:href|src)="https?:/i.test(source),
      false,
      `rendered slide ${slide.slideNumber} contains an external reference`,
    );
  }
  const selected = selectedSlideNumbers(expectedSlideCount);
  const pngResult = await renderPptxToPng(bytes, { slideNumbers: selected });
  assert.deepEqual(
    pngResult.slides.map(({ slideNumber }) => slideNumber),
    selected,
  );
  return {
    representativePngs: pngResult.slides.map((slide) => ({
      ...artifactEvidence(slide.data),
      slideNumber: slide.slideNumber,
      warningCount: slide.warnings.length,
    })),
    svg: {
      byteLength: svgByteLength,
      sha256: svgDigest.digest('hex'),
    },
    warnings: warningCounts(svgResult.slides),
  };
}

function assertSameSlideSize(source, output) {
  assert.ok(
    Math.abs(source.size.width - output.size.width) <= 0.01 &&
      Math.abs(source.size.height - output.size.height) <= 0.01,
    'Google Slides export changed the slide size',
  );
}

async function writeProgress(
  phase,
  completedTemplates,
  currentTemplate = null,
) {
  await writeFile(
    path.join(reportDirectory, 'progress.json'),
    `${JSON.stringify(
      {
        completedTemplates,
        currentTemplate,
        phase,
        schemaVersion: 1,
        templateCount: slidesManiaCorpus.length,
      },
      null,
      2,
    )}\n`,
  );
}

await rm(reportDirectory, { force: true, recursive: true });
await mkdir(reportDirectory, { recursive: true });
await writeProgress('initializing', 0);

assert.equal(slidesManiaCorpus.length, 30);
if (accessToken === undefined || accessToken.trim().length === 0) {
  throw new Error(
    'SlidesMania Google producer verification requires GOOGLE_DRIVE_ACCESS_TOKEN',
  );
}

const corpusEvidence = [];
for (const [offset, corpusEntry] of slidesManiaCorpus.entries()) {
  const index = offset + 1;
  await writeProgress('source-validation', offset, corpusEntry.title);
  const fetched = await fetchSlidesManiaTemplate(corpusEntry, retryingFetch);
  const sourceDocument = await parsePptx(fetched.bytes, {
    audioMode: 'none',
    errorMode: 'strict',
    imageMode: 'none',
    videoMode: 'none',
  });
  assert.ok(sourceDocument.slides.length > 0, 'source deck has no slides');
  assert.equal(await containsSlidesManiaAttribution(fetched.bytes), true);
  const sourceElements = elementCount(sourceDocument);
  assert.ok(sourceElements > 0, 'source deck has no elements');
  const sourceTokens = documentTextTokens(sourceDocument);
  const sourceRender = await renderEvidence(
    fetched.bytes,
    sourceDocument.slides.length,
  );

  await writeProgress('google-import-export', offset, corpusEntry.title);
  const exported = await googleRoundTrip(fetched.bytes, index);
  assert.notEqual(sha256(exported), sha256(fetched.bytes));
  const outputDocument = await parsePptx(exported, {
    audioMode: 'none',
    errorMode: 'strict',
    imageMode: 'none',
    videoMode: 'none',
  });
  assert.equal(outputDocument.slides.length, sourceDocument.slides.length);
  assertSameSlideSize(sourceDocument, outputDocument);
  assert.equal(await containsSlidesManiaAttribution(exported), true);
  const outputElements = elementCount(outputDocument);
  const elementRetention = outputElements / sourceElements;
  assert.ok(
    elementRetention >= 0.75,
    'Google Slides export lost too many elements',
  );
  const textRetention = tokenRetention(
    sourceTokens,
    documentTextTokens(outputDocument),
  );
  assert.ok(textRetention >= 0.7, 'Google Slides export lost too much text');
  const outputRender = await renderEvidence(
    exported,
    outputDocument.slides.length,
  );
  corpusEvidence.push({
    attributionPreserved: true,
    elementRetention,
    elements: { output: outputElements, source: sourceElements },
    index,
    output: {
      pptx: artifactEvidence(exported),
      render: outputRender,
    },
    slideCount: sourceDocument.slides.length,
    source: {
      pptx: artifactEvidence(fetched.bytes),
      render: sourceRender,
    },
    sourcePage: fetched.sourcePage,
    textRetention,
    title: fetched.title,
  });
  await writeProgress('google-import-export', index, null);
  console.log(
    `[${String(index).padStart(2, '0')}/30] verified ${fetched.title} (${sourceDocument.slides.length} slides)`,
  );
}

const aggregateWarningCounts = (side) =>
  corpusEvidence.reduce((counts, entry) => {
    for (const [code, count] of Object.entries(entry[side].render.warnings)) {
      counts[code] = (counts[code] ?? 0) + count;
    }
    return counts;
  }, {});

const evidence = {
  corpus: corpusEvidence,
  execution: {
    repository: process.env.GITHUB_REPOSITORY ?? null,
    revision: process.env.GITHUB_SHA ?? null,
    runAttempt: process.env.GITHUB_RUN_ATTEMPT ?? null,
    runId: process.env.GITHUB_RUN_ID ?? null,
    workflow: process.env.GITHUB_WORKFLOW ?? null,
  },
  platform: {
    architecture: process.arch,
    node: process.version,
    os: process.platform,
  },
  producer: {
    application: 'Google Slides',
    transport: 'Google Drive API v3 controlled import/export',
  },
  provenance: slidesManiaCorpusProvenance,
  schemaVersion: 1,
  summary: {
    allAttributionPreserved: true,
    allGoogleExportsStrictParsed: true,
    allSlidesRenderedWithoutOffice: true,
    minimumElementRetention: Math.min(
      ...corpusEvidence.map(({ elementRetention }) => elementRetention),
    ),
    minimumTextRetention: Math.min(
      ...corpusEvidence.map(({ textRetention }) => textRetention),
    ),
    outputWarnings: aggregateWarningCounts('output'),
    sourceWarnings: aggregateWarningCounts('source'),
    temporaryPresentationsDeleted: true,
    templateCount: corpusEvidence.length,
    totalElements: corpusEvidence.reduce(
      (total, entry) => total + entry.elements.source,
      0,
    ),
    totalSlides: corpusEvidence.reduce(
      (total, entry) => total + entry.slideCount,
      0,
    ),
  },
  verifiedAt: new Date().toISOString(),
};
await writeFile(
  path.join(reportDirectory, 'evidence.json'),
  `${JSON.stringify(evidence, null, 2)}\n`,
);
await writeProgress('complete', slidesManiaCorpus.length);
console.log(JSON.stringify(evidence.summary, null, 2));
