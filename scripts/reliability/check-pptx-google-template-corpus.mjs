import assert from 'node:assert/strict';
import { Buffer } from 'node:buffer';
import { createHash, randomUUID } from 'node:crypto';
import { mkdir, rm, writeFile } from 'node:fs/promises';
import path from 'node:path';
import { setTimeout as delay } from 'node:timers/promises';

import { Resvg } from '@resvg/resvg-js';

import { createPptx, parsePptx } from '../../dist/index.js';
import { renderPptxToPng } from '../../dist/pptx/node.js';
import { roundTripGoogleSlidesPresentation } from './google-slides-drive.mjs';
import { googleTemplateCatalog } from './pptx-google-template-catalog.mjs';

const accessToken = process.env.GOOGLE_DRIVE_ACCESS_TOKEN;

const reportDirectory = path.resolve(
  'reports',
  'reliability',
  'pptx-google-templates',
);
const decksDirectory = path.join(reportDirectory, 'decks');
const imagesDirectory = path.join(reportDirectory, 'images');
const allowedWarningCodes = new Set(['font-substitution']);

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

function googleRoundTripWithRetry(source, name) {
  return roundTripGoogleSlidesPresentation(
    source,
    accessToken,
    `${name}-${randomUUID()}`,
    retryingFetch,
  );
}

function assertSemanticMarker(document, marker, description) {
  const source = JSON.stringify(document);
  assert.match(
    source,
    new RegExp(marker),
    `${description} lost marker ${marker}`,
  );
}

function assertRenderable(result, description) {
  assert.equal(result.slides.length, 1, `${description} must render one slide`);
  const slide = result.slides[0];
  assert.ok(slide, `${description} rendered slide missing`);
  assert.equal(slide.width, 960, `${description} width changed`);
  assert.equal(slide.height, 540, `${description} height changed`);
  assert.ok(
    slide.data.byteLength > 1_000,
    `${description} PNG is suspiciously small`,
  );
  const unexpected = slide.warnings.filter(
    (warning) => !allowedWarningCodes.has(warning.code),
  );
  assert.deepEqual(
    unexpected,
    [],
    `${description} emitted an unsupported visual approximation`,
  );
  return slide;
}

function escapeXml(value) {
  return value
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;')
    .replaceAll("'", '&apos;');
}

function pngDataUri(bytes) {
  return `data:image/png;base64,${Buffer.from(bytes).toString('base64')}`;
}

function contactSheet(title, subtitle, entries) {
  const columns = 5;
  const cardWidth = 224;
  const cardHeight = 164;
  const gap = 12;
  const padding = 20;
  const headerHeight = 82;
  const rows = Math.ceil(entries.length / columns);
  const width = padding * 2 + columns * cardWidth + (columns - 1) * gap;
  const height = headerHeight + rows * cardHeight + (rows - 1) * gap + padding;
  const cards = entries
    .map((entry, index) => {
      const column = index % columns;
      const row = Math.floor(index / columns);
      const x = padding + column * (cardWidth + gap);
      const y = headerHeight + row * (cardHeight + gap);
      return `<g transform="translate(${x} ${y})"><rect width="${cardWidth}" height="${cardHeight}" rx="10" fill="#111827" stroke="#334155"/><image x="8" y="8" width="208" height="117" preserveAspectRatio="xMidYMid meet" href="${pngDataUri(entry.png)}"/><text x="12" y="145" fill="#F8FAFC" font-family="sans-serif" font-size="11" font-weight="700">${escapeXml(entry.index.toString().padStart(2, '0'))} · ${escapeXml(entry.title)}</text><text x="12" y="158" fill="#94A3B8" font-family="sans-serif" font-size="8">${escapeXml(entry.palette.toUpperCase())}</text></g>`;
    })
    .join('');
  const svg = `<svg xmlns="http://www.w3.org/2000/svg" width="${width}" height="${height}" viewBox="0 0 ${width} ${height}"><rect width="${width}" height="${height}" fill="#020617"/><text x="${padding}" y="34" fill="#F8FAFC" font-family="sans-serif" font-size="24" font-weight="700">${escapeXml(title)}</text><text x="${padding}" y="58" fill="#94A3B8" font-family="sans-serif" font-size="12">${escapeXml(subtitle)}</text>${cards}</svg>`;
  return Uint8Array.from(
    new Resvg(svg, { fitTo: { mode: 'original' } }).render().asPng(),
  );
}

await rm(reportDirectory, { force: true, recursive: true });
await Promise.all([
  mkdir(decksDirectory, { recursive: true }),
  mkdir(imagesDirectory, { recursive: true }),
]);

assert.equal(
  googleTemplateCatalog.length,
  30,
  'corpus must contain 30 templates',
);
const evidence = [];
const sourceImages = [];
const googleImages = [];
const authoredTemplates = [];

async function writeProgress(phase, completedGoogleTemplates) {
  await writeFile(
    path.join(reportDirectory, 'progress.json'),
    `${JSON.stringify(
      {
        completedGoogleTemplates,
        completedSourceTemplates: authoredTemplates.length,
        phase,
        schemaVersion: 1,
        templateCount: googleTemplateCatalog.length,
      },
      null,
      2,
    )}\n`,
  );
}

for (const template of googleTemplateCatalog) {
  const elementCount = template.scene.slides[0]?.elements.length ?? 0;
  assert.ok(
    elementCount >= 12,
    `${template.slug} is not visually complex enough`,
  );
  const created = await createPptx(template.scene);
  assert.equal(created.report.level, 'C2');
  const [sourceDocument, sourceRender] = await Promise.all([
    parsePptx(created.data, {
      audioMode: 'none',
      errorMode: 'strict',
      imageMode: 'none',
      videoMode: 'none',
    }),
    renderPptxToPng(created.data),
  ]);
  assertSemanticMarker(
    sourceDocument,
    template.marker,
    `${template.slug} source`,
  );
  assert.ok(
    (sourceDocument.slides[0]?.elements.length ?? 0) >= 12,
    `${template.slug} source lost authored elements`,
  );
  const sourceSlide = assertRenderable(sourceRender, `${template.slug} source`);
  const sourcePptxPath = path.join(
    decksDirectory,
    `${template.slug}-source.pptx`,
  );
  const sourcePngPath = path.join(
    imagesDirectory,
    `${template.slug}-source.png`,
  );
  await Promise.all([
    writeFile(sourcePptxPath, created.data),
    writeFile(sourcePngPath, sourceSlide.data),
  ]);

  const shared = {
    elementCount,
    index: template.index,
    marker: template.marker,
    palette: template.palette,
    slug: template.slug,
    title: template.title,
  };
  sourceImages.push({ ...shared, png: sourceSlide.data });
  authoredTemplates.push({
    created,
    source: {
      documentElementCount: sourceDocument.slides[0]?.elements.length ?? 0,
      png: artifactEvidence(sourceSlide.data),
      pptx: artifactEvidence(created.data),
      warnings: sourceSlide.warnings.map((warning) => warning.code),
    },
    ...shared,
    template,
  });
  console.log(
    `[source ${template.index.toString().padStart(2, '0')}/30] verified ${template.title}`,
  );
}

const sourceContactSheet = contactSheet(
  '30 authored PowerPoint templates',
  'Created through OAKit · strict parsed · rendered without Office',
  sourceImages,
);
await writeFile(
  path.join(reportDirectory, 'contact-sheet-source.png'),
  sourceContactSheet,
);
await writeProgress('google-import-export', 0);

if (accessToken === undefined || accessToken.trim().length === 0) {
  throw new Error(
    'Google Slides template verification requires GOOGLE_DRIVE_ACCESS_TOKEN',
  );
}

for (const authored of authoredTemplates) {
  const { created, source, template, ...shared } = authored;
  const exported = await googleRoundTripWithRetry(
    created.data,
    `oakit-template-${template.index.toString().padStart(2, '0')}`,
  );
  const [googleDocument, googleRender] = await Promise.all([
    parsePptx(exported, {
      audioMode: 'none',
      errorMode: 'strict',
      imageMode: 'none',
      videoMode: 'none',
    }),
    renderPptxToPng(exported),
  ]);
  assertSemanticMarker(
    googleDocument,
    template.marker,
    `${template.slug} Google export`,
  );
  assert.ok(
    (googleDocument.slides[0]?.elements.length ?? 0) >= 12,
    `${template.slug} Google export lost authored elements`,
  );
  const googleSlide = assertRenderable(
    googleRender,
    `${template.slug} Google export`,
  );
  assert.notEqual(
    sha256(exported),
    sha256(created.data),
    `${template.slug} did not pass through a distinct producer export`,
  );
  await Promise.all([
    writeFile(
      path.join(decksDirectory, `${template.slug}-google.pptx`),
      exported,
    ),
    writeFile(
      path.join(imagesDirectory, `${template.slug}-google.png`),
      googleSlide.data,
    ),
  ]);
  googleImages.push({ ...shared, png: googleSlide.data });
  evidence.push({
    ...shared,
    googleExport: {
      documentElementCount: googleDocument.slides[0]?.elements.length ?? 0,
      png: artifactEvidence(googleSlide.data),
      pptx: artifactEvidence(exported),
      warnings: googleSlide.warnings.map((warning) => warning.code),
    },
    source,
  });
  await writeProgress('google-import-export', evidence.length);
  console.log(
    `[google ${template.index.toString().padStart(2, '0')}/30] verified ${template.title}`,
  );
}

const googleContactSheet = contactSheet(
  '30 Google Slides round-trip exports',
  'Controlled Drive import/export · strict parsed · rendered without Office · temporary files deleted',
  googleImages,
);
await writeFile(
  path.join(reportDirectory, 'contact-sheet-google.png'),
  googleContactSheet,
);
await writeProgress('complete', evidence.length);

const manifest = {
  artifacts: {
    googleContactSheet: artifactEvidence(googleContactSheet),
    sourceContactSheet: artifactEvidence(sourceContactSheet),
  },
  corpus: evidence,
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
  schemaVersion: 1,
  summary: {
    allGoogleExportsStrictParsed: true,
    allGoogleExportsVisuallyRendered: true,
    allMarkersPreserved: true,
    allSourcesStrictParsed: true,
    allSourcesVisuallyRendered: true,
    minimumElementsPerTemplate: Math.min(
      ...evidence.map((entry) => entry.elementCount),
    ),
    temporaryPresentationsDeleted: true,
    templateCount: evidence.length,
  },
  verifiedAt: new Date().toISOString(),
};
await writeFile(
  path.join(reportDirectory, 'evidence.json'),
  `${JSON.stringify(manifest, null, 2)}\n`,
);
console.log(JSON.stringify(manifest.summary, null, 2));
