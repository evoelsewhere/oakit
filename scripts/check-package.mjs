import assert from 'node:assert/strict';
import { Buffer } from 'node:buffer';
import { execFile } from 'node:child_process';
import { mkdir, readFile, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import path from 'node:path';
import { createRequire } from 'node:module';
import { mkdtemp } from 'node:fs/promises';
import { promisify } from 'node:util';

import JSZip from 'jszip';

const execFileAsync = promisify(execFile);
const require = createRequire(import.meta.url);
const projectRoot = path.resolve(import.meta.dirname, '..');

async function installPackedArtifact(directory) {
  const packDirectory = path.join(directory, 'pack');
  const consumerDirectory = path.join(directory, 'consumer');
  await Promise.all([
    mkdir(packDirectory, { recursive: true }),
    mkdir(consumerDirectory, { recursive: true }),
  ]);
  const packed = await execFileAsync(
    'npm',
    ['pack', '--ignore-scripts', '--json', '--pack-destination', packDirectory],
    { cwd: projectRoot, maxBuffer: 1024 * 1024 },
  );
  const packResult = JSON.parse(packed.stdout);
  assert.equal(packResult.length, 1);
  const filename = packResult[0]?.filename;
  assert.equal(typeof filename, 'string');
  const tarballPath = path.join(packDirectory, filename);
  await writeFile(
    path.join(consumerDirectory, 'package.json'),
    `${JSON.stringify({ private: true, type: 'module' }, null, 2)}\n`,
  );
  await execFileAsync(
    'npm',
    ['install', '--ignore-scripts', '--no-audit', '--no-fund', tarballPath],
    { cwd: consumerDirectory, maxBuffer: 1024 * 1024 },
  );
  return {
    cliPath: path.join(
      consumerDirectory,
      'node_modules',
      '@evoelsewhere',
      'oakit',
      'dist',
      'cli.js',
    ),
    consumerDirectory,
  };
}

async function createSmokePptx() {
  const zip = new JSZip();
  zip.file(
    '[Content_Types].xml',
    `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
      <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
      <Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
    </Types>`,
  );
  zip.file(
    'ppt/presentation.xml',
    `<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
      xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
      <p:sldIdLst><p:sldId id="256" r:id="rIdSlide1"/></p:sldIdLst>
      <p:sldSz cx="9144000" cy="5143500"/>
    </p:presentation>`,
  );
  zip.file(
    'ppt/_rels/presentation.xml.rels',
    `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Relationship Id="rIdSlide1"
        Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide"
        Target="slides/slide1.xml"/>
    </Relationships>`,
  );
  zip.file(
    'ppt/slides/slide1.xml',
    `<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
      <p:cSld><p:spTree/></p:cSld>
    </p:sld>`,
  );
  return zip.generateAsync({ type: 'uint8array' });
}

const metadata = require('../package.json');
assert.equal(metadata.name, '@evoelsewhere/oakit');
assert.equal(metadata.bin.oakit, './dist/cli.js');

const esm = await import('@evoelsewhere/oakit');
const esmPptx = await import('@evoelsewhere/oakit/pptx');
const esmPptxNode = await import('@evoelsewhere/oakit/pptx/node');
const cjs = require('@evoelsewhere/oakit');
const cjsPptx = require('@evoelsewhere/oakit/pptx');
const cjsPptxNode = require('@evoelsewhere/oakit/pptx/node');
assert.equal(typeof esm.parsePptx, 'function');
assert.equal(typeof esmPptx.parsePptxWithDiagnostics, 'function');
assert.equal(typeof cjs.parsePptx, 'function');
assert.equal(typeof cjsPptx.parsePptxWithDiagnostics, 'function');
assert.equal(typeof esmPptxNode.renderPptxToPng, 'function');
assert.equal(typeof cjsPptxNode.renderPptxDocumentToPng, 'function');

const builtCliPath = path.join(projectRoot, 'dist/cli.js');
assert.equal(
  (await readFile(builtCliPath, 'utf8')).startsWith('#!/usr/bin/env node\n'),
  true,
);

const smokeDirectory = await mkdtemp(
  path.join(tmpdir(), 'oakit-package-smoke-'),
);
try {
  const installed = await installPackedArtifact(smokeDirectory);
  const cliPath = installed.cliPath;
  assert.equal(
    (await readFile(cliPath, 'utf8')).startsWith('#!/usr/bin/env node\n'),
    true,
  );
  const versionResult = await execFileAsync(process.execPath, [
    cliPath,
    '--version',
  ]);
  assert.equal(versionResult.stdout, `oakit ${metadata.version}\n`);
  assert.equal(versionResult.stderr, '');

  const inputPath = path.join(smokeDirectory, 'smoke.pptx');
  const outputPath = path.join(smokeDirectory, 'smoke.json');
  const portablePath = path.join(smokeDirectory, 'smoke.portable.json');
  const renderOutputPath = path.join(smokeDirectory, 'previews');
  const restoredPath = path.join(smokeDirectory, 'restored.pptx');
  const inputBytes = await createSmokePptx();
  await writeFile(inputPath, inputBytes);

  const pngResult = await esmPptxNode.renderPptxToPng(inputBytes);
  assert.equal(pngResult.slides.length, 1);
  assert.deepEqual(
    Array.from(pngResult.slides[0].data.subarray(0, 8)),
    [137, 80, 78, 71, 13, 10, 26, 10],
  );

  const stdoutResult = await execFileAsync(process.execPath, [
    cliPath,
    inputPath,
  ]);
  assert.equal(stdoutResult.stderr, '');
  const stdoutPayload = JSON.parse(stdoutResult.stdout);
  assert.equal(stdoutPayload.format, 'pptx');
  assert.equal(stdoutPayload.document.slides.length, 1);
  assert.deepEqual(stdoutPayload.document.size, { height: 405, width: 720 });

  const fileResult = await execFileAsync(process.execPath, [
    cliPath,
    'convert',
    inputPath,
    '--output',
    outputPath,
    '--pretty',
  ]);
  assert.equal(fileResult.stdout, '');
  assert.equal(fileResult.stderr, '');
  const filePayload = JSON.parse(await readFile(outputPath, 'utf8'));
  assert.equal(filePayload.document.slides.length, 1);

  const snapshotResult = await execFileAsync(process.execPath, [
    cliPath,
    'snapshot',
    inputPath,
    '--output',
    portablePath,
  ]);
  assert.equal(snapshotResult.stdout, '');
  assert.equal(snapshotResult.stderr, '');
  const portablePayload = JSON.parse(await readFile(portablePath, 'utf8'));
  assert.equal(portablePayload.format, 'pptx');
  assert.equal(portablePayload.schemaVersion, 1);
  assert.equal(portablePayload.source.byteLength, inputBytes.byteLength);
  assert.equal(
    portablePayload.source.packageBase64,
    Buffer.from(inputBytes).toString('base64'),
  );

  const restoreResult = await execFileAsync(process.execPath, [
    cliPath,
    'restore',
    portablePath,
    '--output',
    restoredPath,
  ]);
  assert.equal(restoreResult.stdout, '');
  assert.equal(restoreResult.stderr, '');
  assert.deepEqual(
    Array.from(await readFile(restoredPath)),
    Array.from(inputBytes),
  );

  const renderResult = await execFileAsync(process.execPath, [
    cliPath,
    'render',
    restoredPath,
    '--output',
    renderOutputPath,
    '--render-format',
    'png',
    '--slides',
    '1',
    '--scale',
    '0.5',
  ]);
  assert.equal(renderResult.stdout, '');
  assert.equal(renderResult.stderr, '');
  const renderedPng = await readFile(
    path.join(renderOutputPath, 'slide-1.png'),
  );
  assert.deepEqual(
    Array.from(renderedPng.subarray(0, 8)),
    [137, 80, 78, 71, 13, 10, 26, 10],
  );
  const renderManifest = JSON.parse(
    await readFile(path.join(renderOutputPath, 'manifest.json'), 'utf8'),
  );
  assert.equal(renderManifest.format, 'pptx-render');
  assert.equal(renderManifest.renderFormat, 'png');
  assert.deepEqual(renderManifest.slides[0], {
    byteLength: renderedPng.byteLength,
    file: 'slide-1.png',
    format: 'png',
    height: 203,
    mimeType: 'image/png',
    slideNumber: 1,
    warnings: [],
    width: 360,
  });

  const consumerPath = path.join(installed.consumerDirectory, 'consumer.mjs');
  await writeFile(
    consumerPath,
    `import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import { createRequire } from 'node:module';
import { parsePptx } from '@evoelsewhere/oakit';
import { parsePptxWithDiagnostics } from '@evoelsewhere/oakit/pptx';
import { renderPptxToPng } from '@evoelsewhere/oakit/pptx/node';

const input = new Uint8Array(await readFile(process.argv[2]));
assert.equal((await parsePptx(input)).slides.length, 1);
assert.equal((await parsePptxWithDiagnostics(input)).document.slides.length, 1);
assert.equal((await renderPptxToPng(input)).slides.length, 1);
const require = createRequire(import.meta.url);
assert.equal(typeof require('@evoelsewhere/oakit').parsePptx, 'function');
assert.equal(
  typeof require('@evoelsewhere/oakit/pptx').parsePptxWithDiagnostics,
  'function',
);
assert.equal(
  typeof require('@evoelsewhere/oakit/pptx/node').renderPptxToPng,
  'function',
);
`,
  );
  const consumerResult = await execFileAsync(
    process.execPath,
    [consumerPath, restoredPath],
    { cwd: installed.consumerDirectory },
  );
  assert.equal(consumerResult.stdout, '');
  assert.equal(consumerResult.stderr, '');
} finally {
  await rm(smokeDirectory, { force: true, recursive: true });
}

console.log(
  'Package smoke passed: packed install, ESM, CJS, subpath exports, and oakit CLI.',
);
