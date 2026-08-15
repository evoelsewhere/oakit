import assert from 'node:assert/strict';
import { execFile } from 'node:child_process';
import { readFile, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import path from 'node:path';
import { createRequire } from 'node:module';
import { mkdtemp } from 'node:fs/promises';
import { promisify } from 'node:util';

import JSZip from 'jszip';

const execFileAsync = promisify(execFile);
const require = createRequire(import.meta.url);
const projectRoot = path.resolve(import.meta.dirname, '..');

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
assert.equal(metadata.name, 'oakit');
assert.equal(metadata.bin.oakit, './dist/cli.js');

const esm = await import('oakit');
const esmPptx = await import('oakit/pptx');
const cjs = require('oakit');
const cjsPptx = require('oakit/pptx');
assert.equal(typeof esm.parsePptx, 'function');
assert.equal(typeof esmPptx.parsePptxWithDiagnostics, 'function');
assert.equal(typeof cjs.parsePptx, 'function');
assert.equal(typeof cjsPptx.parsePptxWithDiagnostics, 'function');

const cliPath = path.join(projectRoot, 'dist/cli.js');
assert.equal(
  (await readFile(cliPath, 'utf8')).startsWith('#!/usr/bin/env node\n'),
  true,
);

const smokeDirectory = await mkdtemp(
  path.join(tmpdir(), 'oakit-package-smoke-'),
);
try {
  const inputPath = path.join(smokeDirectory, 'smoke.pptx');
  const outputPath = path.join(smokeDirectory, 'smoke.json');
  await writeFile(inputPath, await createSmokePptx());

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
} finally {
  await rm(smokeDirectory, { force: true, recursive: true });
}

console.log('Package smoke passed: ESM, CJS, subpath exports, and oakit CLI.');
