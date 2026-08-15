import { createHash } from 'node:crypto';
import { access, mkdir, readFile, rename, writeFile } from 'node:fs/promises';
import { spawnSync } from 'node:child_process';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';

import JSZip from 'jszip';

const root = resolve(dirname(fileURLToPath(import.meta.url)), '../..');
const cacheDirectory = join(root, '.cache/pptx-corpus');
const manifestPath = join(root, 'test/corpus/pptx-manifest.json');
const includeLarge = process.argv.includes('--include-large');

async function exists(path) {
  try {
    await access(path);
    return true;
  } catch {
    return false;
  }
}

function digest(bytes) {
  return createHash('sha256').update(bytes).digest('hex');
}

async function slideTextDigest(bytes) {
  const archive = await JSZip.loadAsync(bytes);
  const slideNames = Object.keys(archive.files)
    .filter((name) => /^ppt\/slides\/slide\d+\.xml$/.test(name))
    .sort((left, right) => {
      const leftNumber = Number(left.match(/slide(\d+)/)?.[1]);
      const rightNumber = Number(right.match(/slide(\d+)/)?.[1]);
      return leftNumber - rightNumber;
    });
  const textBySlide = [];
  for (const name of slideNames) {
    const xml = await archive.file(name).async('string');
    textBySlide.push(
      [...xml.matchAll(/<a:t>(.*?)<\/a:t>/g)]
        .map((match) => match[1])
        .join('\n'),
    );
  }
  const normalized = textBySlide
    .join('\n---slide---\n')
    .normalize('NFC')
    .replace(/\r\n/g, '\n');
  return digest(normalized);
}

async function matchesFingerprint(entry, bytes) {
  if (entry.sha256) return digest(bytes) === entry.sha256;
  if (entry.slideTextSha256) {
    return (await slideTextDigest(bytes)) === entry.slideTextSha256;
  }
  throw new Error(`Corpus entry ${entry.id} has no integrity fingerprint`);
}

async function fetchEntry(entry) {
  const destination = join(cacheDirectory, `${entry.id}.pptx`);
  if (await exists(destination)) {
    const cached = await readFile(destination);
    if (await matchesFingerprint(entry, cached)) return destination;
  }

  const response = await fetch(entry.url, { redirect: 'follow' });
  if (!response.ok) {
    throw new Error(`Failed to fetch ${entry.id}: HTTP ${response.status}`);
  }
  const bytes = new Uint8Array(await response.arrayBuffer());
  if (entry.maxBytes && bytes.byteLength > entry.maxBytes) {
    throw new Error(
      `Download for ${entry.id} exceeded ${entry.maxBytes} bytes: received ${bytes.byteLength}`,
    );
  }
  if (!(await matchesFingerprint(entry, bytes))) {
    throw new Error(
      `Integrity fingerprint mismatch for ${entry.id}; the remote corpus changed`,
    );
  }
  await writeFile(destination, bytes);
  return destination;
}

async function createLibreOfficeRoundTrip(sourcePath) {
  const destination = join(cacheDirectory, 'libreoffice-roundtrip.pptx');
  if (await exists(destination)) return destination;

  const outputDirectory = join(cacheDirectory, 'libreoffice-output');
  const profileDirectory = join(cacheDirectory, 'libreoffice-profile');
  await mkdir(outputDirectory, { recursive: true });
  await mkdir(profileDirectory, { recursive: true });
  const executable = process.env.SOFFICE_PATH || 'soffice';
  const conversion = spawnSync(
    executable,
    [
      '--headless',
      `-env:UserInstallation=${pathToFileURL(profileDirectory).href}`,
      '--convert-to',
      'pptx',
      '--outdir',
      outputDirectory,
      sourcePath,
    ],
    { encoding: 'utf8' },
  );
  if (conversion.error || conversion.status !== 0) {
    const reason = conversion.error?.message || conversion.stderr || 'unknown';
    throw new Error(`LibreOffice corpus conversion failed: ${reason}`);
  }
  const generated = join(outputDirectory, sourcePath.split('/').pop());
  if (!(await exists(generated))) {
    throw new Error('LibreOffice did not produce the expected PPTX output');
  }
  await rename(generated, destination);
  return destination;
}

await mkdir(cacheDirectory, { recursive: true });
const manifest = JSON.parse(await readFile(manifestPath, 'utf8'));
const selected = manifest.entries.filter(
  (entry) => entry.tier === 'curated' || includeLarge,
);
const resolvedEntries = [];
for (const entry of selected) {
  const path = await fetchEntry(entry);
  resolvedEntries.push({ ...entry, path });
}

const roundTripSource = resolvedEntries.find(
  (entry) => entry.id === 'powerpoint-columns',
);
if (roundTripSource) {
  const path = await createLibreOfficeRoundTrip(roundTripSource.path);
  resolvedEntries.push({
    id: 'libreoffice-roundtrip',
    minimumSlides: 1,
    path,
    producer: 'LibreOffice headless round-trip',
    tier: 'curated',
  });
}

await writeFile(
  join(cacheDirectory, 'resolved.json'),
  `${JSON.stringify({ version: manifest.version, entries: resolvedEntries }, null, 2)}\n`,
);
console.log(`Prepared ${resolvedEntries.length} PPTX corpus files`);
