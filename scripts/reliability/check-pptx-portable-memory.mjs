import { spawn } from 'node:child_process';
import { mkdir, mkdtemp, readFile, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join, resolve } from 'node:path';

import JSZip from 'jszip';

import { createPptx } from '../../dist/index.js';

const MEBIBYTE = 1024 * 1024;
const TIERS = [
  {
    maxPeakRssBytes: 384 * MEBIBYTE,
    maxTotalDurationMs: 15_000,
    name: 'small',
    sourceBytes: 1 * MEBIBYTE,
  },
  {
    maxPeakRssBytes: 768 * MEBIBYTE,
    maxTotalDurationMs: 60_000,
    name: 'normal',
    sourceBytes: 25 * MEBIBYTE,
  },
  {
    maxPeakRssBytes: 2_048 * MEBIBYTE,
    maxTotalDurationMs: 180_000,
    name: 'large',
    sourceBytes: 100 * MEBIBYTE,
  },
];

function scene() {
  return {
    layouts: [],
    masters: [],
    media: [],
    schemaVersion: 2,
    size: { height: 540, width: 960 },
    slides: [
      {
        elements: [
          {
            authored: {
              transform: { height: 80, width: 320, x: 24, y: 32 },
            },
            key: 'memory-text',
            resolved: { hidden: false },
            text: {
              body: {},
              paragraphs: [
                {
                  children: [
                    {
                      key: 'memory-run',
                      text: 'Portable memory evidence',
                      type: 'run',
                    },
                  ],
                  key: 'memory-paragraph',
                },
              ],
            },
            type: 'text',
          },
        ],
        key: 'memory-slide',
      },
    ],
    themes: [],
  };
}

async function createExactPackage(baseBytes, targetBytes) {
  const archive = await JSZip.loadAsync(baseBytes);
  const contentTypes = archive.file('[Content_Types].xml');
  if (!contentTypes) throw new Error('Created package has no content types');
  const contentTypesXml = await contentTypes.async('string');
  archive.file(
    '[Content_Types].xml',
    contentTypesXml.replace(
      '</Types>',
      '<Default Extension="bin" ContentType="application/octet-stream"/></Types>',
    ),
  );
  const firstName = 'ppt/agentData/padding-1.bin';
  const secondName = 'ppt/agentData/padding-2.bin';
  archive.file(firstName, new Uint8Array(), { compression: 'STORE' });
  archive.file(secondName, new Uint8Array(), { compression: 'STORE' });
  const emptyPackage = await archive.generateAsync({
    compression: 'STORE',
    type: 'uint8array',
  });
  const payloadBytes = targetBytes - emptyPackage.byteLength;
  if (payloadBytes <= 0) throw new Error('Memory tier is smaller than PPTX');
  const firstBytes = Math.min(payloadBytes, 60 * MEBIBYTE);
  const secondBytes = payloadBytes - firstBytes;
  if (secondBytes > 60 * MEBIBYTE) {
    throw new Error('Memory tier requires a ZIP part above its safety budget');
  }
  archive.file(firstName, new Uint8Array(firstBytes), {
    compression: 'STORE',
  });
  archive.file(secondName, new Uint8Array(secondBytes), {
    compression: 'STORE',
  });
  const result = await archive.generateAsync({
    compression: 'STORE',
    type: 'uint8array',
  });
  if (result.byteLength !== targetBytes) {
    throw new Error(
      `Expected ${targetBytes} package bytes, received ${result.byteLength}`,
    );
  }
  return result;
}

function runWorker(sourcePath, tier) {
  return new Promise((resolveWorker, rejectWorker) => {
    const workerPath = resolve(
      'scripts/reliability/pptx-portable-memory-worker.mjs',
    );
    const child = spawn(
      process.execPath,
      [
        '--expose-gc',
        '--max-old-space-size=2048',
        workerPath,
        sourcePath,
        tier.name,
      ],
      { stdio: ['ignore', 'pipe', 'pipe'] },
    );
    let stdout = '';
    let stderr = '';
    child.stdout.setEncoding('utf8');
    child.stderr.setEncoding('utf8');
    child.stdout.on('data', (chunk) => {
      stdout += chunk;
    });
    child.stderr.on('data', (chunk) => {
      stderr += chunk;
    });
    child.on('error', rejectWorker);
    child.on('close', (code, signal) => {
      if (code !== 0) {
        rejectWorker(
          new Error(
            `Portable memory worker ${tier.name} failed (${signal ?? code}): ${stderr}`,
          ),
        );
        return;
      }
      try {
        resolveWorker(JSON.parse(stdout.trim()));
      } catch (error) {
        rejectWorker(
          new Error(`Portable memory worker returned invalid evidence`, {
            cause: error,
          }),
        );
      }
    });
  });
}

function assertEvidence(result, tier) {
  if (
    result.byteEqual !== true ||
    result.sourceSha256 !== result.outputSha256 ||
    result.sourceBytes !== tier.sourceBytes ||
    result.outputBytes !== tier.sourceBytes ||
    result.reportLevel !== 'R0'
  ) {
    throw new Error(`Portable memory tier ${tier.name} lost exactness`);
  }
  if (result.peakRssBytes > tier.maxPeakRssBytes) {
    throw new Error(
      `Portable memory tier ${tier.name} exceeded RSS budget: ${result.peakRssBytes} > ${tier.maxPeakRssBytes}`,
    );
  }
  if (result.totalDurationMs > tier.maxTotalDurationMs) {
    throw new Error(
      `Portable memory tier ${tier.name} exceeded duration budget: ${result.totalDurationMs} > ${tier.maxTotalDurationMs}`,
    );
  }
}

const selectedName = process.argv
  .find((argument) => argument.startsWith('--tier='))
  ?.slice('--tier='.length);
const tiers = selectedName
  ? TIERS.filter((tier) => tier.name === selectedName)
  : TIERS;
if (tiers.length === 0) throw new Error(`Unknown memory tier ${selectedName}`);

const temporaryDirectory = await mkdtemp(
  join(tmpdir(), 'oakit-portable-memory-'),
);
try {
  const base = await createPptx(scene());
  const evidence = [];
  for (const tier of tiers) {
    const sourcePath = join(temporaryDirectory, `${tier.name}.pptx`);
    const source = await createExactPackage(base.data, tier.sourceBytes);
    await writeFile(sourcePath, source);
    const result = await runWorker(sourcePath, tier);
    assertEvidence(result, tier);
    evidence.push({
      ...result,
      budgets: {
        maxPeakRssBytes: tier.maxPeakRssBytes,
        maxTotalDurationMs: tier.maxTotalDurationMs,
      },
    });
    await rm(sourcePath, { force: true });
    globalThis.gc?.();
  }
  const reportDirectory = resolve('reports/reliability');
  await mkdir(reportDirectory, { recursive: true });
  const reportPath = join(reportDirectory, 'pptx-portable-memory.json');
  await writeFile(
    reportPath,
    `${JSON.stringify(
      {
        generatedAt: new Date().toISOString(),
        node: process.version,
        platform: `${process.platform}-${process.arch}`,
        tiers: evidence,
        version: 1,
      },
      null,
      2,
    )}\n`,
  );
  const verified = JSON.parse(await readFile(reportPath, 'utf8'));
  for (const tier of verified.tiers) {
    process.stdout.write(
      `${tier.tier}: ${tier.sourceBytes} bytes, peak RSS ${tier.peakRssBytes}, ${tier.totalDurationMs} ms\n`,
    );
  }
  process.stdout.write(`Evidence: ${reportPath}\n`);
} finally {
  await rm(temporaryDirectory, { force: true, recursive: true });
}
