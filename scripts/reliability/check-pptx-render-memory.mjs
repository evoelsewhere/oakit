import { spawn } from 'node:child_process';
import { mkdir, readFile, writeFile } from 'node:fs/promises';
import { join, resolve } from 'node:path';
import { clearTimeout, setTimeout } from 'node:timers';

const MEBIBYTE = 1024 * 1024;
const TIERS = [
  {
    maxPeakRssBytes: 384 * MEBIBYTE,
    maxTotalDurationMs: 15_000,
    maxWorkerMs: 30_000,
    name: 'small',
    slideCount: 1,
  },
  {
    maxPeakRssBytes: 512 * MEBIBYTE,
    maxTotalDurationMs: 120_000,
    maxWorkerMs: 180_000,
    name: 'normal',
    slideCount: 25,
  },
  {
    maxPeakRssBytes: 768 * MEBIBYTE,
    maxTotalDurationMs: 360_000,
    maxWorkerMs: 480_000,
    name: 'large',
    slideCount: 100,
  },
];

function runWorker(tier) {
  return new Promise((resolveWorker, rejectWorker) => {
    const workerPath = resolve(
      'scripts/reliability/pptx-render-memory-worker.mjs',
    );
    const child = spawn(
      process.execPath,
      [
        '--expose-gc',
        '--max-old-space-size=1024',
        workerPath,
        tier.name,
        String(tier.slideCount),
      ],
      {
        env: { ...process.env, PATH: '' },
        stdio: ['ignore', 'pipe', 'pipe'],
      },
    );
    const timer = setTimeout(() => {
      child.kill('SIGKILL');
    }, tier.maxWorkerMs);
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
    child.on('error', (error) => {
      clearTimeout(timer);
      rejectWorker(error);
    });
    child.on('close', (code, signal) => {
      clearTimeout(timer);
      if (code !== 0) {
        rejectWorker(
          new Error(
            `Render memory worker ${tier.name} failed (${signal ?? code}): ${stderr}`,
          ),
        );
        return;
      }
      try {
        resolveWorker(JSON.parse(stdout.trim()));
      } catch (error) {
        rejectWorker(
          new Error('Render memory worker returned invalid evidence', {
            cause: error,
          }),
        );
      }
    });
  });
}

function assertDigest(value, name) {
  if (typeof value !== 'string' || !/^[a-f\d]{64}$/.test(value)) {
    throw new Error(`Render memory tier has invalid ${name}`);
  }
}

function assertEvidence(result, tier) {
  if (
    result.tier !== tier.name ||
    result.slideCount !== tier.slideCount ||
    result.officeRuntimeRequired !== false ||
    result.pathEmpty !== true ||
    result.svgBytes <= 0 ||
    result.pngBytes <= 0 ||
    result.warningCount !== tier.slideCount * 2
  ) {
    throw new Error(`Render memory tier ${tier.name} failed semantic evidence`);
  }
  assertDigest(result.svgSha256, 'SVG digest');
  assertDigest(result.pngSha256, 'PNG digest');
  if (result.peakRssBytes > tier.maxPeakRssBytes) {
    throw new Error(
      `Render memory tier ${tier.name} exceeded RSS budget: ${result.peakRssBytes} > ${tier.maxPeakRssBytes}`,
    );
  }
  if (result.totalDurationMs > tier.maxTotalDurationMs) {
    throw new Error(
      `Render memory tier ${tier.name} exceeded duration budget: ${result.totalDurationMs} > ${tier.maxTotalDurationMs}`,
    );
  }
}

const selectedName = process.argv
  .find((argument) => argument.startsWith('--tier='))
  ?.slice('--tier='.length);
const tiers = selectedName
  ? TIERS.filter((tier) => tier.name === selectedName)
  : TIERS;
if (tiers.length === 0)
  throw new Error(`Unknown render memory tier ${selectedName}`);

const evidence = [];
for (const tier of tiers) {
  const result = await runWorker(tier);
  assertEvidence(result, tier);
  evidence.push({
    ...result,
    budgets: {
      maxPeakRssBytes: tier.maxPeakRssBytes,
      maxTotalDurationMs: tier.maxTotalDurationMs,
      maxWorkerMs: tier.maxWorkerMs,
    },
  });
}

const reportDirectory = resolve('reports/reliability');
await mkdir(reportDirectory, { recursive: true });
const reportPath = join(reportDirectory, 'pptx-render-memory.json');
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
    `${tier.tier}: ${tier.slideCount} slides, SVG ${tier.svgBytes} bytes, PNG ${tier.pngBytes} bytes, peak RSS ${tier.peakRssBytes}, ${tier.totalDurationMs} ms\n`,
  );
}
process.stdout.write(`Evidence: ${reportPath}\n`);
