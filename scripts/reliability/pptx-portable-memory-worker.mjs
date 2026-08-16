import { createHash } from 'node:crypto';
import { Buffer } from 'node:buffer';
import { readFile } from 'node:fs/promises';
import { performance } from 'node:perf_hooks';

import {
  parsePptxRoundTripJson,
  readPptxRoundTrip,
  serializePptxRoundTripJson,
  writePptxRoundTrip,
} from '../../dist/index.js';

function digest(bytes) {
  return createHash('sha256').update(bytes).digest('hex');
}

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

function exactBytes(left, right) {
  if (left.byteLength !== right.byteLength) return false;
  for (let index = 0; index < left.byteLength; index += 1) {
    if (left[index] !== right[index]) return false;
  }
  return true;
}

const sourcePath = process.argv[2];
const tier = process.argv[3];
if (!sourcePath || !tier) {
  throw new Error('Expected a source path and memory tier');
}

const sourceBuffer = await readFile(sourcePath);
const source = new Uint8Array(
  sourceBuffer.buffer,
  sourceBuffer.byteOffset,
  sourceBuffer.byteLength,
);
const sourceSha256 = digest(source);
globalThis.gc?.();
const samples = [memorySample('source-loaded')];
const timings = {};

async function timed(stage, operation) {
  const started = performance.now();
  const result = await operation();
  timings[stage] = Math.round(performance.now() - started);
  samples.push(memorySample(stage));
  return result;
}

const runtime = await timed('read-runtime', () => readPptxRoundTrip(source));
const portable = await timed('serialize-portable', () =>
  serializePptxRoundTripJson(runtime),
);
const wireJson = await timed('stringify-json', () =>
  Promise.resolve(JSON.stringify(portable)),
);
const wireValue = await timed('parse-json', () =>
  Promise.resolve(JSON.parse(wireJson)),
);
const restored = await timed('parse-portable', () =>
  parsePptxRoundTripJson(wireValue),
);
const output = await timed('write-runtime', () => writePptxRoundTrip(restored));
const outputSha256 = digest(output.data);
const byteEqual = exactBytes(output.data, source);
samples.push(memorySample('verified-output'));

const peak = (field) => Math.max(...samples.map((sample) => sample[field]));
const totalDurationMs = Object.values(timings).reduce(
  (total, duration) => total + duration,
  0,
);

process.stdout.write(
  `${JSON.stringify({
    byteEqual,
    outputBytes: output.data.byteLength,
    outputSha256,
    peakArrayBuffersBytes: peak('arrayBuffers'),
    peakExternalBytes: peak('external'),
    peakHeapUsedBytes: peak('heapUsed'),
    peakRssBytes: peak('rss'),
    reportLevel: output.report.level,
    samples,
    sourceBytes: source.byteLength,
    sourceSha256,
    tier,
    timingsMs: timings,
    totalDurationMs,
    wireJsonBytes: Buffer.byteLength(wireJson, 'utf8'),
  })}\n`,
);
