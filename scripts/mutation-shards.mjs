import fs from 'node:fs';
import path from 'node:path';
import process from 'node:process';
import { fileURLToPath } from 'node:url';

import { mutatedFiles } from './mutation-scope.mjs';

const projectRoot = path.resolve(
  path.dirname(fileURLToPath(import.meta.url)),
  '..',
);

function positiveInteger(value, name) {
  if (!/^[1-9]\d*$/.test(value)) {
    throw new Error(`${name} must be a positive integer`);
  }
  return Number(value);
}

export function createMutationShards(files, count, weightOf) {
  if (!Number.isSafeInteger(count) || count <= 0) {
    throw new RangeError('Mutation shard count must be a positive integer');
  }
  if (count > files.length) {
    throw new RangeError('Mutation shard count must not exceed the file count');
  }
  const uniqueFiles = new Set(files);
  if (uniqueFiles.size !== files.length) {
    throw new Error('Mutation shard input contains duplicate files');
  }

  const weighted = files
    .map((file) => ({ file, weight: weightOf(file) }))
    .sort(
      (left, right) =>
        right.weight - left.weight || left.file.localeCompare(right.file),
    );
  if (
    weighted.some(({ weight }) => !Number.isSafeInteger(weight) || weight <= 0)
  ) {
    throw new RangeError('Mutation shard weights must be positive integers');
  }

  const shards = Array.from({ length: count }, () => ({
    files: [],
    weight: 0,
  }));
  for (const entry of weighted) {
    const target = shards.reduce((lightest, candidate) =>
      candidate.weight < lightest.weight ? candidate : lightest,
    );
    target.files.push(entry.file);
    target.weight += entry.weight;
  }
  return shards.map(({ files: shardFiles }) => shardFiles.sort());
}

function sourceBytes(file) {
  return fs.statSync(path.join(projectRoot, file)).size;
}

if (process.argv[1] === fileURLToPath(import.meta.url)) {
  const index = positiveInteger(process.argv[2] ?? '', 'Shard number') - 1;
  const count = positiveInteger(process.argv[3] ?? '', 'Shard count');
  const shards = createMutationShards(mutatedFiles, count, sourceBytes);
  const shard = shards[index];
  if (shard === undefined) {
    throw new RangeError(`Shard number must be between 1 and ${count}`);
  }
  process.stdout.write(`${shard.join(',')}\n`);
}
