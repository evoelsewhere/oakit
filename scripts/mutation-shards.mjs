import fs from 'node:fs';
import path from 'node:path';
import process from 'node:process';
import { fileURLToPath } from 'node:url';

import { fileMutationShardFiles } from './mutation-release-scope.mjs';

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

export function readMutationWorkloads(
  file = path.join(projectRoot, 'scripts/mutation-workloads.json'),
) {
  const value = JSON.parse(fs.readFileSync(file, 'utf8'));
  if (
    value === null ||
    typeof value !== 'object' ||
    Array.isArray(value) ||
    value.schemaVersion !== 1 ||
    value.files === null ||
    typeof value.files !== 'object' ||
    Array.isArray(value.files)
  ) {
    throw new TypeError('Mutation workload history is invalid');
  }
  return value;
}

function mutationWorkload(file, history) {
  const workload = history.files[file];
  if (
    workload === null ||
    typeof workload !== 'object' ||
    Array.isArray(workload)
  ) {
    throw new Error(`Missing mutation workload history for ${file}`);
  }
  const values = [
    workload.mutants,
    workload.staticMutants,
    workload.testsCompleted,
  ];
  if (
    values.some((value) => !Number.isSafeInteger(value) || value < 0) ||
    workload.mutants === 0
  ) {
    throw new RangeError(`Mutation workload history for ${file} is invalid`);
  }
  return workload;
}

export function mutationTestWorkWeight(file, history) {
  const workload = mutationWorkload(file, history);
  return workload.mutants + workload.testsCompleted;
}

export function mutationWorkWeight(file, history) {
  const workload = mutationWorkload(file, history);
  if (workload.estimatedMilliseconds === undefined) {
    return mutationTestWorkWeight(file, history);
  }
  if (
    !Number.isSafeInteger(workload.estimatedMilliseconds) ||
    workload.estimatedMilliseconds <= 0
  ) {
    throw new RangeError(`Mutation timing history for ${file} is invalid`);
  }
  return workload.estimatedMilliseconds;
}

if (process.argv[1] === fileURLToPath(import.meta.url)) {
  const index = positiveInteger(process.argv[2] ?? '', 'Shard number') - 1;
  const count = positiveInteger(process.argv[3] ?? '', 'Shard count');
  const history = readMutationWorkloads();
  const shards = createMutationShards(fileMutationShardFiles, count, (file) =>
    mutationWorkWeight(file, history),
  );
  const shard = shards[index];
  if (shard === undefined) {
    throw new RangeError(`Shard number must be between 1 and ${count}`);
  }
  process.stdout.write(`${shard.join(',')}\n`);
}
