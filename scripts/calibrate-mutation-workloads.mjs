import fs from 'node:fs';
import path from 'node:path';
import process from 'node:process';
import { fileURLToPath } from 'node:url';

import { fileMutationShardFiles } from './mutation-release-scope.mjs';
import {
  createMutationShards,
  mutationTestWorkWeight,
  readMutationWorkloads,
} from './mutation-shards.mjs';

const projectRoot = path.resolve(
  path.dirname(fileURLToPath(import.meta.url)),
  '..',
);

export function calibrateMutationWorkloads(history, files, observations) {
  if (
    observations === null ||
    typeof observations !== 'object' ||
    Array.isArray(observations) ||
    observations.schemaVersion !== 1 ||
    observations.workloadSourceRun !== history.sourceRun ||
    !Number.isSafeInteger(observations.runId) ||
    observations.runId <= 0 ||
    !Array.isArray(observations.shardSeconds) ||
    observations.shardSeconds.length === 0 ||
    observations.shardSeconds.some(
      (seconds) => !Number.isSafeInteger(seconds) || seconds <= 0,
    )
  ) {
    throw new TypeError('Mutation timing observations are invalid');
  }
  const shards = createMutationShards(
    files,
    observations.shardSeconds.length,
    (file) => mutationTestWorkWeight(file, history),
  );
  const calibratedFiles = Object.fromEntries(
    Object.entries(history.files).map(([file, workload]) => [
      file,
      { ...workload },
    ]),
  );
  for (const [index, shardFiles] of shards.entries()) {
    const seconds = observations.shardSeconds[index];
    if (seconds === undefined) {
      throw new Error(`Missing mutation timing for shard ${index + 1}`);
    }
    const shardWeight = shardFiles.reduce(
      (total, file) => total + mutationTestWorkWeight(file, history),
      0,
    );
    for (const file of shardFiles) {
      calibratedFiles[file].estimatedMilliseconds = Math.max(
        1,
        Math.round(
          (seconds * 1000 * mutationTestWorkWeight(file, history)) /
            shardWeight,
        ),
      );
    }
  }
  return {
    ...history,
    files: calibratedFiles,
    timingCalibration: {
      runId: observations.runId,
      shardCount: observations.shardSeconds.length,
    },
  };
}

if (process.argv[1] === fileURLToPath(import.meta.url)) {
  const historyPath = path.resolve(
    projectRoot,
    process.argv[2] ?? 'scripts/mutation-workloads.json',
  );
  const observationsPath = path.resolve(
    projectRoot,
    process.argv[3] ?? 'scripts/mutation-timing-observations.json',
  );
  const history = readMutationWorkloads(historyPath);
  const observations = JSON.parse(fs.readFileSync(observationsPath, 'utf8'));
  const calibrated = calibrateMutationWorkloads(
    history,
    fileMutationShardFiles,
    observations,
  );
  fs.writeFileSync(historyPath, `${JSON.stringify(calibrated, null, 2)}\n`);
  process.stdout.write(
    `Calibrated ${fileMutationShardFiles.length} mutation files from run ${observations.runId}.\n`,
  );
}
