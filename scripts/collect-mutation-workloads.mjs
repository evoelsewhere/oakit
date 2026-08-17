import fs from 'node:fs';
import path from 'node:path';
import process from 'node:process';
import { fileURLToPath } from 'node:url';

import { fileMutationShardFiles } from './mutation-release-scope.mjs';

const projectRoot = path.resolve(
  path.dirname(fileURLToPath(import.meta.url)),
  '..',
);

function record(value, description) {
  if (value === null || typeof value !== 'object' || Array.isArray(value)) {
    throw new TypeError(`${description} must be an object`);
  }
  return value;
}

export function collectMutationWorkloads(reports, expectedFiles, sourceRun) {
  if (!Number.isSafeInteger(sourceRun) || sourceRun <= 0) {
    throw new RangeError('Mutation workload source run must be positive');
  }
  const expected = new Set(expectedFiles);
  const files = {};
  for (const value of reports) {
    const report = record(value, 'Mutation workload report');
    const reportFiles = record(report.files, 'Mutation workload report files');
    for (const [file, fileValue] of Object.entries(reportFiles)) {
      if (!expected.has(file)) continue;
      if (Object.hasOwn(files, file)) {
        throw new Error(`Duplicate mutation workload file: ${file}`);
      }
      const fileReport = record(
        fileValue,
        `Mutation workload file report for ${file}`,
      );
      if (!Array.isArray(fileReport.mutants)) {
        throw new TypeError(
          `Mutation workload mutants for ${file} must be an array`,
        );
      }
      const mutants = fileReport.mutants;
      files[file] = {
        mutants: mutants.length,
        staticMutants: mutants.filter((mutant) => mutant.static === true)
          .length,
        testsCompleted: mutants.reduce((total, mutant) => {
          const count = mutant.testsCompleted ?? 0;
          if (!Number.isSafeInteger(count) || count < 0) {
            throw new RangeError(
              `Mutation workload testsCompleted for ${file} must be non-negative`,
            );
          }
          return total + count;
        }, 0),
      };
    }
  }
  for (const file of expectedFiles) {
    if (!Object.hasOwn(files, file)) {
      throw new Error(`Missing mutation workload file: ${file}`);
    }
  }
  return {
    files: Object.fromEntries(expectedFiles.map((file) => [file, files[file]])),
    schemaVersion: 1,
    sourceRun,
  };
}

if (process.argv[1] === fileURLToPath(import.meta.url)) {
  const inputDirectory = path.resolve(projectRoot, process.argv[2] ?? '');
  const outputPath = path.resolve(
    projectRoot,
    process.argv[3] ?? 'scripts/mutation-workloads.json',
  );
  const sourceRun = Number(process.argv[4]);
  const reportPaths = fs
    .readdirSync(inputDirectory, { recursive: true })
    .filter((file) =>
      /(?:^|\/)mutation-shard-\d+\/shard-\d+\.json$/.test(
        file.split(path.sep).join('/'),
      ),
    )
    .sort()
    .map((file) => path.join(inputDirectory, file));
  const reports = reportPaths.map((reportPath) =>
    JSON.parse(fs.readFileSync(reportPath, 'utf8')),
  );
  const workloads = collectMutationWorkloads(
    reports,
    fileMutationShardFiles,
    sourceRun,
  );
  fs.writeFileSync(outputPath, `${JSON.stringify(workloads, null, 2)}\n`);
  process.stdout.write(
    `Recorded mutation workload for ${fileMutationShardFiles.length} files from run ${sourceRun}.\n`,
  );
}
