import fs from 'node:fs';
import path from 'node:path';
import process from 'node:process';
import { fileURLToPath } from 'node:url';

import { mutatedFiles } from './mutation-scope.mjs';

const projectRoot = path.resolve(
  path.dirname(fileURLToPath(import.meta.url)),
  '..',
);

function reportRecord(value, description) {
  if (value === null || typeof value !== 'object' || Array.isArray(value)) {
    throw new TypeError(`${description} must be an object`);
  }
  return value;
}

export function mergeMutationReports(reports, expectedFiles) {
  if (reports.length === 0) {
    throw new Error('At least one mutation shard report is required');
  }
  const first = reportRecord(reports[0], 'Mutation shard report');
  const files = {};
  const testFiles = {};

  for (const value of reports) {
    const report = reportRecord(value, 'Mutation shard report');
    if (report.schemaVersion !== first.schemaVersion) {
      throw new Error('Mutation shard schema versions do not match');
    }
    const shardFiles = reportRecord(report.files, 'Mutation shard files');
    for (const [file, fileReport] of Object.entries(shardFiles)) {
      if (Object.hasOwn(files, file)) {
        throw new Error(`Duplicate mutation report file: ${file}`);
      }
      files[file] = fileReport;
    }
    const shardTests = reportRecord(
      report.testFiles,
      'Mutation shard test files',
    );
    Object.assign(testFiles, shardTests);
  }

  const expected = new Set(expectedFiles);
  for (const file of Object.keys(files)) {
    if (!expected.has(file)) {
      throw new Error(`Unexpected mutation report file: ${file}`);
    }
  }
  for (const file of expected) {
    if (!Object.hasOwn(files, file)) {
      throw new Error(`Missing mutation report file: ${file}`);
    }
  }

  const orderedFiles = Object.fromEntries(
    expectedFiles.map((file) => [file, files[file]]),
  );
  return {
    ...first,
    config: {
      ...reportRecord(first.config, 'Mutation shard config'),
      mutate: [...expectedFiles],
    },
    files: orderedFiles,
    testFiles,
  };
}

if (process.argv[1] === fileURLToPath(import.meta.url)) {
  const inputDirectory = path.resolve(
    projectRoot,
    process.argv[2] ?? 'reports/mutation/shards',
  );
  const outputPath = path.resolve(
    projectRoot,
    process.argv[3] ?? 'reports/mutation/mutation.json',
  );
  const reportPaths = fs
    .readdirSync(inputDirectory, { recursive: true })
    .filter((file) => file.endsWith('.json'))
    .sort()
    .map((file) => path.join(inputDirectory, file));
  const reports = reportPaths.map((reportPath) =>
    JSON.parse(fs.readFileSync(reportPath, 'utf8')),
  );
  const merged = mergeMutationReports(reports, mutatedFiles);
  fs.mkdirSync(path.dirname(outputPath), { recursive: true });
  fs.writeFileSync(outputPath, `${JSON.stringify(merged)}\n`, 'utf8');
  process.stdout.write(
    `Merged ${reports.length} mutation shards across ${mutatedFiles.length} files.\n`,
  );
}
