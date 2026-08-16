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

function reportString(value, description) {
  if (typeof value !== 'string' || value.length === 0) {
    throw new TypeError(`${description} must be a non-empty string`);
  }
  return value;
}

function reportMutants(value, file) {
  if (!Array.isArray(value)) {
    throw new TypeError(`Mutation shard mutants for ${file} must be an array`);
  }
  return value;
}

export function reportMutantFingerprint(value) {
  const mutant = reportRecord(value, 'Mutation shard mutant');
  const location = reportRecord(mutant.location, 'Mutation shard location');
  reportRecord(location.start, 'Mutation shard location start');
  reportRecord(location.end, 'Mutation shard location end');
  return JSON.stringify([
    location,
    reportString(mutant.mutatorName, 'Mutation shard mutator name'),
    reportString(mutant.replacement, 'Mutation shard replacement'),
  ]);
}

export function mergeMutationReports(reports, expectedFiles) {
  if (reports.length === 0) {
    throw new Error('At least one mutation shard report is required');
  }
  const first = reportRecord(reports[0], 'Mutation shard report');
  const files = {};
  const fileMutants = new Map();
  const testFiles = {};

  for (const value of reports) {
    const report = reportRecord(value, 'Mutation shard report');
    if (report.schemaVersion !== first.schemaVersion) {
      throw new Error('Mutation shard schema versions do not match');
    }
    const shardFiles = reportRecord(report.files, 'Mutation shard files');
    for (const [file, value] of Object.entries(shardFiles)) {
      const fileReport = reportRecord(
        value,
        `Mutation shard file report for ${file}`,
      );
      const source = reportString(
        fileReport.source,
        `Mutation shard source for ${file}`,
      );
      const language = reportString(
        fileReport.language,
        `Mutation shard language for ${file}`,
      );
      let target = files[file];
      let targetMutants = fileMutants.get(file);
      if (target === undefined || targetMutants === undefined) {
        target = { ...fileReport, language, mutants: [], source };
        targetMutants = new Map();
        files[file] = target;
        fileMutants.set(file, targetMutants);
      } else if (target.source !== source || target.language !== language) {
        throw new Error(
          `Mutation shard source metadata does not match: ${file}`,
        );
      }

      for (const mutant of reportMutants(fileReport.mutants, file)) {
        const fingerprint = reportMutantFingerprint(mutant);
        const existing = targetMutants.get(fingerprint);
        if (existing !== undefined) {
          if (existing.status !== mutant.status) {
            throw new Error(
              `Mutation shard status does not match for ${file}: ${fingerprint}`,
            );
          }
          continue;
        }
        targetMutants.set(fingerprint, mutant);
      }
    }
    const shardTests = reportRecord(
      report.testFiles,
      'Mutation shard test files',
    );
    for (const [file, testFile] of Object.entries(shardTests)) {
      if (
        Object.hasOwn(testFiles, file) &&
        JSON.stringify(testFiles[file]) !== JSON.stringify(testFile)
      ) {
        throw new Error(`Mutation shard test metadata does not match: ${file}`);
      }
      testFiles[file] = testFile;
    }
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

  for (const file of expectedFiles) {
    const target = files[file];
    const mutants = fileMutants.get(file);
    target.mutants = [...mutants.entries()]
      .sort(([left], [right]) => left.localeCompare(right))
      .map(([, mutant], index) => ({ ...mutant, id: String(index) }));
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
