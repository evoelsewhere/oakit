import fs from 'node:fs';
import path from 'node:path';
import process from 'node:process';
import { fileURLToPath } from 'node:url';

import strykerConfig from '../stryker.config.mjs';

const projectRoot = path.resolve(
  path.dirname(fileURLToPath(import.meta.url)),
  '..',
);
const reportPath = path.join(projectRoot, 'reports/mutation/mutation.json');

if (!fs.existsSync(reportPath)) {
  throw new Error(`Mutation report is missing: ${reportPath}`);
}

const report = JSON.parse(fs.readFileSync(reportPath, 'utf8'));
const expectedFiles = new Set(strykerConfig.mutate);
const reportedFiles = new Set(Object.keys(report.files));
const failures = [];
const totals = new Map();

for (const expectedFile of expectedFiles) {
  if (!reportedFiles.has(expectedFile)) {
    failures.push(`Missing configured mutation file: ${expectedFile}`);
  }
}

for (const reportedFile of reportedFiles) {
  if (!expectedFiles.has(reportedFile)) {
    failures.push(`Unexpected mutation report file: ${reportedFile}`);
  }
}

for (const [fileName, fileReport] of Object.entries(report.files)) {
  const counts = new Map();
  for (const mutant of fileReport.mutants) {
    counts.set(mutant.status, (counts.get(mutant.status) ?? 0) + 1);
    totals.set(mutant.status, (totals.get(mutant.status) ?? 0) + 1);
  }

  const missed = fileReport.mutants.filter(
    (mutant) => mutant.status !== 'Killed' && mutant.status !== 'CompileError',
  );
  if (missed.length > 0) {
    const details = missed
      .map(
        (mutant) =>
          `${mutant.status}#${mutant.id}@${mutant.location.start.line}:${mutant.location.start.column}`,
      )
      .join(', ');
    failures.push(`${fileName}: ${details}`);
  }

  const killed = counts.get('Killed') ?? 0;
  const scored = fileReport.mutants.filter(
    (mutant) => mutant.status !== 'CompileError',
  ).length;
  if (scored === 0) failures.push(`${fileName}: no scored mutants`);
  if (killed !== scored) {
    failures.push(`${fileName}: killed ${killed}/${scored} scored mutants`);
  }
}

if (failures.length > 0) {
  process.exitCode = 1;
  throw new Error(`Strict mutation audit failed:\n- ${failures.join('\n- ')}`);
}

const killed = totals.get('Killed') ?? 0;
const compileErrors = totals.get('CompileError') ?? 0;
console.log(
  `Strict mutation audit passed: ${killed} killed, ${compileErrors} compile errors, 0 missed across ${reportedFiles.size} files.`,
);
