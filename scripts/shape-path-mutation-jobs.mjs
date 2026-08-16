import fs from 'node:fs';
import path from 'node:path';
import process from 'node:process';
import { fileURLToPath } from 'node:url';

import { Instrumenter } from '@stryker-mutator/instrumenter';

export const shapePathMutationFile = 'src/formats/pptx/internal/shape-path.ts';

const arithmeticJobCount = 4;
const stringJobCount = 2;
const remainingJobCount = 4;

const projectRoot = path.resolve(
  path.dirname(fileURLToPath(import.meta.url)),
  '..',
);

const silentLogger = {
  debug() {},
  error() {},
  info() {},
  isDebugEnabled() {
    return false;
  },
  warn() {},
};

function positiveInteger(value, name) {
  if (!/^[1-9]\d*$/.test(value)) {
    throw new Error(`${name} must be a positive integer`);
  }
  return Number(value);
}

function locationOrder(left, right) {
  return (
    left.location.start.line - right.location.start.line ||
    left.location.start.column - right.location.start.column ||
    left.location.end.line - right.location.end.line ||
    left.location.end.column - right.location.end.column ||
    left.replacement.localeCompare(right.replacement)
  );
}

function createRangeJobs(mutants, mutatorName, count, allMutatorNames) {
  const selected = mutants
    .filter((mutant) => mutant.mutatorName === mutatorName)
    .sort(locationOrder);
  if (selected.length < count) {
    throw new Error(
      `${mutatorName} requires at least ${count} mutants to create balanced jobs`,
    );
  }

  const excludedMutations = allMutatorNames.filter(
    (name) => name !== mutatorName,
  );
  return Array.from({ length: count }, (_, index) => {
    const startIndex = Math.floor((selected.length * index) / count);
    const endIndex = Math.floor((selected.length * (index + 1)) / count) - 1;
    const slice = selected.slice(startIndex, endIndex + 1);
    const startLine = Math.min(
      ...slice.map((mutant) => mutant.location.start.line + 1),
    );
    const endLine = Math.max(
      ...slice.map((mutant) => mutant.location.end.line + 1),
    );
    return {
      allowedMutations: [mutatorName],
      excludedMutations,
      id: `${mutatorName}-${index + 1}`,
      mutate: `${shapePathMutationFile}:${startLine}-${endLine}`,
      range: { endLine, startLine },
    };
  });
}

function createRemainingJobs(mutants, excludedNames, allMutatorNames) {
  const counts = new Map();
  for (const mutant of mutants) {
    if (!excludedNames.has(mutant.mutatorName)) {
      counts.set(mutant.mutatorName, (counts.get(mutant.mutatorName) ?? 0) + 1);
    }
  }
  if (counts.size < remainingJobCount) {
    throw new Error(
      `Shape path requires at least ${remainingJobCount} remaining mutator families`,
    );
  }

  const groups = Array.from({ length: remainingJobCount }, () => ({
    allowedMutations: [],
    weight: 0,
  }));
  const weightedNames = [...counts].sort(
    ([leftName, leftCount], [rightName, rightCount]) =>
      rightCount - leftCount || leftName.localeCompare(rightName),
  );
  for (const [name, count] of weightedNames) {
    const target = groups.reduce((lightest, candidate) =>
      candidate.weight < lightest.weight ? candidate : lightest,
    );
    target.allowedMutations.push(name);
    target.weight += count;
  }

  return groups.map(({ allowedMutations }, index) => ({
    allowedMutations: allowedMutations.sort(),
    excludedMutations: allMutatorNames.filter(
      (name) => !allowedMutations.includes(name),
    ),
    id: `remaining-${index + 1}`,
    mutate: shapePathMutationFile,
    range: null,
  }));
}

export function createShapePathMutationJobs(mutants) {
  if (mutants.length === 0) {
    throw new Error('Shape path mutation jobs require at least one mutant');
  }
  const allMutatorNames = [
    ...new Set(mutants.map((mutant) => mutant.mutatorName)),
  ].sort();
  const rangeNames = new Set(['ArithmeticOperator', 'StringLiteral']);
  for (const name of rangeNames) {
    if (!allMutatorNames.includes(name)) {
      throw new Error(`Missing required shape path mutator family: ${name}`);
    }
  }

  return [
    ...createRangeJobs(
      mutants,
      'ArithmeticOperator',
      arithmeticJobCount,
      allMutatorNames,
    ),
    ...createRangeJobs(
      mutants,
      'StringLiteral',
      stringJobCount,
      allMutatorNames,
    ),
    ...createRemainingJobs(mutants, rangeNames, allMutatorNames),
  ];
}

export function mutationFingerprint(mutant) {
  return JSON.stringify([
    mutant.location,
    mutant.mutatorName,
    mutant.replacement,
  ]);
}

export function selectShapePathJobMutants(mutants, job) {
  const allowed = new Set(job.allowedMutations);
  return mutants.filter((mutant) => {
    if (!allowed.has(mutant.mutatorName)) {
      return false;
    }
    if (job.range === null) {
      return true;
    }
    const startLine = mutant.location.start.line + 1;
    const endLine = mutant.location.end.line + 1;
    return startLine >= job.range.startLine && endLine <= job.range.endLine;
  });
}

export function verifyShapePathMutationJobs(mutants, jobs) {
  const expected = new Set(mutants.map(mutationFingerprint));
  if (expected.size !== mutants.length) {
    throw new Error('Shape path instrumentation produced duplicate mutants');
  }

  const covered = new Set();
  const workloads = [];
  for (const job of jobs) {
    const selected = selectShapePathJobMutants(mutants, job);
    workloads.push(selected.length);
    for (const mutant of selected) {
      covered.add(mutationFingerprint(mutant));
    }
  }
  if (covered.size !== expected.size) {
    throw new Error(
      `Shape path mutation jobs cover ${covered.size} of ${expected.size} mutants`,
    );
  }
  for (const fingerprint of covered) {
    if (!expected.has(fingerprint)) {
      throw new Error('Shape path mutation jobs selected an unexpected mutant');
    }
  }
  return {
    coveredMutants: covered.size,
    duplicateSelections:
      workloads.reduce((total, count) => total + count, 0) - covered.size,
    workloads,
  };
}

export async function instrumentShapePathMutants() {
  const absolutePath = path.join(projectRoot, shapePathMutationFile);
  const content = fs.readFileSync(absolutePath, 'utf8');
  const instrumenter = new Instrumenter(silentLogger);
  const result = await instrumenter.instrument(
    [{ content, mutate: true, name: absolutePath }],
    { excludedMutations: [], ignorers: [], plugins: null },
  );
  return result.mutants;
}

if (process.argv[1] === fileURLToPath(import.meta.url)) {
  const index = positiveInteger(process.argv[2] ?? '', 'Job number') - 1;
  const mutants = await instrumentShapePathMutants();
  const jobs = createShapePathMutationJobs(mutants);
  verifyShapePathMutationJobs(mutants, jobs);
  const job = jobs[index];
  if (job === undefined) {
    throw new RangeError(`Job number must be between 1 and ${jobs.length}`);
  }
  process.stdout.write(`${JSON.stringify(job)}\n`);
}
