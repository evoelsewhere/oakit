import { describe, expect, it } from 'vitest';

import { mergeMutationReports } from '../../scripts/merge-mutation-reports.mjs';
import { mutatedFiles } from '../../scripts/mutation-scope.mjs';
import { createMutationShards } from '../../scripts/mutation-shards.mjs';
import {
  createShapePathMutationJobs,
  instrumentShapePathMutants,
  mutationFingerprint,
  selectShapePathJobMutants,
  verifyShapePathMutationJobs,
} from '../../scripts/shape-path-mutation-jobs.mjs';

function report(files, schemaVersion = '1.0') {
  return {
    config: { mutate: Object.keys(files) },
    files,
    framework: { name: 'vitest' },
    projectRoot: '/workspace',
    schemaVersion,
    testFiles: {},
    thresholds: { high: 100, low: 100 },
  };
}

describe('mutation report sharding', () => {
  it('assigns every configured source to exactly one balanced shard', () => {
    const weights = new Map(
      mutatedFiles.map((file, index) => [file, index + 1]),
    );
    const shards = createMutationShards(mutatedFiles, 8, (file) =>
      weights.get(file),
    );
    const flattened = shards.flat();

    expect(shards).toHaveLength(8);
    expect(new Set(flattened).size).toBe(mutatedFiles.length);
    expect(flattened.toSorted()).toEqual(mutatedFiles.toSorted());
    expect(shards.every((shard) => shard.length > 0)).toBe(true);
  });

  it.each([0, -1, 1.5, Number.MAX_SAFE_INTEGER + 1])(
    'rejects invalid shard count %s',
    (count) => {
      expect(() => createMutationShards(['a.ts'], count, () => 1)).toThrow(
        'Mutation shard count must be a positive integer',
      );
    },
  );

  it('rejects excess shards, duplicate files, and invalid weights', () => {
    expect(() => createMutationShards(['a.ts'], 2, () => 1)).toThrow(
      'Mutation shard count must not exceed the file count',
    );
    expect(() => createMutationShards(['a.ts', 'a.ts'], 1, () => 1)).toThrow(
      'Mutation shard input contains duplicate files',
    );
    expect(() => createMutationShards(['a.ts'], 1, () => 0)).toThrow(
      'Mutation shard weights must be positive integers',
    );
  });

  it('merges complete disjoint reports in configured file order', () => {
    const merged = mergeMutationReports(
      [
        report({ 'b.ts': { mutants: [] } }),
        report({ 'a.ts': { mutants: [] } }),
      ],
      ['a.ts', 'b.ts'],
    );

    expect(Object.keys(merged.files)).toEqual(['a.ts', 'b.ts']);
    expect(merged.config.mutate).toEqual(['a.ts', 'b.ts']);
  });

  it('rejects empty, malformed, mismatched, duplicate, missing, and extra reports', () => {
    expect(() => mergeMutationReports([], ['a.ts'])).toThrow(
      'At least one mutation shard report is required',
    );
    expect(() => mergeMutationReports([null], ['a.ts'])).toThrow(
      'Mutation shard report must be an object',
    );
    expect(() =>
      mergeMutationReports(
        [report({ 'a.ts': {} }), report({ 'b.ts': {} }, '2.0')],
        ['a.ts', 'b.ts'],
      ),
    ).toThrow('Mutation shard schema versions do not match');
    expect(() =>
      mergeMutationReports(
        [report({ 'a.ts': {} }), report({ 'a.ts': {} })],
        ['a.ts'],
      ),
    ).toThrow('Duplicate mutation report file: a.ts');
    expect(() => mergeMutationReports([report({})], ['a.ts'])).toThrow(
      'Missing mutation report file: a.ts',
    );
    expect(() =>
      mergeMutationReports([report({ 'b.ts': {} })], ['a.ts']),
    ).toThrow('Unexpected mutation report file: b.ts');
  });
});

describe('shape path mutation workload', () => {
  it('partitions every instrumented mutant into ten bounded jobs', async () => {
    const mutants = await instrumentShapePathMutants();
    const jobs = createShapePathMutationJobs(mutants);
    const verification = verifyShapePathMutationJobs(mutants, jobs);

    expect(mutants).toHaveLength(8195);
    expect(new Set(mutants.map(mutationFingerprint)).size).toBe(mutants.length);
    expect(jobs).toHaveLength(10);
    expect(verification.coveredMutants).toBe(mutants.length);
    expect(Math.max(...verification.workloads)).toBeLessThanOrEqual(1100);
    expect(verification.workloads.every((count) => count > 0)).toBe(true);
  }, 15_000);

  it('keeps range jobs within their declared mutator family and lines', () => {
    const mutants = [
      {
        location: {
          end: { column: 4, line: 2 },
          start: { column: 1, line: 2 },
        },
        mutatorName: 'ArithmeticOperator',
        replacement: 'left - right',
      },
      {
        location: {
          end: { column: 4, line: 8 },
          start: { column: 1, line: 8 },
        },
        mutatorName: 'StringLiteral',
        replacement: '""',
      },
    ];
    const job = {
      allowedMutations: ['ArithmeticOperator'],
      excludedMutations: ['StringLiteral'],
      id: 'arithmetic-test',
      mutate: 'source.ts:3-3',
      range: { endLine: 3, startLine: 3 },
    };

    expect(selectShapePathJobMutants(mutants, job)).toEqual([mutants[0]]);
  });

  it('rejects empty, incomplete, and duplicate mutation workloads', () => {
    expect(() => createShapePathMutationJobs([])).toThrow(
      'Shape path mutation jobs require at least one mutant',
    );
    const duplicate = {
      location: {
        end: { column: 2, line: 0 },
        start: { column: 1, line: 0 },
      },
      mutatorName: 'ArithmeticOperator',
      replacement: '-',
    };
    expect(() =>
      verifyShapePathMutationJobs([duplicate, duplicate], []),
    ).toThrow('Shape path instrumentation produced duplicate mutants');
    expect(() =>
      verifyShapePathMutationJobs(
        [duplicate],
        [
          {
            allowedMutations: ['StringLiteral'],
            excludedMutations: ['ArithmeticOperator'],
            id: 'missing',
            mutate: 'source.ts',
            range: null,
          },
        ],
      ),
    ).toThrow('Shape path mutation jobs cover 0 of 1 mutants');
  });
});
