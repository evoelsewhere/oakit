import { describe, expect, it } from 'vitest';

import {
  mergeMutationReports,
  reportMutantFingerprint,
} from '../../scripts/merge-mutation-reports.mjs';
import { mutationShardEnvironment } from '../../scripts/mutation-shard-environment.mjs';
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

function fileReport(
  mutants = [],
  source = 'const value = 1;',
  language = 'ts',
) {
  return { language, mutants, source };
}

function reportMutant({
  endLine = 1,
  replacement = '0',
  startLine = 1,
  status = 'Killed',
} = {}) {
  return {
    id: 'source-id',
    location: {
      end: { column: 2, line: endLine },
      start: { column: 1, line: startLine },
    },
    mutatorName: 'ArithmeticOperator',
    replacement,
    status,
  };
}

describe('mutation report sharding', () => {
  it('parses an explicit shard selection and optional mutator exclusions', () => {
    expect(
      mutationShardEnvironment({
        MUTATION_EXCLUDED: 'StringLiteral,BlockStatement',
        MUTATION_FILES: 'a.ts,b.ts:1-20',
        MUTATION_REPORT: 'report.json',
      }),
    ).toEqual({
      excludedMutations: ['StringLiteral', 'BlockStatement'],
      mutate: ['a.ts', 'b.ts:1-20'],
      reportPath: 'report.json',
    });
  });

  it('rejects an empty shard selection or report path', () => {
    expect(() =>
      mutationShardEnvironment({ MUTATION_REPORT: 'report.json' }),
    ).toThrow('MUTATION_FILES must select at least one source file');
    expect(() => mutationShardEnvironment({ MUTATION_FILES: 'a.ts' })).toThrow(
      'MUTATION_REPORT must select a JSON report path',
    );
  });

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
      [report({ 'b.ts': fileReport() }), report({ 'a.ts': fileReport() })],
      ['a.ts', 'b.ts'],
    );

    expect(Object.keys(merged.files)).toEqual(['a.ts', 'b.ts']);
    expect(merged.config.mutate).toEqual(['a.ts', 'b.ts']);
  });

  it('merges and deduplicates matching ranges for one source file', () => {
    const first = reportMutant({ replacement: '-' });
    const second = reportMutant({ replacement: '*', startLine: 2 });
    const merged = mergeMutationReports(
      [
        report({ 'a.ts': fileReport([first]) }),
        report({ 'a.ts': fileReport([first, second]) }),
      ],
      ['a.ts'],
    );

    expect(merged.files['a.ts'].mutants).toHaveLength(2);
    expect(merged.files['a.ts'].mutants.map((mutant) => mutant.id)).toEqual([
      '0',
      '1',
    ]);
    expect(
      new Set(merged.files['a.ts'].mutants.map(reportMutantFingerprint)).size,
    ).toBe(2);
  });

  it('rejects conflicting duplicate mutation results and source metadata', () => {
    const killed = reportMutant();
    const survived = reportMutant({ status: 'Survived' });
    expect(() =>
      mergeMutationReports(
        [
          report({ 'a.ts': fileReport([killed]) }),
          report({ 'a.ts': fileReport([survived]) }),
        ],
        ['a.ts'],
      ),
    ).toThrow('Mutation shard status does not match for a.ts');
    expect(() =>
      mergeMutationReports(
        [
          report({ 'a.ts': fileReport([], 'source one') }),
          report({ 'a.ts': fileReport([], 'source two') }),
        ],
        ['a.ts'],
      ),
    ).toThrow('Mutation shard source metadata does not match: a.ts');
    expect(() =>
      mergeMutationReports(
        [
          report({ 'a.ts': fileReport([], 'source', 'ts') }),
          report({ 'a.ts': fileReport([], 'source', 'js') }),
        ],
        ['a.ts'],
      ),
    ).toThrow('Mutation shard source metadata does not match: a.ts');
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
        [
          report({ 'a.ts': fileReport() }),
          report({ 'b.ts': fileReport() }, '2.0'),
        ],
        ['a.ts', 'b.ts'],
      ),
    ).toThrow('Mutation shard schema versions do not match');
    expect(() => mergeMutationReports([report({})], ['a.ts'])).toThrow(
      'Missing mutation report file: a.ts',
    );
    expect(() =>
      mergeMutationReports([report({ 'b.ts': fileReport() })], ['a.ts']),
    ).toThrow('Unexpected mutation report file: b.ts');
  });
});

describe('shape path mutation workload', () => {
  it('partitions every instrumented mutant into eighteen bounded jobs', async () => {
    const mutants = await instrumentShapePathMutants();
    const jobs = createShapePathMutationJobs(mutants);
    const verification = verifyShapePathMutationJobs(mutants, jobs);

    expect(mutants).toHaveLength(8195);
    expect(new Set(mutants.map(mutationFingerprint)).size).toBe(mutants.length);
    expect(jobs).toHaveLength(18);
    expect(verification.coveredMutants).toBe(mutants.length);
    expect(verification.duplicateSelections).toBe(0);
    expect(Math.max(...verification.workloads)).toBeLessThanOrEqual(1100);
    expect(Math.max(...verification.workloads.slice(6))).toBeLessThanOrEqual(
      450,
    );
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
