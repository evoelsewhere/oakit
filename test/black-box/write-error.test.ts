import { describe, expect, it } from 'vitest';

import { PptxWriteError } from '../../src';

describe('PowerPoint write errors', () => {
  it('exposes a stable public error identity and code', () => {
    const error = new PptxWriteError('package-build-failed', 'Could not build');

    expect(error).toBeInstanceOf(Error);
    expect(error.name).toBe('PptxWriteError');
    expect(error.message).toBe('Could not build');
    expect(error.code).toBe('package-build-failed');
    expect(error.issues).toEqual([]);
  });

  it('preserves the original cause', () => {
    const cause = new Error('ZIP failure');
    const error = new PptxWriteError(
      'package-build-failed',
      'Could not build',
      { cause },
    );

    expect(error.cause).toBe(cause);
  });

  it('captures immutable copies of validation issues', () => {
    const sourceIssue = {
      code: 'invalid-scene-document' as const,
      message: 'Expected an object',
      path: '$',
    };
    const sourceIssues = [sourceIssue];
    const error = new PptxWriteError('invalid-scene', 'Invalid scene', {
      issues: sourceIssues,
    });

    sourceIssue.message = 'changed';
    sourceIssues.push({ ...sourceIssue, path: '$.slides' });

    expect(error.issues).toEqual([
      {
        code: 'invalid-scene-document',
        message: 'Expected an object',
        path: '$',
      },
    ]);
    expect(Object.isFrozen(error.issues)).toBe(true);
    expect(Object.isFrozen(error.issues[0])).toBe(true);
  });

  it('never exposes partial package bytes', () => {
    const error = new PptxWriteError('verification-failed', 'Invalid output');

    expect('data' in error).toBe(false);
  });

  it.each([
    'invalid-snapshot',
    'snapshot-consistency-failed',
    'unsupported-edit-operation',
  ] as const)('exposes round-trip error code %s', (code) => {
    expect(new PptxWriteError(code, 'Round-trip failed').code).toBe(code);
  });
});
