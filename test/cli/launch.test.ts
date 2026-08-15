import { describe, expect, it, vi } from 'vitest';

import {
  formatInternalError,
  launchNodeCli,
  type NodeCliLaunchDependencies,
} from '../../src/cli/launch';
import type { OakitCliIo } from '../../src/cli/run';

function createFixture(runCli: NodeCliLaunchDependencies['runCli']): {
  dependencies: NodeCliLaunchDependencies;
  io: OakitCliIo;
  setExitCode: ReturnType<typeof vi.fn<(exitCode: number) => void>>;
  stderr: string[];
} {
  const stderr: string[] = [];
  const io: OakitCliIo = {
    readFile: vi.fn(() => Promise.resolve(new Uint8Array())),
    readStdin: vi.fn(() => Promise.resolve(new Uint8Array())),
    writeFile: vi.fn(() => Promise.resolve()),
    writeStderr(value) {
      stderr.push(value);
    },
    writeStdout: vi.fn(),
  };
  const setExitCode = vi.fn<(exitCode: number) => void>();
  return {
    dependencies: { io, runCli, setExitCode },
    io,
    setExitCode,
    stderr,
  };
}

describe('Node CLI launch boundary', () => {
  it.each([0, 1, 2])('preserves runner exit code %s', async (exitCode) => {
    const runCli = vi.fn(() => Promise.resolve(exitCode));
    const fixture = createFixture(runCli);

    await expect(
      launchNodeCli(['deck.pptx'], '1.2.3', fixture.dependencies),
    ).resolves.toBeUndefined();

    expect(runCli).toHaveBeenCalledExactlyOnceWith(
      ['deck.pptx'],
      fixture.io,
      '1.2.3',
    );
    expect(fixture.setExitCode).toHaveBeenCalledExactlyOnceWith(exitCode);
    expect(fixture.stderr).toEqual([]);
  });

  it('serializes an unexpected Error without exposing its stack', async () => {
    const fixture = createFixture(() =>
      Promise.reject(new Error('unexpected failure')),
    );

    await launchNodeCli([], '1.2.3', fixture.dependencies);

    expect(fixture.stderr).toEqual([
      '{"error":{"code":"internal-error","message":"unexpected failure"}}\n',
    ]);
    expect(fixture.stderr[0]).not.toContain('stack');
    expect(fixture.setExitCode).toHaveBeenCalledExactlyOnceWith(1);
  });

  it('serializes non-Error failures deterministically', () => {
    expect(formatInternalError('broken')).toBe(
      '{"error":{"code":"internal-error","message":"broken"}}\n',
    );
  });
});
