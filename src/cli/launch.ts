import type { OakitCliIo } from './run';

export interface NodeCliLaunchDependencies {
  readonly io: OakitCliIo;
  readonly runCli: (
    args: readonly string[],
    io: OakitCliIo,
    version: string,
  ) => Promise<number>;
  readonly setExitCode: (exitCode: number) => void;
}

export function formatInternalError(error: unknown): string {
  const message = error instanceof Error ? error.message : String(error);
  return `${JSON.stringify({ error: { code: 'internal-error', message } })}\n`;
}

export async function launchNodeCli(
  args: readonly string[],
  version: string,
  dependencies: NodeCliLaunchDependencies,
): Promise<void> {
  let exitCode = 1;
  try {
    exitCode = await dependencies.runCli(args, dependencies.io, version);
  } catch (error) {
    dependencies.io.writeStderr(formatInternalError(error));
  }
  dependencies.setExitCode(exitCode);
}
