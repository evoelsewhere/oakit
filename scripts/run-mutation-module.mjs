import { spawnSync } from 'node:child_process';
import path from 'node:path';
import process from 'node:process';
import { fileURLToPath } from 'node:url';

import { resolveMutationModule } from './mutation-modules.mjs';

const projectRoot = path.resolve(
  path.dirname(fileURLToPath(import.meta.url)),
  '..',
);
const args = process.argv.slice(2).filter((argument) => argument !== '--');
const name = args.find((argument) => !argument.startsWith('--')) ?? '';
resolveMutationModule(name);
const force = args.includes('--force');
const dynamicOnly = args.includes('--dynamic');
const unknownFlags = args.filter(
  (argument) =>
    argument.startsWith('--') &&
    argument !== '--dynamic' &&
    argument !== '--force',
);
if (unknownFlags.length > 0) {
  throw new Error(`Unknown mutation module option ${unknownFlags.join(', ')}`);
}

const command = process.platform === 'win32' ? 'pnpm.cmd' : 'pnpm';
const result = spawnSync(
  command,
  [
    'exec',
    'stryker',
    'run',
    'stryker.module.config.mjs',
    ...(force ? ['--force'] : []),
  ],
  {
    cwd: projectRoot,
    env: {
      ...process.env,
      MUTATION_DYNAMIC: dynamicOnly ? '1' : '0',
      MUTATION_MODULE: name,
    },
    stdio: 'inherit',
  },
);
if (result.error !== undefined) throw result.error;
process.exitCode = result.status ?? 1;
