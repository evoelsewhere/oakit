#!/usr/bin/env node

import { createNodeCliIo } from './cli/node-io';
import { runOakitCli } from './cli/run';

declare const __OAKIT_VERSION__: string;

runOakitCli(process.argv.slice(2), createNodeCliIo(), __OAKIT_VERSION__).then(
  (exitCode) => {
    process.exitCode = exitCode;
  },
  (error: unknown) => {
    const message = error instanceof Error ? error.message : String(error);
    process.stderr.write(
      `${JSON.stringify({ error: { code: 'internal-error', message } })}\n`,
    );
    process.exitCode = 1;
  },
);
