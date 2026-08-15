#!/usr/bin/env node

import { createNodeCliIo } from './cli/node-io';
import { launchNodeCli } from './cli/launch';
import { runOakitCli } from './cli/run';

declare const __OAKIT_VERSION__: string;

void launchNodeCli(process.argv.slice(2), __OAKIT_VERSION__, {
  io: createNodeCliIo(),
  runCli: runOakitCli,
  setExitCode(exitCode) {
    process.exitCode = exitCode;
  },
});
