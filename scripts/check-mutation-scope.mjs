import fs from 'node:fs';
import path from 'node:path';
import process from 'node:process';
import { fileURLToPath } from 'node:url';

import {
  excludedMutationFiles,
  mutatedFiles,
  pendingMutationFiles,
} from './mutation-scope.mjs';

const projectRoot = path.resolve(
  path.dirname(fileURLToPath(import.meta.url)),
  '..',
);
const sourceRoot = path.join(projectRoot, 'src');
const requireComplete = process.argv.includes('--complete');
const failures = [];

function sourceFiles(directory) {
  const files = [];
  for (const entry of fs.readdirSync(directory, { withFileTypes: true })) {
    const absolutePath = path.join(directory, entry.name);
    if (entry.isDirectory()) files.push(...sourceFiles(absolutePath));
    else if (entry.isFile() && entry.name.endsWith('.ts')) {
      files.push(
        path.relative(projectRoot, absolutePath).split(path.sep).join('/'),
      );
    }
  }
  return files;
}

const excludedFiles = excludedMutationFiles.map(({ file }) => file);
const categories = [
  ['mutated', mutatedFiles],
  ['pending', pendingMutationFiles],
  ['excluded', excludedFiles],
];
const owners = new Map();

for (const [category, files] of categories) {
  for (const file of files) {
    const previous = owners.get(file);
    if (previous !== undefined) {
      failures.push(`${file} appears in both ${previous} and ${category}`);
    } else {
      owners.set(file, category);
    }
  }
}

for (const { file, reason } of excludedMutationFiles) {
  if (reason.trim().length < 20) {
    failures.push(`${file} needs a specific exclusion reason`);
  }
}

const actualFiles = sourceFiles(sourceRoot).sort();
const actualFileSet = new Set(actualFiles);
for (const file of actualFiles) {
  if (!owners.has(file)) failures.push(`Unaccounted source file: ${file}`);
}
for (const file of owners.keys()) {
  if (!actualFileSet.has(file))
    failures.push(`Scope references missing file: ${file}`);
}

if (requireComplete && pendingMutationFiles.length > 0) {
  failures.push(
    `Mutation rollout is incomplete: ${pendingMutationFiles.length} pending file(s)`,
  );
}

if (failures.length > 0) {
  throw new Error(`Mutation scope audit failed:\n- ${failures.join('\n- ')}`);
}

console.log(
  `Mutation scope accounted for ${actualFiles.length} source files: ${mutatedFiles.length} mutated, ${pendingMutationFiles.length} pending, ${excludedFiles.length} justified exclusions.`,
);
