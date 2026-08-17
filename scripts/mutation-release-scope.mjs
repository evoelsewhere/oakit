import { mutationModules } from './mutation-modules.mjs';
import { mutatedFiles } from './mutation-scope.mjs';
import { shapePathMutationFile } from './shape-path-mutation-jobs.mjs';

export const moduleMutationFiles = mutationModules.map(
  (module) => module.source,
);

const separatelyMutatedFiles = new Set([
  shapePathMutationFile,
  ...moduleMutationFiles,
]);

export const fileMutationShardFiles = mutatedFiles.filter(
  (file) => !separatelyMutatedFiles.has(file),
);
