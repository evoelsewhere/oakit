import { readFileSync } from 'node:fs';

import { defineConfig } from 'tsup';

function packageVersion(): string {
  const metadata: unknown = JSON.parse(
    readFileSync(new URL('./package.json', import.meta.url), 'utf8'),
  );
  if (
    typeof metadata !== 'object' ||
    metadata === null ||
    !('version' in metadata) ||
    typeof metadata.version !== 'string'
  ) {
    throw new TypeError('package.json must contain a string version');
  }
  return metadata.version;
}

export default defineConfig({
  entry: {
    cli: 'src/cli.ts',
    index: 'src/index.ts',
    'pptx/index': 'src/formats/pptx/index.ts',
    'pptx/node': 'src/formats/pptx/node.ts',
  },
  format: ['esm', 'cjs'],
  target: 'es2022',
  dts: true,
  sourcemap: true,
  clean: true,
  treeshake: true,
  define: {
    __OAKIT_VERSION__: JSON.stringify(packageVersion()),
  },
});
