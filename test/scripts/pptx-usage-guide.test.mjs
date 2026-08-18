import { readFile } from 'node:fs/promises';
import path from 'node:path';

import { describe, expect, it } from 'vitest';

import * as rootApi from '../../src/index.ts';
import * as nodeApi from '../../src/formats/pptx/node.ts';

const guidePath = path.resolve('docs', 'pptx-usage-guide.md');

describe('PowerPoint usage guide', () => {
  it('documents exported end-to-end APIs and honest capability boundaries', async () => {
    const [guide, metadata, readme] = await Promise.all([
      readFile(guidePath, 'utf8'),
      readFile(path.resolve('package.json'), 'utf8').then(JSON.parse),
      readFile(path.resolve('README.md'), 'utf8'),
    ]);

    for (const name of [
      'createPptx',
      'parsePptx',
      'parsePptxRoundTripJson',
      'parsePptxWithDiagnostics',
      'readPptxRoundTrip',
      'renderPptxDocumentToSvg',
      'renderPptxToSvg',
      'replacePptxRoundTripText',
      'serializePptxRoundTripJson',
      'setPptxRoundTripTextTransform',
      'setPptxRoundTripShapeTransform',
      'validatePptxScene',
      'writePptxRoundTrip',
    ]) {
      expect(rootApi[name], name).toBeTypeOf('function');
      expect(guide, name).toContain(name);
    }
    for (const name of ['renderPptxDocumentToPng', 'renderPptxToPng']) {
      expect(nodeApi[name], name).toBeTypeOf('function');
      expect(guide, name).toContain(name);
    }

    expect(guide).toContain(`OAKit \`${metadata.version}\``);
    expect(guide).toMatch(
      /\|\s*Arbitrary PPTX creation\/editing\s*\|\s*Not claimed\s*\|/,
    );
    expect(guide).toMatch(
      /\|\s*Pixel-identical rendering\s*\|\s*Not claimed\s*\|/,
    );
    expect(guide).toContain('Never edit `snapshot.document`');
    expect(guide).toContain('error.diagnostic.code');
    expect(guide).not.toContain('editableScene');
    expect(guide).toContain('oakit --version');
    expect(readme).toContain(
      '[PowerPoint usage guide](docs/pptx-usage-guide.md)',
    );
    expect((guide.match(/^```/gm)?.length ?? 0) % 2).toBe(0);
  });

  it('keeps every relative guide link resolvable', async () => {
    const guide = await readFile(guidePath, 'utf8');
    const links = [...guide.matchAll(/\]\(([^)]+)\)/g)]
      .map((match) => match[1])
      .filter(
        (target) =>
          !target.startsWith('#') &&
          !target.startsWith('http://') &&
          !target.startsWith('https://'),
      );

    expect(links.length).toBeGreaterThan(0);
    for (const target of links) {
      await expect(
        readFile(path.resolve(path.dirname(guidePath), target)),
        target,
      ).resolves.toBeDefined();
    }
  });
});
