import { readFile } from 'node:fs/promises';
import path from 'node:path';

import { Resvg } from '@resvg/resvg-js';
import { describe, expect, it } from 'vitest';

import { slidesManiaCorpus } from '../../scripts/reliability/slidesmania-corpus.mjs';

const root = path.resolve('docs', 'evidence', '0.0.1', 'slidesmania');

describe('tracked SlidesMania producer evidence', () => {
  it('retains complete aggregate evidence without redistributing templates', async () => {
    const source = await readFile(path.join(root, 'evidence.json'), 'utf8');
    const evidence = JSON.parse(source);
    const progress = JSON.parse(
      await readFile(path.join(root, 'progress.json'), 'utf8'),
    );

    expect(progress).toEqual({
      completedTemplates: 30,
      currentTemplate: null,
      phase: 'complete',
      schemaVersion: 1,
      templateCount: 30,
    });
    expect(evidence.execution).toMatchObject({
      repository: 'evoelsewhere/oakit',
      revision: '3d67c590c41f916b6081e99cd46be40832c4efd5',
      runId: '32036893093',
    });
    expect(evidence.summary).toMatchObject({
      allAttributionPreserved: true,
      allGoogleExportsStrictParsed: true,
      allSlidesRenderedWithoutOffice: true,
      minimumElementRetention: 1,
      minimumTextRetention: 1,
      temporaryPresentationsDeleted: true,
      templateCount: 30,
      totalElements: 9285,
      totalSlides: 733,
    });
    expect(evidence.corpus).toHaveLength(30);
    expect(evidence.corpus.map(({ sourcePage }) => sourcePage)).toEqual(
      slidesManiaCorpus.map(({ sourcePage }) => sourcePage),
    );
    expect(
      evidence.corpus.every(
        ({ source, output }) => source.pptx.sha256 !== output.pptx.sha256,
      ),
    ).toBe(true);
    expect(source).not.toMatch(/\/export\/pptx|\.pptx["']/i);
    expect(source).not.toContain('data:image');
  });

  it('renders the audit graphic and retains a valid tracked PNG', async () => {
    const svg = await readFile(path.join(root, 'producer-audit.svg'));
    const png = await readFile(path.join(root, 'producer-audit.png'));
    const rendered = new Resvg(svg).render().asPng();

    for (const candidate of [rendered, png]) {
      const view = new DataView(
        candidate.buffer,
        candidate.byteOffset,
        candidate.byteLength,
      );
      expect([...candidate.subarray(0, 8)]).toEqual([
        0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a,
      ]);
      expect(view.getUint32(16)).toBe(1600);
      expect(view.getUint32(20)).toBe(900);
      expect(candidate.byteLength).toBeGreaterThan(50_000);
    }
    expect(svg.toString()).toContain('Run 32036893093');
  });

  it('links the evidence, source license, and successful run from the README', async () => {
    const readme = await readFile(path.resolve('README.md'), 'utf8');

    expect(readme).toContain(
      'docs/evidence/0.0.1/slidesmania/producer-audit.png',
    );
    expect(readme).toContain(
      'https://github.com/evoelsewhere/oakit/actions/runs/32036893093',
    );
    expect(readme).toContain(
      'https://slidesmania.com/copyright-and-legal-information/#license',
    );
  });
});
