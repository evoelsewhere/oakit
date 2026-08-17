import { describe, expect, it } from 'vitest';

import { createPptx, parsePptx } from '../../src/index.ts';
import { renderPptxToPng } from '../../src/formats/pptx/node.ts';
import { googleTemplateCatalog } from '../../scripts/reliability/pptx-google-template-catalog.mjs';

describe('Google Slides visual template corpus', () => {
  it('contains 30 unique, complex, deterministic template scenes', () => {
    expect(googleTemplateCatalog).toHaveLength(30);
    expect(new Set(googleTemplateCatalog.map(({ slug }) => slug)).size).toBe(
      30,
    );
    expect(
      new Set(googleTemplateCatalog.map(({ marker }) => marker)).size,
    ).toBe(30);
    expect(
      googleTemplateCatalog.every(
        ({ scene }) => (scene.slides[0]?.elements.length ?? 0) >= 12,
      ),
    ).toBe(true);
    expect(
      new Set(
        googleTemplateCatalog.map(
          ({ scene }) => scene.slides[0]?.backgroundColor,
        ),
      ).size,
    ).toBeGreaterThanOrEqual(10);
  });

  it('strict-parses and Office-free renders every authored source deck', async () => {
    for (const template of googleTemplateCatalog) {
      let created;
      try {
        created = await createPptx(template.scene);
      } catch (error) {
        throw new Error(`${template.slug} could not be created`, {
          cause: error,
        });
      }
      const [document, rendered] = await Promise.all([
        parsePptx(created.data, {
          audioMode: 'none',
          errorMode: 'strict',
          imageMode: 'none',
          videoMode: 'none',
        }),
        renderPptxToPng(created.data),
      ]);
      expect(created.report.level, template.slug).toBe('C2');
      expect(document.slides, template.slug).toHaveLength(1);
      expect(JSON.stringify(document), template.slug).toContain(
        template.marker,
      );
      expect(
        document.slides[0]?.elements.length,
        template.slug,
      ).toBeGreaterThanOrEqual(12);
      expect(rendered.slides, template.slug).toHaveLength(1);
      expect(
        rendered.slides[0]?.data.byteLength,
        template.slug,
      ).toBeGreaterThan(1_000);
      expect(
        rendered.slides[0]?.warnings.filter(
          ({ code }) => code !== 'font-substitution',
        ),
        template.slug,
      ).toEqual([]);
    }
  }, 240_000);
});
