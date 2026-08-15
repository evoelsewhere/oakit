import { readFile } from 'node:fs/promises';
import { resolve } from 'node:path';

import { describe, expect, it } from 'vitest';

import { parsePptxWithDiagnostics } from '../../src';

interface ResolvedCorpusEntry {
  id: string;
  minimumElementCounts?: Record<string, number>;
  minimumSlides: number;
  path: string;
  producer: string;
  tier: 'curated' | 'large';
}

interface ElementLike {
  elements?: ElementLike[];
  type: string;
}

interface ResolvedCorpus {
  entries: ResolvedCorpusEntry[];
  version: number;
}

function expectFiniteNumbers(value: unknown): void {
  if (typeof value === 'number') {
    expect(Number.isFinite(value)).toBe(true);
    return;
  }
  if (!value || typeof value !== 'object') return;
  for (const child of Object.values(value)) expectFiniteNumbers(child);
}

function countElementTypes(
  elements: ElementLike[],
  counts: Record<string, number> = {},
): Record<string, number> {
  for (const element of elements) {
    counts[element.type] = (counts[element.type] ?? 0) + 1;
    if (element.elements) countElementTypes(element.elements, counts);
  }
  return counts;
}

const resolvedManifest = JSON.parse(
  await readFile(resolve('.cache/pptx-corpus/resolved.json'), 'utf8'),
) as ResolvedCorpus;

describe(`real-world PPTX corpus v${resolvedManifest.version}`, () => {
  for (const entry of resolvedManifest.entries) {
    it(`${entry.producer}: ${entry.id}`, async () => {
      const bytes = new Uint8Array(await readFile(entry.path));
      const result = await parsePptxWithDiagnostics(bytes, {
        imageMode: 'none',
        limits: { maxInputBytes: 150 * 1024 * 1024 },
      });

      expect(result.document.slides.length).toBeGreaterThanOrEqual(
        entry.minimumSlides,
      );
      expect(result.document.size.width).toBeGreaterThan(0);
      expect(result.document.size.height).toBeGreaterThan(0);
      expectFiniteNumbers(result.document);
      const elementCounts: Record<string, number> = {};
      for (const slide of result.document.slides) {
        countElementTypes(slide.elements, elementCounts);
        countElementTypes(slide.layoutElements, elementCounts);
      }
      for (const [type, minimum] of Object.entries(
        entry.minimumElementCounts ?? {},
      )) {
        expect(
          elementCounts[type] ?? 0,
          `${entry.id}: ${type}`,
        ).toBeGreaterThanOrEqual(minimum);
      }
      expect(
        result.diagnostics.some((diagnostic) =>
          /^(?:invalid-package|invalid-document-structure|missing-required-part)$/.test(
            diagnostic.code,
          ),
        ),
      ).toBe(false);
    });
  }
});
