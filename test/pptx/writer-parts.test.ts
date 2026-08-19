import { posix } from 'node:path';

import { SaxesParser } from 'saxes';
import { describe, expect, it } from 'vitest';

import type {
  PptxSceneChartElement,
  PptxSceneDocument,
  PptxSceneMedia,
  PptxSceneSlide,
} from '../../src/formats/pptx/scene-types';
import {
  type PptxSerializedPart,
  serializePowerPointParts,
} from '../../src/formats/pptx/writer/parts';

function fieldSlide(key: string, fieldType: string): PptxSceneSlide {
  return {
    elements: [
      {
        authored: {
          transform: { height: 40, width: 160, x: 10, y: 20 },
        },
        key: `${key}-text`,
        resolved: { hidden: false },
        text: {
          body: {},
          paragraphs: [
            {
              children: [
                {
                  fieldType,
                  key: `${key}-field`,
                  text: key,
                  type: 'field',
                },
              ],
              key: `${key}-paragraph`,
            },
          ],
        },
        type: 'text',
      },
    ],
    key,
  };
}

function chartElement(key: string): PptxSceneChartElement {
  return {
    authored: {
      transform: { height: 200, width: 400, x: 20, y: 30 },
    },
    chartType: 'barChart',
    key,
    resolved: { hidden: false },
    series: [
      {
        categories: ['A', 'B'],
        key: `${key}-series`,
        name: 'Series',
        values: [1, 2],
      },
    ],
    type: 'chart',
  };
}

function scene(
  slides: PptxSceneSlide[] = [],
  media: PptxSceneMedia[] = [],
): PptxSceneDocument {
  return {
    layouts: [],
    masters: [],
    media,
    schemaVersion: 2,
    size: { height: 540, width: 960 },
    slides,
    themes: [],
  };
}

function partByPath(
  parts: readonly PptxSerializedPart[],
  path: string,
): PptxSerializedPart {
  const part = parts.find((candidate) => candidate.path === path);
  if (!part) throw new Error(`Missing test package part: ${path}`);
  return part;
}

function xmlByPath(parts: readonly PptxSerializedPart[], path: string): string {
  const data = partByPath(parts, path).data;
  if (typeof data !== 'string') throw new Error(`Expected XML part: ${path}`);
  return data;
}

function relationshipTargets(xml: string): string[] {
  const targets: string[] = [];
  const parser = new SaxesParser({ xmlns: true });
  parser.on('opentag', (tag) => {
    if (tag.local !== 'Relationship') return;
    const target = tag.attributes.Target;
    if (target) targets.push(target.value);
  });
  parser.write(xml).close();
  return targets;
}

function relationshipOwner(path: string): string {
  if (path === '_rels/.rels') return '';
  const match = /^(.*)\/_rels\/([^/]+)\.rels$/.exec(path);
  if (!match?.[1] || !match[2]) {
    throw new Error(`Invalid test relationship path: ${path}`);
  }
  return `${match[1]}/${match[2]}`;
}

describe('PowerPoint package part serialization', () => {
  it('emits the exact zero-slide OPC inventory in canonical order', () => {
    expect(serializePowerPointParts(scene()).map((part) => part.path)).toEqual([
      '[Content_Types].xml',
      '_rels/.rels',
      'ppt/presentation.xml',
      'ppt/_rels/presentation.xml.rels',
      'ppt/slideMasters/slideMaster1.xml',
      'ppt/slideMasters/_rels/slideMaster1.xml.rels',
      'ppt/slideLayouts/slideLayout1.xml',
      'ppt/slideLayouts/_rels/slideLayout1.xml.rels',
      'ppt/theme/theme1.xml',
    ]);
  });

  it('appends each slide beside its owner relationship part', () => {
    const parts = serializePowerPointParts(
      scene([fieldSlide('slide-1', 'slidenum'), fieldSlide('slide-2', 'date')]),
    );

    expect(parts.map((part) => part.path).slice(9)).toEqual([
      'ppt/slides/slide1.xml',
      'ppt/slides/_rels/slide1.xml.rels',
      'ppt/slides/slide2.xml',
      'ppt/slides/_rels/slide2.xml.rels',
    ]);
    expect(xmlByPath(parts, '[Content_Types].xml')).toContain(
      'PartName="/ppt/slides/slide2.xml"',
    );
    expect(xmlByPath(parts, 'ppt/presentation.xml')).toContain(
      '<p:sldId id="257" r:id="rId3"/>',
    );
  });

  it('allocates field identities once across the complete package', () => {
    const parts = serializePowerPointParts(
      scene([fieldSlide('first', 'slidenum'), fieldSlide('second', 'date')]),
    );
    const first = xmlByPath(parts, 'ppt/slides/slide1.xml');
    const second = xmlByPath(parts, 'ppt/slides/slide2.xml');

    expect(first).toContain('id="{00000000-0000-0000-0000-000000000001}"');
    expect(second).toContain('id="{00000000-0000-0000-0000-000000000002}"');
  });

  it('produces deterministic XML without sharing allocator state', () => {
    const input = scene([fieldSlide('slide-1', 'slidenum')]);

    expect(serializePowerPointParts(input)).toEqual(
      serializePowerPointParts(input),
    );
  });

  it('emits only standalone well-formed XML parts', () => {
    const parts = serializePowerPointParts(
      scene([fieldSlide('slide-1', 'slidenum')]),
    );

    for (const part of parts) {
      const parser = new SaxesParser({ xmlns: true });
      if (typeof part.data !== 'string') continue;
      expect(() => parser.write(part.data).close(), part.path).not.toThrow();
    }
  });

  it('resolves every internal relationship target to an emitted part', () => {
    const parts = serializePowerPointParts(
      scene([fieldSlide('slide-1', 'slidenum')]),
    );
    const paths = new Set(parts.map((part) => part.path));
    const resolved: string[] = [];

    for (const part of parts.filter((candidate) =>
      candidate.path.endsWith('.rels'),
    )) {
      const owner = relationshipOwner(part.path);
      if (typeof part.data !== 'string') continue;
      for (const target of relationshipTargets(part.data)) {
        const path = posix.normalize(posix.join(posix.dirname(owner), target));
        resolved.push(path);
        expect(paths.has(path), `${part.path} -> ${path}`).toBe(true);
      }
    }
    expect(resolved).toEqual([
      'ppt/presentation.xml',
      'ppt/slideMasters/slideMaster1.xml',
      'ppt/slides/slide1.xml',
      'ppt/slideLayouts/slideLayout1.xml',
      'ppt/theme/theme1.xml',
      'ppt/slideMasters/slideMaster1.xml',
      'ppt/slideLayouts/slideLayout1.xml',
    ]);
  });

  it('owns native media and binds image elements to binary package parts', () => {
    const callerBytes = new Uint8Array([0x89, 0x50, 0x4e, 0x47]);
    const input = scene(
      [
        {
          elements: [
            {
              authored: {
                transform: { height: 40, width: 50, x: 10, y: 20 },
              },
              key: 'picture',
              mediaKey: 'media',
              resolved: { hidden: false },
              type: 'image',
            },
          ],
          key: 'slide',
        },
      ],
      [{ data: callerBytes, key: 'media', mimeType: 'image/png' }],
    );

    const parts = serializePowerPointParts(input);
    callerBytes.fill(0);

    expect(parts.map(({ path }) => path).slice(9)).toEqual([
      'ppt/media/image1.png',
      'ppt/slides/slide1.xml',
      'ppt/slides/_rels/slide1.xml.rels',
    ]);
    expect(partByPath(parts, 'ppt/media/image1.png').data).toEqual(
      new Uint8Array([0x89, 0x50, 0x4e, 0x47]),
    );
    expect(xmlByPath(parts, 'ppt/slides/slide1.xml')).toContain(
      '<a:blip r:embed="rId2"/>',
    );
    expect(xmlByPath(parts, 'ppt/slides/_rels/slide1.xml.rels')).toContain(
      'Target="../media/image1.png"',
    );
  });

  it('owns native chart parts and binds them after image relationships', () => {
    const input = scene(
      [
        {
          elements: [
            {
              authored: {
                transform: { height: 40, width: 50, x: 10, y: 20 },
              },
              key: 'picture',
              mediaKey: 'media',
              resolved: { hidden: false },
              type: 'image',
            },
            chartElement('chart'),
          ],
          key: 'slide',
        },
      ],
      [
        {
          data: new Uint8Array([0x89, 0x50, 0x4e, 0x47]),
          key: 'media',
          mimeType: 'image/png',
        },
      ],
    );

    const parts = serializePowerPointParts(input);

    expect(parts.map(({ path }) => path).slice(9)).toEqual([
      'ppt/media/image1.png',
      'ppt/charts/chart1.xml',
      'ppt/slides/slide1.xml',
      'ppt/slides/_rels/slide1.xml.rels',
    ]);
    expect(xmlByPath(parts, '[Content_Types].xml')).toContain(
      'PartName="/ppt/charts/chart1.xml"',
    );
    expect(xmlByPath(parts, 'ppt/slides/slide1.xml')).toContain('r:id="rId3"');
    const relationships = xmlByPath(parts, 'ppt/slides/_rels/slide1.xml.rels');
    expect(relationships).toContain('Target="../media/image1.png"');
    expect(relationships).toContain('Target="../charts/chart1.xml"');
    expect(xmlByPath(parts, 'ppt/charts/chart1.xml')).toContain('<c:barChart>');
  });

  it('rejects an image reference missing from the media inventory', () => {
    const input = scene([
      {
        elements: [
          {
            authored: {
              transform: { height: 40, width: 50, x: 10, y: 20 },
            },
            key: 'picture',
            mediaKey: 'missing',
            resolved: { hidden: false },
            type: 'image',
          },
        ],
        key: 'slide',
      },
    ]);

    expect(() => serializePowerPointParts(input)).toThrow(
      'PowerPoint image element picture references missing media missing',
    );
  });

  it('allocates JPEG extensions and rejects an absent image media key', () => {
    const jpeg = scene(
      [],
      [
        {
          data: new Uint8Array([0xff, 0xd8, 0xff, 0xd9]),
          key: 'jpeg',
          mimeType: 'image/jpeg',
        },
      ],
    );
    expect(serializePowerPointParts(jpeg).map(({ path }) => path)).toContain(
      'ppt/media/image1.jpeg',
    );

    const missingKey = scene([
      {
        elements: [
          {
            authored: {
              transform: { height: 40, width: 50, x: 10, y: 20 },
            },
            key: 'picture',
            resolved: { hidden: false },
            type: 'image',
          },
        ],
        key: 'slide',
      },
    ]);
    expect(() => serializePowerPointParts(missingKey)).toThrow(
      'PowerPoint image element picture has no media key',
    );
  });

  it('rejects an oversized internal scene before allocating package parts', () => {
    const input = scene();
    input.slides = new Array<PptxSceneSlide>(10_001);

    expect(() => serializePowerPointParts(input)).toThrow(
      new RangeError(
        'PowerPoint presentation slide count must be an integer from 0 through 10000',
      ),
    );
  });
});
