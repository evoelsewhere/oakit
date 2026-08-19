import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';

import { decodeBase64 } from '../../src/common/binary/base64';
import {
  createPptx,
  parsePptx,
  renderPptxToSvg,
  type PptxSceneDocument,
  type PptxSceneGroupElement,
} from '../../src';

const PNG_BYTES = decodeBase64(
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=',
);

function nestedGroup(): PptxSceneGroupElement {
  return {
    authored: {
      transform: {
        childSpace: { height: 80, width: 80, x: 0, y: 0 },
        height: 80,
        width: 80,
        x: 280,
        y: 60,
      },
    },
    elements: [
      {
        authored: {
          fillColor: '#22C55E',
          geometry: 'ellipse',
          transform: { height: 40, width: 40, x: 20, y: 20 },
        },
        key: 'nested-shape',
        resolved: { hidden: false },
        type: 'shape',
      },
    ],
    key: 'nested-group',
    name: 'Nested group',
    resolved: { hidden: false },
    type: 'group',
  };
}

function scene(): PptxSceneDocument {
  return {
    layouts: [],
    masters: [],
    media: [{ data: PNG_BYTES, key: 'group-media', mimeType: 'image/png' }],
    schemaVersion: 2,
    size: { height: 540, width: 960 },
    slides: [
      {
        elements: [
          {
            authored: {
              transform: {
                childSpace: { height: 200, width: 400, x: 0, y: 0 },
                height: 200,
                width: 400,
                x: 100,
                y: 120,
              },
            },
            elements: [
              {
                authored: {
                  fillColor: '#F97316',
                  geometry: 'roundRect',
                  transform: { height: 60, width: 100, x: 20, y: 30 },
                },
                key: 'group-shape',
                resolved: { hidden: false },
                type: 'shape',
              },
              {
                authored: {
                  transform: { height: 50, width: 120, x: 140, y: 30 },
                },
                key: 'group-text',
                resolved: { hidden: false },
                text: {
                  body: { anchor: 'center' },
                  paragraphs: [
                    {
                      children: [
                        { key: 'group-run', text: 'Grouped', type: 'run' },
                      ],
                      key: 'group-paragraph',
                    },
                  ],
                },
                type: 'text',
              },
              {
                authored: {
                  transform: { height: 60, width: 60, x: 20, y: 110 },
                },
                key: 'group-image',
                mediaKey: 'group-media',
                resolved: { hidden: false },
                type: 'image',
              },
              nestedGroup(),
            ],
            key: 'native-group',
            name: 'Native group',
            resolved: { hidden: false },
            type: 'group',
          },
        ],
        key: 'slide-1',
      },
    ],
    themes: [],
  };
}

describe('native PowerPoint group creation', () => {
  it('creates, strict-parses, and Office-free renders nested native content', async () => {
    const input = scene();
    const before = structuredClone(input);
    const created = await createPptx(input);
    const [parsed, rendered, archive] = await Promise.all([
      parsePptx(created.data, { errorMode: 'strict', imageMode: 'base64' }),
      renderPptxToSvg(created.data, { slideNumbers: [1] }),
      JSZip.loadAsync(created.data),
    ]);
    const group = parsed.slides[0]?.elements[0];

    expect(input).toEqual(before);
    expect(group).toMatchObject({
      childSpace: { height: 200, width: 400, x: 0, y: 0 },
      elements: [
        { id: '3', shapType: 'roundRect', type: 'shape' },
        { id: '4', type: 'text' },
        { id: '5', type: 'image' },
        {
          childSpace: { height: 80, width: 80, x: 0, y: 0 },
          elements: [{ id: '7', shapType: 'ellipse', type: 'shape' }],
          id: '6',
          type: 'group',
        },
      ],
      height: 200,
      id: '2',
      left: 100,
      top: 120,
      type: 'group',
      width: 400,
    });
    expect(archive.file('ppt/media/image1.png')).not.toBeNull();
    const svg = new TextDecoder().decode(rendered.slides[0]?.data);
    expect(svg).toContain('Grouped');
    expect(svg).toContain('#F97316');
    expect(svg).toContain('data:image/png;base64,');
    expect(created.report.supportProfile.id).toBe('pptx-create-native-v1');
  });

  it('is deterministic across concurrent group writes', async () => {
    const [first, second] = await Promise.all([
      createPptx(scene()),
      createPptx(scene()),
    ]);

    expect(second.data).toEqual(first.data);
    expect(second.report).toEqual(first.report);
  });
});
