import { describe, expect, it } from 'vitest';

import {
  createPptx,
  parsePptx,
  renderPptxToSvg,
  type PptxSceneDocument,
  type PptxSceneTableCell,
  type PptxSceneTextBody,
} from '../../src';

function text(key: string, value: string): PptxSceneTextBody {
  return {
    body: { anchor: 'center', wrap: true },
    paragraphs: [
      {
        children: [
          {
            key: `${key}-run`,
            properties: { color: '#0F172A', fontSize: 14 },
            text: value,
            type: 'run',
          },
        ],
        key: `${key}-paragraph`,
      },
    ],
  };
}

function cell(
  key: string,
  value: string,
  fillColor: string,
): PptxSceneTableCell {
  return {
    borders: {
      bottom: { color: '#334155', width: 1 },
      left: { color: '#334155', width: 1 },
      right: { color: '#334155', width: 1 },
      top: { color: '#334155', width: 1 },
    },
    fillColor,
    text: text(key, value),
  };
}

function scene(): PptxSceneDocument {
  return {
    layouts: [],
    masters: [],
    media: [],
    schemaVersion: 2,
    size: { height: 540, width: 960 },
    slides: [
      {
        elements: [
          {
            authored: {
              transform: { height: 100, width: 300, x: 72, y: 90 },
            },
            columns: [120, 180],
            key: 'sales-table',
            name: 'Quarterly sales',
            resolved: { hidden: false },
            rows: [
              {
                cells: [
                  cell('header-product', 'Product', '#E0F2FE'),
                  cell('header-revenue', 'Revenue', '#E0F2FE'),
                ],
                height: 40,
              },
              {
                cells: [
                  cell('value-product', 'Atlas', '#FFFFFF'),
                  cell('value-revenue', '$125K', '#FFFFFF'),
                ],
                height: 60,
              },
            ],
            type: 'table',
          },
        ],
        key: 'slide-1',
      },
    ],
    themes: [],
  };
}

describe('native PowerPoint table creation', () => {
  it('creates, strict-parses, and Office-free renders a styled native table', async () => {
    const input = scene();
    const before = structuredClone(input);
    const created = await createPptx(input);
    const parsed = await parsePptx(created.data, {
      audioMode: 'none',
      errorMode: 'strict',
      imageMode: 'none',
      videoMode: 'none',
    });
    const generated = parsed.slides[0]?.elements[0];

    expect(input).toEqual(before);
    expect(generated).toMatchObject({
      colWidths: [120, 180],
      height: 100,
      left: 72,
      name: 'Quarterly sales',
      rowHeights: [40, 60],
      top: 90,
      type: 'table',
      width: 300,
    });
    if (generated?.type !== 'table') throw new Error('Expected table');
    expect(generated.data.map((row) => row.map((value) => value.text))).toEqual(
      [
        [
          expect.stringContaining('Product'),
          expect.stringContaining('Revenue'),
        ],
        [expect.stringContaining('Atlas'), expect.stringContaining('$125K')],
      ],
    );
    expect(generated.data[0]?.[0]).toMatchObject({
      fillColor: '#E0F2FE',
      vAlign: 'mid',
    });
    expect(generated.data[1]?.[1]?.borders.right).toMatchObject({
      borderColor: '#334155',
      borderType: 'solid',
      borderWidth: 1,
    });

    const rendered = await renderPptxToSvg(created.data, {
      slideNumbers: [1],
    });
    const svg = new TextDecoder().decode(rendered.slides[0]?.data);
    expect(svg).toContain('Product');
    expect(svg).toContain('Revenue');
    expect(svg).toContain('fill="#E0F2FE"');
    expect(svg).not.toContain('<foreignObject');
    expect(created.report.supportProfile.id).toBe('pptx-create-native-v1');
  });

  it('is deterministic across independent native table writes', async () => {
    const [first, second] = await Promise.all([
      createPptx(scene()),
      createPptx(scene()),
    ]);

    expect(second.data).toEqual(first.data);
    expect(second.report).toEqual(first.report);
  });
});
