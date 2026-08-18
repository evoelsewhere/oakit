import { describe, expect, it } from 'vitest';

import type { PptxSceneTransform } from '../../src/formats/pptx/scene-types';
import { createPptxRoundTripTablePreview } from '../../src/formats/pptx/roundtrip/table-preview';
import type { Table } from '../../src/formats/pptx/types';

function table(): Table {
  return {
    borders: {},
    colWidths: [100, 200],
    data: [
      [
        {
          borders: {
            bottom: {
              borderColor: '#0F172A',
              borderType: 'solid',
              borderWidth: 2,
            },
            left: {
              borderColor: '#F97316',
              borderType: 'dotted',
              borderWidth: 3,
            },
            right: {
              borderColor: '#22C55E',
              borderType: 'solid',
              borderWidth: 4,
            },
            top: {
              borderColor: '#334155',
              borderType: 'dashed',
              borderWidth: 1,
            },
          },
          colSpan: 2,
          fillColor: '#E0F2FE',
          fontBold: true,
          fontColor: '#0F172A',
          rowSpan: 2,
          text: 'Alpha',
          vAlign: 'mid',
        },
        { borders: {}, hMerge: 1, text: 'Beta', vAlign: 'down' },
      ],
      [
        {
          borders: {
            top: {
              borderColor: '#000000',
              borderType: 'solid',
              borderWidth: 0,
            },
          },
          text: 'Gamma',
          vAlign: 'up',
          vMerge: 1,
        },
        {
          borders: {},
          fontBold: false,
          hMerge: 1,
          text: 'Delta',
          vAlign: 'mid',
          vMerge: 1,
        },
      ],
    ],
    height: 100,
    id: '2',
    isFlipH: false,
    isFlipV: true,
    left: 10,
    name: 'Native table',
    order: 0,
    rotate: 5,
    rowHeights: [40, 60],
    top: 20,
    type: 'table',
    width: 300,
  };
}

function transform(
  overrides: Partial<PptxSceneTransform> = {},
): PptxSceneTransform {
  return {
    flipHorizontal: false,
    flipVertical: true,
    height: 100,
    rotation: 5,
    width: 300,
    x: 10,
    y: 20,
    ...overrides,
  };
}

describe('PowerPoint native table round-trip preview', () => {
  it('maps every supported grid, cell, border, text, and transform field', () => {
    const result = createPptxRoundTripTablePreview(
      table(),
      0,
      0,
      (value) => `plain:${value}`,
      () => transform(),
    );

    expect(result).toEqual({
      authored: {},
      columns: [100, 200],
      key: 'slide-1-element-1',
      name: 'Native table',
      resolved: { hidden: false, transform: transform() },
      rows: [
        {
          cells: [
            {
              borders: {
                bottom: { color: '#0F172A', style: 'solid', width: 2 },
                left: { color: '#F97316', style: 'dotted', width: 3 },
                right: { color: '#22C55E', style: 'solid', width: 4 },
                top: { color: '#334155', style: 'dashed', width: 1 },
              },
              colSpan: 2,
              fillColor: '#E0F2FE',
              rowSpan: 2,
              text: {
                body: { anchor: 'center' },
                paragraphs: [
                  {
                    children: [
                      {
                        key: 'slide-1-element-1-row-1-cell-1-run-1',
                        properties: { bold: true, color: '#0F172A' },
                        text: 'plain:Alpha',
                        type: 'run',
                      },
                    ],
                    key: 'slide-1-element-1-row-1-cell-1-paragraph-1',
                  },
                ],
              },
            },
            {
              borders: {},
              hMerge: true,
              text: {
                body: { anchor: 'bottom' },
                paragraphs: [
                  {
                    children: [
                      {
                        key: 'slide-1-element-1-row-1-cell-2-run-1',
                        text: 'plain:Beta',
                        type: 'run',
                      },
                    ],
                    key: 'slide-1-element-1-row-1-cell-2-paragraph-1',
                  },
                ],
              },
            },
          ],
          height: 40,
        },
        {
          cells: [
            {
              borders: {},
              text: {
                body: { anchor: 'top' },
                paragraphs: [
                  {
                    children: [
                      {
                        key: 'slide-1-element-1-row-2-cell-1-run-1',
                        text: 'plain:Gamma',
                        type: 'run',
                      },
                    ],
                    key: 'slide-1-element-1-row-2-cell-1-paragraph-1',
                  },
                ],
              },
              vMerge: true,
            },
            {
              borders: {},
              hMerge: true,
              text: {
                body: { anchor: 'center' },
                paragraphs: [
                  {
                    children: [
                      {
                        key: 'slide-1-element-1-row-2-cell-2-run-1',
                        properties: { bold: false },
                        text: 'plain:Delta',
                        type: 'run',
                      },
                    ],
                    key: 'slide-1-element-1-row-2-cell-2-paragraph-1',
                  },
                ],
              },
              vMerge: true,
            },
          ],
          height: 60,
        },
      ],
      type: 'table',
    });
    const plainCell = result?.rows[0]?.cells[1];
    expect(plainCell).not.toHaveProperty('colSpan');
    expect(plainCell).not.toHaveProperty('fillColor');
    expect(plainCell).not.toHaveProperty('rowSpan');
  });

  it('omits an absent optional name without changing native eligibility', () => {
    const value = table();
    delete value.name;
    const result = createPptxRoundTripTablePreview(
      value,
      1,
      2,
      (text) => text,
      () => transform(),
    );

    expect(result).toMatchObject({
      key: 'slide-2-element-3',
      type: 'table',
    });
    expect(result).not.toHaveProperty('name');
  });

  it.each([
    [
      'missing transform',
      (value: Table) => value,
      (): PptxSceneTransform | undefined => undefined,
    ],
    [
      'empty columns',
      (value: Table) => {
        value.colWidths = [];
        return value;
      },
      (): PptxSceneTransform => transform(),
    ],
    [
      'empty data',
      (value: Table) => {
        value.data = [];
        return value;
      },
      (): PptxSceneTransform => transform(),
    ],
    [
      'non-positive column',
      (value: Table) => {
        value.colWidths = [0, 300];
        return value;
      },
      (): PptxSceneTransform => transform(),
    ],
    [
      'non-positive row',
      (value: Table) => {
        value.rowHeights = [0, 100];
        return value;
      },
      (): PptxSceneTransform => transform(),
    ],
    [
      'row size count mismatch',
      (value: Table) => {
        value.rowHeights = [100];
        return value;
      },
      (): PptxSceneTransform => transform(),
    ],
    [
      'cell count mismatch',
      (value: Table) => {
        value.data[1] = [value.data[1]?.[0] as Table['data'][number][number]];
        return value;
      },
      (): PptxSceneTransform => transform(),
    ],
    [
      'column total mismatch',
      (value: Table) => value,
      (): PptxSceneTransform => transform({ width: 301 }),
    ],
    [
      'row total mismatch',
      (value: Table) => value,
      (): PptxSceneTransform => transform({ height: 101 }),
    ],
  ] as const)(
    'rejects %s as a native table target',
    (_name, mutate, resolver) => {
      expect(
        createPptxRoundTripTablePreview(
          mutate(table()),
          0,
          0,
          (value) => value,
          resolver,
        ),
      ).toBeUndefined();
    },
  );
});
