import { describe, expect, it } from 'vitest';

import { renderPptxSvgTable } from '../../src/formats/pptx/render-svg-table';
import type { Table } from '../../src/formats/pptx/types';

const BOX = { height: 40, left: 0, top: 0, width: 100 };

function table(overrides: Partial<Table> = {}): Table {
  return {
    borders: {},
    colWidths: [25, 75],
    data: [
      [
        { borders: {}, text: 'A', vAlign: 'up' },
        { borders: {}, text: 'B', vAlign: 'up' },
      ],
      [
        { borders: {}, text: 'C', vAlign: 'up' },
        { borders: {}, text: 'D', vAlign: 'up' },
      ],
    ],
    height: 40,
    id: 'table',
    left: 0,
    order: 0,
    rowHeights: [10, 30],
    top: 0,
    type: 'table',
    width: 100,
    ...overrides,
  };
}

describe('PowerPoint SVG tables', () => {
  it('uses authored column widths, row heights, fills, and portable text', () => {
    const input = table();
    const cell = input.data[1]?.[1];
    if (cell === undefined) throw new Error('Expected table cell');
    cell.fillColor = '#ABCDEF';
    cell.fontColor = '#123456';
    cell.fontBold = true;
    cell.text = '<b>D &amp; E</b>';

    expect(renderPptxSvgTable(input, BOX)).toBe(
      '<g transform="translate(0 0)"><rect x="0" y="0" width="25" height="10" fill="#ffffff" stroke="none"/><svg x="0" y="0" width="25" height="10" overflow="hidden"><text x="4" y="13" font-family="sans-serif" font-size="11" fill="#111827">A</text></svg></g><g transform="translate(25 0)"><rect x="0" y="0" width="75" height="10" fill="#ffffff" stroke="none"/><svg x="0" y="0" width="75" height="10" overflow="hidden"><text x="4" y="13" font-family="sans-serif" font-size="11" fill="#111827">B</text></svg></g><g transform="translate(0 10)"><rect x="0" y="0" width="25" height="30" fill="#ffffff" stroke="none"/><svg x="0" y="0" width="25" height="30" overflow="hidden"><text x="4" y="13" font-family="sans-serif" font-size="11" fill="#111827">C</text></svg></g><g transform="translate(25 10)"><rect x="0" y="0" width="75" height="30" fill="#ABCDEF" stroke="none"/><svg x="0" y="0" width="75" height="30" overflow="hidden"><text x="4" y="13" font-family="sans-serif" font-size="11" fill="#123456" font-weight="700">D &amp; E</text></svg></g>',
    );
  });

  it('renders merged origins once and skips continuation cells', () => {
    const input = table({
      data: [
        [
          {
            borders: {},
            colSpan: 2,
            rowSpan: 2,
            text: 'Merged',
            vAlign: 'mid',
          },
          { borders: {}, hMerge: 1, text: '', vAlign: 'mid' },
        ],
        [
          { borders: {}, text: '', vAlign: 'mid', vMerge: 1 },
          {
            borders: {},
            hMerge: 1,
            text: '',
            vAlign: 'mid',
            vMerge: 1,
          },
        ],
      ],
    });

    const result = renderPptxSvgTable(input, BOX);

    expect(result.match(/<g transform=/g)).toHaveLength(1);
    expect(result).toContain('width="100" height="40"');
    expect(result).toContain('<text x="4" y="24.5"');
  });

  it('renders exact per-edge border styles', () => {
    const input = table({
      colWidths: [100],
      data: [
        [
          {
            borders: {
              bottom: {
                borderColor: '#00FF00',
                borderType: 'dotted',
                borderWidth: 3,
              },
              left: {
                borderColor: 'unsafe',
                borderType: 'solid',
                borderWidth: 1,
              },
              right: {
                borderColor: '#0000FF',
                borderType: 'dashed',
                borderWidth: 2,
              },
              top: {
                borderColor: '#FF0000',
                borderType: 'solid',
                borderWidth: 1,
              },
            },
            text: '',
            vAlign: 'up',
          },
        ],
      ],
      rowHeights: [40],
    });

    const result = renderPptxSvgTable(input, BOX);
    expect(result).toContain(
      '<line x1="0" y1="0" x2="100" y2="0" stroke="#FF0000" stroke-width="1"/><line x1="100" y1="0" x2="100" y2="40" stroke="#0000FF" stroke-width="2" stroke-dasharray="4 3"/><line x1="0" y1="40" x2="100" y2="40" stroke="#00FF00" stroke-width="3" stroke-dasharray="1 2"/>',
    );
    expect(result).not.toContain('stroke="null"');
  });

  it.each([
    ['up', 13],
    ['mid', 18],
    ['down', 23],
  ])('positions multiline text at vertical alignment %s', (vAlign, firstY) => {
    const input = table({
      colWidths: [100],
      data: [[{ borders: {}, text: 'First<br/>Second', vAlign }]],
      rowHeights: [40],
    });
    const result = renderPptxSvgTable(input, BOX);

    expect(result).toContain(`<text x="4" y="${firstY}"`);
    expect(result).toContain('>First</text>');
    expect(result).toContain('>Second</text>');
    expect(result).toContain(`<text x="4" y="${firstY + 13}"`);
    expect(result).not.toContain('Stryker was here');
  });

  it('falls back to equal dimensions for malformed authored sizes', () => {
    const input = table({
      colWidths: [0, Number.NaN],
      rowHeights: [0, -1],
    });

    const result = renderPptxSvgTable(input, BOX);

    expect(result).toContain('translate(50 20)');
    expect(result).toContain('width="50" height="20"');
  });

  it.each([
    [[25], [50, 50], 'short dimensions'],
    [[0, 1], [50, 50], 'zero dimension'],
    [[1, '2'], [50, 50], 'non-numeric dimension'],
    [[1, Number.NaN], [50, 50], 'non-finite dimension'],
    [[25, 75, Number.NaN], [25, 75], 'ignored trailing dimension'],
  ])('normalizes %j to %j (%s)', (colWidths, expected, name) => {
    expect(name.length).toBeGreaterThan(0);
    const result = renderPptxSvgTable(
      table({ colWidths: colWidths as number[] }),
      BOX,
    );

    expect(
      [...result.matchAll(/<rect x="0" y="0" width="([\d.]+)"/g)]
        .slice(0, 2)
        .map((match) => Number(match[1])),
    ).toEqual(expected);
  });

  it('bounds invalid and oversized spans at the final cell', () => {
    const input = table();
    const topLeft = input.data[0]?.[0];
    const bottomRight = input.data[1]?.[1];
    if (topLeft === undefined || bottomRight === undefined) {
      throw new Error('Expected table cells');
    }
    topLeft.colSpan = 0;
    topLeft.rowSpan = 0;
    bottomRight.colSpan = 99;
    bottomRight.rowSpan = 99;

    const result = renderPptxSvgTable(input, BOX);

    expect(result).toContain(
      '<g transform="translate(0 0)"><rect x="0" y="0" width="25" height="10"',
    );
    expect(result).toContain(
      '<g transform="translate(25 10)"><rect x="0" y="0" width="75" height="30"',
    );
    expect(result).not.toContain('NaN');
  });

  it('omits malformed borders and renders every valid edge endpoint', () => {
    const input = table({ colWidths: [100], rowHeights: [40] });
    input.data = [
      [
        {
          borders: {
            bottom: {
              borderColor: '#00FF00',
              borderType: 'solid',
              borderWidth: 0,
            },
            left: {
              borderColor: '#FF00FF',
              borderType: 'solid',
              borderWidth: 4,
            },
            right: {
              borderColor: '#0000FF',
              borderType: 'solid',
              borderWidth: Number.NaN,
            },
            top: {
              borderColor: '#FF0000',
              borderType: 'solid',
              borderWidth: -1,
            },
          },
          text: '',
          vAlign: 'up',
        },
      ],
    ];

    const result = renderPptxSvgTable(input, BOX);

    expect(result).toContain(
      '<line x1="0" y1="0" x2="0" y2="40" stroke="#FF00FF" stroke-width="4"/>',
    );
    expect(result.match(/<line /g)).toHaveLength(1);
    expect(result).not.toContain('Stryker was here');
  });

  it('omits empty and non-string cell text without extra markup', () => {
    const input = table({
      colWidths: [50, 50],
      data: [
        [
          { borders: null, text: '', vAlign: 'up' },
          { borders: [], text: 7, vAlign: 'up' },
        ],
      ] as unknown as Table['data'],
      rowHeights: [40],
    });

    const result = renderPptxSvgTable(input, BOX);

    expect(result).not.toContain('<text');
    expect(result).not.toContain('<line');
    expect(result).not.toContain('Stryker was here');
  });

  it('normalizes malformed cell containers without interpreting array fields', () => {
    const input = table({
      colWidths: [50, 50],
      data: [[null, []]] as unknown as Table['data'],
      rowHeights: [40],
    });

    const result = renderPptxSvgTable(input, BOX);

    expect(result.match(/<g transform=/g)).toHaveLength(2);
    expect(result).not.toContain('<text');
    expect(result).not.toContain('Stryker was here');
  });

  it('skips malformed rows and merge continuations without separators', () => {
    const input = table({
      data: [
        null,
        [
          { borders: {}, hMerge: 1, text: 'hidden', vAlign: 'up' },
          { borders: {}, text: 'Visible', vAlign: 'up' },
        ],
      ] as unknown as Table['data'],
    });

    const result = renderPptxSvgTable(input, BOX);

    expect(result).toContain('>Visible</text>');
    expect(result).not.toContain('>hidden</text>');
    expect(result).not.toContain('Stryker was here');
  });

  it.each([null, [], [[]]])(
    'renders an empty placeholder for malformed data %#',
    (data) => {
      const input = table({ data: data as Table['data'] });

      expect(renderPptxSvgTable(input, BOX)).toContain('>Empty table</text>');
    },
  );
});
