import { describe, expect, it } from 'vitest';

import { renderPptxSvgRichElement } from '../../src/formats/pptx/render-svg-rich';
import type { Element, Table } from '../../src/formats/pptx/types';

const BOX = { height: 40, left: 0, top: 0, width: 100 };

function render(element: Element) {
  return renderPptxSvgRichElement(element, BOX);
}

describe('PowerPoint rich SVG elements', () => {
  it('renders table cells, portable styles, and escaped text', () => {
    const result = render({
      borders: {},
      colWidths: [50, 50],
      data: [
        [
          {
            borders: {},
            fillColor: '#abcdef',
            fontBold: true,
            fontColor: '#123456',
            text: '<b>A &amp; B</b>',
            vAlign: 'top',
          },
          { borders: {}, text: 'C', vAlign: 'top' },
        ],
        [
          { borders: {}, text: 'D', vAlign: 'top' },
          { borders: {}, text: 'E', vAlign: 'top' },
        ],
      ],
      height: 40,
      id: 'table-1',
      left: 0,
      order: 0,
      rowHeights: [20, 20],
      top: 0,
      type: 'table',
      width: 100,
    });

    expect(result).toEqual({
      body: '<rect x="0" y="0" width="50" height="20" fill="#abcdef" stroke="#9ca3af"/><text x="4" y="14" font-family="sans-serif" font-size="11" fill="#123456" font-weight="700">A &amp; B</text><rect x="50" y="0" width="50" height="20" fill="#ffffff" stroke="#9ca3af"/><text x="54" y="14" font-family="sans-serif" font-size="11" fill="#111827">C</text><rect x="0" y="20" width="50" height="20" fill="#ffffff" stroke="#9ca3af"/><text x="4" y="34" font-family="sans-serif" font-size="11" fill="#111827">D</text><rect x="50" y="20" width="50" height="20" fill="#ffffff" stroke="#9ca3af"/><text x="54" y="34" font-family="sans-serif" font-size="11" fill="#111827">E</text>',
      warningCode: 'approximate-table',
      warningMessage:
        'The preview preserves table text and cells with simplified sizing and styling.',
    });
  });

  it.each([
    { data: [] as Table['data'], name: 'no rows' },
    { data: [[]] as Table['data'], name: 'no columns' },
  ])('renders an empty table placeholder for $name', ({ data }) => {
    const result = render({
      borders: {},
      colWidths: [],
      data,
      height: 40,
      id: 'table-empty',
      left: 0,
      order: 0,
      rowHeights: [],
      top: 0,
      type: 'table',
      width: 100,
    });

    expect(result?.body).toContain('>Empty table</text>');
  });

  it('handles malformed table rows and cells without interpreting values as markup', () => {
    const result = render({
      borders: {},
      colWidths: [],
      data: [
        [null, { text: '', fillColor: 'unsafe', fontColor: 'unsafe' }],
        null,
        [{ text: 'Stryker was here!' }],
      ] as unknown as Table['data'],
      height: 40,
      id: 'table-malformed',
      left: 0,
      order: 0,
      rowHeights: [],
      top: 0,
      type: 'table',
      width: 100,
    });

    expect(result?.body).toBe(
      '<rect x="0" y="0" width="50" height="13.3333" fill="#ffffff" stroke="#9ca3af"/><text x="4" y="14" font-family="sans-serif" font-size="11" fill="#111827"></text><rect x="50" y="0" width="50" height="13.3333" fill="#ffffff" stroke="#9ca3af"/><text x="54" y="14" font-family="sans-serif" font-size="11" fill="#111827"></text><rect x="0" y="26.6667" width="50" height="13.3333" fill="#ffffff" stroke="#9ca3af"/><text x="4" y="40.6667" font-family="sans-serif" font-size="11" fill="#111827">Stryker was here!</text>',
    );
  });

  it('renders an empty table placeholder for a malformed data container', () => {
    const result = render({
      borders: {},
      colWidths: [],
      data: null,
      height: 40,
      id: 'table-invalid',
      left: 0,
      order: 0,
      rowHeights: [],
      top: 0,
      type: 'table',
      width: 100,
    } as unknown as Element);

    expect(result?.body).toBe(
      '<rect x="0" y="0" width="100" height="40" fill="#f3f4f6" stroke="#9ca3af" stroke-dasharray="4 3"/><text x="4" y="16" font-family="sans-serif" font-size="12" fill="#374151">Empty table</text>',
    );
  });

  it('renders common chart values as deterministic bars', () => {
    const result = render({
      chartType: 'barChart',
      colors: ['#ff0000', '#00ff00'],
      data: [
        {
          key: 'Series',
          values: [
            { x: 'A', y: 10 },
            { x: 'B', y: -20 },
          ],
          xlabels: {},
        },
      ],
      height: 40,
      id: 'chart-1',
      left: 0,
      order: 0,
      top: 0,
      type: 'chart',
      width: 100,
    });

    expect(result).toEqual({
      body: '<rect x="0" y="0" width="100" height="40" fill="#ffffff" stroke="#d1d5db"/><rect x="1" y="29" width="48" height="11" fill="#ff0000"/><rect x="51" y="18" width="48" height="22" fill="#00ff00"/><text x="4" y="14" font-family="sans-serif" font-size="11" fill="#374151">barChart</text>',
      warningCode: 'approximate-chart',
      warningMessage:
        'The preview visualizes chart values with simplified portable bars.',
    });
  });

  it('renders finite scatter data and substitutes unsafe chart colors', () => {
    const result = render({
      chartType: 'scatterChart',
      colors: ['url(unsafe)'],
      data: [[Number.NaN, 5]],
      height: 40,
      id: 'scatter-1',
      left: 0,
      order: 0,
      top: 0,
      type: 'chart',
      width: 100,
    });

    expect(result?.body).toContain(
      'x="1" y="18" width="98" height="22" fill="#4f46e5"',
    );
    expect(result?.body).not.toContain('unsafe');
  });

  it('renders a placeholder when chart data has no finite values', () => {
    const result = render({
      chartType: 'scatterChart',
      colors: [],
      data: [[Number.NaN]],
      height: 40,
      id: 'chart-empty',
      left: 0,
      order: 0,
      top: 0,
      type: 'chart',
      width: 100,
    });

    expect(result?.body).toContain('>Chart data unavailable</text>');
  });

  it('ignores malformed common and scatter chart points', () => {
    const result = render({
      chartType: 'barChart',
      colors: [],
      data: [
        null,
        { values: null },
        { values: [null, { y: '5' }, { y: Number.POSITIVE_INFINITY }] },
        ['6', Number.NEGATIVE_INFINITY],
      ],
      height: 40,
      id: 'chart-malformed',
      left: 0,
      order: 0,
      top: 0,
      type: 'chart',
      width: 100,
    } as unknown as Element);

    expect(result?.body).toBe(
      '<rect x="0" y="0" width="100" height="40" fill="#f3f4f6" stroke="#9ca3af" stroke-dasharray="4 3"/><text x="4" y="16" font-family="sans-serif" font-size="12" fill="#374151">Chart data unavailable</text>',
    );
  });

  it('renders a chart placeholder for a malformed data container', () => {
    const result = render({
      chartType: 'barChart',
      colors: [],
      data: null,
      height: 40,
      id: 'chart-invalid',
      left: 0,
      order: 0,
      top: 0,
      type: 'chart',
      width: 100,
    } as unknown as Element);

    expect(result?.body).toContain('>Chart data unavailable</text>');
  });

  it('renders diagram nodes with flattened and escaped labels', () => {
    const result = render({
      elements: [
        {
          borderColor: '',
          borderStrokeDasharray: '',
          borderType: 'solid',
          borderWidth: 0,
          content: '<p>Node &amp; one</p>',
          fill: null,
          height: 20,
          id: 'node-1',
          isFlipH: false,
          isFlipV: false,
          left: 2,
          name: 'Node',
          order: 0,
          rotate: 0,
          shapType: 'rect',
          top: 3,
          type: 'shape',
          vAlign: 'top',
          width: 30,
          wrap: true,
        },
      ],
      height: 40,
      id: 'diagram-1',
      left: 0,
      order: 0,
      textList: [],
      top: 0,
      type: 'diagram',
      width: 100,
    });

    expect(result?.body).toContain('transform="translate(2 3)"');
    expect(result?.body).toContain('>Node &amp; one</text>');
    expect(result?.warningCode).toBe('approximate-diagram');
    expect(result?.warningMessage).toBe(
      'The preview preserves diagram labels with simplified nodes and styling.',
    );
  });

  it('renders diagram fallbacks, skips invalid nodes, and joins valid nodes directly', () => {
    const shape = {
      borderColor: '',
      borderStrokeDasharray: '',
      borderType: 'solid' as const,
      borderWidth: 0,
      content: '',
      fill: null,
      height: 10,
      id: 'shape-node',
      isFlipH: false,
      isFlipV: false,
      left: 1,
      name: 'Named node',
      order: 0,
      rotate: 0,
      shapType: 'rect',
      top: 2,
      type: 'shape' as const,
      vAlign: 'top',
      width: 20,
      wrap: true,
    };
    const text = {
      borderColor: '',
      borderStrokeDasharray: '',
      borderType: 'solid' as const,
      borderWidth: 0,
      content: '',
      fill: null,
      height: 10,
      id: 'text-node',
      isFlipH: false,
      isFlipV: false,
      isVertical: false,
      left: 30,
      name: 'Ignored name',
      order: 1,
      rotate: 0,
      top: 2,
      type: 'text' as const,
      vAlign: 'top',
      width: 20,
      wrap: true,
    };
    const result = render({
      elements: [shape, { ...shape, height: 0 }, text],
      height: 40,
      id: 'diagram-fallbacks',
      left: 0,
      order: 0,
      textList: [],
      top: 0,
      type: 'diagram',
      width: 100,
    });

    expect(result?.body).toContain('>Named node</text>');
    expect(result?.body).toContain('transform="translate(30 2)"');
    expect(result?.body).not.toContain('Ignored name');
    expect(result?.body).not.toContain('Stryker was here!');
    expect(result?.body.match(/<g transform=/g)).toHaveLength(2);
  });

  it.each([
    { elements: [], name: 'empty array' },
    { elements: null, name: 'malformed container' },
  ])('renders an exact diagram placeholder for $name', ({ elements }) => {
    const result = render({
      elements,
      height: 40,
      id: 'diagram-empty',
      left: 0,
      order: 0,
      textList: [],
      top: 0,
      type: 'diagram',
      width: 100,
    } as unknown as Element);

    expect(result?.body).toBe(
      '<rect x="0" y="0" width="100" height="40" fill="#f3f4f6" stroke="#9ca3af" stroke-dasharray="4 3"/><text x="4" y="16" font-family="sans-serif" font-size="12" fill="#374151">Diagram</text>',
    );
  });

  it('uses an embedded raster math preview before a text fallback', () => {
    const image = render({
      height: 40,
      id: 'math-image',
      latex: 'x',
      left: 0,
      order: 0,
      picBase64: 'data:image/png;base64,AA==',
      picBlob: '',
      picRef: '',
      top: 0,
      type: 'math',
      width: 100,
    });
    const fallback = render({
      height: 40,
      id: 'math-text',
      latex: '<x & y>',
      left: 0,
      order: 0,
      picBase64: '',
      picBlob: 'blob:unsafe',
      picRef: 'https://unsafe.example/math.png',
      top: 0,
      type: 'math',
      width: 100,
    });

    expect(image?.body).toContain('href="data:image/png;base64,AA=="');
    expect(fallback?.body).toContain('&lt;x &amp; y&gt;');
    expect(fallback?.body).not.toContain('unsafe.example');
    expect(fallback?.body).not.toContain('blob:unsafe');
    expect(fallback?.warningMessage).toBe(
      'The preview uses an embedded math image or a portable text fallback.',
    );
  });

  it.each([
    {
      expected: 'Authored text',
      latex: 'ignored',
      text: '<b>Authored text</b>',
    },
    { expected: 'Math', latex: '', text: undefined },
    {
      expected: 'Stryker was here!',
      latex: 'Stryker was here!',
      text: undefined,
    },
    { expected: 'Math', latex: 7, text: undefined },
  ])('selects the safe math fallback %#', ({ expected, latex, text }) => {
    const result = render({
      height: 40,
      id: 'math-fallback',
      latex,
      left: 0,
      order: 0,
      picBase64: '',
      picBlob: '',
      picRef: '',
      ...(text === undefined ? {} : { text }),
      top: 0,
      type: 'math',
      width: 100,
    } as unknown as Element);

    expect(result?.body).toContain(`>${expected}</text>`);
  });

  it.each([
    ['audio', 'Audio'],
    ['video', 'Video'],
  ] as const)('renders %s as a non-interactive placeholder', (type, text) => {
    const result = render({
      blob: 'blob:unsafe',
      height: 40,
      id: `${type}-1`,
      left: 0,
      order: 0,
      ref: 'https://unsafe.example/media',
      rotate: 0,
      top: 0,
      type,
      width: 100,
    });

    expect(result?.body).toContain(`>${text}</text>`);
    expect(result?.body).not.toContain('unsafe');
    expect(result?.warningCode).toBe('approximate-media');
    expect(result?.warningMessage).toBe(
      `The preview represents ${type} as a labeled non-interactive placeholder.`,
    );
  });

  it('returns null for an element rendered by the basic SVG layer', () => {
    expect(
      render({
        borderColor: '',
        borderStrokeDasharray: '',
        borderType: 'solid',
        borderWidth: 0,
        content: '',
        fill: null,
        height: 10,
        id: 'text-1',
        isFlipH: false,
        isFlipV: false,
        isVertical: false,
        left: 0,
        name: '',
        order: 0,
        rotate: 0,
        top: 0,
        type: 'text',
        vAlign: 'top',
        width: 10,
        wrap: true,
      }),
    ).toBeNull();
  });
});
