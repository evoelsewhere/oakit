import { describe, expect, it } from 'vitest';

import { renderPptxSvgSlideSource } from '../../src/formats/pptx/render-svg-slide';
import type {
  Element,
  PptxSlide,
  Shape,
  Text,
} from '../../src/formats/pptx/types';

function text(overrides: Partial<Text> = {}): Text {
  return {
    borderColor: '',
    borderStrokeDasharray: '',
    borderType: 'solid',
    borderWidth: 0,
    content: '<p>Hello &amp; welcome</p><p>Second</p>',
    fill: null,
    height: 40,
    id: 'text-1',
    isFlipH: false,
    isFlipV: false,
    isVertical: false,
    left: 10,
    name: 'Text',
    order: 0,
    rotate: 0,
    top: 20,
    type: 'text',
    vAlign: 'top',
    width: 100,
    wrap: true,
    ...overrides,
  };
}

function shape(overrides: Partial<Shape> = {}): Shape {
  return {
    borderColor: '#112233',
    borderStrokeDasharray: '2, 3',
    borderType: 'dashed',
    borderWidth: 1,
    content: '',
    fill: { type: 'color', value: '#abcdef' },
    height: 40,
    id: 'shape-1',
    isFlipH: false,
    isFlipV: false,
    left: 1,
    name: 'Shape',
    order: 0,
    rotate: 0,
    shapType: 'rect',
    top: 2,
    type: 'shape',
    vAlign: 'top',
    width: 80,
    wrap: true,
    ...overrides,
  };
}

function slide(elements: Element[], layoutElements: Element[] = []): PptxSlide {
  return {
    elements,
    fill: { type: 'color', value: '#ffffff' },
    layoutElements,
    note: '',
  };
}

function render(input: PptxSlide) {
  return renderPptxSvgSlideSource(input, {
    outputHeight: 90,
    outputWidth: 160,
    slideNumber: 2,
    sourceHeight: 90,
    sourceWidth: 160,
  });
}

describe('PowerPoint SVG slide source', () => {
  it('renders a deterministic accessible root, background, layout, and text', () => {
    const result = render(slide([text()], [shape({ id: 'layout' })]));

    expect(result.source).toBe(
      '<?xml version="1.0" encoding="UTF-8"?><svg xmlns="http://www.w3.org/2000/svg" role="img" width="160" height="90" viewBox="0 0 160 90"><title>PowerPoint slide 2</title><rect width="160" height="90" fill="#ffffff"/><g transform="translate(1 2)"><rect x="0" y="0" width="80" height="40" rx="0" fill="#abcdef" stroke="#112233" stroke-width="1" stroke-dasharray="2 3"/></g><g transform="translate(10 20)"><rect x="0" y="0" width="100" height="40" fill="none" stroke="none"/><svg x="0" y="0" width="100" height="40" overflow="hidden"><text x="4" y="16" text-anchor="start" font-family="sans-serif"><tspan fill="#111827" font-size="12">Hello &amp; welcome</tspan></text><text x="4" y="32" text-anchor="start" font-family="sans-serif"><tspan fill="#111827" font-size="12">Second</tspan></text></svg></g></svg>',
    );
    expect(result.warnings).toEqual([
      {
        code: 'font-substitution',
        elementId: 'text-1',
        message:
          'The preview uses a portable sans-serif font instead of the authored font.',
        slideNumber: 2,
      },
    ]);
  });

  it('escapes untrusted text and path data without emitting HTML containers', () => {
    const result = render(
      slide([
        shape({
          content: '<p>&lt;script&gt;&amp;</p>',
          path: 'M 0 0 L 10 10" onload="alert(1)',
          pathViewBox: { height: 10, width: 10, x: 0, y: 0 },
        }),
      ]),
    );

    expect(result.source).toContain(
      'd="M 0 0 L 10 10&quot; onload=&quot;alert(1)"',
    );
    expect(result.source).toContain('&lt;script&gt;&amp;');
    expect(result.source).not.toContain('<foreignObject');
    expect(result.warnings).toEqual([
      {
        code: 'font-substitution',
        elementId: 'shape-1',
        message:
          'The preview uses a portable sans-serif font instead of the authored font.',
        slideNumber: 2,
      },
    ]);
  });

  it.each([
    ['ellipse', '<ellipse cx="40" cy="20" rx="40" ry="20"'],
    ['line', '<line x1="0" y1="0" x2="80" y2="40"'],
    ['straightConnector1', '<line x1="0" y1="0" x2="80" y2="40"'],
    ['lineInv', '<line x1="0" y1="40" x2="80" y2="0"'],
    ['roundRect', '<rect x="0" y="0" width="80" height="40" rx="6"'],
  ])('renders the safe %s geometry', (shapType, fragment) => {
    const result = render(slide([shape({ shapType })]));
    expect(result.source).toContain(fragment);
    expect(result.warnings).toEqual([]);
  });

  it.each([
    ['straightConnector1', 80, 0, '<line x1="0" y1="0" x2="80" y2="0"'],
    ['straightConnector1', 0, 40, '<line x1="0" y1="0" x2="0" y2="40"'],
    ['line', 80, 0, '<line x1="0" y1="0" x2="80" y2="0"'],
    ['lineInv', 0, 40, '<line x1="0" y1="40" x2="0" y2="0"'],
  ])(
    'renders a zero-dimension %s with dimensions %sx%s',
    (shapType, width, height, fragment) => {
      const result = render(slide([shape({ height, shapType, width })]));

      expect(result.source).toContain(fragment);
      expect(result.warnings).toEqual([]);
    },
  );

  it.each([
    [Number.NaN, '#112233'],
    [-1, '#112233'],
    [1, 'not-a-color'],
  ])(
    'omits an unsafe border with width %s and color %s',
    (borderWidth, borderColor) => {
      const result = render(slide([shape({ borderColor, borderWidth })]));

      expect(result.source).toContain('fill="#abcdef" stroke="none"');
      expect(result.source).not.toContain('stroke-width=');
    },
  );

  it('omits an empty dash pattern from an otherwise valid border', () => {
    const result = render(
      slide([shape({ borderStrokeDasharray: '', borderType: 'solid' })]),
    );

    expect(result.source).toContain('stroke="#112233" stroke-width="1"');
    expect(result.source).not.toContain('stroke-dasharray=');
    expect(result.source).not.toContain('Stryker was here!');
  });

  it('treats a zero-width colored border as absent', () => {
    const result = render(
      slide([shape({ borderColor: '#112233', borderWidth: 0 })]),
    );

    expect(result.source).toContain('fill="#abcdef" stroke="none"');
    expect(result.source).not.toContain('stroke-width="0"');
  });

  it.each([
    [0, false, false, 'transform="translate(1 2)"'],
    [Number.POSITIVE_INFINITY, false, false, 'transform="translate(1 2)"'],
    ['15', false, false, 'transform="translate(1 2)"'],
    [0, true, false, 'transform="translate(1 2) translate(80 0) scale(-1 1)"'],
    [0, false, true, 'transform="translate(1 2) translate(0 40) scale(1 -1)"'],
  ])(
    'normalizes transform rotate=%s flipH=%s flipV=%s',
    (rotate, isFlipH, isFlipV, expected) => {
      const result = render(
        slide([
          shape({
            isFlipH,
            isFlipV,
            rotate: rotate as number,
          }),
        ]),
      );

      expect(result.source).toContain(expected);
      expect(result.source.match(/rotate\(/g)).toBeNull();
    },
  );

  it('renders a finite nonzero rotation around the element center', () => {
    expect(render(slide([shape({ rotate: -30 })])).source).toContain(
      'transform="translate(1 2) rotate(-30 40 20)"',
    );
  });

  it('treats non-string and empty text content as empty', () => {
    const result = render(
      slide([
        text({ content: '' }),
        text({ content: 7 as unknown as string, id: 'non-string', left: 120 }),
      ]),
    );

    expect(result.source).not.toContain('<tspan');
  });

  it('warns when a shape and fill require an approximation', () => {
    const result = render(
      slide([
        shape({
          fill: {
            type: 'pattern',
            value: {
              backgroundColor: '#ffffff',
              foregroundColor: '#000000',
              type: 'cross',
            },
          },
          shapType: 'star5',
        }),
      ]),
    );

    expect(result.source).toContain('fill="none"');
    expect(result.warnings).toEqual([
      {
        code: 'approximate-fill',
        elementId: 'shape-1',
        message:
          'The preview substituted a fill that SVG cannot safely reproduce yet.',
        slideNumber: 2,
      },
      {
        code: 'approximate-shape',
        elementId: 'shape-1',
        message:
          'The preview represents PowerPoint shape star5 as a rectangle.',
        slideNumber: 2,
      },
    ]);
  });

  it('does not reinterpret a malformed non-color fill as a safe color', () => {
    const result = render(
      slide([
        text({
          fill: {
            type: 'pattern',
            value: '#123456',
          } as unknown as Text['fill'],
        }),
      ]),
    );

    expect(result.source).toContain('fill="none"');
    expect(result.source).not.toContain('fill="#123456"');
    expect(result.warnings.map(({ code }) => code)).toEqual([
      'font-substitution',
      'approximate-fill',
    ]);
  });

  it('keeps stroke-only shapes unfilled', () => {
    const result = render(
      slide([
        shape({ fill: { type: 'color', value: '#abcdef' }, strokeOnly: true }),
      ]),
    );

    expect(result.source).toContain('fill="none" stroke="#112233"');
    expect(result.source).not.toContain('fill="#abcdef"');
  });

  it.each(['', 'Stryker was here!'])(
    'does not render an empty-equivalent path %j',
    (path) => {
      const result = render(
        slide([
          shape({
            path,
            pathViewBox: { height: 10, width: 10, x: 0, y: 0 },
          }),
        ]),
      );

      if (path === '') {
        expect(result.source).not.toContain('<path');
      } else {
        expect(result.source).toContain(`<path d="${path}"`);
      }
    },
  );

  it('renders embedded raster bytes and reports approximated effects', () => {
    const result = render(
      slide([
        {
          base64: 'data:image/png;base64,AA==',
          blob: 'blob:unsafe',
          borderColor: '#010203',
          borderStrokeDasharray: '',
          borderType: 'solid',
          borderWidth: 2,
          filters: { brightness: 10 },
          geom: 'rect',
          height: 30,
          id: 'image-1',
          isFlipH: true,
          isFlipV: true,
          left: 5,
          order: 0,
          rect: { l: 5 },
          ref: 'https://unsafe.example/image.png',
          rotate: 15,
          top: 6,
          type: 'image',
          width: 40,
        },
      ]),
    );

    expect(result.source).toContain(
      'transform="translate(5 6) rotate(15 20 15) translate(40 30) scale(-1 -1)"',
    );
    expect(result.source).toContain('href="data:image/png;base64,AA=="');
    expect(result.source).not.toContain('unsafe.example');
    expect(result.source).not.toContain('blob:unsafe');
    expect(result.warnings).toEqual([
      {
        code: 'approximate-media',
        elementId: 'image-1',
        message:
          'The preview omitted an unsafe image crop or unsupported filter effect.',
        slideNumber: 2,
      },
    ]);
  });

  it.each([
    [{ rect: { l: 5 } }, []],
    [{ filters: { brightness: 10 } }, ['approximate-media']],
    [{}, []],
  ])('reports image effects independently for %#', (effects, warningCodes) => {
    const result = render(
      slide([
        {
          base64: 'data:image/png;base64,AA==',
          blob: '',
          borderColor: '',
          borderStrokeDasharray: '',
          borderType: 'solid',
          borderWidth: 0,
          geom: 'rect',
          height: 30,
          id: 'image-effects',
          isFlipH: false,
          isFlipV: false,
          left: 5,
          order: 0,
          ref: '',
          rotate: 0,
          top: 6,
          type: 'image',
          width: 40,
          ...effects,
        },
      ]),
    );

    expect(result.warnings.map(({ code }) => code)).toEqual(warningCodes);
  });

  it('renders a safe image crop inside a clipped SVG viewport', () => {
    const result = render(
      slide([
        {
          base64: 'data:image/png;base64,AA==',
          blob: '',
          borderColor: '',
          borderStrokeDasharray: '',
          borderType: 'solid',
          borderWidth: 0,
          geom: 'rect',
          height: 80,
          id: 'cropped-image',
          isFlipH: false,
          isFlipV: false,
          left: 0,
          order: 0,
          rect: { b: 0, l: 10, r: 20, t: 25 },
          ref: '',
          rotate: 0,
          top: 0,
          type: 'image',
          width: 100,
        },
      ]),
    );

    expect(result.source).toContain(
      '<svg x="0" y="0" width="100" height="80" overflow="hidden"><image x="-14.2857" y="-26.6667" width="142.8571" height="106.6667" preserveAspectRatio="none" href="data:image/png;base64,AA=="/></svg>',
    );
    expect(result.warnings).toEqual([]);
  });

  it('falls back visibly and warns for an unsafe image crop', () => {
    const image = {
      base64: 'data:image/png;base64,AA==',
      blob: '',
      borderColor: '',
      borderStrokeDasharray: '',
      borderType: 'solid' as const,
      borderWidth: 0,
      geom: 'rect',
      height: 30,
      id: 'unsafe-crop',
      isFlipH: false,
      isFlipV: false,
      left: 0,
      order: 0,
      rect: { l: 101, r: -2 },
      ref: '',
      rotate: 0,
      top: 0,
      type: 'image' as const,
      width: 40,
    };

    const result = render(slide([image]));

    expect(result.source).toContain('preserveAspectRatio="xMidYMid meet"');
    expect(result.warnings.map(({ code }) => code)).toEqual([
      'approximate-media',
    ]);
  });

  it('renders an uncropped image with portable meet semantics', () => {
    const image = {
      base64: 'data:image/png;base64,AA==',
      blob: '',
      borderColor: '',
      borderStrokeDasharray: '',
      borderType: 'solid' as const,
      borderWidth: 0,
      geom: 'rect',
      height: 30,
      id: 'uncropped-image',
      isFlipH: false,
      isFlipV: false,
      left: 0,
      order: 0,
      ref: '',
      rotate: 0,
      top: 0,
      type: 'image' as const,
      width: 40,
    };

    const result = render(slide([image]));

    expect(result.source).toContain(
      '<image x="0" y="0" width="40" height="30" preserveAspectRatio="xMidYMid meet" href="data:image/png;base64,AA=="/>',
    );
    expect(result.warnings).toEqual([]);
  });

  it('uses a visible placeholder instead of an external or missing image', () => {
    const result = render(
      slide([
        {
          base64: '',
          blob: 'blob:unsafe',
          borderColor: '',
          borderStrokeDasharray: '',
          borderType: 'solid',
          borderWidth: 0,
          geom: 'rect',
          height: 30,
          id: 'image-1',
          isFlipH: false,
          isFlipV: false,
          left: 5,
          order: 0,
          ref: 'file:///secret.png',
          rotate: 0,
          top: 6,
          type: 'image',
          width: 40,
        },
      ]),
    );

    expect(result.source).toContain('Image unavailable');
    expect(result.source).not.toContain('file:///');
    expect(result.warnings).toEqual([
      {
        code: 'missing-media',
        elementId: 'image-1',
        message:
          'The preview omitted an image because no safe embedded raster source was available.',
        slideNumber: 2,
      },
    ]);
  });

  it('renders nested groups in local coordinate systems', () => {
    const result = render(
      slide([
        {
          elements: [
            shape({ id: 'one', left: 3, top: 4 }),
            shape({ id: 'two', left: 5, top: 6 }),
          ],
          height: 40,
          id: 'group-1',
          isFlipH: false,
          isFlipV: false,
          left: 20,
          order: 0,
          rotate: 90,
          top: 10,
          type: 'group',
          width: 80,
        },
      ]),
    );

    expect(result.source).toContain(
      '<g transform="translate(20 10) rotate(90 40 20)"><g transform="translate(3 4)">',
    );
    expect(result.source).toContain('</g><g transform="translate(5 6)"><rect');
  });

  it('renders children of a non-empty zero-height group', () => {
    const result = render(
      slide([
        {
          elements: [shape({ height: 0, shapType: 'straightConnector1' })],
          height: 0,
          id: 'line-group',
          isFlipH: false,
          isFlipV: false,
          left: 20,
          order: 0,
          rotate: 0,
          top: 10,
          type: 'group',
          width: 80,
        },
      ]),
    );

    expect(result.source).toContain(
      '<g transform="translate(20 10)"><g transform="translate(1 2)"><line',
    );
    expect(result.warnings).toEqual([]);
  });

  it('rejects an empty zero-height group', () => {
    const result = render(
      slide([
        {
          elements: [],
          height: 0,
          id: 'empty-group',
          isFlipH: false,
          isFlipV: false,
          left: 20,
          order: 0,
          rotate: 0,
          top: 10,
          type: 'group',
          width: 80,
        },
      ]),
    );

    expect(result.warnings.map(({ code }) => code)).toEqual([
      'approximate-shape',
    ]);
  });

  it('skips invalid geometry and omits an empty element id', () => {
    const result = render(slide([text({ height: 0, id: '' })]));

    expect(result.source).toBe(
      '<?xml version="1.0" encoding="UTF-8"?><svg xmlns="http://www.w3.org/2000/svg" role="img" width="160" height="90" viewBox="0 0 160 90"><title>PowerPoint slide 2</title><rect width="160" height="90" fill="#ffffff"/></svg>',
    );
    expect(result.warnings).toEqual([
      {
        code: 'approximate-shape',
        message: 'The preview skipped an element with invalid geometry.',
        slideNumber: 2,
      },
    ]);
  });

  it('delegates charts to the rich portable renderer', () => {
    const result = render(
      slide([
        {
          data: [],
          height: 20,
          id: 'chart-1',
          left: 1,
          order: 0,
          top: 2,
          type: 'chart',
          width: 30,
        } as unknown as Element,
      ]),
    );

    expect(result.source).toContain('>Chart data unavailable</text>');
    expect(result.source).toContain('<g transform="translate(1 2)"><rect');
    expect(result.warnings).toEqual([
      {
        code: 'approximate-chart',
        elementId: 'chart-1',
        message:
          'The preview visualizes chart values with simplified portable bars.',
        slideNumber: 2,
      },
    ]);
  });

  it('uses a safe placeholder for an unknown runtime element type', () => {
    const result = render(
      slide([
        {
          height: 20,
          id: 'widget-1',
          left: 1,
          order: 0,
          top: 2,
          type: 'widget',
          width: 30,
        } as unknown as Element,
      ]),
    );

    expect(result.source).toContain('>widget</text>');
    expect(result.source).toContain('<g transform="translate(1 2)"><rect');
    expect(result.source).not.toContain('scale(');
    expect(result.warnings).toEqual([
      {
        code: 'approximate-shape',
        elementId: 'widget-1',
        message:
          'The preview represents PowerPoint element widget as a placeholder.',
        slideNumber: 2,
      },
    ]);
  });

  it('omits a malformed non-string element id from warnings', () => {
    const result = render(
      slide([
        {
          ...shape({ height: 0 }),
          id: 7,
        } as unknown as Element,
      ]),
    );

    expect(result.warnings).toEqual([
      {
        code: 'approximate-shape',
        message: 'The preview skipped an element with invalid geometry.',
        slideNumber: 2,
      },
    ]);
  });

  it('uses the safe white fallback for an unsupported slide background', () => {
    const input = slide([]);
    input.fill = {
      type: 'gradient',
      value: { colors: [], path: 'line', rot: 0 },
    };

    const result = render(input);

    expect(result.source).toContain(
      '<rect width="160" height="90" fill="#ffffff"/>',
    );
    expect(result.warnings).toEqual([]);
  });

  it('defines and applies safe linear gradients without approximation warnings', () => {
    const first = shape({ id: 'gradient-1' });
    first.fill = {
      type: 'gradient',
      value: {
        colors: [
          { color: '#FFFFFF', pos: '0%' },
          { color: '#434343', pos: '100%' },
        ],
        path: 'line',
        rot: 90,
      },
    };
    const second = shape({ id: 'gradient-2', left: 90 });
    second.fill = first.fill;

    const result = render(slide([first, second]));

    expect(result.source).toContain(
      '<defs><linearGradient id="pptx-gradient-2-1"',
    );
    expect(result.source).toContain('<linearGradient id="pptx-gradient-2-2"');
    expect(result.source).toContain('fill="url(#pptx-gradient-2-1)"');
    expect(result.source).toContain('fill="url(#pptx-gradient-2-2)"');
    expect(result.source).not.toContain('Stryker was here');
    expect(result.warnings).toEqual([]);

    const solid = render(slide([shape()]));
    expect(solid.source).not.toContain('<defs>');
    expect(solid.source).not.toContain('Stryker was here');
  });
});
