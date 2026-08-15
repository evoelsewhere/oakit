import { describe, expect, it } from 'vitest';

import type { XmlLookupValue } from '../../src/common';
import {
  getCustomShapePath,
  identifyShape,
  isStrokeOnlyCustomGeometry,
} from '../../src/formats/pptx/internal/shape';

interface Point {
  x: number;
  y: number;
}

function xml(value: object): XmlLookupValue {
  return value as unknown as XmlLookupValue;
}

function pointNode(point: Point, order: number): object {
  return {
    attrs: {
      order: String(order),
      x: String(point.x),
      y: String(point.y),
    },
  };
}

function polygon(points: readonly Point[]): XmlLookupValue {
  const [first, ...rest] = points;
  if (!first) return xml({ 'a:pathLst': { 'a:path': {} } });

  return xml({
    'a:pathLst': {
      'a:path': {
        attrs: { h: '1000', w: '1000' },
        'a:moveTo': { 'a:pt': pointNode(first, 1) },
        'a:lnTo': rest.map((point, index) => ({
          'a:pt': pointNode(point, index + 2),
        })),
        'a:close': { attrs: { order: String(points.length + 1) } },
      },
    },
  });
}

function regularPolygon(vertexCount: number): Point[] {
  return Array.from({ length: vertexCount }, (_, index) => {
    const angle = -Math.PI / 2 + (index * Math.PI * 2) / vertexCount;
    return {
      x: Math.round(500 + Math.cos(angle) * 400),
      y: Math.round(500 + Math.sin(angle) * 400),
    };
  });
}

function curvePath(options: {
  arcCount?: number;
  close?: boolean;
  cubicEndpoints?: readonly Point[];
  height?: number;
  lineCount?: number;
  quadraticCount?: number;
  width?: number;
}): XmlLookupValue {
  const {
    arcCount = 0,
    close = true,
    cubicEndpoints = [],
    height = 1000,
    lineCount = 0,
    quadraticCount = 0,
    width = 1000,
  } = options;
  let order = 1;
  const path: Record<string, object | object[]> = {
    attrs: { h: String(height), w: String(width) },
    'a:moveTo': { 'a:pt': pointNode({ x: 0, y: 0 }, order++) },
  };

  if (lineCount > 0) {
    path['a:lnTo'] = Array.from({ length: lineCount }, (_, index) => ({
      'a:pt': pointNode(
        { x: (index + 1) * 100, y: index % 2 === 0 ? 200 : 700 },
        order++,
      ),
    }));
  }
  if (cubicEndpoints.length > 0) {
    path['a:cubicBezTo'] = cubicEndpoints.map((endpoint) => ({
      'a:pt': [
        pointNode({ x: endpoint.x / 3, y: endpoint.y / 3 }, order++),
        pointNode(
          { x: (endpoint.x * 2) / 3, y: (endpoint.y * 2) / 3 },
          order++,
        ),
        pointNode(endpoint, order++),
      ],
    }));
  }
  if (quadraticCount > 0) {
    path['a:quadBezTo'] = Array.from(
      { length: quadraticCount },
      (_, index) => ({
        'a:pt': [
          pointNode({ x: 100 + index, y: 200 + index }, order++),
          pointNode({ x: 300 + index, y: 400 + index }, order++),
        ],
      }),
    );
  }
  if (arcCount > 0) {
    path['a:arcTo'] = Array.from({ length: arcCount }, () => ({
      attrs: {
        hR: '500',
        order: String(order++),
        stAng: '0',
        swAng: '10800000',
        wR: '500',
      },
    }));
  }
  if (close) path['a:close'] = { attrs: { order: String(order) } };

  return xml({ 'a:pathLst': { 'a:path': path } });
}

describe('PowerPoint custom shape classification', () => {
  it('keeps missing, empty, open, and degenerate geometry custom', () => {
    expect(identifyShape(xml({}))).toBe('custom');
    expect(identifyShape(xml({ 'a:pathLst': { 'a:path': {} } }))).toBe(
      'custom',
    );
    expect(
      identifyShape(
        xml({
          'a:pathLst': {
            'a:path': {
              attrs: { h: '100', w: '100' },
              'a:moveTo': { 'a:pt': pointNode({ x: 0, y: 0 }, 1) },
              'a:lnTo': [
                { 'a:pt': pointNode({ x: 100, y: 0 }, 2) },
                { 'a:pt': pointNode({ x: 50, y: 100 }, 3) },
              ],
            },
          },
        }),
      ),
    ).toBe('custom');
    expect(
      identifyShape(
        polygon([
          { x: 10, y: 10 },
          { x: 10, y: 10 },
          { x: 10, y: 10 },
        ]),
      ),
    ).toBe('custom');
    expect(
      identifyShape(
        polygon([
          { x: 0, y: 0 },
          { x: 100, y: 100 },
        ]),
      ),
    ).toBe('custom');
  });

  it.each([
    ['triangle', regularPolygon(3)],
    [
      'rect',
      [
        { x: 0, y: 0 },
        { x: 1000, y: 0 },
        { x: 1000, y: 600 },
        { x: 0, y: 600 },
      ],
    ],
    [
      'rhombus',
      [
        { x: 500, y: 0 },
        { x: 1000, y: 500 },
        { x: 500, y: 1000 },
        { x: 0, y: 500 },
      ],
    ],
    [
      'parallelogram',
      [
        { x: 200, y: 0 },
        { x: 1000, y: 0 },
        { x: 800, y: 600 },
        { x: 0, y: 600 },
      ],
    ],
    [
      'trapezoid',
      [
        { x: 200, y: 0 },
        { x: 800, y: 0 },
        { x: 1000, y: 600 },
        { x: 0, y: 600 },
      ],
    ],
    [
      'custom',
      [
        { x: 0, y: 0 },
        { x: 1000, y: 100 },
        { x: 700, y: 900 },
        { x: 100, y: 700 },
      ],
    ],
    ['pentagon', regularPolygon(5)],
    ['hexagon', regularPolygon(6)],
    ['heptagon', regularPolygon(7)],
    ['octagon', regularPolygon(8)],
    ['custom', regularPolygon(9)],
    ['decagon', regularPolygon(10)],
    ['custom', regularPolygon(11)],
    ['dodecagon', regularPolygon(12)],
  ] as const)('classifies a closed %s geometry', (expected, points) => {
    expect(identifyShape(polygon(points))).toBe(expected);
  });

  it('recognizes either traversal direction of an axis-aligned rectangle', () => {
    expect(
      identifyShape(
        polygon([
          { x: 0, y: 0 },
          { x: 0, y: 600 },
          { x: 1000, y: 600 },
          { x: 1000, y: 0 },
        ]),
      ),
    ).toBe('rect');
  });

  it.each([
    [
      'first horizontal edge',
      [
        { x: 0, y: 0 },
        { x: 100, y: 1 },
        { x: 100, y: 100 },
        { x: 0, y: 100 },
      ],
    ],
    [
      'first vertical edge',
      [
        { x: 0, y: 0 },
        { x: 100, y: 0 },
        { x: 99, y: 100 },
        { x: 0, y: 100 },
      ],
    ],
    [
      'opposite horizontal edge',
      [
        { x: 0, y: 0 },
        { x: 100, y: 0 },
        { x: 100, y: 100 },
        { x: 0, y: 99 },
      ],
    ],
    [
      'opposite vertical edge',
      [
        { x: 0, y: 0 },
        { x: 100, y: 0 },
        { x: 100, y: 100 },
        { x: 1, y: 100 },
      ],
    ],
    [
      'first reverse vertical edge',
      [
        { x: 0, y: 0 },
        { x: 1, y: 100 },
        { x: 100, y: 100 },
        { x: 100, y: 0 },
      ],
    ],
    [
      'first reverse horizontal edge',
      [
        { x: 0, y: 0 },
        { x: 0, y: 100 },
        { x: 100, y: 99 },
        { x: 100, y: 0 },
      ],
    ],
    [
      'opposite reverse vertical edge',
      [
        { x: 0, y: 0 },
        { x: 0, y: 100 },
        { x: 100, y: 100 },
        { x: 99, y: 0 },
      ],
    ],
    [
      'opposite reverse horizontal edge',
      [
        { x: 0, y: 0 },
        { x: 0, y: 100 },
        { x: 100, y: 100 },
        { x: 100, y: 1 },
      ],
    ],
  ] as const)(
    'rejects a near rectangle with a changed %s',
    (_label, points) => {
      expect(identifyShape(polygon(points))).not.toBe('rect');
    },
  );

  it('removes only exactly repeated vertices', () => {
    expect(
      identifyShape(
        polygon([
          { x: 0, y: 0 },
          { x: 1000, y: 0 },
          { x: 0, y: 1000 },
          { x: 0, y: 0 },
        ]),
      ),
    ).toBe('triangle');
    expect(
      identifyShape(
        polygon([
          { x: 0, y: 0 },
          { x: 1, y: 0 },
          { x: 0, y: 1 },
        ]),
      ),
    ).toBe('triangle');
  });

  it('does not mistake a four-curve path missing two extrema for an ellipse', () => {
    const geometry = xml({
      'a:pathLst': {
        'a:path': {
          attrs: { h: '1000', w: '1000' },
          'a:moveTo': { 'a:pt': pointNode({ x: 500, y: 0 }, 1) },
          'a:cubicBezTo': [
            [
              { x: 800, y: 0 },
              { x: 1000, y: 200 },
              { x: 1000, y: 500 },
            ],
            [
              { x: 1000, y: 800 },
              { x: 800, y: 1000 },
              { x: 500, y: 1000 },
            ],
            [
              { x: 800, y: 900 },
              { x: 900, y: 800 },
              { x: 1000, y: 500 },
            ],
            [
              { x: 900, y: 200 },
              { x: 800, y: 100 },
              { x: 500, y: 0 },
            ],
          ].map((points, commandIndex) => ({
            'a:pt': points.map((point, pointIndex) =>
              pointNode(point, commandIndex * 3 + pointIndex + 2),
            ),
          })),
          'a:close': { attrs: { order: '14' } },
        },
      },
    });

    expect(identifyShape(geometry)).toBe('custom');
  });

  it('recognizes a closed four-cubic ellipse from all four extrema', () => {
    const geometry = xml({
      'a:pathLst': {
        'a:path': {
          attrs: { h: '1000', w: '1000' },
          'a:moveTo': { 'a:pt': pointNode({ x: 500, y: 0 }, 1) },
          'a:cubicBezTo': [
            [
              { x: 776, y: 0 },
              { x: 1000, y: 224 },
              { x: 1000, y: 500 },
            ],
            [
              { x: 1000, y: 776 },
              { x: 776, y: 1000 },
              { x: 500, y: 1000 },
            ],
            [
              { x: 224, y: 1000 },
              { x: 0, y: 776 },
              { x: 0, y: 500 },
            ],
            [
              { x: 0, y: 224 },
              { x: 224, y: 0 },
              { x: 500, y: 0 },
            ],
          ].map((points, commandIndex) => ({
            'a:pt': points.map((point, pointIndex) =>
              pointNode(point, commandIndex * 3 + pointIndex + 2),
            ),
          })),
          'a:close': { attrs: { order: '14' } },
        },
      },
    });

    expect(identifyShape(geometry)).toBe('ellipse');
  });

  it.each([
    [
      'top threshold',
      [
        { x: 500, y: 100 },
        { x: 1000, y: 500 },
        { x: 500, y: 1000 },
        { x: 0, y: 500 },
      ],
    ],
    [
      'bottom threshold',
      [
        { x: 500, y: 0 },
        { x: 1000, y: 500 },
        { x: 500, y: 900 },
        { x: 0, y: 500 },
      ],
    ],
    [
      'left threshold',
      [
        { x: 500, y: 0 },
        { x: 1000, y: 500 },
        { x: 500, y: 1000 },
        { x: 100, y: 500 },
      ],
    ],
    [
      'right threshold',
      [
        { x: 500, y: 0 },
        { x: 900, y: 500 },
        { x: 500, y: 1000 },
        { x: 0, y: 500 },
      ],
    ],
    [
      'outside the top tolerance',
      [
        { x: 500, y: 200 },
        { x: 1000, y: 500 },
        { x: 500, y: 1000 },
        { x: 0, y: 500 },
      ],
    ],
  ] as const)(
    'does not accept an ellipse endpoint on the %s',
    (_label, endpoints) => {
      expect(identifyShape(curvePath({ cubicEndpoints: endpoints }))).toBe(
        'custom',
      );
    },
  );

  it('requires positive path dimensions before classifying cubic extrema', () => {
    const endpoints = [
      { x: 0, y: 0 },
      { x: 0, y: 0 },
      { x: 0, y: 0 },
      { x: 0, y: 0 },
    ];
    expect(
      identifyShape(curvePath({ cubicEndpoints: endpoints, width: 0 })),
    ).toBe('custom');
    expect(
      identifyShape(curvePath({ cubicEndpoints: endpoints, height: 0 })),
    ).toBe('custom');
  });

  it('does not classify extra, open, or line-mixed cubic curves as circles', () => {
    const extrema = [
      { x: 500, y: 0 },
      { x: 1000, y: 500 },
      { x: 500, y: 1000 },
      { x: 0, y: 500 },
    ];
    expect(
      identifyShape(
        curvePath({
          cubicEndpoints: [...extrema, { x: 500, y: 500 }],
        }),
      ),
    ).toBe('custom');
    expect(
      identifyShape(curvePath({ close: false, cubicEndpoints: extrema })),
    ).toBe('custom');
    expect(
      identifyShape(curvePath({ cubicEndpoints: extrema, lineCount: 1 })),
    ).toBe('custom');
    expect(
      identifyShape(curvePath({ cubicEndpoints: extrema, quadraticCount: 1 })),
    ).toBe('custom');
  });

  it('recognizes a closed pair of arcs as an ellipse', () => {
    expect(
      identifyShape(
        xml({
          'a:pathLst': {
            'a:path': {
              attrs: { h: '1000', w: '1000' },
              'a:arcTo': [
                {
                  attrs: {
                    hR: '500',
                    order: '1',
                    stAng: '0',
                    swAng: '10800000',
                    wR: '500',
                  },
                },
                {
                  attrs: {
                    hR: '500',
                    order: '2',
                    stAng: '10800000',
                    swAng: '10800000',
                    wR: '500',
                  },
                },
              ],
              'a:close': { attrs: { order: '3' } },
            },
          },
        }),
      ),
    ).toBe('ellipse');
  });

  it('requires two closed arcs without line segments for an ellipse', () => {
    expect(identifyShape(curvePath({ arcCount: 1 }))).toBe('custom');
    expect(identifyShape(curvePath({ arcCount: 2, close: false }))).toBe(
      'custom',
    );
    expect(identifyShape(curvePath({ arcCount: 2, lineCount: 1 }))).toBe(
      'custom',
    );
    expect(identifyShape(curvePath({ arcCount: 1, lineCount: 2 }))).toBe(
      'custom',
    );
  });

  it('recognizes line and curve combinations without promoting open paths', () => {
    const curvedTriangle = {
      attrs: { h: '1000', w: '1000' },
      'a:moveTo': { 'a:pt': pointNode({ x: 0, y: 1000 }, 1) },
      'a:lnTo': [
        { 'a:pt': pointNode({ x: 500, y: 0 }, 2) },
        { 'a:pt': pointNode({ x: 1000, y: 1000 }, 3) },
        { 'a:pt': pointNode({ x: 750, y: 900 }, 4) },
      ],
      'a:quadBezTo': {
        'a:pt': [
          pointNode({ x: 500, y: 850 }, 5),
          pointNode({ x: 0, y: 1000 }, 6),
        ],
      },
    };

    expect(
      identifyShape(
        xml({
          'a:pathLst': {
            'a:path': {
              ...curvedTriangle,
              'a:close': { attrs: { order: '7' } },
            },
          },
        }),
      ),
    ).toBe('triangle');
    expect(
      identifyShape(xml({ 'a:pathLst': { 'a:path': curvedTriangle } })),
    ).toBe('custom');
  });

  it.each([
    ['triangle', 3],
    ['roundRect', 4],
    ['pentagon', 5],
    ['hexagon', 6],
    ['heptagon', 7],
    ['octagon', 8],
  ] as const)(
    'uses the %s family for a closed path with %i lines and one curve',
    (expected, lineCount) => {
      expect(identifyShape(curvePath({ lineCount, quadraticCount: 1 }))).toBe(
        expected,
      );
    },
  );

  it('requires the complete mixed-curve contract', () => {
    expect(identifyShape(curvePath({ lineCount: 2, quadraticCount: 1 }))).toBe(
      'custom',
    );
    expect(identifyShape(curvePath({ lineCount: 3, quadraticCount: 4 }))).toBe(
      'custom',
    );
    expect(identifyShape(curvePath({ lineCount: 3, quadraticCount: 3 }))).toBe(
      'triangle',
    );
    expect(identifyShape(curvePath({ lineCount: 4, quadraticCount: 4 }))).toBe(
      'roundRect',
    );
    expect(identifyShape(curvePath({ lineCount: 4, quadraticCount: 5 }))).toBe(
      'custom',
    );
    expect(
      identifyShape(
        curvePath({ close: false, lineCount: 4, quadraticCount: 4 }),
      ),
    ).toBe('custom');
    expect(
      identifyShape(
        curvePath({ close: false, lineCount: 3, quadraticCount: 1 }),
      ),
    ).toBe('custom');
    expect(identifyShape(curvePath({ arcCount: 1, lineCount: 3 }))).toBe(
      'custom',
    );
    expect(identifyShape(curvePath({ lineCount: 9, quadraticCount: 1 }))).toBe(
      'custom',
    );
  });

  it('ignores malformed cubic and quadratic commands', () => {
    const malformed = (elementName: 'a:cubicBezTo' | 'a:quadBezTo') =>
      xml({
        'a:pathLst': {
          'a:path': {
            attrs: { h: '1000', w: '1000' },
            [elementName]: {
              'a:pt': pointNode({ x: 500, y: 500 }, 1),
            },
            'a:close': { attrs: { order: '2' } },
          },
        },
      });

    expect(identifyShape(malformed('a:cubicBezTo'))).toBe('custom');
    expect(identifyShape(malformed('a:quadBezTo'))).toBe('custom');
  });

  it('does not count malformed curves beside a valid curved polygon', () => {
    const geometry = (elementName: 'a:cubicBezTo' | 'a:quadBezTo') => {
      const result = curvePath({ lineCount: 3, quadraticCount: 3 });
      const pathList = result as unknown as Record<string, unknown>;
      const container = pathList['a:pathLst'] as
        Record<string, unknown> | undefined;
      const path = container?.['a:path'] as Record<string, unknown> | undefined;
      if (!path) throw new Error('Expected a generated path');
      const existing = path[elementName];
      const existingCommands = Array.isArray(existing)
        ? (existing as unknown[])
        : [];
      path[elementName] = [
        ...existingCommands,
        { 'a:pt': pointNode({ x: 500, y: 500 }, 99) },
      ];
      return result;
    };

    expect(identifyShape(geometry('a:cubicBezTo'))).toBe('triangle');
    expect(identifyShape(geometry('a:quadBezTo'))).toBe('triangle');
  });

  it('recognizes rhombi with horizontal or vertical sides', () => {
    expect(
      identifyShape(
        polygon([
          { x: 0, y: 0 },
          { x: 5, y: 0 },
          { x: 8, y: 4 },
          { x: 3, y: 4 },
        ]),
      ),
    ).toBe('rhombus');
    expect(
      identifyShape(
        polygon([
          { x: 0, y: 0 },
          { x: 0, y: 5 },
          { x: 4, y: 8 },
          { x: 4, y: 3 },
        ]),
      ),
    ).toBe('rhombus');
  });
});

describe('PowerPoint custom path collection', () => {
  it('returns an empty path when custom geometry has no paths', () => {
    expect(getCustomShapePath(xml({}), 20, 10)).toBe('');
  });

  it.each([
    ['zero width', { h: '100', w: '0' }],
    ['zero height', { h: '0', w: '100' }],
    ['missing width', { h: '100' }],
    ['missing height', { w: '100' }],
    ['non-numeric width', { h: '100', w: 'wide' }],
    ['non-numeric height', { h: 'tall', w: '100' }],
  ] as const)('skips a path with %s', (_label, attrs) => {
    const geometry = xml({
      'a:pathLst': {
        'a:path': {
          attrs,
          'a:moveTo': {
            attrs: { order: '1' },
            'a:pt': pointNode({ x: 10, y: 10 }, 2),
          },
        },
      },
    });

    expect(getCustomShapePath(geometry, 20, 10)).toBe('');
  });

  it('renders every path with its own coordinate system', () => {
    const geometry = xml({
      'a:pathLst': {
        'a:path': [
          {
            attrs: { h: '100', order: '1', w: '100' },
            'a:moveTo': {
              attrs: { order: '2' },
              'a:pt': pointNode({ x: 0, y: 0 }, 3),
            },
            'a:lnTo': {
              attrs: { order: '4' },
              'a:pt': pointNode({ x: 100, y: 100 }, 5),
            },
            'a:close': { attrs: { order: '6' } },
          },
          {
            attrs: { h: '100', order: '7', w: '200' },
            'a:moveTo': {
              attrs: { order: '8' },
              'a:pt': pointNode({ x: 200, y: 0 }, 9),
            },
            'a:lnTo': {
              attrs: { order: '10' },
              'a:pt': pointNode({ x: 0, y: 100 }, 11),
            },
          },
        ],
      },
    });

    expect(getCustomShapePath(geometry, 20, 10)).toBe(
      ' M0,0 L20,10z M20,0 L0,10',
    );
  });

  it('preserves interleaved command and subpath order', () => {
    const geometry = xml({
      'a:pathLst': {
        'a:path': {
          attrs: { h: '100', w: '100' },
          'a:moveTo': [
            {
              attrs: { order: '1' },
              'a:pt': pointNode({ x: 0, y: 0 }, 90),
            },
            {
              attrs: { order: '5' },
              'a:pt': pointNode({ x: 10, y: 10 }, 91),
            },
          ],
          'a:lnTo': {
            attrs: { order: '3' },
            'a:pt': pointNode({ x: 100, y: 0 }, 92),
          },
          'a:quadBezTo': {
            attrs: { order: '2' },
            'a:pt': [
              pointNode({ x: 25, y: 25 }, 93),
              pointNode({ x: 50, y: 50 }, 94),
            ],
          },
          'a:cubicBezTo': {
            attrs: { order: '6' },
            'a:pt': [
              pointNode({ x: 20, y: 20 }, 95),
              pointNode({ x: 30, y: 30 }, 96),
              pointNode({ x: 40, y: 40 }, 97),
            ],
          },
          'a:close': [{ attrs: { order: '4' } }, { attrs: { order: '7' } }],
        },
      },
    });

    expect(getCustomShapePath(geometry, 10, 10)).toBe(
      ' M0,0 Q2.5,2.5 5,5 L10,0z M1,1 C2,2 3,3 4,4z',
    );
  });

  it('skips point and Bézier commands with missing control points', () => {
    const geometry = xml({
      'a:pathLst': {
        'a:path': {
          attrs: { h: '100', w: '100' },
          'a:moveTo': { attrs: { order: '1' } },
          'a:lnTo': { attrs: { order: '2' } },
          'a:quadBezTo': {
            attrs: { order: '3' },
            'a:pt': pointNode({ x: 10, y: 10 }, 4),
          },
          'a:cubicBezTo': {
            attrs: { order: '5' },
            'a:pt': [
              pointNode({ x: 10, y: 10 }, 6),
              pointNode({ x: 20, y: 20 }, 7),
            ],
          },
          'a:close': { attrs: { order: '8' } },
        },
      },
    });

    expect(getCustomShapePath(geometry, 10, 10)).toBe('z');
  });

  it('marks custom geometry stroke-only only when every path disables fill', () => {
    const paths = (fills: readonly (string | undefined)[]) =>
      xml({
        'a:pathLst': {
          'a:path': fills.map((fill) => ({
            attrs: fill === undefined ? {} : { fill },
          })),
        },
      });

    expect(isStrokeOnlyCustomGeometry(xml({}))).toBe(false);
    expect(isStrokeOnlyCustomGeometry(paths(['none']))).toBe(true);
    expect(isStrokeOnlyCustomGeometry(paths(['none', 'none']))).toBe(true);
    expect(isStrokeOnlyCustomGeometry(paths(['none', 'norm']))).toBe(false);
    expect(isStrokeOnlyCustomGeometry(paths([undefined]))).toBe(false);
  });
});
