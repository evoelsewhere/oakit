import { describe, expect, it } from 'vitest';

import { validatePptxScene } from '../../src';

function creationScene(): Record<string, unknown> {
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
              transform: {
                height: 80,
                rotation: 15,
                width: 300,
                x: 20,
                y: 30,
              },
            },
            key: 'text-1',
            resolved: { hidden: false },
            text: {
              body: {},
              paragraphs: [
                {
                  children: [
                    {
                      key: 'run-1',
                      properties: { fontSize: 18 },
                      text: 'Hello',
                      type: 'run',
                    },
                  ],
                  key: 'paragraph-1',
                },
              ],
            },
            type: 'text',
          },
        ],
        key: 'slide-1',
      },
    ],
    themes: [],
  };
}

function element(scene: Record<string, unknown>): Record<string, unknown> {
  const slides = scene.slides as Record<string, unknown>[];
  const slide = slides[0] as Record<string, unknown>;
  const elements = slide.elements as Record<string, unknown>[];
  return elements[0] as Record<string, unknown>;
}

function validateCreation(value: unknown) {
  return validatePptxScene(value, { profile: 'create-text-v1' });
}

function validateNativeCreation(value: unknown) {
  return validatePptxScene(value, { profile: 'create-native-v1' });
}

function firstRun(scene: Record<string, unknown>): Record<string, unknown> {
  const text = element(scene).text as Record<string, unknown>;
  const paragraphs = text.paragraphs as Record<string, unknown>[];
  const paragraph = paragraphs[0] as Record<string, unknown>;
  const children = paragraph.children as Record<string, unknown>[];
  return children[0] as Record<string, unknown>;
}

describe('PowerPoint creation scene validation', () => {
  it('accepts the bounded source-free text profile', () => {
    expect(validateCreation(creationScene())).toEqual({
      issues: [],
      valid: true,
    });
  });

  it('accepts bounded visual styling for rich text templates', () => {
    const scene = creationScene();
    const slides = scene.slides as Record<string, unknown>[];
    const slide = slides[0] as Record<string, unknown>;
    slide.backgroundColor = '#0F172A';
    const authored = element(scene).authored as Record<string, unknown>;
    authored.fillColor = '#1E293B';
    authored.geometry = 'roundRect';
    authored.lineColor = '#38BDF8';
    authored.lineWidth = 1.5;
    firstRun(scene).properties = {
      bold: true,
      color: '#F8FAFC',
      fontSize: 18,
    };

    expect(validateCreation(scene)).toEqual({ issues: [], valid: true });
  });

  it('accepts native shapes with bounded geometry and styling', () => {
    const scene = creationScene();
    const slide = (scene.slides as Record<string, unknown>[])[0];
    if (slide === undefined) throw new Error('Expected slide');
    slide.elements = [
      {
        authored: {
          fillColor: '#1E293B',
          geometry: 'roundRect',
          lineColor: '#38BDF8',
          lineWidth: 1.5,
          transform: { height: 120, width: 240, x: 40, y: 50 },
        },
        key: 'shape-1',
        resolved: { hidden: false },
        type: 'shape',
      },
    ];

    expect(validateNativeCreation(scene)).toEqual({ issues: [], valid: true });
    expect(validateCreation(scene).issues).toContainEqual({
      code: 'unsupported-feature',
      message: 'Creation profile create-text-v1 supports text elements only',
      path: '$.slides[0].elements[0]',
    });
  });

  it('requires a serializable authored transform for native shapes', () => {
    const scene = creationScene();
    const slide = (scene.slides as Record<string, unknown>[])[0];
    if (slide === undefined) throw new Error('Expected slide');
    slide.elements = [
      {
        authored: { geometry: 'ellipse' },
        key: 'shape-1',
        resolved: { hidden: false },
        type: 'shape',
      },
    ];

    expect(validateNativeCreation(scene).issues).toContainEqual({
      code: 'unsupported-feature',
      message:
        'Creation profile create-native-v1 requires an authored shape transform',
      path: '$.slides[0].elements[0].authored.transform',
    });
  });

  it('accepts signature-checked native image media and references', () => {
    const scene = creationScene();
    scene.media = [
      {
        data: new Uint8Array([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]),
        key: 'image-1',
        mimeType: 'image/png',
      },
    ];
    const slide = (scene.slides as Record<string, unknown>[])[0];
    if (slide === undefined) throw new Error('Expected slide');
    slide.elements = [
      {
        authored: {
          transform: { height: 100, width: 160, x: 40, y: 50 },
        },
        key: 'picture-1',
        mediaKey: 'image-1',
        resolved: { hidden: false },
        type: 'image',
      },
    ];

    expect(validateNativeCreation(scene)).toEqual({ issues: [], valid: true });
    expect(validateCreation(scene).issues).toEqual(
      expect.arrayContaining([
        expect.objectContaining({
          code: 'unsupported-feature',
          path: '$.media',
        }),
        expect.objectContaining({
          code: 'unsupported-feature',
          path: '$.slides[0].elements[0]',
        }),
      ]),
    );
  });

  it.each([
    ['image/png', new Uint8Array([0x89, 0x50]), 'PNG'],
    ['image/jpeg', new Uint8Array([0xff, 0xd8, 0xff, 0x00]), 'JPEG'],
  ])('rejects invalid %s media signatures', (mimeType, data, label) => {
    const scene = creationScene();
    scene.media = [{ data, key: 'image-1', mimeType }];

    expect(validateNativeCreation(scene).issues).toContainEqual({
      code: 'invalid-scene-document',
      message: `${label} media data has an invalid signature`,
      path: '$.media[0].data',
    });
  });

  it('rejects dangling media references and unsupported image styling', () => {
    const scene = creationScene();
    const slide = (scene.slides as Record<string, unknown>[])[0];
    if (slide === undefined) throw new Error('Expected slide');
    slide.elements = [
      {
        authored: {
          fillColor: '#FFFFFF',
          transform: { height: 100, width: 160, x: 40, y: 50 },
        },
        key: 'picture-1',
        mediaKey: 'missing-image',
        resolved: { hidden: false },
        type: 'image',
      },
    ];

    expect(validateNativeCreation(scene).issues).toEqual(
      expect.arrayContaining([
        {
          code: 'unsupported-feature',
          message:
            'Creation profile create-native-v1 does not apply shape styling to images',
          path: '$.slides[0].elements[0].authored',
        },
        {
          code: 'invalid-hierarchy-reference',
          message: 'Reference points to an unknown public key',
          path: '$.slides[0].elements[0].mediaKey',
        },
      ]),
    );
  });

  it.each(['ellipse', 'rect', 'roundRect'])(
    'accepts text shape geometry %s',
    (geometry) => {
      const scene = creationScene();
      const authored = element(scene).authored as Record<string, unknown>;
      authored.geometry = geometry;

      expect(validateCreation(scene)).toEqual({ issues: [], valid: true });
    },
  );

  it.each([
    [
      '$.slides[0].backgroundColor',
      (scene: Record<string, unknown>) => {
        const slides = scene.slides as Record<string, unknown>[];
        const slide = slides[0] as Record<string, unknown>;
        slide.backgroundColor = '0F172A';
      },
    ],
    [
      '$.slides[0].elements[0].authored.fillColor',
      (scene: Record<string, unknown>) => {
        const authored = element(scene).authored as Record<string, unknown>;
        authored.fillColor = '#XYZ123';
      },
    ],
    [
      '$.slides[0].elements[0].authored.lineColor',
      (scene: Record<string, unknown>) => {
        const authored = element(scene).authored as Record<string, unknown>;
        authored.lineColor = '#1234';
      },
    ],
    [
      '$.slides[0].elements[0].text.paragraphs[0].children[0].properties.color',
      (scene: Record<string, unknown>) => {
        firstRun(scene).properties = { color: 'red' };
      },
    ],
  ])('rejects a malformed visual color at %s', (path, mutate) => {
    const scene = creationScene();
    mutate(scene);

    expect(validateCreation(scene).issues).toContainEqual({
      code: 'invalid-scene-document',
      message: 'Expected a #RRGGBB color',
      path,
    });
  });

  it.each(['x#0F172A', '#0F172Ax'])(
    'rejects a color with data outside the #RRGGBB boundary: %s',
    (backgroundColor) => {
      const scene = creationScene();
      const slides = scene.slides as Record<string, unknown>[];
      const slide = slides[0] as Record<string, unknown>;
      slide.backgroundColor = backgroundColor;

      expect(validateCreation(scene).issues).toContainEqual({
        code: 'invalid-scene-document',
        message: 'Expected a #RRGGBB color',
        path: '$.slides[0].backgroundColor',
      });
    },
  );

  it('rejects unsupported geometry and unsafe line widths', () => {
    const scene = creationScene();
    const authored = element(scene).authored as Record<string, unknown>;
    authored.geometry = 'star';
    authored.lineWidth = 0;

    expect(validateCreation(scene).issues).toEqual(
      expect.arrayContaining([
        {
          code: 'invalid-scene-document',
          message: 'Unknown text shape geometry',
          path: '$.slides[0].elements[0].authored.geometry',
        },
        {
          code: 'invalid-numeric-value',
          message: 'Expected a positive finite number',
          path: '$.slides[0].elements[0].authored.lineWidth',
        },
      ]),
    );
  });

  it.each([
    [1_000_000_000_000, 'Value exceeds the safe OOXML integer range'],
    [0.000_039, 'Value must round to a positive OOXML integer'],
  ])(
    'applies serialized line width bounds only to creation for %s',
    (lineWidth, message) => {
      const scene = creationScene();
      const authored = element(scene).authored as Record<string, unknown>;
      authored.lineWidth = lineWidth;

      expect(validatePptxScene(scene)).toEqual({ issues: [], valid: true });
      expect(validateCreation(scene).issues).toContainEqual({
        code: 'invalid-numeric-value',
        message,
        path: '$.slides[0].elements[0].authored.lineWidth',
      });
    },
  );

  it('requires authored geometry without changing the general scene profile', () => {
    const scene = creationScene();
    element(scene).authored = {};

    expect(validatePptxScene(scene)).toEqual({ issues: [], valid: true });
    expect(validateCreation(scene)).toEqual({
      issues: [
        {
          code: 'unsupported-feature',
          message:
            'Creation profile create-text-v1 requires an authored text transform',
          path: '$.slides[0].elements[0].authored.transform',
        },
      ],
      valid: false,
    });
  });

  it('rejects explicit hierarchy until its writer mapping ships', () => {
    const scene = creationScene();
    scene.themes = [{ key: 'theme-1' }];
    scene.masters = [{ elements: [], key: 'master-1', themeKey: 'theme-1' }];
    scene.layouts = [{ elements: [], key: 'layout-1', masterKey: 'master-1' }];
    const slide = (scene.slides as Record<string, unknown>[])[0];
    if (slide) slide.layoutKey = 'layout-1';

    expect(validatePptxScene(scene)).toEqual({ issues: [], valid: true });
    expect(validateCreation(scene).issues).toContainEqual({
      code: 'unsupported-feature',
      message:
        'Creation profile create-text-v1 generates its own minimal hierarchy',
      path: '$',
    });
  });

  it('rejects placeholders until owner and inheritance serialization ships', () => {
    const scene = creationScene();
    element(scene).placeholder = { role: 'slide-instance' };

    expect(validatePptxScene(scene)).toEqual({ issues: [], valid: true });
    expect(validateCreation(scene).issues).toContainEqual({
      code: 'unsupported-feature',
      message:
        'Creation profile create-text-v1 does not support placeholders yet',
      path: '$.slides[0].elements[0].placeholder',
    });
  });

  it('does not apply text geometry requirements to preservation-only elements', () => {
    const scene = creationScene();
    const slide = (scene.slides as Record<string, unknown>[])[0];
    if (slide) {
      slide.elements = [
        {
          authored: {},
          feature: 'chart',
          key: 'chart-1',
          resolved: { hidden: false },
          type: 'unsupported',
        },
      ];
    }

    expect(validateCreation(scene)).toEqual({
      issues: [
        {
          code: 'unsupported-feature',
          message:
            'Creation profile create-text-v1 supports text elements only',
          path: '$.slides[0].elements[0]',
        },
      ],
      valid: false,
    });
  });

  it.each([
    [
      '$.size.width',
      (scene: Record<string, unknown>) => {
        scene.size = { height: 540, width: 1_000_000_000_000 };
      },
    ],
    [
      '$.size.height',
      (scene: Record<string, unknown>) => {
        scene.size = { height: 1_000_000_000_000, width: 960 };
      },
    ],
    [
      '$.slides[0].elements[0].authored.transform.x',
      (scene: Record<string, unknown>) => {
        const authored = element(scene).authored as Record<string, unknown>;
        const transform = authored.transform as Record<string, unknown>;
        transform.x = 1_000_000_000_000;
      },
    ],
    [
      '$.slides[0].elements[0].authored.transform.y',
      (scene: Record<string, unknown>) => {
        const authored = element(scene).authored as Record<string, unknown>;
        const transform = authored.transform as Record<string, unknown>;
        transform.y = -1_000_000_000_000;
      },
    ],
    [
      '$.slides[0].elements[0].authored.transform.width',
      (scene: Record<string, unknown>) => {
        const authored = element(scene).authored as Record<string, unknown>;
        const transform = authored.transform as Record<string, unknown>;
        transform.width = 1_000_000_000_000;
      },
    ],
    [
      '$.slides[0].elements[0].authored.transform.height',
      (scene: Record<string, unknown>) => {
        const authored = element(scene).authored as Record<string, unknown>;
        const transform = authored.transform as Record<string, unknown>;
        transform.height = 1_000_000_000_000;
      },
    ],
    [
      '$.slides[0].elements[0].authored.transform.rotation',
      (scene: Record<string, unknown>) => {
        const authored = element(scene).authored as Record<string, unknown>;
        const transform = authored.transform as Record<string, unknown>;
        transform.rotation = 1_000_000_000_000;
      },
    ],
    [
      '$.slides[0].elements[0].text.paragraphs[0].children[0].properties.fontSize',
      (scene: Record<string, unknown>) => {
        const text = element(scene).text as Record<string, unknown>;
        const paragraphs = text.paragraphs as Record<string, unknown>[];
        const paragraph = paragraphs[0] as Record<string, unknown>;
        const children = paragraph.children as Record<string, unknown>[];
        const run = children[0] as Record<string, unknown>;
        run.properties = { fontSize: 1_000_000_000_000_000 };
      },
    ],
  ] as const)(
    'rejects a value outside the serializable integer range at %s',
    (path, mutate) => {
      const scene = creationScene();
      mutate(scene);

      expect(validatePptxScene(scene)).toEqual({ issues: [], valid: true });
      expect(validateCreation(scene).issues).toContainEqual({
        code: 'invalid-numeric-value',
        message: 'Value exceeds the safe OOXML integer range',
        path,
      });
    },
  );

  it('checks resolved transform ranges before a future writer can depend on them', () => {
    const scene = creationScene();
    element(scene).resolved = {
      hidden: false,
      transform: {
        height: 80,
        rotation: -1_000_000_000_000,
        width: 300,
        x: 20,
        y: 30,
      },
    };

    expect(validateCreation(scene).issues).toContainEqual({
      code: 'invalid-numeric-value',
      message: 'Value exceeds the safe OOXML integer range',
      path: '$.slides[0].elements[0].resolved.transform.rotation',
    });
  });

  it('does not add duplicate range errors for non-numeric values', () => {
    const scene = creationScene();
    scene.size = { height: 540, width: Number.POSITIVE_INFINITY };

    expect(validateCreation(scene)).toEqual({
      issues: [
        {
          code: 'invalid-numeric-value',
          message: 'Expected a positive finite number',
          path: '$.size.width',
        },
      ],
      valid: false,
    });
  });

  it.each([
    [
      '$.size.width',
      (scene: Record<string, unknown>) => {
        scene.size = { height: 540, width: 0.000_039 };
      },
    ],
    [
      '$.size.height',
      (scene: Record<string, unknown>) => {
        scene.size = { height: 0.000_039, width: 960 };
      },
    ],
    [
      '$.slides[0].elements[0].authored.transform.width',
      (scene: Record<string, unknown>) => {
        const authored = element(scene).authored as Record<string, unknown>;
        const transform = authored.transform as Record<string, unknown>;
        transform.width = 0.000_039;
      },
    ],
    [
      '$.slides[0].elements[0].authored.transform.height',
      (scene: Record<string, unknown>) => {
        const authored = element(scene).authored as Record<string, unknown>;
        const transform = authored.transform as Record<string, unknown>;
        transform.height = 0.000_039;
      },
    ],
    [
      '$.slides[0].elements[0].resolved.transform.width',
      (scene: Record<string, unknown>) => {
        element(scene).resolved = {
          hidden: false,
          transform: {
            height: 80,
            width: 0.000_039,
            x: 20,
            y: 30,
          },
        };
      },
    ],
    [
      '$.slides[0].elements[0].text.paragraphs[0].children[0].properties.fontSize',
      (scene: Record<string, unknown>) => {
        const text = element(scene).text as Record<string, unknown>;
        const paragraphs = text.paragraphs as Record<string, unknown>[];
        const paragraph = paragraphs[0] as Record<string, unknown>;
        const children = paragraph.children as Record<string, unknown>[];
        const run = children[0] as Record<string, unknown>;
        run.properties = { fontSize: 0.004 };
      },
    ],
  ] as const)(
    'rejects a positive value that quantizes to zero at %s',
    (path, mutate) => {
      const scene = creationScene();
      mutate(scene);

      expect(validatePptxScene(scene)).toEqual({ issues: [], valid: true });
      expect(validateCreation(scene).issues).toContainEqual({
        code: 'invalid-numeric-value',
        message: 'Value must round to a positive OOXML integer',
        path,
      });
    },
  );

  it('accepts the smallest positive values that quantize to one', () => {
    const scene = creationScene();
    scene.size = { height: 0.000_04, width: 0.000_04 };
    const authored = element(scene).authored as Record<string, unknown>;
    const transform = authored.transform as Record<string, unknown>;
    transform.height = 0.000_04;
    transform.width = 0.000_04;
    const text = element(scene).text as Record<string, unknown>;
    const paragraphs = text.paragraphs as Record<string, unknown>[];
    const paragraph = paragraphs[0] as Record<string, unknown>;
    paragraph.endProperties = { fontSize: 0.006 };

    expect(validateCreation(scene)).toEqual({ issues: [], valid: true });
  });

  it('accepts signed positions and rotation in the creation profile', () => {
    const scene = creationScene();
    const authored = element(scene).authored as Record<string, unknown>;
    authored.transform = {
      height: 80,
      rotation: -45,
      width: 300,
      x: -20,
      y: -0.000_04,
    };

    expect(validateCreation(scene)).toEqual({ issues: [], valid: true });
  });

  it('rejects a scene beyond the bounded creation slide count', () => {
    const scene = creationScene();
    scene.slides = new Array(10_001);

    expect(validatePptxScene(scene)).toEqual({ issues: [], valid: true });
    expect(validateCreation(scene)).toEqual({
      issues: [
        {
          code: 'unsupported-feature',
          message:
            'Creation profile create-text-v1 supports at most 10000 slides',
          path: '$.slides',
        },
      ],
      valid: false,
    });
  });

  it('rejects an oversized element graph only in the creation profile', () => {
    const scene = creationScene();
    const slides = scene.slides as Record<string, unknown>[];
    const firstSlide = slides[0];
    if (firstSlide) firstSlide.elements = new Array(5_001);

    expect(validatePptxScene(scene)).toEqual({ issues: [], valid: true });
    expect(validateCreation(scene)).toEqual({
      issues: [
        {
          code: 'resource-limit-exceeded',
          message:
            'Creation profile create-text-v1 supports at most 5000 elements',
          path: '$.slides',
        },
      ],
      valid: false,
    });
  });

  it('does not traverse creation resources after structural validation fails', () => {
    const scene = creationScene();
    scene.schemaVersion = 1;
    const slides = scene.slides as Record<string, unknown>[];
    const firstSlide = slides[0];
    if (firstSlide) firstSlide.elements = new Array(5_001);

    expect(validateCreation(scene)).toEqual({
      issues: [
        {
          code: 'unsupported-schema-version',
          message: 'Only PowerPoint scene schema version 2 is supported',
          path: '$.schemaVersion',
        },
      ],
      valid: false,
    });
  });
});
