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

describe('PowerPoint creation scene validation', () => {
  it('accepts the bounded source-free text profile', () => {
    expect(validateCreation(creationScene())).toEqual({
      issues: [],
      valid: true,
    });
  });

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
});
