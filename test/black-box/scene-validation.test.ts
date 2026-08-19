import { describe, expect, it } from 'vitest';

import { validatePptxScene } from '../../src';

function minimalScene(): Record<string, unknown> {
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
              transform: { height: 80, width: 300, x: 20, y: 30 },
            },
            key: 'text-1',
            resolved: {
              hidden: false,
              transform: { height: 80, width: 300, x: 20, y: 30 },
            },
            text: {
              body: { anchor: 'top', autoFit: 'shape', wrap: true },
              paragraphs: [
                {
                  children: [
                    {
                      key: 'run-1',
                      preserveSpace: true,
                      properties: {
                        bold: true,
                        fontFamily: 'Aptos',
                        fontSize: 18,
                        italic: false,
                        language: 'en-US',
                      },
                      text: 'Hello world',
                      type: 'run',
                    },
                    { key: 'break-1', type: 'break' },
                    {
                      fieldType: 'slidenum',
                      key: 'field-1',
                      text: '1',
                      type: 'field',
                    },
                  ],
                  endProperties: { language: 'en-US' },
                  key: 'paragraph-1',
                  properties: { alignment: 'left', level: 0 },
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

function mutableSlide(scene: Record<string, unknown>): Record<string, unknown> {
  return (scene.slides as Record<string, unknown>[])[0] ?? {};
}

function mutableElement(
  scene: Record<string, unknown>,
): Record<string, unknown> {
  return (mutableSlide(scene).elements as Record<string, unknown>[])[0] ?? {};
}

function mutableTextBody(
  scene: Record<string, unknown>,
): Record<string, unknown> {
  return mutableElement(scene).text as Record<string, unknown>;
}

function mutableParagraph(
  scene: Record<string, unknown>,
): Record<string, unknown> {
  return (
    (mutableTextBody(scene).paragraphs as Record<string, unknown>[])[0] ?? {}
  );
}

function mutableTextNode(
  scene: Record<string, unknown>,
  index: number,
): Record<string, unknown> {
  return (
    (mutableParagraph(scene).children as Record<string, unknown>[])[index] ?? {}
  );
}

function convertFirstElementToGroup(
  scene: Record<string, unknown>,
): Record<string, unknown> {
  const element = mutableElement(scene);
  delete element.text;
  element.type = 'group';
  element.elements = [];
  const transform = {
    childSpace: { height: 40, width: 160, x: 0, y: 0 },
    height: 40,
    width: 160,
    x: 10,
    y: 20,
  };
  element.authored = { transform: structuredClone(transform) };
  element.resolved = {
    hidden: false,
    transform: structuredClone(transform),
  };
  return element;
}

function issuePairs(value: unknown): Array<[string, string]> {
  const result = validatePptxScene(value);
  expect(result.valid).toBe(result.issues.length === 0);
  return result.issues.map((issue) => [issue.code, issue.path]);
}

function expectSingleIssue(
  value: unknown,
  expected: { code: string; message: string; path: string },
): void {
  expect(validatePptxScene(value)).toEqual({
    issues: [expected],
    valid: false,
  });
}

describe('PowerPoint scene validation', () => {
  it('accepts the complete initial structured-text contract', () => {
    expect(validatePptxScene(minimalScene())).toEqual({
      issues: [],
      valid: true,
    });
    expect(
      validatePptxScene(minimalScene(), { profile: 'create-text-v1' }),
    ).toEqual({ issues: [], valid: true });
  });

  it('rejects an unknown ordinary transform property', () => {
    const scene = minimalScene();
    const authored = mutableElement(scene).authored as Record<string, unknown>;
    const transform = authored.transform as Record<string, unknown>;
    transform['Stryker was here'] = true;

    expect(validatePptxScene(scene).issues).toContainEqual({
      code: 'invalid-scene-document',
      message: 'Unknown property',
      path: '$.slides[0].elements[0].authored.transform.Stryker was here',
    });
  });

  it.each([
    ['flipHorizontal', 'bad'],
    ['rotation', Number.MAX_VALUE],
  ])(
    'does not treat group child-space %s as a transform field',
    (key, value) => {
      const scene = minimalScene();
      const group = convertFirstElementToGroup(scene);
      const authored = group.authored as Record<string, unknown>;
      const transform = authored.transform as Record<string, unknown>;
      const childSpace = transform.childSpace as Record<string, unknown>;
      childSpace[key] = value;
      const path = `$.slides[0].elements[0].authored.transform.childSpace.${key}`;

      expect(
        validatePptxScene(scene, { profile: 'create-native-v1' }).issues.filter(
          (issue) => issue.path === path,
        ),
      ).toEqual([
        {
          code: 'invalid-scene-document',
          message: 'Unknown property',
          path,
        },
      ]);
    },
  );

  it('does not mutate caller-owned scene data', () => {
    const scene = minimalScene();
    const before = structuredClone(scene);

    validatePptxScene(scene, { profile: 'create-text-v1' });

    expect(scene).toEqual(before);
  });

  it('rejects non-object input and unknown root properties', () => {
    expect(issuePairs(null)).toEqual([['invalid-scene-document', '$']]);
    expect(issuePairs([])).toEqual([['invalid-scene-document', '$']]);

    const scene = minimalScene();
    scene.rawXml = '<p:sld/>';
    expect(issuePairs(scene)).toContainEqual([
      'invalid-scene-document',
      '$.rawXml',
    ]);
  });

  it('requires schema version 2 and every root collection', () => {
    const scene = minimalScene();
    scene.schemaVersion = 1;
    delete scene.media;

    expect(issuePairs(scene)).toEqual([
      ['unsupported-schema-version', '$.schemaVersion'],
      ['invalid-scene-document', '$.media'],
    ]);
  });

  it.each([0, -1, Number.NaN, Number.POSITIVE_INFINITY, '960'])(
    'rejects invalid positive slide dimensions: %s',
    (width) => {
      const scene = minimalScene();
      scene.size = { height: 540, width };
      expect(issuePairs(scene)).toContainEqual([
        'invalid-numeric-value',
        '$.size.width',
      ]);
    },
  );

  it('rejects invalid and duplicate public keys across the graph', () => {
    const scene = minimalScene();
    mutableSlide(scene).key = 'unsafe key';
    const element = mutableElement(scene);
    const text = element.text as Record<string, unknown>;
    const paragraph = (text.paragraphs as Record<string, unknown>[])[0] ?? {};
    const run = (paragraph.children as Record<string, unknown>[])[0] ?? {};
    run.key = 'text-1';

    expect(issuePairs(scene)).toEqual(
      expect.arrayContaining([
        ['invalid-scene-document', '$.slides[0].key'],
        [
          'duplicate-public-key',
          '$.slides[0].elements[0].text.paragraphs[0].children[0].key',
        ],
      ]),
    );
  });

  it('accepts a complete explicit hierarchy and placeholder source', () => {
    const scene = minimalScene();
    scene.themes = [{ key: 'theme-1' }];
    scene.masters = [
      {
        elements: [
          {
            authored: {},
            feature: 'decorative-shape',
            key: 'master-element-1',
            resolved: { hidden: false },
            type: 'unsupported',
          },
        ],
        key: 'master-1',
        themeKey: 'theme-1',
      },
    ];
    scene.layouts = [{ elements: [], key: 'layout-1', masterKey: 'master-1' }];
    mutableSlide(scene).layoutKey = 'layout-1';
    mutableElement(scene).placeholder = {
      hasCustomPrompt: false,
      index: 0,
      orientation: 'horizontal',
      role: 'slide-instance',
      size: 'full',
      sourceKey: 'master-element-1',
      type: 'title',
    };

    expect(validatePptxScene(scene)).toEqual({ issues: [], valid: true });
  });

  it('rejects partial and dangling hierarchy references', () => {
    const scene = minimalScene();
    scene.themes = [{ key: 'theme-1' }];
    mutableSlide(scene).layoutKey = 'missing-layout';
    mutableElement(scene).placeholder = {
      role: 'slide-instance',
      sourceKey: 'missing-placeholder',
    };

    expect(issuePairs(scene)).toEqual(
      expect.arrayContaining([
        ['invalid-hierarchy-reference', '$'],
        ['invalid-hierarchy-reference', '$.slides[0].layoutKey'],
        [
          'invalid-hierarchy-reference',
          '$.slides[0].elements[0].placeholder.sourceKey',
        ],
      ]),
    );
  });

  it('forbids naming a layout when requesting the generated hierarchy', () => {
    const scene = minimalScene();
    mutableSlide(scene).layoutKey = 'layout-default';
    expect(issuePairs(scene)).toContainEqual([
      'invalid-hierarchy-reference',
      '$.slides[0].layoutKey',
    ]);
  });

  it('keeps preservation-only elements out of the creation profile', () => {
    const scene = minimalScene();
    mutableSlide(scene).elements = [
      {
        authored: {},
        feature: 'chart',
        key: 'chart-1',
        previewText: 'Quarterly revenue',
        resolved: { hidden: false },
        type: 'unsupported',
      },
    ];

    expect(validatePptxScene(scene)).toEqual({ issues: [], valid: true });
    expect(
      validatePptxScene(scene, { profile: 'create-text-v1' }).issues,
    ).toContainEqual(
      expect.objectContaining({
        code: 'unsupported-feature',
        path: '$.slides[0].elements[0]',
      }),
    );
  });

  it('accepts represented media and keeps it outside the text creation profile', () => {
    const scene = minimalScene();
    scene.media = [
      {
        data: new Uint8Array([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]),
        key: 'image-1',
        mimeType: 'image/png',
      },
    ];

    expect(issuePairs(scene)).toEqual([]);
    expect(
      validatePptxScene(scene, { profile: 'create-text-v1' }).issues.map(
        (issue) => [issue.code, issue.path],
      ),
    ).toContainEqual(['unsupported-feature', '$.media']);
  });

  it('allows preservation-only images to omit creation media ownership', () => {
    const scene = minimalScene();
    mutableSlide(scene).elements = [
      {
        authored: { fillColor: '#FFFFFF' },
        key: 'preserved-image',
        resolved: { hidden: false },
        type: 'image',
      },
    ];

    expect(validatePptxScene(scene)).toEqual({ issues: [], valid: true });
    expect(
      validatePptxScene(scene, { profile: 'create-native-v1' }).issues,
    ).toEqual(
      expect.arrayContaining([
        expect.objectContaining({
          path: '$.slides[0].elements[0].mediaKey',
        }),
        expect.objectContaining({
          path: '$.slides[0].elements[0].authored',
        }),
      ]),
    );
  });

  it('rejects unknown element fields, kinds, and malformed base state', () => {
    const scene = minimalScene();
    const element = mutableElement(scene);
    element.type = 'widget';
    element.raw = {};
    element.authored = { hidden: 'false' };
    element.resolved = { hidden: 0 };

    expect(issuePairs(scene)).toEqual(
      expect.arrayContaining([
        ['invalid-scene-document', '$.slides[0].elements[0].type'],
        ['invalid-scene-document', '$.slides[0].elements[0].authored.hidden'],
        ['invalid-scene-document', '$.slides[0].elements[0].resolved.hidden'],
      ]),
    );
  });

  it.each([
    ['x', Number.NaN],
    ['y', Number.NEGATIVE_INFINITY],
    ['width', 0],
    ['height', -1],
    ['rotation', '90'],
  ])('rejects invalid transform %s values', (field, value) => {
    const scene = minimalScene();
    const element = mutableElement(scene);
    const authored = element.authored as Record<string, unknown>;
    const transform = authored.transform as Record<string, unknown>;
    transform[field] = value;

    expect(issuePairs(scene)).toContainEqual([
      'invalid-numeric-value',
      `$.slides[0].elements[0].authored.transform.${field}`,
    ]);
  });

  it('rejects empty paragraphs and unsafe XML text characters', () => {
    const scene = minimalScene();
    const element = mutableElement(scene);
    const text = element.text as Record<string, unknown>;
    text.paragraphs = [];
    expect(issuePairs(scene)).toContainEqual([
      'invalid-scene-document',
      '$.slides[0].elements[0].text.paragraphs',
    ]);

    const unsafe = minimalScene();
    const unsafeText = mutableElement(unsafe).text as Record<string, unknown>;
    const paragraph =
      (unsafeText.paragraphs as Record<string, unknown>[])[0] ?? {};
    const run = (paragraph.children as Record<string, unknown>[])[0] ?? {};
    run.text = 'before\u0000after';
    expect(issuePairs(unsafe)).toContainEqual([
      'invalid-office-text-escape',
      '$.slides[0].elements[0].text.paragraphs[0].children[0].text',
    ]);
  });

  it('rejects malformed text discriminants and formatting boundaries', () => {
    const scene = minimalScene();
    const text = mutableElement(scene).text as Record<string, unknown>;
    const paragraph = (text.paragraphs as Record<string, unknown>[])[0] ?? {};
    paragraph.properties = { alignment: 'middle', level: 9 };
    const children = paragraph.children as Record<string, unknown>[];
    (children[0] ?? {}).type = 'tab';
    (children[1] ?? {}).properties = { fontSize: 0 };
    (children[2] ?? {}).fieldType = '';

    expect(issuePairs(scene)).toEqual(
      expect.arrayContaining([
        [
          'invalid-scene-document',
          '$.slides[0].elements[0].text.paragraphs[0].children[0].type',
        ],
        [
          'invalid-numeric-value',
          '$.slides[0].elements[0].text.paragraphs[0].children[1].properties.fontSize',
        ],
        [
          'invalid-scene-document',
          '$.slides[0].elements[0].text.paragraphs[0].children[2].fieldType',
        ],
        [
          'invalid-scene-document',
          '$.slides[0].elements[0].text.paragraphs[0].properties.alignment',
        ],
        [
          'invalid-numeric-value',
          '$.slides[0].elements[0].text.paragraphs[0].properties.level',
        ],
      ]),
    );
  });

  it('returns stable codes, paths, and messages for primitive failures', () => {
    expectSingleIssue(null, {
      code: 'invalid-scene-document',
      message: 'Expected an object',
      path: '$',
    });

    const unknown = minimalScene();
    unknown.rawXml = true;
    expectSingleIssue(unknown, {
      code: 'invalid-scene-document',
      message: 'Unknown property',
      path: '$.rawXml',
    });

    const invalidSize = minimalScene();
    invalidSize.size = { height: 540, width: 0 };
    expectSingleIssue(invalidSize, {
      code: 'invalid-numeric-value',
      message: 'Expected a positive finite number',
      path: '$.size.width',
    });

    const invalidCoordinate = minimalScene();
    const authored = mutableElement(invalidCoordinate).authored as Record<
      string,
      unknown
    >;
    const transform = authored.transform as Record<string, unknown>;
    transform.x = Number.NaN;
    expectSingleIssue(invalidCoordinate, {
      code: 'invalid-numeric-value',
      message: 'Expected a finite number',
      path: '$.slides[0].elements[0].authored.transform.x',
    });
  });

  it.each(['themes', 'masters', 'layouts', 'slides', 'media'])(
    'requires root %s to be an array',
    (property) => {
      const scene = minimalScene();
      scene[property] = {};
      expect(validatePptxScene(scene).issues).toEqual([
        {
          code: 'invalid-scene-document',
          message: 'Expected an array',
          path: `$.${property}`,
        },
      ]);
    },
  );

  it.each([
    ['slides', '$.slides[0]'],
    ['themes', '$.themes[0]'],
    ['masters', '$.masters[0]'],
    ['layouts', '$.layouts[0]'],
  ])('requires every %s item to be an object', (collection, path) => {
    const scene = minimalScene();
    scene[collection] = [null];
    expect(validatePptxScene(scene).issues).toContainEqual({
      code: 'invalid-scene-document',
      message: 'Expected an object',
      path,
    });
  });

  it('requires nested scene collections and objects at their exact paths', () => {
    const cases: Array<{
      expectedPath: string;
      mutate: (scene: Record<string, unknown>) => void;
    }> = [
      {
        expectedPath: '$.size',
        mutate: (scene) => {
          scene.size = [];
        },
      },
      {
        expectedPath: '$.slides[0].elements',
        mutate: (scene) => {
          mutableSlide(scene).elements = {};
        },
      },
      {
        expectedPath: '$.slides[0].elements[0]',
        mutate: (scene) => {
          mutableSlide(scene).elements = [null];
        },
      },
      {
        expectedPath: '$.slides[0].elements[0].authored',
        mutate: (scene) => {
          mutableElement(scene).authored = [];
        },
      },
      {
        expectedPath: '$.slides[0].elements[0].resolved',
        mutate: (scene) => {
          mutableElement(scene).resolved = null;
        },
      },
      {
        expectedPath: '$.slides[0].elements[0].authored.transform',
        mutate: (scene) => {
          mutableElement(scene).authored = { transform: [] };
        },
      },
      {
        expectedPath: '$.slides[0].elements[0].text',
        mutate: (scene) => {
          mutableElement(scene).text = [];
        },
      },
      {
        expectedPath: '$.slides[0].elements[0].text.body',
        mutate: (scene) => {
          mutableTextBody(scene).body = [];
        },
      },
      {
        expectedPath: '$.slides[0].elements[0].text.paragraphs',
        mutate: (scene) => {
          mutableTextBody(scene).paragraphs = {};
        },
      },
      {
        expectedPath: '$.slides[0].elements[0].text.paragraphs[0]',
        mutate: (scene) => {
          mutableTextBody(scene).paragraphs = [null];
        },
      },
      {
        expectedPath: '$.slides[0].elements[0].text.paragraphs[0].children',
        mutate: (scene) => {
          mutableParagraph(scene).children = {};
        },
      },
      {
        expectedPath: '$.slides[0].elements[0].text.paragraphs[0].children[0]',
        mutate: (scene) => {
          mutableParagraph(scene).children = [null];
        },
      },
      {
        expectedPath:
          '$.slides[0].elements[0].text.paragraphs[0].endProperties',
        mutate: (scene) => {
          mutableParagraph(scene).endProperties = [];
        },
      },
      {
        expectedPath: '$.slides[0].elements[0].text.paragraphs[0].properties',
        mutate: (scene) => {
          mutableParagraph(scene).properties = [];
        },
      },
      {
        expectedPath:
          '$.slides[0].elements[0].text.paragraphs[0].children[0].properties',
        mutate: (scene) => {
          mutableTextNode(scene, 0).properties = [];
        },
      },
      {
        expectedPath: '$.slides[0].elements[0].placeholder',
        mutate: (scene) => {
          mutableElement(scene).placeholder = [];
        },
      },
    ];

    for (const { expectedPath, mutate } of cases) {
      const scene = minimalScene();
      mutate(scene);
      expect(validatePptxScene(scene).issues).toContainEqual(
        expect.objectContaining({
          code: 'invalid-scene-document',
          path: expectedPath,
        }),
      );
    }
  });

  it('accepts every currently declared optional property', () => {
    const scene = minimalScene();
    const slide = mutableSlide(scene);
    slide.hidden = false;
    slide.name = 'Intro';
    const element = mutableElement(scene);
    element.description = 'Description';
    element.name = 'Title 1';
    element.title = 'Accessible title';
    element.placeholder = {
      hasCustomPrompt: true,
      index: 1,
      orientation: 'vertical',
      prompt: 'Click to edit',
      role: 'slide-instance',
      size: 'quarter',
      type: 'title',
    };
    const authored = element.authored as Record<string, unknown>;
    authored.hidden = false;
    const authoredTransform = authored.transform as Record<string, unknown>;
    authoredTransform.flipHorizontal = false;
    authoredTransform.flipVertical = true;
    authoredTransform.rotation = 0;
    const resolved = element.resolved as Record<string, unknown>;
    const resolvedTransform = resolved.transform as Record<string, unknown>;
    resolvedTransform.flipHorizontal = true;
    resolvedTransform.flipVertical = false;
    resolvedTransform.rotation = 45;
    const body = mutableTextBody(scene).body as Record<string, unknown>;
    body.vertical = false;
    const field = mutableTextNode(scene, 2);
    field.properties = { bold: false };
    const lineBreak = mutableTextNode(scene, 1);
    lineBreak.properties = { italic: true };

    expect(validatePptxScene(scene)).toEqual({ issues: [], valid: true });
  });

  it.each([
    [
      'slide hidden',
      '$.slides[0].hidden',
      (scene: Record<string, unknown>) => {
        mutableSlide(scene).hidden = 'false';
      },
    ],
    [
      'slide name',
      '$.slides[0].name',
      (scene: Record<string, unknown>) => {
        mutableSlide(scene).name = false;
      },
    ],
    [
      'description',
      '$.slides[0].elements[0].description',
      (scene: Record<string, unknown>) => {
        mutableElement(scene).description = false;
      },
    ],
    [
      'element name',
      '$.slides[0].elements[0].name',
      (scene: Record<string, unknown>) => {
        mutableElement(scene).name = false;
      },
    ],
    [
      'title',
      '$.slides[0].elements[0].title',
      (scene: Record<string, unknown>) => {
        mutableElement(scene).title = false;
      },
    ],
    [
      'authored hidden',
      '$.slides[0].elements[0].authored.hidden',
      (scene: Record<string, unknown>) => {
        const authored = mutableElement(scene).authored as Record<
          string,
          unknown
        >;
        authored.hidden = 'false';
      },
    ],
    [
      'flip horizontal',
      '$.slides[0].elements[0].authored.transform.flipHorizontal',
      (scene: Record<string, unknown>) => {
        const authored = mutableElement(scene).authored as Record<
          string,
          unknown
        >;
        const transform = authored.transform as Record<string, unknown>;
        transform.flipHorizontal = 0;
      },
    ],
    [
      'flip vertical',
      '$.slides[0].elements[0].authored.transform.flipVertical',
      (scene: Record<string, unknown>) => {
        const authored = mutableElement(scene).authored as Record<
          string,
          unknown
        >;
        const transform = authored.transform as Record<string, unknown>;
        transform.flipVertical = 0;
      },
    ],
    [
      'text vertical',
      '$.slides[0].elements[0].text.body.vertical',
      (scene: Record<string, unknown>) => {
        const body = mutableTextBody(scene).body as Record<string, unknown>;
        body.vertical = 0;
      },
    ],
    [
      'text wrap',
      '$.slides[0].elements[0].text.body.wrap',
      (scene: Record<string, unknown>) => {
        const body = mutableTextBody(scene).body as Record<string, unknown>;
        body.wrap = 0;
      },
    ],
    [
      'preserve space',
      '$.slides[0].elements[0].text.paragraphs[0].children[0].preserveSpace',
      (scene: Record<string, unknown>) => {
        mutableTextNode(scene, 0).preserveSpace = 0;
      },
    ],
    [
      'bold',
      '$.slides[0].elements[0].text.paragraphs[0].children[0].properties.bold',
      (scene: Record<string, unknown>) => {
        const properties = mutableTextNode(scene, 0).properties as Record<
          string,
          unknown
        >;
        properties.bold = 0;
      },
    ],
    [
      'italic',
      '$.slides[0].elements[0].text.paragraphs[0].children[0].properties.italic',
      (scene: Record<string, unknown>) => {
        const properties = mutableTextNode(scene, 0).properties as Record<
          string,
          unknown
        >;
        properties.italic = 0;
      },
    ],
    [
      'font family',
      '$.slides[0].elements[0].text.paragraphs[0].children[0].properties.fontFamily',
      (scene: Record<string, unknown>) => {
        const properties = mutableTextNode(scene, 0).properties as Record<
          string,
          unknown
        >;
        properties.fontFamily = 0;
      },
    ],
    [
      'language',
      '$.slides[0].elements[0].text.paragraphs[0].children[0].properties.language',
      (scene: Record<string, unknown>) => {
        const properties = mutableTextNode(scene, 0).properties as Record<
          string,
          unknown
        >;
        properties.language = 0;
      },
    ],
    [
      'custom prompt',
      '$.slides[0].elements[0].placeholder.hasCustomPrompt',
      (scene: Record<string, unknown>) => {
        mutableElement(scene).placeholder = {
          hasCustomPrompt: 0,
          role: 'slide-instance',
        };
      },
    ],
    [
      'prompt',
      '$.slides[0].elements[0].placeholder.prompt',
      (scene: Record<string, unknown>) => {
        mutableElement(scene).placeholder = {
          prompt: 0,
          role: 'slide-instance',
        };
      },
    ],
    [
      'placeholder type',
      '$.slides[0].elements[0].placeholder.type',
      (scene: Record<string, unknown>) => {
        mutableElement(scene).placeholder = { role: 'slide-instance', type: 0 };
      },
    ],
  ] as const)('rejects an invalid optional %s value', (_name, path, mutate) => {
    const scene = minimalScene();
    mutate(scene);
    expect(validatePptxScene(scene).issues).toContainEqual({
      code: 'invalid-scene-document',
      message:
        path.endsWith('hidden') ||
        path.includes('flip') ||
        path.endsWith('vertical') ||
        path.endsWith('wrap') ||
        path.endsWith('preserveSpace') ||
        path.endsWith('bold') ||
        path.endsWith('italic') ||
        path.endsWith('hasCustomPrompt')
          ? 'Expected a boolean'
          : 'Expected a string',
      path,
    });
  });

  it.each(['bottom', 'center', 'distributed', 'justified', 'top'])(
    'accepts text anchor %s',
    (anchor) => {
      const scene = minimalScene();
      const body = mutableTextBody(scene).body as Record<string, unknown>;
      body.anchor = anchor;
      expect(validatePptxScene(scene).valid).toBe(true);
    },
  );

  it.each(['none', 'shape', 'text'])('accepts text auto-fit %s', (autoFit) => {
    const scene = minimalScene();
    const body = mutableTextBody(scene).body as Record<string, unknown>;
    body.autoFit = autoFit;
    expect(validatePptxScene(scene).valid).toBe(true);
  });

  it.each(['center', 'distributed', 'justify', 'left', 'right'])(
    'accepts paragraph alignment %s',
    (alignment) => {
      const scene = minimalScene();
      mutableParagraph(scene).properties = { alignment, level: 8 };
      expect(validatePptxScene(scene).valid).toBe(true);
    },
  );

  it.each(['layout-definition', 'master-definition', 'slide-instance'])(
    'accepts placeholder role %s',
    (role) => {
      const scene = minimalScene();
      mutableElement(scene).placeholder = { role };
      expect(validatePptxScene(scene).valid).toBe(true);
    },
  );

  it.each(['horizontal', 'vertical'])(
    'accepts placeholder orientation %s',
    (orientation) => {
      const scene = minimalScene();
      mutableElement(scene).placeholder = {
        orientation,
        role: 'slide-instance',
      };
      expect(validatePptxScene(scene).valid).toBe(true);
    },
  );

  it.each(['full', 'half', 'quarter'])(
    'accepts placeholder size %s',
    (size) => {
      const scene = minimalScene();
      mutableElement(scene).placeholder = { role: 'slide-instance', size };
      expect(validatePptxScene(scene).valid).toBe(true);
    },
  );

  it('rejects invalid text-body and placeholder enum values exactly', () => {
    const cases: Array<{
      expected: { code: string; message: string; path: string };
      mutate: (scene: Record<string, unknown>) => void;
    }> = [
      {
        expected: {
          code: 'invalid-scene-document',
          message: 'Unknown text anchor',
          path: '$.slides[0].elements[0].text.body.anchor',
        },
        mutate: (scene) => {
          const body = mutableTextBody(scene).body as Record<string, unknown>;
          body.anchor = 'middle';
        },
      },
      {
        expected: {
          code: 'invalid-scene-document',
          message: 'Unknown text auto-fit mode',
          path: '$.slides[0].elements[0].text.body.autoFit',
        },
        mutate: (scene) => {
          const body = mutableTextBody(scene).body as Record<string, unknown>;
          body.autoFit = 'grow';
        },
      },
      {
        expected: {
          code: 'invalid-scene-document',
          message: 'Unknown placeholder role',
          path: '$.slides[0].elements[0].placeholder.role',
        },
        mutate: (scene) => {
          mutableElement(scene).placeholder = { role: 'unknown' };
        },
      },
      {
        expected: {
          code: 'invalid-scene-document',
          message: 'Unknown placeholder orientation',
          path: '$.slides[0].elements[0].placeholder.orientation',
        },
        mutate: (scene) => {
          mutableElement(scene).placeholder = {
            orientation: 'diagonal',
            role: 'slide-instance',
          };
        },
      },
      {
        expected: {
          code: 'invalid-scene-document',
          message: 'Unknown placeholder size',
          path: '$.slides[0].elements[0].placeholder.size',
        },
        mutate: (scene) => {
          mutableElement(scene).placeholder = {
            role: 'slide-instance',
            size: 'large',
          };
        },
      },
    ];

    for (const { expected, mutate } of cases) {
      const scene = minimalScene();
      mutate(scene);
      expectSingleIssue(scene, expected);
    }
  });

  it.each([
    ['\u0000', false],
    ['\u0001', false],
    ['\u0008', false],
    ['\u0009', true],
    ['\u000A', true],
    ['\u000D', true],
    ['\u000E', false],
    ['\u001F', false],
    ['\u0020', true],
    ['\uD7FF', true],
    ['\uD800', false],
    ['\uDBFF', false],
    ['\uDC00', false],
    ['\uDFFF', false],
    ['\uD800A', false],
    ['\uD800\uDC00', true],
    ['\uE000', true],
    ['\uFFFE', false],
    ['\uFFFF', false],
    ['😀', true],
  ])('classifies XML text boundary %j as valid=%s', (text, valid) => {
    const scene = minimalScene();
    mutableTextNode(scene, 0).text = text;
    const result = validatePptxScene(scene);
    expect(result.valid).toBe(valid);
    expect(result.issues.map((issue) => issue.code)).toEqual(
      valid ? [] : ['invalid-office-text-escape'],
    );
  });

  it('distinguishes non-string text, invalid text, and field identity errors', () => {
    const nonString = minimalScene();
    mutableTextNode(nonString, 0).text = 7;
    expect(validatePptxScene(nonString).issues).toContainEqual({
      code: 'invalid-scene-document',
      message: 'Expected text content',
      path: '$.slides[0].elements[0].text.paragraphs[0].children[0].text',
    });

    const unsafe = minimalScene();
    mutableTextNode(unsafe, 0).text = '\uFFFF';
    expect(validatePptxScene(unsafe).issues).toContainEqual({
      code: 'invalid-office-text-escape',
      message: 'Text contains a character that cannot be serialized safely',
      path: '$.slides[0].elements[0].text.paragraphs[0].children[0].text',
    });

    const field = minimalScene();
    mutableTextNode(field, 2).fieldType = '   ';
    expect(validatePptxScene(field).issues).toContainEqual({
      code: 'invalid-scene-document',
      message: 'Expected a non-empty field type',
      path: '$.slides[0].elements[0].text.paragraphs[0].children[2].fieldType',
    });
  });

  it.each([-1, 1.5, Number.NaN])(
    'rejects placeholder index boundary %s',
    (index) => {
      const scene = minimalScene();
      mutableElement(scene).placeholder = { index, role: 'slide-instance' };
      expect(validatePptxScene(scene).issues).toContainEqual({
        code: 'invalid-numeric-value',
        message: 'Placeholder index must be a non-negative integer',
        path: '$.slides[0].elements[0].placeholder.index',
      });
    },
  );

  it.each([-1, 1.5, 9])('rejects paragraph level boundary %s', (level) => {
    const scene = minimalScene();
    mutableParagraph(scene).properties = { level };
    expect(validatePptxScene(scene).issues).toContainEqual({
      code: 'invalid-numeric-value',
      message: 'Paragraph level must be an integer from 0 through 8',
      path: '$.slides[0].elements[0].text.paragraphs[0].properties.level',
    });
  });

  it('validates hierarchy references and reference types independently', () => {
    const complete = minimalScene();
    complete.themes = [{ key: 'theme-1', name: 'Office' }];
    complete.masters = [
      {
        elements: [],
        key: 'master-1',
        name: 'Master',
        themeKey: 'theme-1',
      },
    ];
    complete.layouts = [
      {
        elements: [],
        key: 'layout-1',
        masterKey: 'master-1',
        name: 'Title',
      },
    ];
    mutableSlide(complete).layoutKey = 'layout-1';
    expect(validatePptxScene(complete)).toEqual({ issues: [], valid: true });

    const unknownTheme = structuredClone(complete);
    const unknownThemeMaster =
      (unknownTheme.masters as Record<string, unknown>[])[0] ?? {};
    unknownThemeMaster.themeKey = 'missing';
    expect(validatePptxScene(unknownTheme).issues).toContainEqual({
      code: 'invalid-hierarchy-reference',
      message: 'Master references an unknown theme',
      path: '$.masters[0].themeKey',
    });

    const unknownMaster = structuredClone(complete);
    const unknownMasterLayout =
      (unknownMaster.layouts as Record<string, unknown>[])[0] ?? {};
    unknownMasterLayout.masterKey = 'missing';
    expect(validatePptxScene(unknownMaster).issues).toContainEqual({
      code: 'invalid-hierarchy-reference',
      message: 'Layout references an unknown master',
      path: '$.layouts[0].masterKey',
    });

    const invalidThemeType = structuredClone(complete);
    const invalidThemeMaster =
      (invalidThemeType.masters as Record<string, unknown>[])[0] ?? {};
    invalidThemeMaster.themeKey = 7;
    expect(validatePptxScene(invalidThemeType).issues).toContainEqual({
      code: 'invalid-scene-document',
      message: 'Expected a theme key',
      path: '$.masters[0].themeKey',
    });

    const invalidMasterType = structuredClone(complete);
    const invalidMasterLayout =
      (invalidMasterType.layouts as Record<string, unknown>[])[0] ?? {};
    invalidMasterLayout.masterKey = 7;
    expect(validatePptxScene(invalidMasterType).issues).toContainEqual({
      code: 'invalid-scene-document',
      message: 'Expected a master key',
      path: '$.layouts[0].masterKey',
    });

    const invalidLayoutType = structuredClone(complete);
    mutableSlide(invalidLayoutType).layoutKey = 7;
    expect(validatePptxScene(invalidLayoutType).issues).toContainEqual({
      code: 'invalid-scene-document',
      message: 'Expected a layout key',
      path: '$.slides[0].layoutKey',
    });
  });

  it.each([
    ['themes', [{ key: 'theme-only' }]],
    ['masters', [{ elements: [], key: 'master-only', themeKey: 'missing' }]],
    ['layouts', [{ elements: [], key: 'layout-only', masterKey: 'missing' }]],
  ])('rejects a hierarchy containing only %s', (property, value) => {
    const scene = minimalScene();
    scene[property] = value;
    expect(validatePptxScene(scene).issues).toContainEqual({
      code: 'invalid-hierarchy-reference',
      message: 'A declared hierarchy needs themes, masters, and layouts',
      path: '$',
    });
  });

  it('returns stable hierarchy and capability messages', () => {
    const namedGeneratedLayout = minimalScene();
    mutableSlide(namedGeneratedLayout).layoutKey = 'layout-default';
    expect(validatePptxScene(namedGeneratedLayout).issues).toContainEqual({
      code: 'invalid-hierarchy-reference',
      message: 'A generated minimal hierarchy must not name a layout',
      path: '$.slides[0].layoutKey',
    });

    const unsupported = minimalScene();
    mutableSlide(unsupported).elements = [
      {
        authored: {},
        feature: 'chart',
        key: 'chart-1',
        resolved: { hidden: false },
        type: 'unsupported',
      },
    ];
    expect(
      validatePptxScene(unsupported, { profile: 'create-text-v1' }).issues,
    ).toContainEqual({
      code: 'unsupported-feature',
      message: 'Creation profile create-text-v1 supports text elements only',
      path: '$.slides[0].elements[0]',
    });

    const media = minimalScene();
    media.media = [{}];
    expect(validatePptxScene(media).issues).toEqual(
      expect.arrayContaining([
        expect.objectContaining({ path: '$.media[0].data' }),
        expect.objectContaining({ path: '$.media[0].key' }),
      ]),
    );
  });

  it.each([undefined, [], new Date(0), new Uint8Array([1])])(
    'rejects non-plain object roots',
    (value) => {
      expectSingleIssue(value, {
        code: 'invalid-scene-document',
        message: 'Expected an object',
        path: '$',
      });
    },
  );

  it('validates both size dimensions and permits negative positions', () => {
    const invalid = minimalScene();
    invalid.size = { height: -1, width: 960 };
    expectSingleIssue(invalid, {
      code: 'invalid-numeric-value',
      message: 'Expected a positive finite number',
      path: '$.size.height',
    });

    const valid = minimalScene();
    const authored = mutableElement(valid).authored as Record<string, unknown>;
    const transform = authored.transform as Record<string, unknown>;
    transform.x = -20;
    transform.y = -30;
    expect(validatePptxScene(valid)).toEqual({ issues: [], valid: true });
  });

  it('returns exact key diagnostics for invalid and duplicate identities', () => {
    const invalid = minimalScene();
    mutableElement(invalid).key = '';
    expect(validatePptxScene(invalid).issues).toContainEqual({
      code: 'invalid-scene-document',
      message: 'Expected a non-empty portable key of at most 128 characters',
      path: '$.slides[0].elements[0].key',
    });

    const duplicate = minimalScene();
    mutableTextNode(duplicate, 0).key = 'text-1';
    expect(validatePptxScene(duplicate).issues).toContainEqual({
      code: 'duplicate-public-key',
      message: 'Duplicate public key: text-1',
      path: '$.slides[0].elements[0].text.paragraphs[0].children[0].key',
    });

    const paragraph = minimalScene();
    mutableParagraph(paragraph).key = '';
    expect(validatePptxScene(paragraph).issues).toContainEqual(
      expect.objectContaining({
        path: '$.slides[0].elements[0].text.paragraphs[0].key',
      }),
    );
  });

  it('validates field text and every text-node branch independently', () => {
    const fieldText = minimalScene();
    mutableTextNode(fieldText, 2).text = 1;
    expect(validatePptxScene(fieldText).issues).toContainEqual({
      code: 'invalid-scene-document',
      message: 'Expected text content',
      path: '$.slides[0].elements[0].text.paragraphs[0].children[2].text',
    });

    const lineBreak = minimalScene();
    mutableTextNode(lineBreak, 1).unexpected = true;
    expect(validatePptxScene(lineBreak).issues).toContainEqual({
      code: 'invalid-scene-document',
      message: 'Unknown property',
      path: '$.slides[0].elements[0].text.paragraphs[0].children[1].unexpected',
    });

    const unknown = minimalScene();
    mutableTextNode(unknown, 0).type = 'tab';
    expect(validatePptxScene(unknown).issues).toContainEqual({
      code: 'invalid-scene-document',
      message: 'Unknown text node type',
      path: '$.slides[0].elements[0].text.paragraphs[0].children[0].type',
    });
  });

  it('accepts omitted optional paragraph and text-body properties', () => {
    const scene = minimalScene();
    const paragraph = mutableParagraph(scene);
    delete paragraph.endProperties;
    delete paragraph.properties;
    const body = mutableTextBody(scene).body as Record<string, unknown>;
    delete body.anchor;
    delete body.autoFit;

    expect(validatePptxScene(scene)).toEqual({ issues: [], valid: true });

    const withoutAlignment = minimalScene();
    mutableParagraph(withoutAlignment).properties = { level: 0 };
    expect(validatePptxScene(withoutAlignment).valid).toBe(true);

    const withoutLevel = minimalScene();
    mutableParagraph(withoutLevel).properties = { alignment: 'left' };
    expect(validatePptxScene(withoutLevel).valid).toBe(true);
  });

  it('returns exact paragraph and text-body diagnostics', () => {
    const alignment = minimalScene();
    mutableParagraph(alignment).properties = { alignment: 'middle' };
    expectSingleIssue(alignment, {
      code: 'invalid-scene-document',
      message: 'Unknown paragraph alignment',
      path: '$.slides[0].elements[0].text.paragraphs[0].properties.alignment',
    });

    const paragraphProperties = minimalScene();
    mutableParagraph(paragraphProperties).properties = {
      alignment: 'left',
      unexpected: true,
    };
    expect(validatePptxScene(paragraphProperties).issues).toContainEqual({
      code: 'invalid-scene-document',
      message: 'Unknown property',
      path: '$.slides[0].elements[0].text.paragraphs[0].properties.unexpected',
    });

    const body = minimalScene();
    const bodyProperties = mutableTextBody(body).body as Record<
      string,
      unknown
    >;
    bodyProperties.unexpected = true;
    expect(validatePptxScene(body).issues).toContainEqual({
      code: 'invalid-scene-document',
      message: 'Unknown property',
      path: '$.slides[0].elements[0].text.body.unexpected',
    });

    const empty = minimalScene();
    mutableTextBody(empty).paragraphs = [];
    expectSingleIssue(empty, {
      code: 'invalid-scene-document',
      message: 'A text body needs at least one paragraph',
      path: '$.slides[0].elements[0].text.paragraphs',
    });
  });

  it('validates placeholder source keys and unsupported metadata', () => {
    const invalidSource = minimalScene();
    mutableElement(invalidSource).placeholder = {
      role: 'slide-instance',
      sourceKey: 1,
    };
    expect(validatePptxScene(invalidSource).issues).toContainEqual({
      code: 'invalid-scene-document',
      message: 'Expected a public key',
      path: '$.slides[0].elements[0].placeholder.sourceKey',
    });

    for (const feature of [undefined, 1, '', '   ']) {
      const scene = minimalScene();
      mutableSlide(scene).elements = [
        {
          authored: {},
          feature,
          key: 'unsupported-1',
          resolved: { hidden: false },
          type: 'unsupported',
        },
      ];
      expect(validatePptxScene(scene).issues).toContainEqual({
        code: 'invalid-scene-document',
        message: 'Expected a non-empty unsupported feature name',
        path: '$.slides[0].elements[0].feature',
      });
    }

    const invalidPreview = minimalScene();
    mutableSlide(invalidPreview).elements = [
      {
        authored: {},
        feature: 'chart',
        key: 'unsupported-1',
        previewText: 1,
        resolved: { hidden: false },
        type: 'unsupported',
      },
    ];
    expect(validatePptxScene(invalidPreview).issues).toContainEqual({
      code: 'invalid-scene-document',
      message: 'Expected a string',
      path: '$.slides[0].elements[0].previewText',
    });
  });

  it('returns exact element-base and resolved-transform diagnostics', () => {
    const unknownType = minimalScene();
    mutableElement(unknownType).type = 'widget';
    expect(validatePptxScene(unknownType).issues).toContainEqual({
      code: 'invalid-scene-document',
      message: 'Unknown scene element type',
      path: '$.slides[0].elements[0].type',
    });

    const authored = minimalScene();
    const authoredState = mutableElement(authored).authored as Record<
      string,
      unknown
    >;
    authoredState.unexpected = true;
    expect(validatePptxScene(authored).issues).toContainEqual({
      code: 'invalid-scene-document',
      message: 'Unknown property',
      path: '$.slides[0].elements[0].authored.unexpected',
    });

    const resolved = minimalScene();
    const resolvedState = mutableElement(resolved).resolved as Record<
      string,
      unknown
    >;
    resolvedState.unexpected = true;
    expect(validatePptxScene(resolved).issues).toContainEqual({
      code: 'invalid-scene-document',
      message: 'Unknown property',
      path: '$.slides[0].elements[0].resolved.unexpected',
    });

    const hidden = minimalScene();
    mutableElement(hidden).resolved = { hidden: 0 };
    expectSingleIssue(hidden, {
      code: 'invalid-scene-document',
      message: 'Expected a boolean',
      path: '$.slides[0].elements[0].resolved.hidden',
    });

    const resolvedTransform = minimalScene();
    mutableElement(resolvedTransform).resolved = {
      hidden: false,
      transform: { height: 80, width: 0, x: 20, y: 30 },
    };
    expect(validatePptxScene(resolvedTransform).issues).toContainEqual({
      code: 'invalid-numeric-value',
      message: 'Expected a positive finite number',
      path: '$.slides[0].elements[0].resolved.transform.width',
    });
  });

  it('returns exact schema and owner diagnostics', () => {
    const schema = minimalScene();
    schema.schemaVersion = 3;
    expectSingleIssue(schema, {
      code: 'unsupported-schema-version',
      message: 'Only PowerPoint scene schema version 2 is supported',
      path: '$.schemaVersion',
    });

    const cases: Array<{
      expectedPath: string;
      mutate: (scene: Record<string, unknown>) => void;
    }> = [
      {
        expectedPath: '$.themes[0].key',
        mutate: (scene) => {
          scene.themes = [{ key: '' }];
        },
      },
      {
        expectedPath: '$.themes[0].name',
        mutate: (scene) => {
          scene.themes = [{ key: 'theme-1', name: 1 }];
        },
      },
      {
        expectedPath: '$.masters[0].key',
        mutate: (scene) => {
          scene.masters = [{ elements: [], key: '', themeKey: 'theme-1' }];
        },
      },
      {
        expectedPath: '$.masters[0].name',
        mutate: (scene) => {
          scene.masters = [
            { elements: [], key: 'master-1', name: 1, themeKey: 'theme-1' },
          ];
        },
      },
      {
        expectedPath: '$.masters[0].elements',
        mutate: (scene) => {
          scene.masters = [
            { elements: {}, key: 'master-1', themeKey: 'theme-1' },
          ];
        },
      },
      {
        expectedPath: '$.layouts[0].key',
        mutate: (scene) => {
          scene.layouts = [{ elements: [], key: '', masterKey: 'master-1' }];
        },
      },
      {
        expectedPath: '$.layouts[0].name',
        mutate: (scene) => {
          scene.layouts = [
            { elements: [], key: 'layout-1', masterKey: 'master-1', name: 1 },
          ];
        },
      },
      {
        expectedPath: '$.layouts[0].elements',
        mutate: (scene) => {
          scene.layouts = [
            { elements: {}, key: 'layout-1', masterKey: 'master-1' },
          ];
        },
      },
    ];

    for (const { expectedPath, mutate } of cases) {
      const scene = minimalScene();
      mutate(scene);
      expect(validatePptxScene(scene).issues).toContainEqual(
        expect.objectContaining({ path: expectedPath }),
      );
    }
  });

  it.each(['themes', 'masters', 'layouts'])(
    'detects a hierarchy missing only %s',
    (missing) => {
      const scene = minimalScene();
      scene.themes = [{ key: 'theme-1' }];
      scene.masters = [{ elements: [], key: 'master-1', themeKey: 'theme-1' }];
      scene.layouts = [
        { elements: [], key: 'layout-1', masterKey: 'master-1' },
      ];
      mutableSlide(scene).layoutKey = 'layout-1';
      scene[missing] = [];

      expect(validatePptxScene(scene).issues).toContainEqual({
        code: 'invalid-hierarchy-reference',
        message: 'A declared hierarchy needs themes, masters, and layouts',
        path: '$',
      });
    },
  );

  it('returns exact dangling layout and placeholder reference messages', () => {
    const layout = minimalScene();
    layout.themes = [{ key: 'theme-1' }];
    layout.masters = [{ elements: [], key: 'master-1', themeKey: 'theme-1' }];
    layout.layouts = [{ elements: [], key: 'layout-1', masterKey: 'master-1' }];
    mutableSlide(layout).layoutKey = 'missing';
    expect(validatePptxScene(layout).issues).toContainEqual({
      code: 'invalid-hierarchy-reference',
      message: 'Slide references an unknown layout',
      path: '$.slides[0].layoutKey',
    });

    const placeholder = minimalScene();
    mutableElement(placeholder).placeholder = {
      role: 'slide-instance',
      sourceKey: 'missing',
    };
    expect(validatePptxScene(placeholder).issues).toContainEqual({
      code: 'invalid-hierarchy-reference',
      message: 'Reference points to an unknown public key',
      path: '$.slides[0].elements[0].placeholder.sourceKey',
    });
  });
});
