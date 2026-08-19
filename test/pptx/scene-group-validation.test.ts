import { describe, expect, it } from 'vitest';

import { validatePptxScene } from '../../src/formats/pptx/scene-validation';

type Mutable = Record<string, unknown>;

const GROUP_PATH = '$.slides[0].elements[0]';

function shape(key = 'child-shape'): Mutable {
  return {
    authored: {
      transform: { height: 40, width: 60, x: 10, y: 20 },
    },
    key,
    resolved: { hidden: false },
    type: 'shape',
  };
}

function group(): Mutable {
  return {
    authored: {
      transform: {
        childSpace: { height: 100, width: 200, x: 0, y: 0 },
        height: 100,
        width: 200,
        x: 40,
        y: 50,
      },
    },
    elements: [shape()],
    key: 'group-1',
    resolved: { hidden: false },
    type: 'group',
  };
}

function document(element: Mutable = group()) {
  return {
    layouts: [],
    masters: [],
    media: [],
    schemaVersion: 2,
    size: { height: 540, width: 960 },
    slides: [{ elements: [element], key: 'slide-1' }],
    themes: [],
  };
}

function validate(
  element: Mutable,
  profile: 'create-native-v1' | 'create-text-v1' | 'scene' = 'create-native-v1',
) {
  return validatePptxScene(document(element), { profile });
}

function authoredTransform(element: Mutable): Mutable {
  return ((element.authored as Mutable).transform ?? {}) as Mutable;
}

describe('native PowerPoint group scene validation', () => {
  it('accepts nested child-space transforms in create and scene profiles', () => {
    expect(validate(group())).toEqual({ issues: [], valid: true });
    expect(validate(group(), 'scene')).toEqual({ issues: [], valid: true });
  });

  it('requires a group elements array and validates every nested element', () => {
    const primitive = group();
    primitive.elements = {};
    expect(validate(primitive).issues).toContainEqual({
      code: 'invalid-scene-document',
      message: 'Expected an array',
      path: `${GROUP_PATH}.elements`,
    });

    const nested = group();
    nested.elements = [null];
    expect(validate(nested).issues).toContainEqual({
      code: 'invalid-scene-document',
      message: 'Expected an object',
      path: `${GROUP_PATH}.elements[0]`,
    });
  });

  it('requires an exact child coordinate space', () => {
    const missing = group();
    delete authoredTransform(missing).childSpace;
    expect(validate(missing).issues).toContainEqual({
      code: 'invalid-scene-document',
      message: 'Expected an object',
      path: `${GROUP_PATH}.authored.transform.childSpace`,
    });

    const invalid = group();
    const childSpace = authoredTransform(invalid).childSpace as Mutable;
    childSpace.width = 0;
    childSpace.x = Number.NaN;
    const issues = validate(invalid).issues;
    expect(issues).toEqual(
      expect.arrayContaining([
        {
          code: 'invalid-numeric-value',
          message: 'Expected a positive finite number',
          path: `${GROUP_PATH}.authored.transform.childSpace.width`,
        },
        {
          code: 'invalid-numeric-value',
          message: 'Expected a finite number',
          path: `${GROUP_PATH}.authored.transform.childSpace.x`,
        },
      ]),
    );
  });

  it('rejects child-space rotation, flips, and nested childSpace fields', () => {
    for (const key of [
      'rotation',
      'flipHorizontal',
      'flipVertical',
      'childSpace',
    ]) {
      const value = group();
      const childSpace = authoredTransform(value).childSpace as Mutable;
      childSpace[key] =
        key === 'rotation' ? 10 : key === 'childSpace' ? {} : true;
      expect(validate(value).issues).toContainEqual({
        code: 'invalid-scene-document',
        message: 'Unknown property',
        path: `${GROUP_PATH}.authored.transform.childSpace.${key}`,
      });
    }
  });

  it('requires authored group geometry and rejects shape styling', () => {
    const missing = group();
    delete (missing.authored as Mutable).transform;
    expect(validate(missing).issues).toContainEqual({
      code: 'unsupported-feature',
      message:
        'Creation profile create-native-v1 requires an authored group transform',
      path: `${GROUP_PATH}.authored.transform`,
    });

    for (const key of ['fillColor', 'geometry', 'lineColor', 'lineWidth']) {
      const value = group();
      (value.authored as Mutable)[key] =
        key === 'geometry' ? 'rect' : key === 'lineWidth' ? 1 : '#123456';
      expect(validate(value).issues).toContainEqual({
        code: 'unsupported-feature',
        message:
          'Creation profile create-native-v1 does not apply shape styling to groups',
        path: `${GROUP_PATH}.authored`,
      });
    }
  });

  it('rejects groups from the text-only creation profile', () => {
    expect(validate(group(), 'create-text-v1').issues).toContainEqual({
      code: 'unsupported-feature',
      message: 'Creation profile create-text-v1 supports text elements only',
      path: GROUP_PATH,
    });
  });

  it('rejects duplicate, repeated, and cyclic nested elements', () => {
    const duplicate = group();
    duplicate.elements = [shape('same'), shape('same')];
    expect(validate(duplicate).issues).toContainEqual({
      code: 'duplicate-public-key',
      message: 'Duplicate public key: same',
      path: `${GROUP_PATH}.elements[1].key`,
    });

    const repeated = group();
    const shared = shape('shared');
    repeated.elements = [shared, shared];
    expect(validate(repeated).issues).toContainEqual({
      code: 'invalid-scene-document',
      message:
        'Scene elements must not contain repeated or cyclic object references',
      path: `${GROUP_PATH}.elements[1]`,
    });

    const cyclic = group();
    cyclic.elements = [cyclic];
    expect(validate(cyclic).issues).toContainEqual({
      code: 'invalid-scene-document',
      message:
        'Scene elements must not contain repeated or cyclic object references',
      path: `${GROUP_PATH}.elements[0]`,
    });
  });
});
