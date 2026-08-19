import { describe, expect, it, vi } from 'vitest';

import type { PptxSceneGroupElement } from '../../src/formats/pptx/scene-types';
import type { Group } from '../../src/formats/pptx/types';
import {
  verifyPowerPointGroupElement,
  type PptxGroupVerificationDependencies,
} from '../../src/formats/pptx/writer/group-verify';

function expected(): PptxSceneGroupElement {
  return {
    authored: {
      transform: {
        childSpace: { height: 80, width: 160, x: 5, y: 10 },
        height: 100,
        width: 200,
        x: 20,
        y: 30,
      },
    },
    elements: [
      {
        authored: { transform: { height: 20, width: 30, x: 1, y: 2 } },
        key: 'child',
        resolved: { hidden: false },
        type: 'shape',
      },
    ],
    key: 'group',
    resolved: { hidden: false },
    type: 'group',
  };
}

function generated(): Group {
  return {
    childSpace: { height: 80, width: 160, x: 5, y: 10 },
    elements: [
      {
        borderColor: '#000000',
        borderStrokeDasharray: '0',
        borderType: 'solid',
        borderWidth: 0,
        content: '',
        fill: null,
        height: 20,
        id: '3',
        isFlipH: false,
        isFlipV: false,
        left: 1,
        name: 'Shape 3',
        order: 0,
        rotate: 0,
        shapType: 'rect',
        top: 2,
        type: 'shape',
        vAlign: 'up',
        width: 30,
        wrap: true,
      },
    ],
    height: 100,
    id: '2',
    isFlipH: false,
    isFlipV: false,
    left: 20,
    order: 0,
    rotate: 0,
    top: 30,
    type: 'group',
    width: 200,
  };
}

function dependencies() {
  const verifyChild = vi.fn();
  const verifyTransform = vi.fn();
  const value: PptxGroupVerificationDependencies = {
    expectedPointValue: (input) => input,
    verifyChild,
    verifyTransform,
  };
  return { value, verifyChild, verifyTransform };
}

describe('native PowerPoint group verification', () => {
  it('verifies transform, child space, and every child in order', () => {
    const generatedValue = generated();
    const expectedValue = expected();
    const { value, verifyChild, verifyTransform } = dependencies();

    expect(() =>
      verifyPowerPointGroupElement(
        generatedValue,
        expectedValue,
        'slide 1, element 2',
        value,
      ),
    ).not.toThrow();
    expect(verifyTransform).toHaveBeenCalledWith(
      generatedValue,
      expectedValue,
      'slide 1, element 2',
    );
    expect(verifyChild).toHaveBeenCalledWith(
      generatedValue.elements[0],
      expectedValue.elements[0],
      0,
    );
  });

  it.each([undefined, { type: 'shape' }])(
    'rejects missing or non-group output %#',
    (value) => {
      expect(() =>
        verifyPowerPointGroupElement(
          value as never,
          expected(),
          'target',
          dependencies().value,
        ),
      ).toThrow('Generated PowerPoint group missing at target');
    },
  );

  it('requires the expected authored child space', () => {
    const expectedValue = expected();
    const transform = expectedValue.authored.transform;
    if (transform === undefined) throw new Error('Expected transform');
    delete (transform as { childSpace?: unknown }).childSpace;

    expect(() =>
      verifyPowerPointGroupElement(
        generated(),
        expectedValue,
        'target',
        dependencies().value,
      ),
    ).toThrow('Expected PowerPoint group child space missing at target');
  });

  it.each([
    ['missing', (value: Group) => delete value.childSpace],
    [
      'x',
      (value: Group) =>
        ((value.childSpace as NonNullable<Group['childSpace']>).x = 6),
    ],
    [
      'y',
      (value: Group) =>
        ((value.childSpace as NonNullable<Group['childSpace']>).y = 11),
    ],
    [
      'width',
      (value: Group) =>
        ((value.childSpace as NonNullable<Group['childSpace']>).width = 161),
    ],
    [
      'height',
      (value: Group) =>
        ((value.childSpace as NonNullable<Group['childSpace']>).height = 81),
    ],
  ])('rejects a %s child-space mismatch', (_name, mutate) => {
    const value = generated();
    mutate(value);

    expect(() =>
      verifyPowerPointGroupElement(
        value,
        expected(),
        'target',
        dependencies().value,
      ),
    ).toThrow('Generated PowerPoint group child space mismatch at target');
  });

  it('rejects child-count mismatches without dispatching children', () => {
    const value = generated();
    value.elements = [];
    const deps = dependencies();

    expect(() =>
      verifyPowerPointGroupElement(value, expected(), 'target', deps.value),
    ).toThrow('Generated PowerPoint group child count mismatch at target');
    expect(deps.verifyChild).not.toHaveBeenCalled();
  });
});
