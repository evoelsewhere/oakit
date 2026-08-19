import { describe, expect, it, vi } from 'vitest';

import type { PptxSceneElement } from '../../src/formats/pptx/scene-types';
import {
  createPptxRoundTripGroupPreview,
  type PptxGroupPreviewDependencies,
} from '../../src/formats/pptx/roundtrip/group-preview';
import type { Element, Group } from '../../src/formats/pptx/types';

function group(): Group {
  return {
    childSpace: { height: 80, width: 160, x: 5, y: 10 },
    elements: [
      { id: 'child-1', type: 'shape' } as Element,
      { id: 'child-2', type: 'image' } as Element,
    ],
    height: 100,
    id: 'group',
    isFlipH: false,
    isFlipV: true,
    left: 20,
    order: 0,
    rotate: 5,
    top: 30,
    type: 'group',
    width: 200,
  };
}

function mapped(key: string): PptxSceneElement {
  return {
    authored: {},
    feature: 'fixture',
    key,
    resolved: { hidden: false },
    type: 'unsupported',
  };
}

function dependencies() {
  const mapChild = vi.fn((_child: Element, _index: number, key: string) =>
    mapped(key),
  );
  const value: PptxGroupPreviewDependencies = {
    mapChild,
    resolveTransform: () => ({
      flipHorizontal: false,
      flipVertical: true,
      height: 100,
      rotation: 5,
      width: 200,
      x: 20,
      y: 30,
    }),
  };
  return { mapChild, value };
}

describe('PowerPoint native group round-trip preview', () => {
  it('binds child space, nested keys, transforms, and child order', () => {
    const input = group();
    const { mapChild, value } = dependencies();
    const result = createPptxRoundTripGroupPreview(input, 1, 2, value);

    expect(result).toEqual({
      authored: {},
      elements: [
        mapped('slide-2-element-3-element-1'),
        mapped('slide-2-element-3-element-2'),
      ],
      key: 'slide-2-element-3',
      resolved: {
        hidden: false,
        transform: {
          childSpace: { height: 80, width: 160, x: 5, y: 10 },
          flipHorizontal: false,
          flipVertical: true,
          height: 100,
          rotation: 5,
          width: 200,
          x: 20,
          y: 30,
        },
      },
      type: 'group',
    });
    expect(mapChild).toHaveBeenNthCalledWith(
      1,
      input.elements[0],
      0,
      'slide-2-element-3-element-1',
    );
    expect(mapChild).toHaveBeenNthCalledWith(
      2,
      input.elements[1],
      1,
      'slide-2-element-3-element-2',
    );
  });

  it('uses an owner-provided key for nested groups', () => {
    const result = createPptxRoundTripGroupPreview(
      group(),
      0,
      0,
      dependencies().value,
      'parent-element-4',
    );

    expect(result).toMatchObject({
      elements: [
        { key: 'parent-element-4-element-1' },
        { key: 'parent-element-4-element-2' },
      ],
      key: 'parent-element-4',
    });
  });

  it.each([
    ['missing transform', (value: Group) => value, true],
    [
      'missing child space',
      (value: Group) => {
        delete value.childSpace;
        return value;
      },
      false,
    ],
    [
      'non-finite child x',
      (value: Group) => {
        (value.childSpace as NonNullable<Group['childSpace']>).x = Number.NaN;
        return value;
      },
      false,
    ],
    [
      'non-finite child y',
      (value: Group) => {
        (value.childSpace as NonNullable<Group['childSpace']>).y = Infinity;
        return value;
      },
      false,
    ],
    [
      'zero child width',
      (value: Group) => {
        (value.childSpace as NonNullable<Group['childSpace']>).width = 0;
        return value;
      },
      false,
    ],
    [
      'non-finite child width',
      (value: Group) => {
        (value.childSpace as NonNullable<Group['childSpace']>).width =
          Number.NaN;
        return value;
      },
      false,
    ],
    [
      'zero child height',
      (value: Group) => {
        (value.childSpace as NonNullable<Group['childSpace']>).height = 0;
        return value;
      },
      false,
    ],
    [
      'non-finite child height',
      (value: Group) => {
        (value.childSpace as NonNullable<Group['childSpace']>).height =
          Infinity;
        return value;
      },
      false,
    ],
  ] as const)('keeps %s preservation-only', (_name, mutate, noTransform) => {
    const input = mutate(group());
    const deps = dependencies().value;
    if (noTransform) deps.resolveTransform = () => undefined;

    expect(createPptxRoundTripGroupPreview(input, 0, 0, deps)).toBeUndefined();
  });
});
