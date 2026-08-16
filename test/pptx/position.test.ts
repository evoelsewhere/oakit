import { describe, expect, it } from 'vitest';

import type { XmlLookupValue } from '../../src/common';
import { getPosition, getSize } from '../../src/formats/pptx/internal/position';

function transform(
  values: Record<string, unknown>,
  kind: 'a:ext' | 'a:off',
): XmlLookupValue {
  return { [kind]: { attrs: values } } as unknown as XmlLookupValue;
}

describe('PPTX position and size normalization', () => {
  it('returns zero coordinates when no transform exists', () => {
    expect(getPosition()).toEqual({ left: 0, top: 0 });
    expect(getSize()).toEqual({ height: 0, width: 0 });
  });

  it('converts positive and negative EMU coordinates to points', () => {
    expect(
      getPosition(transform({ x: '-12700', y: '25400' }, 'a:off')),
    ).toEqual({ left: -1, top: 2 });
    expect(getSize(transform({ cx: '914400', cy: '457200' }, 'a:ext'))).toEqual(
      { height: 36, width: 72 },
    );
    expect(getPosition(transform({ x: '+12700', y: '-0' }, 'a:off'))).toEqual({
      left: 1,
      top: 0,
    });
    expect(getSize(transform({ cx: '+12700', cy: '0' }, 'a:ext'))).toEqual({
      height: 0,
      width: 1,
    });
  });

  it('prefers slide values over layout and master values', () => {
    expect(
      getPosition(
        transform({ x: '12700', y: '25400' }, 'a:off'),
        transform({ x: '38100', y: '50800' }, 'a:off'),
        transform({ x: '63500', y: '76200' }, 'a:off'),
      ),
    ).toEqual({ left: 1, top: 2 });
    expect(
      getSize(
        transform({ cx: '12700', cy: '25400' }, 'a:ext'),
        transform({ cx: '38100', cy: '50800' }, 'a:ext'),
        transform({ cx: '63500', cy: '76200' }, 'a:ext'),
      ),
    ).toEqual({ height: 2, width: 1 });
  });

  it('inherits layout and then master transforms', () => {
    const layoutPosition = transform({ x: '12700', y: '25400' }, 'a:off');
    const masterPosition = transform({ x: '38100', y: '50800' }, 'a:off');
    const layoutSize = transform({ cx: '63500', cy: '76200' }, 'a:ext');
    const masterSize = transform({ cx: '88900', cy: '101600' }, 'a:ext');

    expect(getPosition(undefined, layoutPosition, masterPosition)).toEqual({
      left: 1,
      top: 2,
    });
    expect(getPosition(undefined, undefined, masterPosition)).toEqual({
      left: 3,
      top: 4,
    });
    expect(getSize(undefined, layoutSize, masterSize)).toEqual({
      height: 6,
      width: 5,
    });
    expect(getSize(undefined, undefined, masterSize)).toEqual({
      height: 8,
      width: 7,
    });
  });

  it('defaults each absent coordinate independently to zero', () => {
    expect(getPosition(transform({ x: '12700' }, 'a:off'))).toEqual({
      left: 1,
      top: 0,
    });
    expect(getSize(transform({ cy: '25400' }, 'a:ext'))).toEqual({
      height: 2,
      width: 0,
    });
    expect(getPosition(transform({ x: null, y: undefined }, 'a:off'))).toEqual({
      left: 0,
      top: 0,
    });
    expect(getSize(transform({ cx: null, cy: undefined }, 'a:ext'))).toEqual({
      height: 0,
      width: 0,
    });
  });

  it.each([
    '',
    ' ',
    'not-a-number',
    '12700x',
    'Infinity',
    '-Infinity',
    '1.5',
    '1e3',
    '01',
    '9007199254740992',
  ])('exposes malformed geometry %j to validation', (value) => {
    const position = getPosition(transform({ x: value, y: value }, 'a:off'));
    const size = getSize(transform({ cx: value, cy: value }, 'a:ext'));

    expect(Number.isFinite(position.left)).toBe(false);
    expect(Number.isFinite(position.top)).toBe(false);
    expect(Number.isFinite(size.width)).toBe(false);
    expect(Number.isFinite(size.height)).toBe(false);
  });

  it('exposes negative sizes while preserving negative positions', () => {
    expect(
      getPosition(transform({ x: '-12700', y: '-25400' }, 'a:off')),
    ).toEqual({
      left: -1,
      top: -2,
    });
    const size = getSize(transform({ cx: '-12700', cy: '-25400' }, 'a:ext'));
    expect(Number.isFinite(size.width)).toBe(false);
    expect(Number.isFinite(size.height)).toBe(false);
  });
});
