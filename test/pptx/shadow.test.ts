import { describe, expect, it } from 'vitest';

import type { XmlLookupValue } from '../../src/common';
import type { PptxParserContext } from '../../src/formats/pptx/internal/context';
import { getShadow } from '../../src/formats/pptx/internal/shadow';

function xml(value: object): XmlLookupValue {
  return value as unknown as XmlLookupValue;
}

function context(): PptxParserContext {
  return {
    slideContent: xml({}),
    slideLayoutContent: xml({}),
    slideMasterContent: xml({}),
    themeContent: xml({}),
  } as unknown as PptxParserContext;
}

describe('PPTX shadow parsing', () => {
  it('returns finite neutral values for an empty shadow', () => {
    expect(getShadow(xml({}), context())).toEqual({
      blur: 0,
      color: '',
      h: 0,
      v: 0,
    });
  });

  it('converts distance, blur, direction, and color exactly', () => {
    expect(
      getShadow(
        xml({
          attrs: { blurRad: '25400', dir: '0', dist: '12700' },
          'a:srgbClr': { attrs: { val: '12abef' } },
        }),
        context(),
      ),
    ).toEqual({ blur: 2, color: '#12abef', h: 1, v: 0 });
  });

  it.each([
    ['5400000', 0, 1],
    ['10800000', -1, 0],
    ['16200000', 0, -1],
    ['21600000', 1, 0],
  ] as const)(
    'projects direction %s onto horizontal and vertical offsets',
    (direction, horizontal, vertical) => {
      const shadow = getShadow(
        xml({ attrs: { dir: direction, dist: '12700' } }),
        context(),
      );

      expect(shadow.h).toBeCloseTo(horizontal, 12);
      expect(shadow.v).toBeCloseTo(vertical, 12);
    },
  );

  it.each(['not-a-number', '60000x', '1.5', '-60000', 'Infinity'])(
    'normalizes invalid direction %j to zero degrees',
    (direction) => {
      const shadow = getShadow(
        xml({ attrs: { dir: direction, dist: '12700' } }),
        context(),
      );

      expect(shadow.h).toBe(1);
      expect(shadow.v).toBe(0);
    },
  );

  it.each(['not-a-number', '12700x', '1.5', '-12700', 'Infinity'])(
    'normalizes invalid distance %j to zero points',
    (distance) => {
      const shadow = getShadow(
        xml({ attrs: { dir: '5400000', dist: distance } }),
        context(),
      );

      expect(shadow.h).toBe(0);
      expect(shadow.v).toBe(0);
    },
  );

  it.each(['not-a-number', '25400x', '1.5', '-25400', 'Infinity'])(
    'normalizes invalid blur radius %j to zero points',
    (blurRadius) => {
      expect(
        getShadow(xml({ attrs: { blurRad: blurRadius } }), context()).blur,
      ).toBe(0);
    },
  );
});
