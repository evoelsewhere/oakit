import { describe, expect, it } from 'vitest';

import type { XmlLookupValue } from '../../src/common';
import { getTextInsets } from '../../src/formats/pptx/internal/text-insets';

function xml(value: object): XmlLookupValue {
  return value as unknown as XmlLookupValue;
}

function shapeWithInsets(
  attrs: Record<string, number | string> = {},
): XmlLookupValue {
  return {
    'p:txBody': { 'a:bodyPr': { attrs } },
  } as unknown as XmlLookupValue;
}

const emptyNode = {} as unknown as XmlLookupValue;

describe('PPTX text inset normalization', () => {
  it('returns null when no shape level defines text body properties', () => {
    expect(getTextInsets(emptyNode)).toBeNull();
    expect(getTextInsets(emptyNode, xml({ unrelated: true }))).toBeNull();
    expect(
      getTextInsets(emptyNode, undefined, xml({ unrelated: true })),
    ).toBeNull();
  });

  it('uses DrawingML defaults when a text body has no inset attributes', () => {
    expect(getTextInsets(shapeWithInsets())).toEqual({
      b: 3.6,
      l: 7.2,
      r: 7.2,
      t: 3.6,
    });
  });

  it('resolves every inset independently across slide, layout, and master', () => {
    const slide = shapeWithInsets({ lIns: '12700' });
    const layout = shapeWithInsets({ lIns: '25400', tIns: '38100' });
    const master = shapeWithInsets({
      bIns: '63500',
      lIns: '76200',
      rIns: '50800',
      tIns: '88900',
    });

    expect(getTextInsets(slide, layout, master)).toEqual({
      b: 5,
      l: 1,
      r: 4,
      t: 3,
    });
  });

  it('inherits from layout or master when the slide has no text body', () => {
    expect(
      getTextInsets(emptyNode, shapeWithInsets({ lIns: '12700' })),
    ).toEqual({ b: 3.6, l: 1, r: 7.2, t: 3.6 });
    expect(
      getTextInsets(emptyNode, undefined, shapeWithInsets({ rIns: '25400' })),
    ).toEqual({ b: 3.6, l: 7.2, r: 2, t: 3.6 });
  });

  it('preserves explicit zero and accepts numeric attribute values', () => {
    expect(
      getTextInsets(
        shapeWithInsets({ bIns: 50_800, lIns: 0, rIns: '0', tIns: 25_400 }),
      ),
    ).toEqual({ b: 4, l: 0, r: 0, t: 2 });
  });

  it.each(['not-a-number', '12700x', 'Infinity', '-Infinity'])(
    'normalizes malformed inset %s without returning a non-finite number',
    (value) => {
      expect(
        getTextInsets(
          shapeWithInsets({
            bIns: value,
            lIns: value,
            rIns: value,
            tIns: value,
          }),
        ),
      ).toEqual({ b: 0, l: 0, r: 0, t: 0 });
    },
  );

  it('treats empty attributes as absent and applies defaults', () => {
    expect(
      getTextInsets(
        shapeWithInsets({ bIns: '', lIns: '', rIns: '', tIns: '' }),
      ),
    ).toEqual({ b: 3.6, l: 7.2, r: 7.2, t: 3.6 });
  });

  it('inherits past empty slide and layout attributes', () => {
    expect(
      getTextInsets(
        shapeWithInsets({ lIns: '' }),
        shapeWithInsets({ lIns: '' }),
        shapeWithInsets({ lIns: '12700' }),
      ),
    ).toEqual({ b: 3.6, l: 1, r: 7.2, t: 3.6 });
  });
});
