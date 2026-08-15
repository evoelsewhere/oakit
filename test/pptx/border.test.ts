import { describe, expect, it } from 'vitest';

import type { XmlLookupValue } from '../../src/common';
import { getBorder } from '../../src/formats/pptx/internal/border';
import type { PptxParserContext } from '../../src/formats/pptx/internal/context';

function xml(value: object): XmlLookupValue {
  return value as unknown as XmlLookupValue;
}

function context(
  themeLines?: XmlLookupValue | XmlLookupValue[],
  accentColor = 'ff0000',
): PptxParserContext {
  return {
    slideContent: xml({}),
    slideLayoutContent: xml({}),
    slideMasterContent: xml({}),
    themeContent: xml({
      'a:theme': {
        'a:themeElements': {
          'a:clrScheme': {
            'a:accent1': {
              'a:srgbClr': { attrs: { val: accentColor } },
            },
          },
          'a:fmtScheme': {
            'a:lnStyleLst': {
              ...(themeLines === undefined ? {} : { 'a:ln': themeLines }),
            },
          },
        },
      },
    }),
  } as unknown as PptxParserContext;
}

function line(value: object): XmlLookupValue {
  return xml(value);
}

describe('PPTX border parsing', () => {
  it('returns a stable invisible default when no line exists', () => {
    expect(getBorder(undefined, undefined, context())).toEqual({
      borderColor: '#000000',
      borderType: 'solid',
      borderWidth: 0,
      strokeDasharray: '0',
    });
  });

  it('parses a direct shape line with width and RGB color', () => {
    expect(
      getBorder(
        xml({
          'p:spPr': {
            'a:ln': {
              attrs: { w: '25400' },
              'a:solidFill': {
                'a:srgbClr': { attrs: { val: '12abef' } },
              },
            },
          },
        }),
        'shape',
        context(),
      ),
    ).toMatchObject({
      borderColor: '#12abef',
      borderWidth: 2,
    });
  });

  it('gives an authored shape line precedence over its theme reference', () => {
    const themeLine = line({
      attrs: { w: '12700' },
      'a:solidFill': { 'a:srgbClr': { attrs: { val: '111111' } } },
    });

    expect(
      getBorder(
        xml({
          'p:spPr': {
            'a:ln': {
              attrs: { w: '25400' },
              'a:solidFill': {
                'a:srgbClr': { attrs: { val: '222222' } },
              },
            },
          },
          'p:style': { 'a:lnRef': { attrs: { idx: '1' } } },
        }),
        undefined,
        context(themeLine),
      ),
    ).toMatchObject({ borderColor: '#222222', borderWidth: 2 });
  });

  it('accepts a raw line node and honors noFill', () => {
    expect(
      getBorder(
        line({
          attrs: { w: '25400' },
          'a:noFill': {},
          'a:solidFill': {
            'a:srgbClr': { attrs: { val: 'ffffff' } },
          },
        }),
        undefined,
        context(),
      ),
    ).toMatchObject({ borderColor: '#ffffff', borderWidth: 0 });
  });

  it.each(['not-a-number', '12700x', '-12700', 'Infinity'])(
    'normalizes invalid line width %j to zero',
    (width) => {
      expect(
        getBorder(line({ attrs: { w: width } }), 'obj', context()).borderWidth,
      ).toBe(0);
    },
  );

  it('resolves a referenced theme line by one-based index', () => {
    const first = line({
      attrs: { w: '12700' },
      'a:solidFill': { 'a:srgbClr': { attrs: { val: '111111' } } },
    });
    const second = line({
      attrs: { w: '38100' },
      'a:prstDash': { attrs: { val: 'dash' } },
      'a:solidFill': { 'a:srgbClr': { attrs: { val: '222222' } } },
    });

    expect(
      getBorder(
        xml({ 'p:style': { 'a:lnRef': { attrs: { idx: '2' } } } }),
        undefined,
        context([first, second]),
      ),
    ).toMatchObject({
      borderColor: '#222222',
      borderType: 'dashed',
      borderWidth: 3,
      strokeDasharray: '5',
    });
  });

  it('supports a singleton theme line entry', () => {
    const themeLine = line({
      attrs: { w: '12700' },
      'a:solidFill': { 'a:srgbClr': { attrs: { val: 'abcdef' } } },
    });

    expect(
      getBorder(
        xml({ 'p:style': { 'a:lnRef': { attrs: { idx: '1' } } } }),
        undefined,
        context(themeLine),
      ),
    ).toMatchObject({ borderColor: '#abcdef', borderWidth: 1 });
  });

  it('does not reuse a singleton theme line for another index', () => {
    const themeLine = line({
      attrs: { w: '12700' },
      'a:solidFill': { 'a:srgbClr': { attrs: { val: 'abcdef' } } },
    });

    expect(
      getBorder(
        xml({ 'p:style': { 'a:lnRef': { attrs: { idx: '2' } } } }),
        undefined,
        context(themeLine),
      ),
    ).toMatchObject({ borderColor: '#000000', borderWidth: 0 });
  });

  it('uses the line reference scheme color and a valid shade', () => {
    const node = xml({
      'p:style': {
        'a:lnRef': {
          'a:schemeClr': {
            attrs: { val: 'accent1' },
            'a:shade': { attrs: { val: '50000' } },
          },
        },
      },
    });

    expect(getBorder(node, undefined, context()).borderColor).toBe('#800000');
  });

  it.each(['50000x', '50000.5', '-1', '100001', 'Infinity'])(
    'ignores malformed shade %j instead of partially parsing it',
    (shade) => {
      const node = xml({
        'p:style': {
          'a:lnRef': {
            'a:schemeClr': {
              attrs: { val: 'accent1' },
              'a:shade': { attrs: { val: shade } },
            },
          },
        },
      });

      expect(
        getBorder(node, undefined, context(undefined, '808080')).borderColor,
      ).toBe('#808080');
    },
  );

  it.each([
    ['0', '#000000'],
    ['100000', '#808080'],
    ['200000', '#808080'],
  ] as const)(
    'clamps shade %s to its observable color range',
    (shade, color) => {
      const node = xml({
        'p:style': {
          'a:lnRef': {
            'a:schemeClr': {
              attrs: { val: 'accent1' },
              'a:shade': { attrs: { val: shade } },
            },
          },
        },
      });

      expect(
        getBorder(node, undefined, context(undefined, '808080')).borderColor,
      ).toBe(color);
    },
  );

  it.each([
    ['solid', 'solid', '0'],
    ['dash', 'dashed', '5'],
    ['dashDot', 'dashed', '5, 5, 1, 5'],
    ['dot', 'dotted', '1, 5'],
    ['lgDash', 'dashed', '10, 5'],
    ['lgDashDotDot', 'dotted', '10, 5, 1, 5, 1, 5'],
    ['sysDash', 'dashed', '5, 2'],
    ['sysDashDot', 'dotted', '5, 2, 1, 5'],
    ['sysDashDotDot', 'dotted', '5, 2, 1, 5, 1, 5'],
    ['sysDot', 'dotted', '2, 5'],
    ['unknown', 'solid', '0'],
  ] as const)('maps dash style %s', (dash, borderType, strokeDasharray) => {
    expect(
      getBorder(
        line({ 'a:prstDash': { attrs: { val: dash } } }),
        undefined,
        context(),
      ),
    ).toMatchObject({ borderType, strokeDasharray });
  });

  it('parses line endings and defaults an absent ending type', () => {
    expect(
      getBorder(
        line({
          'a:headEnd': { attrs: { len: 'lg', type: 'triangle', w: 'sm' } },
          'a:tailEnd': { attrs: {} },
        }),
        undefined,
        context(),
      ),
    ).toMatchObject({
      headEnd: { length: 'lg', type: 'triangle', width: 'sm' },
      tailEnd: { type: 'none' },
    });
  });

  it.each(['none', 'triangle', 'stealth', 'diamond', 'oval', 'arrow'] as const)(
    'accepts line ending type %s',
    (type) => {
      expect(
        getBorder(
          line({ 'a:headEnd': { attrs: { type } } }),
          undefined,
          context(),
        ).headEnd,
      ).toEqual({ type });
    },
  );

  it.each(['sm', 'med', 'lg'] as const)(
    'accepts line ending size %s for width and length',
    (size) => {
      expect(
        getBorder(
          line({
            'a:headEnd': {
              attrs: { len: size, type: 'arrow', w: size },
            },
          }),
          undefined,
          context(),
        ).headEnd,
      ).toEqual({ length: size, type: 'arrow', width: size });
    },
  );

  it('normalizes unsupported line ending enum values', () => {
    expect(
      getBorder(
        line({
          'a:headEnd': {
            attrs: { len: 'huge', type: 'script', w: 'wide' },
          },
        }),
        undefined,
        context(),
      ).headEnd,
    ).toEqual({ type: 'none' });
  });

  it('normalizes missing and malformed colors to opaque black', () => {
    expect(getBorder(line({}), undefined, context()).borderColor).toBe(
      '#000000',
    );
    expect(
      getBorder(
        line({
          'a:solidFill': {
            'a:srgbClr': { attrs: { val: 'not-a-color' } },
          },
        }),
        undefined,
        context(),
      ).borderColor,
    ).toBe('#000000');

    const schemeNode = xml({
      'p:style': {
        'a:lnRef': {
          'a:schemeClr': { attrs: { val: 'accent1' } },
        },
      },
    });
    expect(
      getBorder(schemeNode, undefined, context(undefined, 'invalid'))
        .borderColor,
    ).toBe('#000000');
  });
});
