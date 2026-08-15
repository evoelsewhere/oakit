import { describe, expect, it } from 'vitest';

import type { XmlLookupValue } from '../../src/common';
import type { PptxParserContext } from '../../src/formats/pptx/internal/context';
import { getSchemeColorFromTheme } from '../../src/formats/pptx/internal/scheme-color';

function xml(value: object): XmlLookupValue {
  return value as unknown as XmlLookupValue;
}

function context(
  values: {
    layoutMap?: Record<string, string>;
    masterMap?: Record<string, string>;
    slideMap?: Record<string, string>;
  } = {},
): PptxParserContext {
  const themeColors: Record<string, XmlLookupValue> = {};
  for (const [name, color] of Object.entries({
    accent1: '111111',
    accent2: '222222',
    accent3: '333333',
    accent4: '444444',
    dk2: '202020',
    lt1: 'eeeeee',
    lt2: 'dddddd',
    tx1: '999999',
    undefined: '888888',
  })) {
    themeColors[`a:${name}`] = xml({ 'a:srgbClr': { attrs: { val: color } } });
  }
  themeColors['a:dk1'] = xml({
    'a:sysClr': { attrs: { lastClr: '101010' } },
  });

  return {
    slideContent: values.slideMap
      ? xml({
          'p:sld': {
            'p:clrMapOvr': {
              'a:overrideClrMapping': { attrs: values.slideMap },
            },
          },
        })
      : xml({}),
    slideLayoutContent: values.layoutMap
      ? xml({
          'p:sldLayout': {
            'p:clrMapOvr': {
              'a:overrideClrMapping': { attrs: values.layoutMap },
            },
          },
        })
      : xml({}),
    slideMasterContent: values.masterMap
      ? xml({
          'p:sldMaster': { 'p:clrMap': { attrs: values.masterMap } },
        })
      : xml({}),
    themeContent: xml({
      'a:theme': {
        'a:themeElements': { 'a:clrScheme': themeColors },
      },
    }),
  } as unknown as PptxParserContext;
}

describe('PPTX theme scheme color resolution', () => {
  it('uses the placeholder color when phClr is supplied', () => {
    expect(
      getSchemeColorFromTheme('a:phClr', context(), undefined, 'abcdef'),
    ).toBe('abcdef');
  });

  it('ignores a placeholder argument for non-placeholder schemes', () => {
    expect(
      getSchemeColorFromTheme('a:accent1', context(), undefined, 'abcdef'),
    ).toBe('111111');
  });

  it.each([
    ['a:tx1', '101010'],
    ['a:tx2', '202020'],
    ['a:bg1', 'eeeeee'],
    ['a:bg2', 'dddddd'],
    ['a:accent1', '111111'],
  ])('resolves fallback scheme %s as %s', (scheme, expected) => {
    expect(getSchemeColorFromTheme(scheme, context())).toBe(expected);
  });

  it('prefers an explicit color map over document-level maps', () => {
    const parserContext = context({
      layoutMap: { tx1: 'accent3' },
      masterMap: { tx1: 'accent4' },
      slideMap: { tx1: 'accent2' },
    });

    expect(
      getSchemeColorFromTheme('a:tx1', parserContext, xml({ tx1: 'accent1' })),
    ).toBe('111111');
  });

  it.each([
    ['a:tx1', 'accent1', '111111'],
    ['a:tx2', 'accent2', '222222'],
    ['a:bg1', 'accent3', '333333'],
    ['a:bg2', 'accent4', '444444'],
  ])('maps %s through %s', (scheme, mappedScheme, expected) => {
    expect(
      getSchemeColorFromTheme(
        scheme,
        context(),
        xml({ [scheme.slice(2)]: mappedScheme }),
      ),
    ).toBe(expected);
  });

  it('returns undefined when an explicit map omits the requested role', () => {
    expect(
      getSchemeColorFromTheme('a:tx1', context(), xml({ accent1: 'accent2' })),
    ).toBeUndefined();
  });

  it('uses slide, layout, and master maps in that precedence order', () => {
    expect(
      getSchemeColorFromTheme(
        'a:tx1',
        context({
          layoutMap: { tx1: 'accent3' },
          masterMap: { tx1: 'accent4' },
          slideMap: { tx1: 'accent2' },
        }),
      ),
    ).toBe('222222');
    expect(
      getSchemeColorFromTheme(
        'a:tx1',
        context({
          layoutMap: { tx1: 'accent3' },
          masterMap: { tx1: 'accent4' },
        }),
      ),
    ).toBe('333333');
    expect(
      getSchemeColorFromTheme(
        'a:tx1',
        context({ masterMap: { tx1: 'accent4' } }),
      ),
    ).toBe('444444');
  });

  it('does not remap direct accent scheme colors', () => {
    expect(
      getSchemeColorFromTheme(
        'a:accent1',
        context({ slideMap: { accent1: 'accent4' } }),
      ),
    ).toBe('111111');
  });

  it('returns undefined for a missing theme entry or placeholder', () => {
    expect(getSchemeColorFromTheme('a:accent9', context())).toBeUndefined();
    expect(getSchemeColorFromTheme('a:phClr', context())).toBeUndefined();
  });
});
