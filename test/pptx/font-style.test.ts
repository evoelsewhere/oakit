import { describe, expect, it } from 'vitest';

import type { XmlLookupValue } from '../../src/common';
import type { PptxParserContext } from '../../src/formats/pptx/internal/context';
import {
  getFontBold,
  getFontColor,
  getFontDecoration,
  getFontDecorationLine,
  getFontItalic,
  getFontShadow,
  getFontSize,
  getFontSpace,
  getFontSubscript,
  getFontType,
} from '../../src/formats/pptx/internal/font-style';

interface FontFixture {
  defaultTextStyle?: XmlLookupValue;
  layout?: XmlLookupValue;
  level?: number | string;
  master?: XmlLookupValue;
  masterTextStyles?: XmlLookupValue;
  node?: XmlLookupValue;
  paragraph?: XmlLookupValue;
  textBody?: XmlLookupValue;
  type?: string;
}

function xml(value: object): XmlLookupValue {
  return value as unknown as XmlLookupValue;
}

function context(themeContent: object = {}): PptxParserContext {
  return {
    slideContent: xml({}),
    slideLayoutContent: xml({}),
    slideMasterContent: xml({}),
    themeContent: xml(themeContent),
  } as unknown as PptxParserContext;
}

function styleArgs(fixture: FontFixture = {}): Parameters<typeof getFontBold> {
  return [
    fixture.node ?? xml({}),
    fixture.paragraph ?? xml({}),
    fixture.textBody ?? xml({}),
    fixture.layout,
    fixture.master,
    fixture.type ?? 'body',
    fixture.masterTextStyles,
    fixture.level ?? 1,
  ];
}

function typeface(value: string): XmlLookupValue {
  return xml({ 'a:latin': { attrs: { typeface: value } } });
}

function solidColor(value: string): XmlLookupValue {
  return xml({ 'a:solidFill': { 'a:srgbClr': { attrs: { val: value } } } });
}

function fontReferenceColor(value: string): XmlLookupValue {
  return xml({ 'a:srgbClr': { attrs: { val: value } } });
}

function themeFontScheme(
  major: object = { 'a:latin': { attrs: { typeface: 'Major' } } },
  minor: object = { 'a:latin': { attrs: { typeface: 'Minor' } } },
): object {
  return {
    'a:theme': {
      'a:themeElements': {
        'a:fontScheme': {
          'a:majorFont': major,
          'a:minorFont': minor,
        },
      },
    },
  };
}

describe('PowerPoint font style resolution', () => {
  it.each([
    ['run properties', { node: xml({ 'a:rPr': typeface('Run') }) }, 'Run'],
    [
      'end-paragraph properties when a run has no properties',
      { paragraph: xml({ 'a:endParaRPr': typeface('End') }) },
      'End',
    ],
    [
      'paragraph defaults before body defaults',
      {
        paragraph: xml({ 'a:pPr': { 'a:defRPr': typeface('Paragraph') } }),
        textBody: xml({
          'a:lstStyle': { 'a:lvl1pPr': { 'a:defRPr': typeface('Body') } },
        }),
      },
      'Paragraph',
    ],
    [
      'text-body list defaults',
      {
        textBody: xml({
          'a:lstStyle': { 'a:lvl1pPr': { 'a:defRPr': typeface('Body') } },
        }),
      },
      'Body',
    ],
    [
      'layout list defaults before layout paragraph defaults',
      {
        layout: xml({
          'p:txBody': {
            'a:lstStyle': {
              'a:lvl1pPr': { 'a:defRPr': typeface('Layout list') },
            },
            'a:p': {
              'a:pPr': { 'a:defRPr': typeface('Layout paragraph') },
            },
          },
        }),
      },
      'Layout list',
    ],
    [
      'layout paragraph defaults',
      {
        layout: xml({
          'p:txBody': {
            'a:p': {
              'a:pPr': { 'a:defRPr': typeface('Layout paragraph') },
            },
          },
        }),
      },
      'Layout paragraph',
    ],
    [
      'master list defaults before master paragraph defaults',
      {
        master: xml({
          'p:txBody': {
            'a:lstStyle': {
              'a:lvl1pPr': { 'a:defRPr': typeface('Master list') },
            },
            'a:p': {
              'a:pPr': { 'a:defRPr': typeface('Master paragraph') },
            },
          },
        }),
      },
      'Master list',
    ],
    [
      'master paragraph defaults',
      {
        master: xml({
          'p:txBody': {
            'a:p': {
              'a:pPr': { 'a:defRPr': typeface('Master paragraph') },
            },
          },
        }),
      },
      'Master paragraph',
    ],
    [
      'master title styles',
      {
        masterTextStyles: xml({
          'p:titleStyle': {
            'a:lvl1pPr': { 'a:defRPr': typeface('Master title') },
          },
        }),
        type: 'title',
      },
      'Master title',
    ],
    [
      'master centered-title styles',
      {
        masterTextStyles: xml({
          'p:titleStyle': {
            'a:lvl1pPr': { 'a:defRPr': typeface('Master centered title') },
          },
        }),
        type: 'ctrTitle',
      },
      'Master centered title',
    ],
    [
      'master subtitle body fallback',
      {
        masterTextStyles: xml({
          'p:bodyStyle': {
            'a:lvl1pPr': { 'a:defRPr': typeface('Master subtitle') },
          },
        }),
        type: 'subTitle',
      },
      'Master subtitle',
    ],
    [
      'master body styles',
      {
        masterTextStyles: xml({
          'p:bodyStyle': {
            'a:lvl1pPr': { 'a:defRPr': typeface('Master body') },
          },
        }),
        type: 'body',
      },
      'Master body',
    ],
    [
      'master other styles',
      {
        masterTextStyles: xml({
          'p:otherStyle': {
            'a:lvl1pPr': { 'a:defRPr': typeface('Master other') },
          },
        }),
        type: 'obj',
      },
      'Master other',
    ],
  ] as const)('inherits typeface from %s', (_name, fixture, expected) => {
    expect(getFontType(...styleArgs(fixture), context(themeFontScheme()))).toBe(
      expected,
    );
  });

  it('treats explicit false values as overrides instead of falling through', () => {
    const fixture: FontFixture = {
      node: xml({ 'a:rPr': { attrs: { b: '0', i: 'false' } } }),
      paragraph: xml({
        'a:pPr': { 'a:defRPr': { attrs: { b: '1', i: '1' } } },
      }),
    };

    expect(getFontBold(...styleArgs(fixture))).toBe('');
    expect(getFontItalic(...styleArgs(fixture))).toBe('');
  });

  it.each(['1', 'true'])('accepts DrawingML true value %j', (value) => {
    const fixture = {
      node: xml({ 'a:rPr': { attrs: { b: value, i: value } } }),
    };

    expect(getFontBold(...styleArgs(fixture))).toBe('bold');
    expect(getFontItalic(...styleArgs(fixture))).toBe('italic');
  });

  it.each(['0', 'false', '', 'yes', 'TRUE'])(
    'rejects DrawingML false or malformed value %j',
    (value) => {
      const fixture = {
        node: xml({ 'a:rPr': { attrs: { b: value, i: value } } }),
      };

      expect(getFontBold(...styleArgs(fixture))).toBe('');
      expect(getFontItalic(...styleArgs(fixture))).toBe('');
    },
  );

  it.each([
    'dash',
    'dashHeavy',
    'dashLong',
    'dashLongHeavy',
    'dbl',
    'dotDash',
    'dotDashHeavy',
    'dotDotDash',
    'dotDotDashHeavy',
    'dotted',
    'dottedHeavy',
    'heavy',
    'sng',
    'wavy',
    'wavyDbl',
    'wavyHeavy',
    'words',
  ])('maps underline style %s to CSS underline', (value) => {
    const fixture = { node: xml({ 'a:rPr': { attrs: { u: value } } }) };

    expect(getFontDecoration(...styleArgs(fixture))).toBe('underline');
  });

  it.each(['', 'none', 'invalid'])(
    'does not decorate underline value %j',
    (value) => {
      const fixture = { node: xml({ 'a:rPr': { attrs: { u: value } } }) };

      expect(getFontDecoration(...styleArgs(fixture))).toBe('');
    },
  );

  it.each(['sngStrike', 'dblStrike'])(
    'maps strike style %s to CSS line-through',
    (value) => {
      const fixture = {
        node: xml({ 'a:rPr': { attrs: { strike: value } } }),
      };

      expect(getFontDecorationLine(...styleArgs(fixture))).toBe('line-through');
    },
  );

  it.each(['', 'noStrike', 'invalid'])(
    'does not decorate strike value %j',
    (value) => {
      const fixture = {
        node: xml({ 'a:rPr': { attrs: { strike: value } } }),
      };

      expect(getFontDecorationLine(...styleArgs(fixture))).toBe('');
    },
  );

  it.each([
    ['125', '1.25pt'],
    ['-125', '-1.25pt'],
    ['+125', '1.25pt'],
    ['0', ''],
    ['125junk', ''],
    ['junk125', ''],
    ['1e2', ''],
    ['0x10', ''],
    [' 125', ''],
    ['125 ', ''],
  ])('normalizes character spacing %j', (value, expected) => {
    const fixture = { node: xml({ 'a:rPr': { attrs: { spc: value } } }) };

    expect(getFontSpace(...styleArgs(fixture))).toBe(expected);
  });

  it.each([
    ['25000', 'super'],
    ['-25000', 'sub'],
    ['0', ''],
    ['25000junk', ''],
    ['junk25000', ''],
  ])('normalizes baseline %j', (value, expected) => {
    const fixture = {
      node: xml({ 'a:rPr': { attrs: { baseline: value } } }),
    };

    expect(getFontSubscript(...styleArgs(fixture))).toBe(expected);
  });

  it.each([
    ['2400', 'body', '24pt'],
    ['+2400', 'body', '24pt'],
    ['0', 'body', '18pt'],
    ['-100', 'body', '18pt'],
    ['1e3', 'body', '18pt'],
    ['', 'dt', '12pt'],
    ['', 'sldNum', '12pt'],
    ['', 'body', '18pt'],
  ])('normalizes font size %j for %s', (value, type, expected) => {
    const fixture = {
      node: xml({ 'a:rPr': { attrs: { sz: value } } }),
      type,
    };

    expect(getFontSize(...styleArgs(fixture), undefined)).toBe(expected);
  });

  it('uses level-specific default text size before the default paragraph', () => {
    const fixture: FontFixture = {
      defaultTextStyle: xml({
        'a:lvl1pPr': { 'a:defRPr': { attrs: { sz: '2200' } } },
        'a:defPPr': { 'a:defRPr': { attrs: { sz: '3300' } } },
      }),
    };

    expect(getFontSize(...styleArgs(fixture), fixture.defaultTextStyle)).toBe(
      '22pt',
    );
  });

  it('falls back to default paragraph size when no level size exists', () => {
    const fixture: FontFixture = {
      defaultTextStyle: xml({
        'a:defPPr': { 'a:defRPr': { attrs: { sz: '3300' } } },
      }),
    };

    expect(getFontSize(...styleArgs(fixture), fixture.defaultTextStyle)).toBe(
      '33pt',
    );
  });

  it('resolves solid and gradient run colors', () => {
    const solidFixture = { node: xml({ 'a:rPr': solidColor('123456') }) };
    const gradientFixture = {
      node: xml({
        'a:rPr': {
          'a:gradFill': {
            'a:gsLst': {
              'a:gs': [
                {
                  attrs: { pos: '0' },
                  'a:srgbClr': { attrs: { val: '000000' } },
                },
                {
                  attrs: { pos: '100000' },
                  'a:srgbClr': { attrs: { val: 'ffffff' } },
                },
              ],
            },
            'a:lin': { attrs: { ang: '0' } },
          },
        },
      }),
    };

    expect(getFontColor(...styleArgs(solidFixture), undefined, context())).toBe(
      '#123456',
    );
    expect(
      getFontColor(...styleArgs(gradientFixture), undefined, context()),
    ).toEqual({
      colors: [
        { color: '#000000', pos: '0%' },
        { color: '#ffffff', pos: '100%' },
      ],
      path: 'line',
      rot: 0,
    });
  });

  it.each([
    ['shape font reference', {}, fontReferenceColor('111111'), '#111111'],
    [
      'layout font reference',
      {
        layout: xml({
          'p:style': { 'a:fontRef': fontReferenceColor('222222') },
        }),
      },
      undefined,
      '#222222',
    ],
    [
      'master font reference',
      {
        master: xml({
          'p:style': { 'a:fontRef': fontReferenceColor('333333') },
        }),
      },
      undefined,
      '#333333',
    ],
    [
      'master text style',
      {
        masterTextStyles: xml({
          'p:bodyStyle': {
            'a:lvl1pPr': { 'a:defRPr': solidColor('444444') },
          },
        }),
      },
      undefined,
      '#444444',
    ],
  ] as const)(
    'resolves color from %s',
    (_name, fixture, fontReference, expected) => {
      expect(
        getFontColor(...styleArgs(fixture), fontReference, context()),
      ).toBe(expected);
    },
  );

  it('uses the first available color source and returns empty when none exists', () => {
    const fixture: FontFixture = {
      layout: xml({
        'p:style': { 'a:fontRef': fontReferenceColor('222222') },
      }),
      master: xml({
        'p:style': { 'a:fontRef': fontReferenceColor('333333') },
      }),
      masterTextStyles: xml({
        'p:bodyStyle': {
          'a:lvl1pPr': { 'a:defRPr': solidColor('444444') },
        },
      }),
    };

    expect(
      getFontColor(
        ...styleArgs(fixture),
        fontReferenceColor('111111'),
        context(),
      ),
    ).toBe('#111111');
    expect(getFontColor(...styleArgs(), undefined, context())).toBe('');
  });

  it('formats shadows with and without optional blur and color', () => {
    const complete = {
      node: xml({
        'a:rPr': {
          'a:effectLst': {
            'a:outerShdw': {
              attrs: { blurRad: '25400', dir: '0', dist: '12700' },
              'a:srgbClr': { attrs: { val: '123456' } },
            },
          },
        },
      }),
    };
    const minimal = {
      node: xml({
        'a:rPr': {
          'a:effectLst': {
            'a:outerShdw': { attrs: { dir: '0', dist: '12700' } },
          },
        },
      }),
    };

    expect(getFontShadow(...styleArgs(complete), context())).toBe(
      '1pt 0pt 2pt #123456',
    );
    expect(getFontShadow(...styleArgs(minimal), context())).toBe('1pt 0pt');
    expect(getFontShadow(...styleArgs(), context())).toBe('');
  });

  it('falls back through East Asian and complex theme defaults', () => {
    const eastAsianTheme = context(
      themeFontScheme(
        { 'a:ea': { attrs: { typeface: 'Major East Asian' } } },
        { 'a:ea': { attrs: { typeface: 'Minor East Asian' } } },
      ),
    );
    const complexTheme = context(
      themeFontScheme(
        { 'a:cs': { attrs: { typeface: 'Major Complex' } } },
        { 'a:cs': { attrs: { typeface: 'Minor Complex' } } },
      ),
    );

    expect(getFontType(...styleArgs({ type: 'title' }), eastAsianTheme)).toBe(
      'Major East Asian',
    );
    expect(getFontType(...styleArgs({ type: 'body' }), eastAsianTheme)).toBe(
      'Minor East Asian',
    );
    expect(getFontType(...styleArgs({ type: 'title' }), complexTheme)).toBe(
      'Major Complex',
    );
    expect(getFontType(...styleArgs({ type: 'body' }), complexTheme)).toBe(
      'Minor Complex',
    );
  });

  it('uses direct symbol typefaces and strips unknown theme markers', () => {
    const symbol = {
      node: xml({ 'a:rPr': { 'a:sym': { attrs: { typeface: 'Symbol' } } } }),
    };
    const unknownTheme = {
      node: xml({
        'a:rPr': { 'a:latin': { attrs: { typeface: '+custom' } } },
      }),
    };

    expect(getFontType(...styleArgs(symbol), context(themeFontScheme()))).toBe(
      'Symbol',
    );
    expect(
      getFontType(...styleArgs(unknownTheme), context(themeFontScheme())),
    ).toBe('custom');
  });
});
