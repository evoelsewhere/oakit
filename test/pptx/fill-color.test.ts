import { describe, expect, it } from 'vitest';

import {
  getBgGradientFill,
  getFillType,
  getGradientFill,
  getPatternFill,
  getSolidFill,
} from '../../src/formats/pptx/internal/fill';
import { fillContext, xml } from './fill-fixture';

function colorContext() {
  return fillContext({
    themeContent: xml({
      'a:theme': {
        'a:themeElements': {
          'a:clrScheme': {
            'a:accent1': { 'a:srgbClr': { attrs: { val: '112233' } } },
            'a:accent2': { 'a:srgbClr': { attrs: { val: '445566' } } },
          },
        },
      },
    }),
  }).context;
}

function rgb(value: string, transforms: object = {}) {
  return xml({
    'a:srgbClr': { attrs: { val: value }, ...transforms },
  });
}

describe('PPTX fill classification', () => {
  it.each([
    ['a:noFill', 'NO_FILL'],
    ['a:solidFill', 'SOLID_FILL'],
    ['a:gradFill', 'GRADIENT_FILL'],
    ['a:pattFill', 'PATTERN_FILL'],
    ['a:blipFill', 'PIC_FILL'],
    ['a:grpFill', 'GROUP_FILL'],
  ] as const)('classifies %s', (key, expected) => {
    expect(getFillType(xml({ [key]: {} }))).toBe(expected);
  });

  it('uses explicit no-fill precedence and returns empty for no fill node', () => {
    expect(getFillType(xml({ 'a:noFill': {}, 'a:solidFill': {} }))).toBe(
      'NO_FILL',
    );
    expect(getFillType(undefined)).toBe('');
    expect(getFillType(xml({}))).toBe('');
  });
});

describe('PPTX solid colors', () => {
  const context = colorContext();

  it('returns empty for missing, unknown, and malformed color nodes', () => {
    expect(getSolidFill(undefined, undefined, undefined, context)).toBe('');
    expect(getSolidFill(xml({}), undefined, undefined, context)).toBe('');
    expect(
      getSolidFill(rgb('not-a-color'), undefined, undefined, context),
    ).toBe('');
    expect(
      getSolidFill(
        rgb('not-a-color', {
          'a:alpha': { attrs: { val: '50000' } },
        }),
        undefined,
        undefined,
        context,
      ),
    ).toBe('');
    expect(
      getSolidFill(
        xml({ 'a:srgbClr': { attrs: {} } }),
        undefined,
        undefined,
        context,
      ),
    ).toBe('');
    expect(
      getSolidFill(
        xml({ 'a:schemeClr': { attrs: {} } }),
        undefined,
        undefined,
        context,
      ),
    ).toBe('');
    expect(
      getSolidFill(
        xml({ 'a:schemeClr': { attrs: {} } }),
        undefined,
        undefined,
        fillContext({
          themeContent: xml({
            'a:theme': {
              'a:themeElements': {
                'a:clrScheme': {
                  'a:undefined': {
                    'a:srgbClr': { attrs: { val: 'deadbe' } },
                  },
                },
              },
            },
          }),
        }).context,
      ),
    ).toBe('');
    expect(
      getSolidFill(
        xml({ 'a:prstClr': { attrs: {} } }),
        undefined,
        undefined,
        context,
      ),
    ).toBe('');
    expect(
      getSolidFill(
        xml({ 'a:sysClr': { attrs: {} } }),
        undefined,
        undefined,
        context,
      ),
    ).toBe('');
  });

  it.each([
    [rgb('AABBCC'), '#AABBCC'],
    [xml({ 'a:prstClr': { attrs: { val: 'dkBlue' } } }), '#00008b'],
    [xml({ 'a:sysClr': { attrs: { lastClr: 'ABCDEF' } } }), '#ABCDEF'],
    [
      xml({
        'a:scrgbClr': { attrs: { b: '0', g: '50000', r: '100000' } },
      }),
      '#ff8000',
    ],
    [
      xml({
        'a:scrgbClr': { attrs: { b: '0%', g: '50%', r: '100%' } },
      }),
      '#ff8000',
    ],
    [
      xml({
        'a:hslClr': {
          attrs: { hue: '7200000', lum: '50000', sat: '100000' },
        },
      }),
      '#00ff00',
    ],
  ] as const)('resolves a DrawingML color form', (node, expected) => {
    expect(getSolidFill(node, undefined, undefined, context)).toBe(expected);
  });

  it('clamps malformed and out-of-range numeric color components', () => {
    expect(
      getSolidFill(
        xml({
          'a:scrgbClr': {
            attrs: { b: 'Infinity', g: '-100000', r: '200000' },
          },
        }),
        undefined,
        undefined,
        context,
      ),
    ).toBe('#ff0000');
    expect(
      getSolidFill(
        xml({
          'a:scrgbClr': { attrs: { b: '75000', g: '50000', r: '25000' } },
        }),
        undefined,
        undefined,
        context,
      ),
    ).toBe('#4080bf');
    expect(
      getSolidFill(
        xml({
          'a:hslClr': {
            attrs: { hue: 'invalid', lum: '50000', sat: '100000' },
          },
        }),
        undefined,
        undefined,
        context,
      ),
    ).toBe('#ff0000');
  });

  it('resolves theme mappings and placeholder colors', () => {
    expect(
      getSolidFill(
        xml({ 'a:schemeClr': { attrs: { val: 'accent1' } } }),
        undefined,
        undefined,
        context,
      ),
    ).toBe('#112233');
    expect(
      getSolidFill(
        xml({ 'a:schemeClr': { attrs: { val: 'tx1' } } }),
        xml({ tx1: 'accent2' }),
        undefined,
        context,
      ),
    ).toBe('#445566');
    expect(
      getSolidFill(
        xml({ 'a:schemeClr': { attrs: { val: 'phClr' } } }),
        undefined,
        'abcdef',
        context,
      ),
    ).toBe('#abcdef');
  });

  it.each([
    ['a:hueMod', '200000', '#999933'],
    ['a:lumMod', '50000', '#1a334c'],
    ['a:lumOff', '25000', '#79a6d2'],
    ['a:satMod', '50000', '#4d667f'],
    ['a:shade', '50000', '#1a334c'],
    ['a:tint', '50000', '#8cb2d9'],
  ] as const)('applies %s', (name, value, expected) => {
    expect(
      getSolidFill(
        rgb('336699', { [name]: { attrs: { val: value } } }),
        undefined,
        undefined,
        context,
      ),
    ).toBe(expected);
  });

  it('ignores malformed transformations instead of partially parsing them', () => {
    expect(
      getSolidFill(
        rgb('336699', {
          'a:shade': { attrs: { val: '50000x' } },
          'a:tint': { attrs: { val: 'Infinity' } },
        }),
        undefined,
        undefined,
        context,
      ),
    ).toBe('#336699');
  });

  it('applies exact, multiplicative, and offset alpha operations', () => {
    expect(
      getSolidFill(
        rgb('336699', {
          'a:alpha': { attrs: { val: '50000' } },
        }),
        undefined,
        undefined,
        context,
      ),
    ).toBe('#33669980');
    expect(
      getSolidFill(
        rgb('336699', {
          'a:alphaMod': { attrs: { val: '50000' } },
        }),
        undefined,
        undefined,
        context,
      ),
    ).toBe('#33669980');
    expect(
      getSolidFill(
        rgb('336699', {
          'a:alpha': { attrs: { order: '1', val: '50000' } },
          'a:alphaMod': { attrs: { order: '2', val: '50000' } },
        }),
        undefined,
        undefined,
        context,
      ),
    ).toBe('#33669940');
    expect(
      getSolidFill(
        rgb('336699', {
          'a:alpha': { attrs: { order: '1', val: '50000' } },
          'a:alphaOff': { attrs: { order: '2', val: '25000' } },
        }),
        undefined,
        undefined,
        context,
      ),
    ).toBe('#336699bf');
  });

  it('clamps alpha and preserves it through later color transforms', () => {
    expect(
      getSolidFill(
        rgb('336699', {
          'a:alpha': { attrs: { order: '1', val: '200000' } },
          'a:shade': { attrs: { order: '2', val: '50000' } },
        }),
        undefined,
        undefined,
        context,
      ),
    ).toBe('#1a334cff');
    expect(
      getSolidFill(
        rgb('336699', {
          'a:alpha': { attrs: { val: '-100000' } },
        }),
        undefined,
        undefined,
        context,
      ),
    ).toBe('#33669900');
  });

  it('applies color transformations in authored XML order', () => {
    expect(
      getSolidFill(
        rgb('336699', {
          'a:shade': { attrs: { order: '2', val: '50000' } },
          'a:tint': { attrs: { order: '1', val: '50000' } },
        }),
        undefined,
        undefined,
        context,
      ),
    ).toBe('#2c5986');
    expect(
      getSolidFill(
        rgb('336699', {
          'a:shade': { attrs: { order: '1', val: '50000' } },
          'a:tint': { attrs: { order: '2', val: '50000' } },
        }),
        undefined,
        undefined,
        context,
      ),
    ).toBe('#6799cb');
    expect(
      getSolidFill(
        rgb('336699', {
          'a:shade': { attrs: { order: '1', val: '50000' } },
          'a:tint': { attrs: { order: '1', val: '50000' } },
        }),
        undefined,
        undefined,
        context,
      ),
    ).toBe('#6799cb');
    expect(
      getSolidFill(
        rgb('336699', {
          'a:shade': { attrs: { val: '50000' } },
          'a:tint': { attrs: { val: '50000' } },
        }),
        undefined,
        undefined,
        context,
      ),
    ).toBe('#6799cb');
  });
});

describe('PPTX gradients and patterns', () => {
  const context = colorContext();

  it('sorts and clamps gradient stops while preserving an explicit zero', () => {
    expect(
      getGradientFill(
        xml({
          'a:gsLst': {
            'a:gs': [
              {
                attrs: { pos: '100000' },
                'a:srgbClr': { attrs: { val: 'ff0000' } },
              },
              {
                attrs: { pos: '0' },
                'a:srgbClr': { attrs: { val: '0000ff' } },
              },
              {
                attrs: { pos: '50000' },
                'a:srgbClr': { attrs: { val: '00ff00' } },
              },
              {
                attrs: { pos: '-1000' },
                'a:srgbClr': { attrs: { val: '111111' } },
              },
              {
                attrs: { pos: '200000' },
                'a:srgbClr': { attrs: { val: 'eeeeee' } },
              },
            ],
          },
          'a:lin': { attrs: { ang: '120000' } },
        }),
        context,
      ),
    ).toEqual({
      colors: [
        { color: '#0000ff', pos: '0%' },
        { color: '#111111', pos: '0%' },
        { color: '#00ff00', pos: '50%' },
        { color: '#ff0000', pos: '100%' },
        { color: '#eeeeee', pos: '100%' },
      ],
      path: 'line',
      rot: 2,
    });
  });

  it.each(['circle', 'rect', 'shape'] as const)(
    'preserves gradient path %s',
    (path) => {
      expect(
        getGradientFill(xml({ 'a:path': { attrs: { path } } }), context).path,
      ).toBe(path);
    },
  );

  it('uses linear defaults for missing or unsupported gradient geometry', () => {
    expect(getGradientFill(xml({}), context)).toEqual({
      colors: [],
      path: 'line',
      rot: 0,
    });
    expect(
      getGradientFill(
        xml({ 'a:path': { attrs: { path: 'unsupported' } } }),
        context,
      ).path,
    ).toBe('line');
  });

  it('parses pattern colors and applies neutral defaults independently', () => {
    expect(
      getPatternFill(
        xml({
          'a:pattFill': {
            attrs: { prst: 'cross' },
            'a:bgClr': { 'a:srgbClr': { attrs: { val: 'abcdef' } } },
            'a:fgClr': { 'a:srgbClr': { attrs: { val: '123456' } } },
          },
        }),
        context,
      ),
    ).toEqual({
      backgroundColor: '#abcdef',
      foregroundColor: '#123456',
      type: 'cross',
    });
    expect(getPatternFill(xml({ 'a:pattFill': {} }), context)).toEqual({
      backgroundColor: '#ffffff',
      foregroundColor: '#000000',
      type: '',
    });
    expect(getPatternFill(xml({}), context)).toBeNull();
  });

  it('normalizes a background placeholder and rejects unsafe values', () => {
    const master = xml({});
    expect(getBgGradientFill(xml({}), 'ABCDEF', master, context)).toBe(
      '#ABCDEF',
    );
    expect(getBgGradientFill(xml({}), '#123456', master, context)).toBe(
      '#123456',
    );
    expect(
      getBgGradientFill(xml({}), 'red;position:fixed', master, context),
    ).toBeNull();
    expect(getBgGradientFill(xml({}), undefined, master, context)).toBeNull();
  });

  it('builds an authored background gradient instead of its placeholder', () => {
    expect(
      getBgGradientFill(
        xml({
          'a:gradFill': {
            'a:gsLst': {
              'a:gs': {
                attrs: { pos: '50000' },
                'a:schemeClr': { attrs: { val: 'phClr' } },
              },
            },
          },
        }),
        'abcdef',
        xml({}),
        context,
      ),
    ).toEqual({
      colors: [{ color: '#abcdef', pos: '50%' }],
      path: 'line',
      rot: 0,
    });
  });

  it('uses the master color map while building a background gradient', () => {
    expect(
      getBgGradientFill(
        xml({
          'a:gradFill': {
            'a:gsLst': {
              'a:gs': {
                attrs: { pos: '0' },
                'a:schemeClr': { attrs: { val: 'tx1' } },
              },
            },
          },
        }),
        undefined,
        xml({
          'p:sldMaster': {
            'p:clrMap': { attrs: { tx1: 'accent2' } },
          },
        }),
        context,
      ),
    ).toEqual({
      colors: [{ color: '#445566', pos: '0%' }],
      path: 'line',
      rot: 0,
    });
  });
});
