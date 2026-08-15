import { describe, expect, it } from 'vitest';

import {
  getShapeFill,
  getSlideBackgroundFill,
} from '../../src/formats/pptx/internal/fill';
import { fillContext, xml } from './fill-fixture';

function solid(value: string) {
  return { 'a:solidFill': { 'a:srgbClr': { attrs: { val: value } } } };
}

function scheme(value: string) {
  return { 'a:solidFill': { 'a:schemeClr': { attrs: { val: value } } } };
}

function background(
  root: 'p:sld' | 'p:sldLayout' | 'p:sldMaster',
  fill: object,
) {
  return xml({
    [root]: { 'p:cSld': { 'p:bg': { 'p:bgPr': fill } } },
  });
}

function backgroundReference(
  root: 'p:sld' | 'p:sldLayout',
  index: string,
  placeholder = 'abcdef',
) {
  return xml({
    [root]: {
      'p:cSld': {
        'p:bg': {
          'p:bgRef': {
            attrs: { idx: index },
            'a:srgbClr': { attrs: { val: placeholder } },
          },
        },
      },
    },
  });
}

function shape(properties?: object, fillReference?: object) {
  return xml({
    ...(properties === undefined ? {} : { 'p:spPr': properties }),
    ...(fillReference === undefined
      ? {}
      : { 'p:style': { 'a:fillRef': fillReference } }),
  });
}

describe('PPTX slide background inheritance', () => {
  it('uses white when no background is authored', async () => {
    await expect(
      getSlideBackgroundFill(fillContext().context),
    ).resolves.toEqual({ type: 'color', value: '#fff' });
  });

  it('uses slide, layout, and master backgrounds in that precedence order', async () => {
    const slide = background('p:sld', solid('111111'));
    const layout = background('p:sldLayout', solid('222222'));
    const master = background('p:sldMaster', solid('333333'));

    await expect(
      getSlideBackgroundFill(
        fillContext({
          slideContent: slide,
          slideLayoutContent: layout,
          slideMasterContent: master,
        }).context,
      ),
    ).resolves.toEqual({ type: 'color', value: '#111111' });
    await expect(
      getSlideBackgroundFill(
        fillContext({
          slideLayoutContent: layout,
          slideMasterContent: master,
        }).context,
      ),
    ).resolves.toEqual({ type: 'color', value: '#222222' });
    await expect(
      getSlideBackgroundFill(
        fillContext({ slideMasterContent: master }).context,
      ),
    ).resolves.toEqual({ type: 'color', value: '#333333' });
  });

  it('uses slide, layout, and master color maps in that precedence order', async () => {
    const themeContent = xml({
      'a:theme': {
        'a:themeElements': {
          'a:clrScheme': {
            'a:accent1': { 'a:srgbClr': { attrs: { val: '111111' } } },
            'a:accent2': { 'a:srgbClr': { attrs: { val: '222222' } } },
            'a:accent3': { 'a:srgbClr': { attrs: { val: '333333' } } },
          },
        },
      },
    });
    const masterWithMap = xml({
      'p:sldMaster': {
        'p:clrMap': { attrs: { tx1: 'accent1' } },
        'p:cSld': { 'p:bg': { 'p:bgPr': scheme('tx1') } },
      },
    });
    const layoutWithOverride = xml({
      'p:sldLayout': {
        'p:clrMapOvr': {
          'a:overrideClrMapping': { attrs: { tx1: 'accent2' } },
        },
        'p:cSld': { 'p:bg': { 'p:bgPr': scheme('tx1') } },
      },
    });
    const slideWithOverride = xml({
      'p:sld': {
        'p:clrMapOvr': {
          'a:overrideClrMapping': { attrs: { tx1: 'accent3' } },
        },
        'p:cSld': { 'p:bg': { 'p:bgPr': scheme('tx1') } },
      },
    });

    await expect(
      getSlideBackgroundFill(
        fillContext({
          slideMasterContent: masterWithMap,
          themeContent,
        }).context,
      ),
    ).resolves.toEqual({ type: 'color', value: '#111111' });
    await expect(
      getSlideBackgroundFill(
        fillContext({
          slideLayoutContent: layoutWithOverride,
          slideMasterContent: masterWithMap,
          themeContent,
        }).context,
      ),
    ).resolves.toEqual({ type: 'color', value: '#222222' });
    await expect(
      getSlideBackgroundFill(
        fillContext({
          slideContent: slideWithOverride,
          slideLayoutContent: layoutWithOverride,
          slideMasterContent: masterWithMap,
          themeContent,
        }).context,
      ),
    ).resolves.toEqual({ type: 'color', value: '#333333' });
    await expect(
      getSlideBackgroundFill(
        fillContext({
          slideContent: xml({
            'p:sld': {
              'p:clrMapOvr': {
                'a:overrideClrMapping': { attrs: { tx1: 'accent3' } },
              },
            },
          }),
          slideMasterContent: masterWithMap,
          themeContent,
        }).context,
      ),
    ).resolves.toEqual({ type: 'color', value: '#111111' });
  });

  it('treats an explicit no-fill or invalid fill as a white background', async () => {
    const layout = background('p:sldLayout', solid('222222'));
    await expect(
      getSlideBackgroundFill(
        fillContext({
          slideContent: background('p:sld', { 'a:noFill': {} }),
          slideLayoutContent: layout,
        }).context,
      ),
    ).resolves.toEqual({ type: 'color', value: '#fff' });
    await expect(
      getSlideBackgroundFill(
        fillContext({
          slideContent: background('p:sld', solid('invalid')),
          slideLayoutContent: layout,
        }).context,
      ),
    ).resolves.toEqual({ type: 'color', value: '#fff' });
  });

  it('resolves gradient and pattern background values', async () => {
    await expect(
      getSlideBackgroundFill(
        fillContext({
          slideContent: background('p:sld', {
            'a:gradFill': {
              'a:gsLst': {
                'a:gs': {
                  attrs: { pos: '50000' },
                  'a:srgbClr': { attrs: { val: '123456' } },
                },
              },
            },
          }),
        }).context,
      ),
    ).resolves.toEqual({
      type: 'gradient',
      value: {
        colors: [{ color: '#123456', pos: '50%' }],
        path: 'line',
        rot: 0,
      },
    });
    await expect(
      getSlideBackgroundFill(
        fillContext({
          slideContent: background('p:sld', {
            'a:pattFill': { attrs: { prst: 'cross' } },
          }),
        }).context,
      ),
    ).resolves.toEqual({
      type: 'pattern',
      value: {
        backgroundColor: '#ffffff',
        foregroundColor: '#000000',
        type: 'cross',
      },
    });
  });

  it('resolves picture backgrounds against their owning part', async () => {
    const fixture = fillContext({
      options: { imageMode: 'none' },
      slideContent: background('p:sld', {
        'a:blipFill': {
          'a:blip': {
            attrs: { 'r:embed': 'rIdImage' },
            'a:alphaModFix': { attrs: { amt: '25000' } },
          },
        },
      }),
      slideRelationships: {
        rIdImage: { target: 'ppt/media/bg.png', type: 'image' },
      },
    });

    await expect(getSlideBackgroundFill(fixture.context)).resolves.toEqual({
      type: 'image',
      value: {
        base64: '',
        blob: '',
        opacity: 0.25,
        ref: 'ppt/media/bg.png',
      },
    });
  });

  it('uses layout and master relationship maps for inherited pictures', async () => {
    const picture = {
      'a:blipFill': {
        'a:blip': { attrs: { 'r:embed': 'rIdImage' } },
      },
    };
    await expect(
      getSlideBackgroundFill(
        fillContext({
          layoutRelationships: {
            rIdImage: { target: 'ppt/media/layout.png', type: 'image' },
          },
          options: { imageMode: 'none' },
          slideLayoutContent: background('p:sldLayout', picture),
        }).context,
      ),
    ).resolves.toMatchObject({
      type: 'image',
      value: { ref: 'ppt/media/layout.png' },
    });
    await expect(
      getSlideBackgroundFill(
        fillContext({
          masterRelationships: {
            rIdImage: { target: 'ppt/media/master.png', type: 'image' },
          },
          options: { imageMode: 'none' },
          slideMasterContent: background('p:sldMaster', picture),
        }).context,
      ),
    ).resolves.toMatchObject({
      type: 'image',
      value: { ref: 'ppt/media/master.png' },
    });
  });

  it('selects theme background fills by document order and substitutes phClr', async () => {
    const themeContent = xml({
      'a:theme': {
        'a:themeElements': {
          'a:fmtScheme': {
            'a:bgFillStyleLst': {
              'a:solidFill': [
                {
                  attrs: { order: '20' },
                  'a:srgbClr': { attrs: { val: '222222' } },
                },
                {
                  attrs: { order: '10' },
                  'a:schemeClr': { attrs: { val: 'phClr' } },
                },
              ],
            },
          },
        },
      },
    });
    await expect(
      getSlideBackgroundFill(
        fillContext({
          slideContent: backgroundReference('p:sld', '1001'),
          themeContent,
        }).context,
      ),
    ).resolves.toEqual({ type: 'color', value: '#abcdef' });
    await expect(
      getSlideBackgroundFill(
        fillContext({
          slideContent: backgroundReference('p:sld', '1002'),
          themeContent,
        }).context,
      ),
    ).resolves.toEqual({ type: 'color', value: '#222222' });
  });

  it('resolves a theme picture using theme relationships', async () => {
    const themeContent = xml({
      'a:theme': {
        'a:themeElements': {
          'a:fmtScheme': {
            'a:bgFillStyleLst': {
              'a:blipFill': {
                'a:blip': { attrs: { 'r:embed': 'rIdThemeImage' } },
              },
            },
          },
        },
      },
    });

    await expect(
      getSlideBackgroundFill(
        fillContext({
          options: { imageMode: 'none' },
          slideContent: backgroundReference('p:sld', '1001'),
          themeContent,
          themeRelationships: {
            rIdThemeImage: {
              target: 'ppt/media/theme.png',
              type: 'image',
            },
          },
        }).context,
      ),
    ).resolves.toMatchObject({
      type: 'image',
      value: { ref: 'ppt/media/theme.png' },
    });
  });

  it('uses white when the selected theme fill has no valid value', async () => {
    await expect(
      getSlideBackgroundFill(
        fillContext({
          slideContent: backgroundReference('p:sld', '1001'),
          themeContent: xml({
            'a:theme': {
              'a:themeElements': {
                'a:fmtScheme': {
                  'a:bgFillStyleLst': {
                    'a:solidFill': {
                      'a:srgbClr': { attrs: { val: 'invalid' } },
                    },
                  },
                },
              },
            },
          }),
        }).context,
      ),
    ).resolves.toEqual({ type: 'color', value: '#fff' });
  });

  it.each(['999', '1000', '1003', '1001.5', 'Infinity', 'invalid'])(
    'uses white for invalid or missing theme background index %s',
    async (index) => {
      await expect(
        getSlideBackgroundFill(
          fillContext({
            slideContent: backgroundReference('p:sld', index),
            themeContent: xml({
              'a:theme': {
                'a:themeElements': {
                  'a:fmtScheme': {
                    'a:bgFillStyleLst': {
                      'a:solidFill': {
                        'a:srgbClr': { attrs: { val: '111111' } },
                      },
                    },
                  },
                },
              },
            }),
          }).context,
        ),
      ).resolves.toEqual({ type: 'color', value: '#fff' });
    },
  );

  it('ignores metadata keys and gives malformed theme order lowest priority', async () => {
    const themeContent = xml({
      'a:theme': {
        'a:themeElements': {
          'a:fmtScheme': {
            'a:bgFillStyleLst': {
              attrs: { order: '0' },
              metadata: { attrs: { order: '0' } },
              'a:solidFill': [
                {
                  attrs: { order: 'invalid' },
                  'a:srgbClr': { attrs: { val: '222222' } },
                },
                {
                  attrs: { order: '1' },
                  'a:srgbClr': { attrs: { val: '111111' } },
                },
              ],
            },
          },
        },
      },
    });

    await expect(
      getSlideBackgroundFill(
        fillContext({
          slideContent: backgroundReference('p:sld', '1001'),
          themeContent,
        }).context,
      ),
    ).resolves.toEqual({ type: 'color', value: '#111111' });
  });
});

describe('PPTX shape fill inheritance', () => {
  it('returns null when neither properties nor a style reference define fill', async () => {
    await expect(
      getShapeFill(shape(), fillContext().context, 'slide'),
    ).resolves.toBeNull();
  });

  it('uses a direct solid fill before style, layout, and master candidates', async () => {
    await expect(
      getShapeFill(shape(solid('111111')), fillContext().context, 'slide', {
        slideLayoutSpNode: shape(solid('222222')),
        slideMasterSpNode: shape(solid('333333')),
      }),
    ).resolves.toEqual({ type: 'color', value: '#111111' });
  });

  it('falls through missing direct values to style and inherited candidates', async () => {
    const context = fillContext().context;
    await expect(
      getShapeFill(
        shape(solid('invalid'), {
          'a:srgbClr': { attrs: { val: '123456' } },
        }),
        context,
        'slide',
      ),
    ).resolves.toEqual({ type: 'color', value: '#123456' });
    await expect(
      getShapeFill(shape(), context, 'slide', {
        slideLayoutSpNode: shape(solid('222222')),
        slideMasterSpNode: shape(solid('333333')),
      }),
    ).resolves.toEqual({ type: 'color', value: '#222222' });
    await expect(
      getShapeFill(shape(), context, 'slide', {
        slideMasterSpNode: shape(solid('333333')),
      }),
    ).resolves.toEqual({ type: 'color', value: '#333333' });
  });

  it('uses explicit no-fill to stop style and placeholder inheritance', async () => {
    await expect(
      getShapeFill(
        shape(
          { 'a:noFill': {} },
          { 'a:srgbClr': { attrs: { val: '123456' } } },
        ),
        fillContext().context,
        'slide',
        { slideLayoutSpNode: shape(solid('222222')) },
      ),
    ).resolves.toBeNull();
    await expect(
      getShapeFill(shape(), fillContext().context, 'slide', {
        slideLayoutSpNode: shape({ 'a:noFill': {} }),
        slideMasterSpNode: shape(solid('333333')),
      }),
    ).resolves.toBeNull();
  });

  it('resolves gradient and pattern fills from shape properties', async () => {
    await expect(
      getShapeFill(
        shape({
          'a:gradFill': {
            'a:gsLst': {
              'a:gs': {
                attrs: { pos: '0' },
                'a:srgbClr': { attrs: { val: '123456' } },
              },
            },
          },
        }),
        fillContext().context,
        'slide',
      ),
    ).resolves.toEqual({
      type: 'gradient',
      value: {
        colors: [{ color: '#123456', pos: '0%' }],
        path: 'line',
        rot: 0,
      },
    });
    await expect(
      getShapeFill(
        shape({ 'a:pattFill': { attrs: { prst: 'cross' } } }),
        fillContext().context,
        'slide',
      ),
    ).resolves.toEqual({
      type: 'pattern',
      value: {
        backgroundColor: '#ffffff',
        foregroundColor: '#000000',
        type: 'cross',
      },
    });
  });

  it('resolves a picture fill with opacity from the candidate source', async () => {
    const fixture = fillContext({
      options: { imageMode: 'none' },
      slideRelationships: {
        rIdImage: { target: 'ppt/media/image.png', type: 'image' },
      },
    });
    await expect(
      getShapeFill(
        shape({
          'a:blipFill': {
            'a:blip': {
              attrs: { 'r:embed': 'rIdImage' },
              'a:alphaModFix': { attrs: { amt: '75000' } },
            },
          },
        }),
        fixture.context,
        'slide',
      ),
    ).resolves.toEqual({
      type: 'image',
      value: {
        base64: '',
        blob: '',
        opacity: 0.75,
        ref: 'ppt/media/image.png',
      },
    });
  });

  it('uses inherited picture relationship maps for layout and master shapes', async () => {
    const pictureShape = shape({
      'a:blipFill': {
        'a:blip': { attrs: { 'r:embed': 'rIdImage' } },
      },
    });
    const fixture = fillContext({
      layoutRelationships: {
        rIdImage: { target: 'ppt/media/layout.png', type: 'image' },
      },
      masterRelationships: {
        rIdImage: { target: 'ppt/media/master.png', type: 'image' },
      },
      options: { imageMode: 'none' },
    });

    await expect(
      getShapeFill(shape(), fixture.context, 'slide', {
        slideLayoutSpNode: pictureShape,
      }),
    ).resolves.toMatchObject({
      type: 'image',
      value: { ref: 'ppt/media/layout.png' },
    });
    await expect(
      getShapeFill(shape(), fixture.context, 'slide', {
        slideMasterSpNode: pictureShape,
      }),
    ).resolves.toMatchObject({
      type: 'image',
      value: { ref: 'ppt/media/master.png' },
    });
  });

  it('uses the closest authored group fill without mutating the hierarchy', async () => {
    const outer = xml({ 'p:grpSpPr': solid('111111') });
    const inner = xml({ 'p:grpSpPr': solid('222222') });
    const hierarchy = [outer, inner];

    await expect(
      getShapeFill(shape({ 'a:grpFill': {} }), fillContext().context, 'slide', {
        groupHierarchy: hierarchy,
      }),
    ).resolves.toEqual({ type: 'color', value: '#222222' });
    expect(hierarchy).toEqual([outer, inner]);
  });

  it('continues past an unfilled group but stops at explicit group no-fill', async () => {
    const outer = xml({ 'p:grpSpPr': solid('111111') });
    await expect(
      getShapeFill(shape({ 'a:grpFill': {} }), fillContext().context, 'slide', {
        groupHierarchy: [outer, xml({ 'p:grpSpPr': {} })],
      }),
    ).resolves.toEqual({ type: 'color', value: '#111111' });
    await expect(
      getShapeFill(shape({ 'a:grpFill': {} }), fillContext().context, 'slide', {
        groupHierarchy: [outer, xml({ 'p:grpSpPr': { 'a:noFill': {} } })],
      }),
    ).resolves.toBeNull();
  });
});
