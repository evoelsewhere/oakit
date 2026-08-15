import { describe, expect, it } from 'vitest';

import type { XmlLookupValue } from '../../src/common';
import type { PptxParserContext } from '../../src/formats/pptx/internal/context';
import {
  genTextBody,
  getListLevel,
  getListType,
  getSpanStyleInfo,
  getTextNodeValue,
} from '../../src/formats/pptx/internal/text';

function xml(value: object): XmlLookupValue {
  return value as unknown as XmlLookupValue;
}

function context(
  overrides: Partial<PptxParserContext> = {},
): PptxParserContext {
  return {
    defaultTextStyle: xml({}),
    slideContent: xml({}),
    slideLayoutContent: xml({}),
    slideMasterContent: xml({}),
    slideMasterTextStyles: xml({}),
    slideResObj: {},
    themeContent: xml({}),
    ...overrides,
  } as PptxParserContext;
}

function run(text: string, order: number): XmlLookupValue {
  return xml({ attrs: { order }, 'a:rPr': {}, 'a:t': text });
}

function field(text: string, order: number): XmlLookupValue {
  return xml({ attrs: { order }, 'a:rPr': {}, 'a:t': text });
}

function lineBreak(order: number): XmlLookupValue {
  return xml({ attrs: { order }, 'a:rPr': {} });
}

function renderBody(
  textBody: XmlLookupValue,
  parserContext: PptxParserContext = context(),
  shape: XmlLookupValue = xml({}),
  type = 'body',
): string {
  return genTextBody(
    textBody,
    shape,
    undefined,
    undefined,
    type,
    parserContext,
  );
}

function render(paragraph: object): string {
  return renderBody(xml({ 'a:p': paragraph }));
}

function listParagraph(
  text: string,
  level: string,
  type: 'ol' | 'ul' = 'ul',
): XmlLookupValue {
  return xml({
    'a:pPr': {
      attrs: { lvl: level },
      [type === 'ul' ? 'a:buChar' : 'a:buAutoNum']: {},
    },
    'a:r': run(text, 10),
  });
}

describe('PowerPoint text primitives', () => {
  it.each([
    ['direct string', 'Text', 'Text'],
    ['wrapped text', xml({ value: 'Wrapped' }), 'Wrapped'],
    ['missing text', xml({}), undefined],
  ])('reads %s without coercing absent values', (_name, node, expected) => {
    expect(getTextNodeValue(node)).toBe(expected);
  });

  it.each([
    [undefined, 0],
    ['0', 0],
    ['1', 1],
    ['2', 2],
    ['3', 3],
    ['4', 4],
    ['5', 5],
    ['6', 6],
    ['7', 7],
    ['8', 8],
    ['-1', 0],
    ['9', 0],
    ['1junk', 0],
    ['1.5', 0],
    ['Infinity', 0],
    ['', 0],
  ])('normalizes list level %s to %i', (level, expected) => {
    const attrs = level === undefined ? {} : { lvl: level };
    expect(getListLevel(xml({ 'a:pPr': { attrs } }))).toBe(expected);
  });

  it.each([undefined, 'malformed'])(
    'uses the first text style level when paragraph level is %s',
    (level) => {
      const paragraphProperties =
        level === undefined ? {} : { attrs: { lvl: level } };
      expect(
        renderBody(
          xml({
            'a:lstStyle': {
              'a:lvl1pPr': { 'a:defRPr': { attrs: { sz: '2200' } } },
            },
            'a:p': {
              'a:pPr': paragraphProperties,
              'a:r': run('Level', 10),
            },
          }),
        ),
      ).toContain('font-size: 22pt;');
    },
  );

  it.each([
    ['character bullet', { 'a:buChar': {} }, 'ul'],
    ['automatic number', { 'a:buAutoNum': {} }, 'ol'],
    [
      'character bullet before automatic number',
      { 'a:buChar': {}, 'a:buAutoNum': {} },
      'ul',
    ],
    ['unrelated paragraph properties', { attrs: { algn: 'ctr' } }, ''],
    ['missing paragraph properties', undefined, ''],
  ])('recognizes %s as %j', (_name, properties, expected) => {
    expect(
      getListType(xml(properties === undefined ? {} : { 'a:pPr': properties })),
    ).toBe(expected);
  });

  it('returns no HTML for a text body without paragraphs', () => {
    expect(renderBody(xml({}))).toBe('');
  });

  it('preserves an empty paragraph with a non-breaking glyph', () => {
    expect(render({})).toBe(
      '<p style="text-align: left;"><span style="font-size: 18pt;">&nbsp;</span></p>',
    );
  });

  it('escapes text and preserves tabs and whitespace deterministically', () => {
    expect(render({ 'a:r': run('A\t B\nC &amp; <x>', 10) })).toBe(
      '<p style="text-align: left;"><span style="font-size: 18pt;">A&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;B&nbsp;C&nbsp;&amp;&nbsp;&lt;x&gt;</span></p>',
    );
  });
});

describe('PowerPoint text run ordering', () => {
  it('renders runs, fields, and breaks in document order', () => {
    expect(
      render({
        'a:r': [run('First', 10), run('Last', 40)],
        'a:fld': field('Field', 30),
        'a:br': lineBreak(20),
      }),
    ).toBe(
      '<p style="text-align: left;"><span style="font-size: 18pt;">First</span><br><span style="font-size: 18pt;">FieldLast</span></p>',
    );
  });

  it('renders every field when a paragraph contains no ordinary run', () => {
    expect(
      render({
        'a:fld': [field('One', 10), field('Two', 20)],
      }),
    ).toBe(
      '<p style="text-align: left;"><span style="font-size: 18pt;">OneTwo</span></p>',
    );
  });

  it('renders a break-only paragraph as a semantic line break', () => {
    expect(render({ 'a:br': lineBreak(10) })).toBe(
      '<p style="text-align: left;"><br></p>',
    );
  });

  it('preserves consecutive authored line breaks', () => {
    expect(
      render({
        'a:r': run('Before', 10),
        'a:br': [lineBreak(20), lineBreak(30)],
      }),
    ).toBe(
      '<p style="text-align: left;"><span style="font-size: 18pt;">Before</span><br><br></p>',
    );
  });

  it('merges adjacent equal styles but splits distinct run styles', () => {
    const plain = run('Plain', 10);
    const same = run('Same', 20);
    const bold = xml({
      attrs: { order: 30 },
      'a:rPr': { attrs: { b: '1' } },
      'a:t': 'Bold',
    });

    expect(render({ 'a:r': [plain, same, bold] })).toBe(
      '<p style="text-align: left;"><span style="font-size: 18pt;">PlainSame</span><span style="font-size: 18pt;font-weight: bold;">Bold</span></p>',
    );
  });

  it('keeps safe links isolated from adjacent unlinked text', () => {
    const linked = xml({
      attrs: { order: 20 },
      'a:rPr': { 'a:hlinkClick': { attrs: { 'r:id': 'link' } } },
      'a:t': 'Link',
    });
    expect(
      genTextBody(
        xml({
          'a:p': {
            'a:r': [run('Before', 10), linked, run('After', 30)],
          },
        }),
        xml({}),
        undefined,
        undefined,
        'body',
        context({
          slideResObj: {
            link: { target: 'https://example.com', type: 'hyperlink' },
          },
        }),
      ),
    ).toBe(
      '<p style="text-align: left;"><span style="font-size: 18pt;">Before</span><span style="font-size: 18pt;"><a href="https://example.com" target="_blank" rel="noopener noreferrer">Link</a></span><span style="font-size: 18pt;">After</span></p>',
    );
  });
});

describe('PowerPoint text lists', () => {
  it('renders nested list levels inside their parent list item', () => {
    expect(
      renderBody(
        xml({
          'a:p': [
            listParagraph('Parent', '0'),
            listParagraph('Child', '1'),
            listParagraph('Sibling', '0'),
          ],
        }),
      ),
    ).toBe(
      '<ul><li><p style="text-align: left;"><span style="font-size: 18pt;">Parent</span></p><ul><li><p style="text-align: left;"><span style="font-size: 18pt;">Child</span></p></li></ul></li><li><p style="text-align: left;"><span style="font-size: 18pt;">Sibling</span></p></li></ul>',
    );
  });

  it('switches list type without leaving an item or list open', () => {
    expect(
      renderBody(
        xml({
          'a:p': [
            listParagraph('Bullet', '0'),
            listParagraph('Number', '0', 'ol'),
            { 'a:r': run('Plain', 10) },
          ],
        }),
      ),
    ).toBe(
      '<ul><li><p style="text-align: left;"><span style="font-size: 18pt;">Bullet</span></p></li></ul><ol><li><p style="text-align: left;"><span style="font-size: 18pt;">Number</span></p></li></ol><p style="text-align: left;"><span style="font-size: 18pt;">Plain</span></p>',
    );
  });

  it('normalizes an initial deep level into the first list level', () => {
    expect(
      renderBody(
        xml({
          'a:p': [listParagraph('Deep', '8'), listParagraph('Root', '0')],
        }),
      ),
    ).toBe(
      '<ul><li><p style="text-align: left;"><span style="font-size: 18pt;">Deep</span></p></li><li><p style="text-align: left;"><span style="font-size: 18pt;">Root</span></p></li></ul>',
    );
  });

  it('does not apply paragraph indentation CSS to list items', () => {
    const paragraph = listParagraph('List', '0') as unknown as Record<
      string,
      XmlLookupValue
    >;
    paragraph['a:pPr'] = xml({
      attrs: { indent: '-6350', lvl: '0', marL: '12700' },
      'a:buChar': {},
    });

    const html = renderBody(xml({ 'a:p': paragraph }));
    expect(html).not.toContain('margin-left:');
    expect(html).not.toContain('text-indent:');
  });

  it('closes every nested list from deepest to shallowest before plain text', () => {
    const html = renderBody(
      xml({
        'a:p': [
          listParagraph('Root', '0'),
          listParagraph('Child', '1'),
          listParagraph('Grand', '2', 'ol'),
          { 'a:r': run('Plain', 10) },
        ],
      }),
    );

    expect(html).toContain(
      'Grand</span></p></li></ol></li></ul></li></ul><p style="text-align: left;">',
    );
  });

  it('closes only deeper lists in reverse order before an ancestor sibling', () => {
    const html = renderBody(
      xml({
        'a:p': [
          listParagraph('Root', '0'),
          listParagraph('Child', '1'),
          listParagraph('Grand', '2', 'ol'),
          listParagraph('Sibling', '0'),
        ],
      }),
    );

    expect(html).toContain(
      'Grand</span></p></li></ol></li></ul></li><li><p style="text-align: left;">',
    );
  });
});

describe('PowerPoint text paragraph CSS', () => {
  it('assembles spacing and indentation from resolved paragraph metrics', () => {
    const paragraphProperties = {
      attrs: { indent: '-6350', marL: '12700' },
      'a:lnSpc': { 'a:spcPct': { attrs: { val: '150000' } } },
      'a:spcBef': { 'a:spcPct': { attrs: { val: '50000' } } },
      'a:spcAft': { 'a:spcPts': { attrs: { val: '1200' } } },
    };

    expect(
      render({ 'a:pPr': paragraphProperties, 'a:r': run('Metrics', 10) }),
    ).toBe(
      '<p style="text-align: left;line-height: 1.5;margin-top: 0.5em;margin-bottom: 12pt;margin-left: 1pt;text-indent: -0.5pt;"><span style="font-size: 18pt;">Metrics</span></p>',
    );
  });

  it.each([
    [
      'line spacing only',
      { 'a:lnSpc': { 'a:spcPct': { attrs: { val: '125000' } } } },
      'text-align: left;line-height: 1.25;',
    ],
    [
      'space before only',
      { 'a:spcBef': { 'a:spcPts': { attrs: { val: '600' } } } },
      'text-align: left;margin-top: 6pt;',
    ],
    [
      'space after only',
      { 'a:spcAft': { 'a:spcPts': { attrs: { val: '700' } } } },
      'text-align: left;margin-bottom: 7pt;',
    ],
  ])('emits only %s', (_name, properties, expectedStyle) => {
    const html = render({
      'a:pPr': properties,
      'a:r': run('Spacing', 10),
    });
    expect(html).toContain(`<p style="${expectedStyle}">`);
    expect(html).not.toContain('undefined');
  });

  it.each([
    ['left margin only', { marL: '12700' }, 'margin-left: 1pt;'],
    ['text indent only', { indent: '-6350' }, 'text-indent: -0.5pt;'],
  ])('emits only %s', (_name, attrs, expectedDeclaration) => {
    const html = render({
      'a:pPr': { attrs },
      'a:r': run('Indent', 10),
    });
    expect(html).toContain(expectedDeclaration);
    expect(html).not.toContain('undefined');
  });
});

describe('PowerPoint span CSS assembly', () => {
  it('serializes every supported authored run style in a stable order', () => {
    const style = getSpanStyleInfo(
      xml({
        'a:rPr': {
          attrs: {
            b: '1',
            baseline: '-25000',
            i: '1',
            spc: '125',
            strike: 'sngStrike',
            sz: '2400',
            u: 'sng',
          },
          'a:effectLst': {
            'a:outerShdw': {
              attrs: { blurRad: '25400', dir: '0', dist: '12700' },
              'a:srgbClr': { attrs: { val: '654321' } },
            },
          },
          'a:latin': { attrs: { typeface: 'Font Name' } },
          'a:solidFill': { 'a:srgbClr': { attrs: { val: '123456' } } },
        },
        'a:t': 'Styled',
      }),
      xml({}),
      xml({}),
      undefined,
      undefined,
      undefined,
      'body',
      undefined,
      undefined,
      context(),
    );

    expect(style).toEqual({
      hasLink: false,
      linkURL: null,
      styleText:
        'color: #123456;font-size: 24pt;font-family: "Font Name";font-weight: bold;font-style: italic;text-decoration: underline;text-decoration-line: line-through;letter-spacing: 1.25pt;vertical-align: sub;text-shadow: 1pt 0pt 2pt #654321;',
      text: 'Styled',
    });
  });

  it('serializes gradient text with exact stops and rotation', () => {
    const style = getSpanStyleInfo(
      xml({
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
        'a:t': 'Gradient',
      }),
      xml({}),
      xml({}),
      undefined,
      undefined,
      undefined,
      'body',
      undefined,
      undefined,
      context(),
    );

    expect(style.styleText).toBe(
      'background: linear-gradient(90deg, #000000 0%, #ffffff 100%); background-clip: text; color: transparent;font-size: 18pt;',
    );
  });

  it('uses a shape font reference through the text-body boundary', () => {
    const html = renderBody(
      xml({ 'a:p': { 'a:r': run('Referenced', 10) } }),
      context(),
      xml({
        'p:style': {
          'a:fontRef': {
            'a:srgbClr': { attrs: { val: 'abcdef' } },
          },
        },
      }),
    );

    expect(html).toContain('color: #abcdef;');
  });

  it('uses default text style for table cells instead of master text style', () => {
    const parserContext = context({
      defaultTextStyle: xml({
        'a:lvl1pPr': { 'a:defRPr': { attrs: { sz: '3300' } } },
      }),
      slideMasterTextStyles: xml({
        'p:bodyStyle': {
          'a:lvl1pPr': { 'a:defRPr': { attrs: { sz: '4400' } } },
        },
      }),
    });
    const textBody = xml({ 'a:p': { 'a:r': run('Cell', 10) } });
    const tableHtml = renderBody(
      textBody,
      parserContext,
      xml({ 'a:tcPr': {} }),
    );
    const shapeHtml = renderBody(textBody, parserContext, xml({}));

    expect(tableHtml).toContain('font-size: 33pt;');
    expect(tableHtml).not.toContain('font-size: 44pt;');
    expect(shapeHtml).toContain('font-size: 44pt;');
    expect(shapeHtml).not.toContain('font-size: 33pt;');
  });

  it('does not let a default table style affect an ordinary text shape', () => {
    const html = renderBody(
      xml({ 'a:p': { 'a:r': run('Shape', 10) } }),
      context({
        defaultTextStyle: xml({
          'a:lvl1pPr': { 'a:defRPr': { attrs: { sz: '3300' } } },
        }),
      }),
      xml({}),
    );

    expect(html).toContain('font-size: 18pt;');
    expect(html).not.toContain('font-size: 33pt;');
  });
});
