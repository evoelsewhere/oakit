import { describe, expect, it } from 'vitest';

import type { XmlLookupValue } from '../../src/common';
import type { PptxParserContext } from '../../src/formats/pptx/internal/context';
import {
  genTextBody,
  getListLevel,
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
): string {
  return genTextBody(
    textBody,
    xml({}),
    undefined,
    undefined,
    'body',
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
});
