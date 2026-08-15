import { describe, expect, it } from 'vitest';

import type { XmlLookupValue } from '../../src/common';
import type { PptxParserContext } from '../../src/formats/pptx/internal/context';
import {
  getTableBorders,
  getTableCellParams,
  getTableRowParams,
} from '../../src/formats/pptx/internal/table';

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

function solid(color: string) {
  return { 'a:srgbClr': { attrs: { val: color } } };
}

function line(color: string, width = '12700') {
  return { attrs: { w: width }, 'a:solidFill': solid(color) };
}

function section(fill: string, font: string, bold: boolean | string = false) {
  const boldValue =
    typeof bold === 'string' ? bold : bold === true ? 'on' : undefined;
  return {
    'a:tcStyle': {
      'a:fill': { 'a:solidFill': solid(fill) },
      'a:tcBdr': {
        'a:bottom': { 'a:ln': line(fill) },
        'a:left': { 'a:ln': line(fill) },
        'a:right': { 'a:ln': line(fill) },
        'a:top': { 'a:ln': line(fill) },
      },
    },
    'a:tcTxStyle': {
      attrs: { ...(boldValue === undefined ? {} : { b: boldValue }) },
      'a:solidFill': solid(font),
    },
  };
}

const noTableFlags = {
  isBandColAttr: 0,
  isBandRowAttr: 0,
  isFrstColAttr: 0,
  isFrstRowAttr: 0,
  isLstColAttr: 0,
  isLstRowAttr: 0,
};

describe('PPTX table borders', () => {
  it('parses all four table-level border directions', () => {
    const borders = getTableBorders(
      xml({
        'a:bottom': { 'a:ln': line('111111', '12700') },
        'a:left': { 'a:ln': line('222222', '25400') },
        'a:right': { 'a:ln': line('333333', '38100') },
        'a:top': { 'a:ln': line('444444', '50800') },
      }),
      context(),
    );

    expect(borders).toMatchObject({
      bottom: { borderColor: '#111111', borderWidth: 1 },
      left: { borderColor: '#222222', borderWidth: 2 },
      right: { borderColor: '#333333', borderWidth: 3 },
      top: { borderColor: '#444444', borderWidth: 4 },
    });
  });

  it('omits directions that are not authored', () => {
    expect(getTableBorders(xml({}), context())).toEqual({});
  });
});

describe('PPTX table cell styles', () => {
  it('uses section fill, font, boldness, and borders', () => {
    const style = xml({ 'a:firstRow': section('112233', '445566', true) });
    const cell = getTableCellParams(xml({}), style, 'a:firstRow', context());

    expect(cell).toMatchObject({
      borders: {
        bottom: { borderColor: '#112233' },
        left: { borderColor: '#112233' },
        right: { borderColor: '#112233' },
        top: { borderColor: '#112233' },
      },
      fillColor: '#112233',
      fontBold: true,
      fontColor: '#445566',
    });
  });

  it('gives authored cell fill and borders precedence over table styles', () => {
    const style = xml({
      'a:firstRow': section('111111', '222222'),
      'a:wholeTbl': section('333333', '444444'),
    });
    const cell = getTableCellParams(
      xml({
        'a:tcPr': {
          'a:lnB': line('aaaaaa'),
          'a:lnL': line('bbbbbb'),
          'a:lnR': line('cccccc'),
          'a:lnT': line('dddddd'),
          'a:solidFill': solid('abcdef'),
        },
      }),
      style,
      'a:firstRow',
      context(),
    );

    expect(cell).toMatchObject({
      borders: {
        bottom: { borderColor: '#aaaaaa' },
        left: { borderColor: '#bbbbbb' },
        right: { borderColor: '#cccccc' },
        top: { borderColor: '#dddddd' },
      },
      fillColor: '#abcdef',
      fontColor: '#222222',
    });
  });

  it('falls back to whole-table borders when a section omits them', () => {
    const style = xml({
      'a:firstRow': {
        'a:tcStyle': { 'a:fill': { 'a:solidFill': solid('111111') } },
      },
      'a:wholeTbl': section('333333', '444444'),
    });
    const cell = getTableCellParams(xml({}), style, 'a:firstRow', context());

    expect(cell.borders).toMatchObject({
      bottom: { borderColor: '#333333' },
      left: { borderColor: '#333333' },
      right: { borderColor: '#333333' },
      top: { borderColor: '#333333' },
    });
  });

  it('returns neutral optional styling when no source or style exists', () => {
    expect(
      getTableCellParams(xml({}), undefined, undefined, context()),
    ).toEqual({ borders: {}, vAlign: 'up' });
  });

  it('falls back to the section fill when the cell has no solid fill', () => {
    const cell = getTableCellParams(
      xml({ 'a:tcPr': { 'a:noFill': {} } }),
      xml({ 'a:firstRow': section('123456', 'abcdef') }),
      'a:firstRow',
      context(),
    );

    expect(cell.fillColor).toBe('#123456');
  });

  it.each([
    ['1', true],
    ['on', true],
    ['true', true],
    ['0', false],
    ['off', false],
    ['false', false],
  ] as const)('parses cell bold value %s', (bold, expected) => {
    const cell = getTableCellParams(
      xml({}),
      xml({ 'a:firstRow': section('123456', 'abcdef', bold) }),
      'a:firstRow',
      context(),
    );

    expect(cell.fontBold).toBe(expected);
  });

  it('omits an unrecognized cell bold value', () => {
    const cell = getTableCellParams(
      xml({}),
      xml({ 'a:firstRow': section('123456', 'abcdef', 'yes') }),
      'a:firstRow',
      context(),
    );

    expect(cell).not.toHaveProperty('fontBold');
  });
});

describe('PPTX table row styles', () => {
  const rows = [xml({}), xml({}), xml({}), xml({})];
  const style = xml({
    'a:band1H': section('111111', '121212', true),
    'a:band2H': section('222222', '232323'),
    'a:firstRow': section('333333', '343434', true),
    'a:lastRow': section('444444', '454545', true),
    'a:wholeTbl': section('555555', '565656'),
  });

  it('uses whole-table defaults for an ordinary row', () => {
    expect(getTableRowParams(rows, 1, noTableFlags, style, context())).toEqual({
      fillColor: '#555555',
      fontColor: '#565656',
    });
  });

  it('allows the first-row style to override whole-table defaults', () => {
    expect(
      getTableRowParams(
        rows,
        0,
        { ...noTableFlags, isFrstRowAttr: 1 },
        style,
        context(),
      ),
    ).toEqual({
      fillColor: '#333333',
      fontBold: true,
      fontColor: '#343434',
    });
  });

  it.each([
    [1, '#111111', '#121212', true],
    [2, '#222222', '#232323', undefined],
    [3, '#111111', '#121212', true],
  ] as const)(
    'applies row banding at index %d',
    (index, fillColor, fontColor, fontBold) => {
      expect(
        getTableRowParams(
          rows,
          index,
          { ...noTableFlags, isBandRowAttr: 1 },
          style,
          context(),
        ),
      ).toEqual({
        fillColor,
        ...(fontBold === undefined ? {} : { fontBold }),
        fontColor,
      });
    },
  );

  it('applies last-row styling after banding', () => {
    expect(
      getTableRowParams(
        rows,
        3,
        { ...noTableFlags, isBandRowAttr: 1, isLstRowAttr: 1 },
        style,
        context(),
      ),
    ).toEqual({
      fillColor: '#444444',
      fontBold: true,
      fontColor: '#454545',
    });
  });

  it('preserves whole-table values when a band style is partial', () => {
    const partialStyle = xml({
      'a:band1H': {
        'a:tcTxStyle': {
          attrs: { b: 'off' },
          'a:solidFill': solid('121212'),
        },
      },
      'a:wholeTbl': section('555555', '565656', true),
    });

    expect(
      getTableRowParams(
        rows,
        1,
        { ...noTableFlags, isBandRowAttr: 1 },
        partialStyle,
        context(),
      ),
    ).toEqual({
      fillColor: '#555555',
      fontBold: false,
      fontColor: '#121212',
    });
  });

  it('preserves whole-table values when the selected band is absent', () => {
    const partialStyle = xml({
      'a:wholeTbl': section('555555', '565656', true),
    });

    expect(
      getTableRowParams(
        rows,
        2,
        { ...noTableFlags, isBandRowAttr: 1 },
        partialStyle,
        context(),
      ),
    ).toEqual({
      fillColor: '#555555',
      fontBold: true,
      fontColor: '#565656',
    });
  });

  it('does not apply row banding to the first row', () => {
    expect(
      getTableRowParams(
        rows,
        0,
        { ...noTableFlags, isBandRowAttr: 1 },
        style,
        context(),
      ),
    ).toEqual({ fillColor: '#555555', fontColor: '#565656' });
  });

  it('does not treat non-one style flags as enabled', () => {
    expect(
      getTableRowParams(
        rows,
        1,
        {
          ...noTableFlags,
          isBandRowAttr: 2,
          isFrstRowAttr: 2,
          isLstRowAttr: 2,
        },
        style,
        context(),
      ),
    ).toEqual({ fillColor: '#555555', fontColor: '#565656' });
  });

  it('does not apply first-row styling to a later row', () => {
    expect(
      getTableRowParams(
        rows,
        1,
        { ...noTableFlags, isFrstRowAttr: 1 },
        style,
        context(),
      ),
    ).toEqual({ fillColor: '#555555', fontColor: '#565656' });
  });

  it('does not apply last-row styling before the final row', () => {
    expect(
      getTableRowParams(
        rows,
        2,
        { ...noTableFlags, isLstRowAttr: 1 },
        style,
        context(),
      ),
    ).toEqual({ fillColor: '#555555', fontColor: '#565656' });
  });

  it('lets last-row style win when a single row is both first and last', () => {
    expect(
      getTableRowParams(
        [xml({})],
        0,
        {
          ...noTableFlags,
          isFrstRowAttr: 1,
          isLstRowAttr: 1,
        },
        style,
        context(),
      ),
    ).toEqual({
      fillColor: '#444444',
      fontBold: true,
      fontColor: '#454545',
    });
  });

  it('merges partial first-row and last-row styles with defaults', () => {
    const partialStyle = xml({
      'a:firstRow': {
        'a:tcStyle': { 'a:fill': { 'a:solidFill': solid('111111') } },
      },
      'a:lastRow': {
        'a:tcTxStyle': {
          attrs: { b: 'false' },
          'a:solidFill': solid('222222'),
        },
      },
      'a:wholeTbl': section('555555', '565656', true),
    });

    expect(
      getTableRowParams(
        [xml({})],
        0,
        {
          ...noTableFlags,
          isFrstRowAttr: 1,
          isLstRowAttr: 1,
        },
        partialStyle,
        context(),
      ),
    ).toEqual({
      fillColor: '#111111',
      fontBold: false,
      fontColor: '#222222',
    });
  });

  it('returns no optional row styling without a table style', () => {
    expect(
      getTableRowParams(rows, 0, noTableFlags, undefined, context()),
    ).toEqual({});
  });
});
