import { describe, expect, it } from 'vitest';

import type { XmlLookupValue } from '../../src/common';
import {
  findOMath,
  latexFormart,
  parseAccent,
  parseBar,
  parseBox,
  parseDelimiter,
  parseEqArr,
  parseFraction,
  parseFunction,
  parseGroupChr,
  parseLimit,
  parseMatrix,
  parseNary,
  parseOMath,
  parseRadical,
  parseSubscript,
  parseSuperscript,
} from '../../src/formats/pptx/internal/math';

function xml(value: object): XmlLookupValue {
  return value as unknown as XmlLookupValue;
}

function expression(value: string): XmlLookupValue {
  return xml({ 'm:r': { 'm:t': value } });
}

describe('Office Math discovery', () => {
  it('finds singleton and repeated equations recursively in document order', () => {
    const first = expression('x');
    const second = expression('y');
    const third = expression('z');

    expect(
      findOMath({
        before: null,
        nested: [
          { 'm:oMath': first },
          { deeper: { 'm:oMath': [second, third] } },
        ],
      }),
    ).toEqual([first, second, third]);
  });

  it.each([undefined, null, 'text', 3, true])(
    'ignores non-container input %j',
    (value) => expect(findOMath(value)).toEqual([]),
  );
});

describe('Office Math structures', () => {
  it('parses fractions, superscripts, and subscripts', () => {
    expect(
      parseFraction(
        xml({ 'm:num': expression('1'), 'm:den': expression('2') }),
      ),
    ).toBe('\\frac{1}{2}');
    expect(
      parseSuperscript(
        xml({ 'm:e': expression('x'), 'm:sup': expression('2') }),
      ),
    ).toBe('x^{2}');
    expect(
      parseSubscript(xml({ 'm:e': expression('x'), 'm:sub': expression('i') })),
    ).toBe('x_{i}');
  });

  it('parses square and indexed radicals', () => {
    expect(parseRadical(xml({ 'm:e': expression('x') }))).toBe('\\sqrt{x}');
    expect(
      parseRadical(xml({ 'm:deg': expression('3'), 'm:e': expression('x') })),
    ).toBe('\\sqrt[3]{x}');
  });

  it('parses matrices and equation arrays', () => {
    expect(
      parseMatrix(
        xml({
          'm:mr': [
            { 'm:e': [expression('a'), expression('b')] },
            { 'm:e': [expression('c'), expression('d')] },
          ],
        }),
      ),
    ).toBe('\\begin{matrix} a & b \\\\ c & d \\end{matrix}');
    expect(
      parseEqArr(xml({ 'm:e': [expression('x=1'), expression('y=2')] })),
    ).toBe('\\begin{cases} x=1 \\\\ y=2 \\end{cases}');
  });

  it('parses n-ary operators with explicit and default characters', () => {
    const values = {
      'm:e': expression('x'),
      'm:sub': expression('0'),
      'm:sup': expression('n'),
    };
    expect(parseNary(xml(values))).toBe('∫_{0}^{n}{x}');
    expect(
      parseNary(
        xml({
          ...values,
          'm:naryPr': { 'm:chr': { attrs: { 'm:val': '∑' } } },
        }),
      ),
    ).toBe('∑_{0}^{n}{x}');
  });

  it('parses upper and lower limits', () => {
    const limit = xml({ 'm:e': expression('lim'), 'm:lim': expression('x→0') });
    expect(parseLimit(limit, 'low')).toBe('lim_{x→0}');
    expect(parseLimit(limit, 'upp')).toBe('lim^{x→0}');
  });

  it('parses default, paired, and one-sided delimiters', () => {
    expect(parseDelimiter(xml({ 'm:e': expression('x') }))).toBe(
      '\\left(x\\right)',
    );
    expect(
      parseDelimiter(
        xml({
          'm:dPr': {
            'm:begChr': { attrs: { 'm:val': '[' } },
            'm:endChr': { attrs: { 'm:val': ']' } },
          },
          'm:e': expression('x'),
        }),
      ),
    ).toBe('\\left[x\\right]');
    expect(
      parseDelimiter(
        xml({
          'm:dPr': { 'm:begChr': { attrs: { 'm:val': '|' } } },
          'm:e': expression('x'),
        }),
      ),
    ).toBe('|x');
    expect(
      parseDelimiter(
        xml({
          'm:dPr': { 'm:endChr': { attrs: { 'm:val': '|' } } },
          'm:e': expression('x'),
        }),
      ),
    ).toBe('x|');
  });

  it('parses functions, group characters, boxes, and bars', () => {
    expect(
      parseFunction(
        xml({ 'm:fName': expression('sin'), 'm:e': expression('x') }),
      ),
    ).toBe('\\sin{x}');
    expect(
      parseGroupChr(
        xml({
          'm:groupChrPr': { 'm:chr': { attrs: { 'm:val': '⏞' } } },
          'm:e': expression('x'),
        }),
      ),
    ).toBe('⏞x⏞');
    expect(parseGroupChr(xml({ 'm:e': expression('x') }))).toBe('x');
    expect(parseBox(xml({ 'm:e': expression('x') }))).toBe('\\boxed{x}');
    expect(
      parseBar(
        xml({
          'm:barPr': { 'm:pos': { attrs: { 'm:val': 'top' } } },
          'm:e': expression('x'),
        }),
      ),
    ).toBe('\\overline{x}');
    expect(parseBar(xml({ 'm:e': expression('x') }))).toBe('\\underline{x}');
  });

  it.each([
    ['\u0301', 'acute'],
    ['\u0300', 'grave'],
    ['\u0302', 'hat'],
    ['\u0303', 'tilde'],
    ['\u0304', 'bar'],
    ['\u0306', 'breve'],
    ['\u0307', 'dot'],
    ['\u0308', 'ddot'],
    ['\u030A', 'mathring'],
    ['\u030B', 'H'],
    ['\u030C', 'check'],
    ['\u0327', 'c'],
  ] as const)('maps accent %s to \\%s', (character, command) => {
    expect(
      parseAccent(
        xml({
          'm:accPr': { 'm:chr': { attrs: { 'm:val': character } } },
          'm:e': expression('x'),
        }),
      ),
    ).toBe(`\\${command}{x}`);
  });

  it('preserves custom accents and defaults to a caret', () => {
    expect(
      parseAccent(
        xml({
          'm:accPr': { 'm:chr': { attrs: { 'm:val': '→' } } },
          'm:e': expression('x'),
        }),
      ),
    ).toBe('\\→{x}');
    expect(parseAccent(xml({ 'm:e': expression('x') }))).toBe('\\^{x}');
  });
});

describe('Office Math composition', () => {
  it.each([
    [
      'm:nary',
      {
        'm:e': expression('x'),
        'm:sub': expression('0'),
        'm:sup': expression('n'),
      },
      '∫_{0}^{n}{x}',
    ],
    [
      'm:limLow',
      { 'm:e': expression('lim'), 'm:lim': expression('0') },
      'lim_{0}',
    ],
    [
      'm:limUpp',
      { 'm:e': expression('lim'), 'm:lim': expression('∞') },
      'lim^{∞}',
    ],
    ['m:d', { 'm:e': expression('x') }, '\\left(x\\right)'],
    [
      'm:func',
      { 'm:fName': expression('cos'), 'm:e': expression('x') },
      '\\cos{x}',
    ],
    ['m:groupChr', { 'm:e': expression('x') }, 'x'],
    [
      'm:eqArr',
      { 'm:e': [expression('x'), expression('y')] },
      '\\begin{cases} x \\\\ y \\end{cases}',
    ],
    ['m:bar', { 'm:e': expression('x') }, '\\underline{x}'],
    ['m:acc', { 'm:e': expression('x') }, '\\^{x}'],
    [
      'm:m',
      { 'm:mr': { 'm:e': [expression('a'), expression('b')] } },
      '\\begin{matrix} a & b \\end{matrix}',
    ],
  ] as const)('dispatches %s through parseOMath', (key, value, expected) => {
    expect(parseOMath(xml({ [key]: value }))).toBe(expected);
  });

  it('dispatches every supported part and ignores unknown metadata', () => {
    const formula = parseOMath(
      xml({
        unknown: { value: 'ignored' },
        'm:f': { 'm:num': expression('1'), 'm:den': expression('2') },
        'm:sSup': { 'm:e': expression('x'), 'm:sup': expression('2') },
        'm:sSub': { 'm:e': expression('y'), 'm:sub': expression('i') },
        'm:rad': { 'm:e': expression('z') },
        'm:borderBox': { 'm:e': expression('b') },
      }),
    );

    expect(formula).toBe('\\frac{1}{2}x^{2}y_{i}\\sqrt{z}\\boxed{b}');
  });

  it('orders parts by direct and control run order', () => {
    const ordered = parseOMath(
      xml({
        'm:f': {
          'm:fPr': { 'm:ctrlPr': { 'a:rPr': { attrs: { order: '30' } } } },
          'm:num': expression('1'),
          'm:den': expression('2'),
        },
        'm:r': [
          { 'a:rPr': { attrs: { order: '20' } }, 'm:t': 'second' },
          { 'a:rPr': { attrs: { order: '10' } }, 'm:t': 'first' },
        ],
      }),
    );

    expect(ordered).toBe('firstsecond\\frac{1}{2}');
  });

  it('handles arrays, empty values, text scalars, and empty structures', () => {
    expect(
      parseOMath([expression('a'), expression('b')] as XmlLookupValue),
    ).toBe('ab');
    expect(parseOMath(xml({ 'm:t': 'plain' }))).toBe('plain');
    expect(parseOMath(xml({ 'm:t': { value: 'not-scalar' } }))).toBe('');
    expect(parseOMath(undefined)).toBe('');
    expect(parseOMath(xml({}))).toBe('');
  });

  it('decodes each XML entity exactly once', () => {
    expect(latexFormart('&lt;&gt;&amp;&apos;&quot;')).toBe('<>&\'"');
    expect(latexFormart('&amp;lt;')).toBe('&lt;');
    expect(latexFormart('unchanged')).toBe('unchanged');
  });
});
