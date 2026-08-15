import { describe, expect, it } from 'vitest';

import { createDrawingGuideResolver } from '../../src/formats/pptx/internal/custom-geometry-guide';

function resolver(
  formulas: Readonly<Record<string, string>> = {},
  width = 320,
  height = 160,
) {
  return createDrawingGuideResolver(
    width,
    height,
    new Map(Object.entries(formulas)),
  );
}

describe('DrawingML custom geometry guide formulas', () => {
  it.each([
    ['3cd4', 16_200_000],
    ['3cd8', 8_100_000],
    ['5cd8', 13_500_000],
    ['7cd8', 18_900_000],
    ['b', 160],
    ['cd2', 10_800_000],
    ['cd4', 5_400_000],
    ['cd8', 2_700_000],
    ['h', 160],
    ['hc', 160],
    ['hd2', 80],
    ['hd3', 160 / 3],
    ['hd4', 40],
    ['hd5', 32],
    ['hd6', 160 / 6],
    ['hd8', 20],
    ['l', 0],
    ['ls', 320],
    ['r', 320],
    ['ss', 160],
    ['ssd2', 80],
    ['ssd4', 40],
    ['ssd6', 160 / 6],
    ['ssd8', 20],
    ['ssd16', 10],
    ['ssd32', 5],
    ['t', 0],
    ['vc', 80],
    ['w', 320],
    ['wd2', 160],
    ['wd3', 320 / 3],
    ['wd4', 80],
    ['wd5', 64],
    ['wd6', 320 / 6],
    ['wd8', 40],
    ['wd10', 32],
    ['wd32', 10],
  ] as const)('resolves the built-in %s guide', (name, expected) => {
    expect(resolver()(name)).toBe(expected);
  });

  it.each([
    ['abs -9', 9],
    ['+/ 8 4 3', 4],
    ['+- 8 4 3', 9],
    ['at2 0 1', 5_400_000],
    ['cos 10 cd4', 0],
    ['cat2 10 3 4', 6],
    ['?: 1 7 9', 7],
    ['?: 0 7 9', 9],
    ['?: -1 7 9', 9],
    ['max 7 9', 9],
    ['min 7 9', 7],
    ['mod 3 4 12', 13],
    ['*/ 6 7 3', 14],
    ['*/ 6 7 0', 0],
    ['pin 5 2 10', 5],
    ['pin 5 7 10', 7],
    ['pin 5 12 10', 10],
    ['sat2 10 3 4', 8],
    ['sin 10 cd4', 10],
    ['sqrt 81', 9],
    ['tan 10 cd8', 10],
    ['val -17', -17],
    [`val ${String(Number.MAX_SAFE_INTEGER)}`, Number.MAX_SAFE_INTEGER],
    [' \t +-\t8  4\n3 ', 9],
    ['sin 10 cd8', Math.SQRT1_2 * 10],
  ] as const)('evaluates %s', (formula, expected) => {
    const actual = resolver({ result: formula })('result');
    expect(actual).toBeCloseTo(expected, 10);
  });

  it('resolves forward, nested, and repeated guide references', () => {
    const resolve = resolver({
      first: '+- second 5 2',
      second: '*/ third 4 2',
      third: 'val 10',
    });

    expect(resolve('first')).toBe(23);
    expect(resolve('second')).toBe(20);
    expect(resolve('third')).toBe(10);
    expect(resolve('first')).toBe(23);
  });

  it.each([
    ['0', 0],
    ['+42', 42],
    ['-42', -42],
    [String(Number.MAX_SAFE_INTEGER), Number.MAX_SAFE_INTEGER],
  ] as const)('accepts the integer literal %s', (token, expected) => {
    expect(resolver()(token)).toBe(expected);
  });

  it.each([
    '',
    '1.5',
    '1e3',
    '0x10',
    String(Number.MAX_SAFE_INTEGER + 1),
    'unknown',
  ])('rejects the non-coordinate token %s', (token) => {
    expect(resolver()(token)).toBeUndefined();
  });

  it.each([
    '',
    'unknown 1',
    'abs',
    'abs 1 2',
    '*/ 1 2',
    'val missing',
    '?: missing 7 9',
    '?: 1 missing 9',
    'sqrt -1',
    '*/ 9007199254740991 2 1',
    'tan 9007199254740991 cd4',
  ])('rejects the invalid formula %s', (formula) => {
    expect(resolver({ result: formula })('result')).toBeUndefined();
  });

  it('breaks self-references, cycles, and excessive guide chains', () => {
    expect(resolver({ self: 'val self' })('self')).toBeUndefined();
    expect(
      resolver({ first: 'val second', second: 'val first' })('first'),
    ).toBeUndefined();

    const entries = Array.from(
      { length: 257 },
      (_, index) =>
        [`g${index}`, index === 256 ? 'val 1' : `val g${index + 1}`] as const,
    );
    const resolve = createDrawingGuideResolver(320, 160, new Map(entries));
    expect(resolve('g0')).toBeUndefined();
    expect(resolve('g256')).toBe(1);
  });

  it('rejects non-finite built-in dimensions', () => {
    expect(createDrawingGuideResolver(Number.NaN, 160)('w')).toBeUndefined();
    expect(
      createDrawingGuideResolver(320, Number.POSITIVE_INFINITY)('h'),
    ).toBeUndefined();
  });
});
