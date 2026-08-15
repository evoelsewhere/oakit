import { describe, expect, it } from 'vitest';

import {
  normalizeHexColor,
  serializeCssFontFamily,
} from '../../src/common/text/css';

describe('CSS value serialization', () => {
  it('quotes font names so declaration delimiters remain text', () => {
    expect(
      serializeCssFontFamily(
        'Arial; background-image:url(https://attacker.example/pixel)',
      ),
    ).toBe('"Arial; background-image:url(https://attacker.example/pixel)"');
  });

  it('escapes quotes, backslashes, entities, and control characters', () => {
    expect(serializeCssFontFamily('A&quot;B\\C\nD')).toBe('"A\\"B\\\\C\\a D"');
    expect(serializeCssFontFamily(` A\u001fB\u007fC `)).toBe('"A\\1f B\\7f C"');
    expect(serializeCssFontFamily('A\u0000B')).toBe('"A\uFFFDB"');
  });

  it('trims the font name and rejects names without visible content', () => {
    expect(serializeCssFontFamily('  Aptos  ')).toBe('"Aptos"');
    expect(serializeCssFontFamily(' \n\t ')).toBeNull();
  });

  it.each(['112233', '#aabbcc', '11223344'])(
    'accepts valid DrawingML hexadecimal color %s',
    (value) => {
      expect(normalizeHexColor(value)).toBe(`#${value.replace('#', '')}`);
    },
  );

  it.each([
    'red',
    '#123',
    '12345g',
    'fff;position:fixed',
    'x112233',
    '112233x',
    '11#2233',
  ])('rejects unsafe or unsupported color %s', (value) => {
    expect(normalizeHexColor(value)).toBeNull();
  });

  it('trims a single optional leading hash from a valid color', () => {
    expect(normalizeHexColor('  #A1b2C3  ')).toBe('#A1b2C3');
  });
});
