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
  });

  it.each(['112233', '#aabbcc', '11223344'])(
    'accepts valid DrawingML hexadecimal color %s',
    (value) => {
      expect(normalizeHexColor(value)).toBe(`#${value.replace('#', '')}`);
    },
  );

  it.each(['red', '#123', '12345g', 'fff;position:fixed'])(
    'rejects unsafe or unsupported color %s',
    (value) => {
      expect(normalizeHexColor(value)).toBeNull();
    },
  );
});
