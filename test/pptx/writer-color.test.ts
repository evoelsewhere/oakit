import { describe, expect, it } from 'vitest';

import {
  powerPointRgbValue,
  serializeSolidColorFill,
} from '../../src/formats/pptx/writer/color';

describe('PowerPoint writer colors', () => {
  it('normalizes a validated color to uppercase DrawingML RGB', () => {
    expect(powerPointRgbValue('#0fa2Bc')).toBe('0FA2BC');
    expect(serializeSolidColorFill('#0fa2Bc')).toBe(
      '<a:solidFill><a:srgbClr val="0FA2BC"/></a:solidFill>',
    );
  });

  it.each(['0FA2BC', 'x#0FA2BC', '#0FA2BCx', '#1234', '#GG0000', ''])(
    'rejects invalid color %j',
    (color) => {
      expect(() => powerPointRgbValue(color)).toThrow(
        'PowerPoint color must use #RRGGBB',
      );
    },
  );
});
