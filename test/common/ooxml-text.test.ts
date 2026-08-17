import { describe, expect, it } from 'vitest';

import { decodeOfficeTextEscapes } from '../../src/common/ooxml/text';

describe('Office Open XML text escapes', () => {
  it.each([
    ['plain', 'plain'],
    ['_x0041_', 'A'],
    ['_X0062_', 'b'],
    ['_x005F_x0041_', '_x0041_'],
    ['_xD83D__xDE00_', '😀'],
    ['before_x000A_after', 'before\nafter'],
  ])('decodes %j as %j', (source, expected) => {
    expect(decodeOfficeTextEscapes(source)).toBe(expected);
  });
});
