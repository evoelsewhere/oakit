import { describe, expect, it } from 'vitest';

import {
  escapeXmlAttribute,
  escapeXmlText,
  serializeDrawingText,
} from '../../src/formats/pptx/writer/xml';

describe('PowerPoint writer XML primitives', () => {
  it.each([
    ['', ''],
    ['plain text', 'plain text'],
    ['<&>', '&lt;&amp;&gt;'],
    ['"quoted" and \'apostrophe\'', '"quoted" and \'apostrophe\''],
    ['_x0041_', '_x005F_x0041_'],
    ['_XabCD_', '_x005F_XabCD_'],
    ['_x005F__x000A_', '_x005F_x005F__x005F_x000A_'],
    ['&_x003C_<', '&amp;_x005F_x003C_&lt;'],
  ])('escapes text %j exactly', (input, expected) => {
    expect(escapeXmlText(input)).toBe(expected);
  });

  it('escapes every XML attribute delimiter after Office sequences', () => {
    expect(escapeXmlAttribute(`<&>"'_x0022_`)).toBe(
      '&lt;&amp;&gt;&quot;&apos;_x005F_x0022_',
    );
  });

  it.each([
    ['text', '<a:t>text</a:t>'],
    ['two words', '<a:t>two words</a:t>'],
    ['', '<a:t></a:t>'],
    [' leading', '<a:t xml:space="preserve"> leading</a:t>'],
    ['trailing ', '<a:t xml:space="preserve">trailing </a:t>'],
    ['two  spaces', '<a:t xml:space="preserve">two  spaces</a:t>'],
    ['a\tb', '<a:t xml:space="preserve">a\tb</a:t>'],
    ['a\nb', '<a:t xml:space="preserve">a\nb</a:t>'],
    ['a\rb', '<a:t xml:space="preserve">a\rb</a:t>'],
    ['<&', '<a:t>&lt;&amp;</a:t>'],
    ['_x000A_', '<a:t>_x005F_x000A_</a:t>'],
  ])('serializes DrawingML text %j', (input, expected) => {
    expect(serializeDrawingText(input)).toBe(expected);
  });

  it('honors an authored whitespace-preservation flag', () => {
    expect(serializeDrawingText('plain', true)).toBe(
      '<a:t xml:space="preserve">plain</a:t>',
    );
    expect(serializeDrawingText('plain', false)).toBe('<a:t>plain</a:t>');
    expect(serializeDrawingText(' leading', false)).toBe(
      '<a:t xml:space="preserve"> leading</a:t>',
    );
  });
});
