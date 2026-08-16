import { describe, expect, it } from 'vitest';

import {
  escapeSvgText,
  plainTextFromPowerPointHtml,
} from '../../src/formats/pptx/render-text';

describe('PowerPoint render text', () => {
  it('flattens rich paragraphs, lists, breaks, whitespace, and entities', () => {
    expect(
      plainTextFromPowerPointHtml(
        '<p><span>Quarterly&nbsp;&amp; annual</span><br/> plan</p>' +
          '<ul><li>First&nbsp;&nbsp;item</li><li>Second&#x20;item</li></ul>',
      ),
    ).toBe('Quarterly & annual\nplan\n• First  item\n• Second item');
  });

  it('ignores tag attributes containing greater-than characters', () => {
    expect(
      plainTextFromPowerPointHtml(
        '<p data-value="1 > 0"><span title=\'x > y\'>Safe</span></p>',
      ),
    ).toBe('Safe');
  });

  it('recognizes closing tags with leading whitespace', () => {
    expect(plainTextFromPowerPointHtml('left<  /P>right')).toBe('left\nright');
  });

  it('does not infer a tag name after punctuation or whitespace', () => {
    expect(plainTextFromPowerPointHtml('left< / p>middle<!---->right')).toBe(
      'leftmiddleright',
    );
  });

  it('turns hostile markup into harmless visible text', () => {
    const flattened = plainTextFromPowerPointHtml(
      '<script>alert(1)</script><img src=x onerror=alert(2)><p>&lt;safe&gt;</p>',
    );

    expect(flattened).toBe('alert(1)\n<safe>');
    expect(escapeSvgText(flattened)).toBe('alert(1)\n&lt;safe&gt;');
  });

  it('preserves a dangling less-than sign as text', () => {
    expect(plainTextFromPowerPointHtml('value < unfinished')).toBe(
      'value < unfinished',
    );
  });

  it.each([
    ['<br>', '\n'],
    ['<br/>', '\n'],
    ['</div>', '\n'],
    ['</li>', '\n'],
    ['</p>', '\n'],
    ['<span>', ''],
    ['</span>', ''],
  ])('maps tag %s to its exact boundary', (tag, boundary) => {
    expect(plainTextFromPowerPointHtml(`left${tag}right`)).toBe(
      `left${boundary}right`,
    );
  });

  it('normalizes CR, repeated spaces, tabs, and excessive blank lines', () => {
    expect(
      plainTextFromPowerPointHtml('  one\t two\r\n\r\r\n\r\n  three  '),
    ).toBe('one two\n\nthree');
    expect(plainTextFromPowerPointHtml('before\rafter')).toBe('before\nafter');
  });

  it('escapes every XML delimiter and replaces invalid XML characters', () => {
    expect(escapeSvgText(`<&>"'\u0000\u0008\t\n\r😀`)).toBe(
      '&lt;&amp;&gt;&quot;&#039;��\t\n\r😀',
    );
  });

  it('retains every XML 1.0 code-point boundary', () => {
    expect(
      escapeSvgText(
        String.fromCodePoint(
          0x20,
          0xd7ff,
          0xe000,
          0xfffc,
          0xfffd,
          0x10_000,
          0x10_ffff,
        ),
      ),
    ).toBe(
      String.fromCodePoint(
        0x20,
        0xd7ff,
        0xe000,
        0xfffc,
        0xfffd,
        0x10_000,
        0x10_ffff,
      ),
    );
  });

  it('replaces lone surrogates and forbidden noncharacters', () => {
    expect(escapeSvgText('\uD800\uFFFE\uFFFF')).toBe('���');
  });
});
