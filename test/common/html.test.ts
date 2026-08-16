import { afterEach, describe, expect, it, vi } from 'vitest';

import {
  decodeXmlEntities,
  escapeHtml,
  hasValidText,
  sanitizeHyperlink,
} from '../../src/common/text/html';

describe('HTML helpers', () => {
  afterEach(() => {
    vi.unstubAllGlobals();
  });

  it('escapes text and attribute delimiters', () => {
    expect(escapeHtml(`<script data-value="'&">`)).toBe(
      '&lt;script data-value=&quot;&#039;&amp;&quot;&gt;',
    );
  });

  it('decodes named and numeric XML entities before HTML escaping', () => {
    expect(decodeXmlEntities('&LT;&amp;&apos;&gt;&quot;&#39;&#x3E;')).toBe(
      `<&'>"'>`,
    );
  });

  it('keeps invalid numeric entities and accepts Unicode boundary values', () => {
    expect(decodeXmlEntities('&#1114112; &#x110000;')).toBe(
      '&#1114112; &#x110000;',
    );
    expect(decodeXmlEntities('&#0;&#1114111;')).toBe(
      `\u0000${String.fromCodePoint(0x10ffff)}`,
    );
  });

  it.each([
    'https://example.com/path?a=1#result',
    'http://example.com',
    'mailto:hello@example.com',
  ])('allows safe hyperlink protocols: %s', (value) => {
    expect(sanitizeHyperlink(value)).toBe(value);
  });

  it.each([
    'javascript:alert(1)',
    'data:text/html,<script>alert(1)</script>',
    'file:///etc/passwd',
    'relative/path',
  ])('rejects unsafe or ambiguous hyperlinks: %s', (value) => {
    expect(sanitizeHyperlink(value)).toBeNull();
  });

  it('trims a safe link and rejects an empty candidate', () => {
    expect(sanitizeHyperlink('  HTTPS://example.com/a  ')).toBe(
      'HTTPS://example.com/a',
    );
    expect(sanitizeHyperlink(' \n\t ')).toBeNull();
  });

  it.each([
    ['', false],
    ['<p> \n <br>\t</p>', false],
    ['<p>&nbsp;</p>', false],
    ['<p>&#160;&#xA0;</p>', false],
    ['<p>Visible</p>', true],
    ['<span title="metadata">A   B</span>', true],
  ])('recognizes visible text in %j', (html, expected) => {
    expect(hasValidText(html)).toBe(expected);
  });

  it.each([
    ['Visible text', true],
    [' \n\t ', false],
    [null, false],
  ])(
    'uses DOMParser text content when it is available: %j',
    (textContent, expected) => {
      vi.stubGlobal(
        'DOMParser',
        class {
          parseFromString(_html: string, mediaType: string) {
            expect(mediaType).toBe('text/html');
            return { body: { textContent } };
          }
        },
      );

      expect(hasValidText('<p>ignored by the parser stub</p>')).toBe(expected);
    },
  );
});
