import { describe, expect, it } from 'vitest';

import {
  decodeXmlEntities,
  escapeHtml,
  sanitizeHyperlink,
} from '../../src/common/text/html';

describe('HTML helpers', () => {
  it('escapes text and attribute delimiters', () => {
    expect(escapeHtml(`<script data-value="'&">`)).toBe(
      '&lt;script data-value=&quot;&#039;&amp;&quot;&gt;',
    );
  });

  it('decodes named and numeric XML entities before HTML escaping', () => {
    expect(decodeXmlEntities('&lt;&amp;&apos;&#39;&#x3E;')).toBe("<&''>");
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
});
