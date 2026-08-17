import { describe, expect, it } from 'vitest';

import {
  escapeSvgText,
  plainTextFromPowerPointHtml,
  renderedTextFromPowerPointHtml,
} from '../../src/formats/pptx/render-text';

describe('PowerPoint render text', () => {
  it('extracts bounded rich span styles and paragraph alignment', () => {
    expect(
      renderedTextFromPowerPointHtml(
        '<p data-kind="hero" STYLE="text-align: center;"><span data-kind="title" STYLE="color: #F8FAFC;font-family: &quot;Aptos Display&quot;;font-size: 18pt;font-weight: 700;font-style: italic;">Hello&nbsp;AI</span></p>',
      ),
    ).toEqual([
      {
        alignment: 'center',
        runs: [
          {
            bold: true,
            color: '#F8FAFC',
            fontFamily: 'Aptos Display',
            fontSize: 18,
            italic: true,
            text: 'Hello AI',
          },
        ],
      },
    ]);
  });

  it('drops hostile or unbounded style values while preserving visible text', () => {
    expect(
      renderedTextFromPowerPointHtml(
        '<p style="text-align: evil"><span style="color: url(https://evil);font-size: 9999pt">Safe</span></p>',
      ),
    ).toEqual([
      {
        alignment: 'left',
        runs: [{ bold: false, italic: false, text: 'Safe' }],
      },
    ]);
  });

  it.each([
    ['font-family: Arial', 'Arial'],
    ['font-family: "Urbanist"', 'Urbanist'],
    ["font-family: 'Aptos Display'", 'Aptos Display'],
    ['font-family: Arial, sans-serif', undefined],
    ['font-family: url(evil)', undefined],
    [`font-family: ${'A'.repeat(80)}`, 'A'.repeat(80)],
    [`font-family: ${'A'.repeat(81)}`, undefined],
    ['font-family: ""', undefined],
    ['font-family: "Arial', undefined],
    ["font-family: Arial'", undefined],
    ['font-family: "Arial\'', undefined],
    ["font-family: '", undefined],
  ])('bounds authored font family %j', (style, expected) => {
    expect(
      renderedTextFromPowerPointHtml(
        `<p><span style="${style.replaceAll('"', '&quot;')}">Font</span></p>`,
      )[0]?.runs[0]?.fontFamily,
    ).toBe(expected);
  });

  it.each([
    ['left', 'left'],
    ['center', 'center'],
    ['right', 'right'],
    ['justify', 'left'],
  ])('maps paragraph alignment %s to %s', (source, expected) => {
    expect(
      renderedTextFromPowerPointHtml(
        `<p style='text-align: ${source}'><span>Aligned</span></p>`,
      )[0]?.alignment,
    ).toBe(expected);
  });

  it('keeps multiple rich runs and paragraphs in authored order', () => {
    expect(
      renderedTextFromPowerPointHtml(
        '<p><span style="color: #112233;font-weight: 600">First</span><span style="font-size: 12.5pt;font-style: italic">Second</span></p>' +
          '<li><span style="font-weight: 900">Third</span></li>',
      ),
    ).toEqual([
      {
        alignment: 'left',
        runs: [
          {
            bold: true,
            color: '#112233',
            italic: false,
            text: 'First',
          },
          {
            bold: false,
            fontSize: 12.5,
            italic: true,
            text: 'Second',
          },
        ],
      },
      {
        alignment: 'left',
        runs: [{ bold: true, italic: false, text: 'Third' }],
      },
    ]);
  });

  it('falls back for plain paragraph bodies and non-paragraph HTML', () => {
    expect(renderedTextFromPowerPointHtml('<p>Plain&nbsp;text</p>')).toEqual([
      {
        alignment: 'left',
        runs: [{ bold: false, italic: false, text: 'Plain text' }],
      },
    ]);
    expect(renderedTextFromPowerPointHtml('<div>Loose</div>')).toEqual([
      {
        alignment: 'left',
        runs: [{ bold: false, italic: false, text: 'Loose' }],
      },
    ]);
    expect(renderedTextFromPowerPointHtml('<p><span></span></p>')).toEqual([]);
    expect(renderedTextFromPowerPointHtml('')).toEqual([]);
    expect(
      renderedTextFromPowerPointHtml(
        '<p style="text-align: center">Centered plain</p>',
      ),
    ).toEqual([
      {
        alignment: 'center',
        runs: [{ bold: false, italic: false, text: 'Centered plain' }],
      },
    ]);
  });

  it.each([
    ['font-size: 18pt', { fontSize: 18 }],
    ['font-size: 18.25pt', { fontSize: 18.25 }],
    ['font-size: 512pt', { fontSize: 512 }],
    ['font-size: 0pt', {}],
    ['font-size: 513pt', {}],
    ['font-size: x18pt', {}],
    ['font-size: 18ptx', {}],
    ['font-size: 18.apt', {}],
    ['color: #A1B2C3', { color: '#A1B2C3' }],
    ['color: x#A1B2C3', {}],
    ['color: #A1B2C3x', {}],
    ['font-weight: 500', { bold: false }],
    ['font-weight: 600', { bold: true }],
    ['font-weight: bold', { bold: true }],
    ['font-weight: xbold', { bold: false }],
    ['font-weight: bolder', { bold: false }],
    ['font-style: italic', { italic: true }],
    ['font-style: xitalic', { italic: false }],
    ['font-style: italicized', { italic: false }],
  ])('bounds rich declaration %j', (style, expected) => {
    const run = renderedTextFromPowerPointHtml(
      `<p><span style="${style}">Value</span></p>`,
    )[0]?.runs[0];
    expect(run).toEqual({
      bold: false,
      italic: false,
      text: 'Value',
      ...expected,
    });
  });

  it('ignores malformed and empty CSS declarations deterministically', () => {
    expect(
      renderedTextFromPowerPointHtml(
        '<p style="broken; : empty; text-align: right; text-align: center"><span style="broken; color: #123456; color:">Value</span></p>',
      ),
    ).toEqual([
      {
        alignment: 'center',
        runs: [
          {
            bold: false,
            color: '#123456',
            italic: false,
            text: 'Value',
          },
        ],
      },
    ]);
  });

  it('requires an exact CSS property name and keeps the complete value', () => {
    expect(
      renderedTextFromPowerPointHtml(
        '<p style="text-align:center trailing"><span style="!color:#123456;color!!!:#654321;color:#ABCDEF">Value</span></p>',
      ),
    ).toEqual([
      {
        alignment: 'left',
        runs: [
          {
            bold: false,
            color: '#ABCDEF',
            italic: false,
            text: 'Value',
          },
        ],
      },
    ]);
  });

  it('ignores declarations without a separator instead of overriding a valid value', () => {
    expect(
      renderedTextFromPowerPointHtml(
        '<p><span style="font-weight:700;font-weightx">Bold</span></p>',
      )[0]?.runs[0]?.bold,
    ).toBe(true);
  });

  it('retains explicit false style defaults for spans without style', () => {
    expect(
      renderedTextFromPowerPointHtml('<p><span>Plain run</span></p>'),
    ).toEqual([
      {
        alignment: 'left',
        runs: [{ bold: false, italic: false, text: 'Plain run' }],
      },
    ]);
  });

  it('parses multiline spans and closing-tag whitespace without widening markup', () => {
    expect(
      renderedTextFromPowerPointHtml(
        '<p><span style="color: #123456">Line\nbreak</span   ></p>',
      ),
    ).toEqual([
      {
        alignment: 'left',
        runs: [
          {
            bold: false,
            color: '#123456',
            italic: false,
            text: 'Line\nbreak',
          },
        ],
      },
    ]);
  });

  it('preserves inline spacing, entities, links, and breaks across runs', () => {
    expect(
      renderedTextFromPowerPointHtml(
        '<p><span>First</span><span>&nbsp;second &amp; <a href="https://example.test">linked</a><br/>line</span></p>',
      )[0]?.runs.map(({ text }) => text),
    ).toEqual(['First', ' second & linked\nline']);
  });

  it.each([
    '<br>',
    '<br/>',
    '<br />',
    '<br data-kind="x">',
    "<br data-kind='>'/>",
  ])('maps rich inline break %s without retaining markup', (tag) => {
    expect(
      renderedTextFromPowerPointHtml(`<p><span>A${tag}B</span></p>`)[0]?.runs[0]
        ?.text,
    ).toBe('A\nB');
  });

  it('strips single-quoted inline tags and normalizes standalone carriage return', () => {
    expect(
      renderedTextFromPowerPointHtml(
        "<p><span>A<a title='x > y'>B</a>\rC</span></p>",
      )[0]?.runs[0]?.text,
    ).toBe('AB\nC');
  });
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
