import { decodeXmlEntities, escapeHtml } from '../../common';

function appendLineBreak(output: string): string {
  return output.endsWith('\n') ? output : `${output}\n`;
}

interface HtmlTag {
  closing: boolean;
  name?: string;
}

function htmlTag(source: string): HtmlTag {
  const trimmed = source.trimStart();
  const closing = trimmed.startsWith('/');
  const name = (closing ? trimmed.slice(1) : trimmed).match(/^[\da-z]+/i);
  return name ? { closing, name: name.toString().toLowerCase() } : { closing };
}

function appendTagBoundary(output: string, tag: HtmlTag): string {
  if (!tag.closing && tag.name === 'li') return `${appendLineBreak(output)}• `;
  if (
    tag.name === 'br' ||
    tag.name === 'div' ||
    tag.name === 'p' ||
    (tag.closing && tag.name === 'li')
  ) {
    return appendLineBreak(output);
  }
  return output;
}

export function plainTextFromPowerPointHtml(input: string): string {
  const tags = Array.from(input.matchAll(/<(?:[^"'<>]|"[^"]*"|'[^']*')*>/g));
  const flattened = tags.reduce(
    (state, match) => {
      const source = match[0];
      const start = match.index;
      const output = state.output + input.slice(state.cursor, start);
      return {
        cursor: start + source.length,
        output: appendTagBoundary(output, htmlTag(source.slice(1))),
      };
    },
    { cursor: 0, output: '' },
  );
  const output = flattened.output + input.slice(flattened.cursor);
  return decodeXmlEntities(output.replace(/&nbsp;/gi, '&#160;'))
    .replace(/\r\n?/g, '\n')
    .split('\n')
    .map((line) => line.replace(/[\t ]+/g, ' ').trim())
    .join('\n')
    .replace(/\n{3,}/g, '\n\n')
    .trim();
}

function isXmlCodePoint(codePoint: number): boolean {
  return (
    codePoint === 0x9 ||
    codePoint === 0xa ||
    codePoint === 0xd ||
    (codePoint >= 0x20 && codePoint <= 0xd7ff) ||
    (codePoint >= 0xe000 && codePoint <= 0xfffc) ||
    codePoint >= 0x10_000
  );
}

export function escapeSvgText(input: string): string {
  let valid = '';
  for (const character of input) {
    const codePoint = character.codePointAt(0);
    valid +=
      codePoint !== undefined && isXmlCodePoint(codePoint)
        ? character
        : '\uFFFD';
  }
  return escapeHtml(valid);
}
