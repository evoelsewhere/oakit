import { decodeXmlEntities, escapeHtml } from '../../common';

function appendLineBreak(output: string): string {
  return output.endsWith('\n') ? output : `${output}\n`;
}

interface HtmlTag {
  closing: boolean;
  name?: string;
}

export interface PptxRenderedTextRun {
  bold: boolean;
  color?: string;
  fontSize?: number;
  italic: boolean;
  text: string;
}

export interface PptxRenderedTextParagraph {
  alignment: 'center' | 'left' | 'right';
  runs: PptxRenderedTextRun[];
}

const RENDERED_CSS_PROPERTIES = new Set([
  'color',
  'font-size',
  'font-style',
  'font-weight',
  'text-align',
]);

function attribute(source: string, name: string): string | undefined {
  const match = new RegExp(
    `(?:^|\\s)${name}\\s*=\\s*(?:"([^"]*)"|'([^']*)')`,
    'i',
  ).exec(source);
  const value = match?.[1] ?? match?.[2];
  return value === undefined ? undefined : decodeXmlEntities(value);
}

function cssDeclarations(source: string): Map<string, string> {
  const result = new Map<string, string>();
  for (const declaration of source.split(';')) {
    const separator = declaration.indexOf(':');
    if (separator === -1) continue;
    const name = declaration.slice(0, separator).trim();
    const value = declaration.slice(separator + 1).trim();
    const normalizedName = name.toLowerCase();
    if (!RENDERED_CSS_PROPERTIES.has(normalizedName) || value === '') continue;
    result.set(normalizedName, value);
  }
  return result;
}

function paragraphAlignment(attributes: string) {
  const styleSource = attribute(attributes, 'style');
  if (styleSource === undefined) return 'left';
  const style = cssDeclarations(styleSource);
  const alignment = style.get('text-align');
  return alignment === 'center' || alignment === 'right' ? alignment : 'left';
}

function inlineTextFromPowerPointHtml(input: string): string {
  return decodeXmlEntities(
    input
      .replace(/<(?:[^"'<>]|"[^"]*"|'[^']*')*>/g, (source) =>
        htmlTag(source.slice(1)).name === 'br' ? '\n' : '',
      )
      .replace(/&nbsp;/gi, '&#160;'),
  ).replace(/\r\n?/g, '\n');
}

function renderedRun(attributes: string, body: string): PptxRenderedTextRun {
  const styleSource = attribute(attributes, 'style');
  if (styleSource === undefined) {
    return {
      bold: false,
      italic: false,
      text: inlineTextFromPowerPointHtml(body),
    };
  }
  const style = cssDeclarations(styleSource);
  const color = style.get('color');
  const fontSizeValue = style.get('font-size');
  const fontSizeMatch =
    fontSizeValue === undefined
      ? null
      : /^([0-9]+(?:\.[0-9]+)?)pt$/i.exec(fontSizeValue);
  const fontSize = Number(fontSizeMatch?.[1]);
  const fontWeight = style.get('font-weight');
  const fontStyle = style.get('font-style');
  return {
    bold: fontWeight !== undefined && /^(?:[6-9]00|bold)$/i.test(fontWeight),
    ...(typeof color === 'string' && /^#[0-9a-f]{6}$/i.test(color)
      ? { color }
      : {}),
    ...(Number.isFinite(fontSize) && fontSize > 0 && fontSize <= 512
      ? { fontSize }
      : {}),
    italic: fontStyle !== undefined && /^italic$/i.test(fontStyle),
    text: inlineTextFromPowerPointHtml(body),
  };
}

export function renderedTextFromPowerPointHtml(
  input: string,
): PptxRenderedTextParagraph[] {
  const paragraphs = [
    ...input.matchAll(/<(p|li)\b([^>]*)>([\s\S]*?)<\/\1\s*>/gi),
  ]
    .map((paragraph) => {
      const body = paragraph[3] as string;
      const runs = [...body.matchAll(/<span\b([^>]*)>([\s\S]*?)<\/span\s*>/gi)]
        .map((span) => renderedRun(span[1] as string, span[2] as string))
        .filter((run) => run.text !== '');
      if (runs.length === 0) {
        const fallback = plainTextFromPowerPointHtml(body);
        if (fallback !== '') {
          runs.push({ bold: false, italic: false, text: fallback });
        }
      }
      return {
        alignment: paragraphAlignment(paragraph[2] as string),
        runs,
      } satisfies PptxRenderedTextParagraph;
    })
    .filter((paragraph) => paragraph.runs.length > 0);
  if (paragraphs.length > 0) return paragraphs;
  const fallback = plainTextFromPowerPointHtml(input);
  return fallback === ''
    ? []
    : [
        {
          alignment: 'left',
          runs: [{ bold: false, italic: false, text: fallback }],
        },
      ];
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
