import type { XmlLookupValue } from '../../../common';

import { getTextByPathList } from '../../../common';

interface MathPart {
  key: string;
  value: XmlLookupValue;
}

function asArray(value: XmlLookupValue | undefined): XmlLookupValue[] {
  if (!value) return [];
  return Array.isArray(value) ? value : [value];
}

function child(node: XmlLookupValue, key: string): XmlLookupValue | undefined {
  return getTextByPathList<XmlLookupValue>(node, [key]);
}

function* nestedValues(value: unknown): Generator<unknown> {
  switch (Object.prototype.toString.call(value)) {
    case '[object Array]':
    case '[object Object]':
      yield* Object.values(Object(value) as Record<string, unknown>);
  }
}

export function findOMath(value: unknown): XmlLookupValue[] {
  const record = Object(value) as Record<string, unknown>;
  const results: XmlLookupValue[] = [];
  const directMath = record['m:oMath'];
  if (directMath) {
    results.push(
      ...(Array.isArray(directMath)
        ? (directMath as XmlLookupValue[])
        : [directMath as XmlLookupValue]),
    );
  }
  for (const nested of nestedValues(value)) results.push(...findOMath(nested));
  return results;
}

export function parseFraction(fraction: XmlLookupValue): string {
  return `\\frac{${parseOMath(child(fraction, 'm:num'))}}{${parseOMath(
    child(fraction, 'm:den'),
  )}}`;
}

export function parseSuperscript(superscript: XmlLookupValue): string {
  return `${parseOMath(child(superscript, 'm:e'))}^{${parseOMath(
    child(superscript, 'm:sup'),
  )}}`;
}

export function parseSubscript(subscript: XmlLookupValue): string {
  return `${parseOMath(child(subscript, 'm:e'))}_{${parseOMath(
    child(subscript, 'm:sub'),
  )}}`;
}

export function parseRadical(radical: XmlLookupValue): string {
  const degree = parseOMath(child(radical, 'm:deg'));
  const expression = parseOMath(child(radical, 'm:e'));
  return degree ? `\\sqrt[${degree}]{${expression}}` : `\\sqrt{${expression}}`;
}

export function parseMatrix(matrix: XmlLookupValue): string {
  const rows = asArray(child(matrix, 'm:mr')).map((row) =>
    asArray(child(row, 'm:e')).map(parseOMath).join(' & '),
  );
  return `\\begin{matrix} ${rows.join(' \\\\ ')} \\end{matrix}`;
}

export function parseNary(nary: XmlLookupValue): string {
  const operator =
    getTextByPathList<string>(nary, ['m:naryPr', 'm:chr', 'attrs', 'm:val']) ??
    '∫';
  return `${operator}_{${parseOMath(child(nary, 'm:sub'))}}^{${parseOMath(
    child(nary, 'm:sup'),
  )}}{${parseOMath(child(nary, 'm:e'))}}`;
}

export function parseLimit(limit: XmlLookupValue, type: 'low' | 'upp'): string {
  const base = parseOMath(child(limit, 'm:e'));
  const value = parseOMath(child(limit, 'm:lim'));
  return type === 'low' ? `${base}_{${value}}` : `${base}^{${value}}`;
}

export function parseDelimiter(delimiter: XmlLookupValue): string {
  let left = getTextByPathList<string>(delimiter, [
    'm:dPr',
    'm:begChr',
    'attrs',
    'm:val',
  ]);
  let right = getTextByPathList<string>(delimiter, [
    'm:dPr',
    'm:endChr',
    'attrs',
    'm:val',
  ]);
  if (!left && !right) {
    left = '(';
    right = ')';
  }
  if (left && right) {
    left = `\\left${left}`;
    right = `\\right${right}`;
  }
  return `${left ?? ''}${parseOMath(child(delimiter, 'm:e'))}${right ?? ''}`;
}

export function parseFunction(func: XmlLookupValue): string {
  return `\\${parseOMath(child(func, 'm:fName'))}{${parseOMath(
    child(func, 'm:e'),
  )}}`;
}

export function parseGroupChr(group: XmlLookupValue): string {
  const character =
    getTextByPathList<string>(group, [
      'm:groupChrPr',
      'm:chr',
      'attrs',
      'm:val',
    ]) ?? '';
  return `${character}${parseOMath(child(group, 'm:e'))}${character}`;
}

export function parseEqArr(equationArray: XmlLookupValue): string {
  const equations = asArray(child(equationArray, 'm:e'))
    .map(parseOMath)
    .join(' \\\\ ');
  return `\\begin{cases} ${equations} \\end{cases}`;
}

export function parseBar(bar: XmlLookupValue): string {
  const expression = parseOMath(child(bar, 'm:e'));
  const position = getTextByPathList<string>(bar, [
    'm:barPr',
    'm:pos',
    'attrs',
    'm:val',
  ]);
  return position === 'top'
    ? `\\overline{${expression}}`
    : `\\underline{${expression}}`;
}

export function parseAccent(accent: XmlLookupValue): string {
  const character =
    getTextByPathList<string>(accent, ['m:accPr', 'm:chr', 'attrs', 'm:val']) ??
    '^';
  const expression = parseOMath(child(accent, 'm:e'));
  const commands: Readonly<Record<string, string>> = {
    '\u0301': 'acute',
    '\u0300': 'grave',
    '\u0302': 'hat',
    '\u0303': 'tilde',
    '\u0304': 'bar',
    '\u0306': 'breve',
    '\u0307': 'dot',
    '\u0308': 'ddot',
    '\u030A': 'mathring',
    '\u030B': 'H',
    '\u030C': 'check',
    '\u0327': 'c',
  };
  return `\\${commands[character] ?? character}{${expression}}`;
}

export function parseBox(box: XmlLookupValue): string {
  return `\\boxed{${parseOMath(child(box, 'm:e'))}}`;
}

function partOrder(part: MathPart): number {
  const directOrder = getTextByPathList<string>(part.value, [
    'a:rPr',
    'attrs',
    'order',
  ]);
  const controlOrder = getTextByPathList<string>(part.value, [
    `${part.key}Pr`,
    'm:ctrlPr',
    'a:rPr',
    'attrs',
    'order',
  ]);
  return Number(directOrder ?? controlOrder ?? 0);
}

function parseMathPart({ key, value }: MathPart): string {
  switch (key) {
    case 'm:f':
      return parseFraction(value);
    case 'm:sSup':
      return parseSuperscript(value);
    case 'm:sSub':
      return parseSubscript(value);
    case 'm:rad':
      return parseRadical(value);
    case 'm:nary':
      return parseNary(value);
    case 'm:limLow':
      return parseLimit(value, 'low');
    case 'm:limUpp':
      return parseLimit(value, 'upp');
    case 'm:d':
      return parseDelimiter(value);
    case 'm:func':
      return parseFunction(value);
    case 'm:groupChr':
      return parseGroupChr(value);
    case 'm:eqArr':
      return parseEqArr(value);
    case 'm:bar':
      return parseBar(value);
    case 'm:acc':
      return parseAccent(value);
    case 'm:borderBox':
      return parseBox(value);
    case 'm:m':
      return parseMatrix(value);
    case 'm:r':
      return parseOMath(value);
    case 'm:t':
      return typeof value === 'string' ? value : '';
    default:
      return '';
  }
}

export function parseOMath(oMath: XmlLookupValue | undefined): string {
  if (!oMath) return '';
  if (Array.isArray(oMath)) return oMath.map(parseOMath).join('');

  const parts: MathPart[] = [];
  for (const key of Object.keys(oMath)) {
    const value = oMath[key];
    for (const item of asArray(value)) parts.push({ key, value: item });
  }

  return parts
    .sort((a, b) => partOrder(a) - partOrder(b))
    .map(parseMathPart)
    .join('');
}

export function latexFormart(latex: string): string {
  return latex
    .replaceAll(/&lt;/g, '<')
    .replaceAll(/&gt;/g, '>')
    .replaceAll(/&amp;/g, '&')
    .replaceAll(/&apos;/g, "'")
    .replaceAll(/&quot;/g, '"');
}
