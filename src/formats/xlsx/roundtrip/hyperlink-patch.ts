import type { XlsxHyperlinkTarget } from '../types';
import { XlsxWriteError } from './errors';
import type { ResolvedXlsxWriteLimits } from './types';
import {
  decodeXlsxXml,
  encodeXlsxXml,
  tokenizeXlsxXml,
  xlsxXmlLocalName,
  type XlsxXmlTagToken,
} from './worksheet-patch';
import { writeLimitFailure } from './write-limits';

export interface XlsxInternalHyperlinkPatch {
  cell: string;
  operationId: string;
  target: Extract<XlsxHyperlinkTarget, { kind: 'internal' }> | null;
}

export interface XlsxHyperlinkPatchResult {
  data: Uint8Array;
  patchBytes: number;
  patchCount: number;
}

interface TextPatch {
  end: number;
  replacement: string;
  start: number;
}

function failure(
  message: string,
  part: string,
  request?: XlsxInternalHyperlinkPatch,
  featureClass = 'hyperlink-xml',
): never {
  throw new XlsxWriteError('preservation-conflict', message, {
    ...(request === undefined
      ? {}
      : { cell: request.cell, operationId: request.operationId }),
    featureClass,
    part,
  });
}

function escapeAttribute(
  value: string,
  part: string,
  request: XlsxInternalHyperlinkPatch,
): string {
  let output = '';
  for (const character of value) {
    const code = character.codePointAt(0)!;
    if (
      code !== 0x09 &&
      code !== 0x0a &&
      code !== 0x0d &&
      (code < 0x20 ||
        (code > 0xd7ff && code < 0xe000) ||
        (code > 0xfffd && code < 0x1_0000))
    ) {
      throw new XlsxWriteError(
        'invalid-roundtrip-json',
        'XLSX hyperlink location contains an invalid XML character',
        { cell: request.cell, operationId: request.operationId, part },
      );
    }
    output +=
      character === '&'
        ? '&amp;'
        : character === '<'
          ? '&lt;'
          : character === '"'
            ? '&quot;'
            : character === '\r'
              ? '&#13;'
              : character === '\n'
                ? '&#10;'
                : character === '\t'
                  ? '&#9;'
                  : character;
  }
  return output;
}

function matchingClose(
  tokens: readonly XlsxXmlTagToken[],
  openIndex: number,
): XlsxXmlTagToken {
  const open = tokens[openIndex]!;
  if (open.selfClosing) return open;
  return tokens.find(
    (token) => token.depth === open.depth && token.start >= open.end,
  )!;
}

function directCollection(
  tokens: readonly XlsxXmlTagToken[],
  root: XlsxXmlTagToken,
  name: string,
): XlsxXmlTagToken | undefined {
  const prefix = root.name.slice(0, -xlsxXmlLocalName(root.name).length);
  return tokens.find(
    (token) =>
      !token.closing &&
      token.depth === root.depth + 1 &&
      token.name === `${prefix}${name}`,
  );
}

function hyperlinkReplacement(
  text: string,
  token: XlsxXmlTagToken,
  request: XlsxInternalHyperlinkPatch,
  part: string,
): string {
  if (token.attributes.some((attribute) => attribute.name.includes(':'))) {
    failure(
      'XLSX internal hyperlink edit cannot replace an external relationship target',
      part,
      request,
      'external-hyperlink',
    );
  }
  const allowed = new Set(['display', 'location', 'ref', 'tooltip']);
  if (token.attributes.some((attribute) => !allowed.has(attribute.name))) {
    failure(
      'XLSX hyperlink element contains an unsupported attribute',
      part,
      request,
    );
  }
  if (request.target === null) return '';
  const authored = token.attributes
    .filter((attribute) => attribute.name !== 'location')
    .map((attribute) => text.slice(attribute.start, attribute.end))
    .join('');
  return `<${token.name}${authored} location="${escapeAttribute(
    request.target.location,
    part,
    request,
  )}"/>`;
}

function newHyperlink(
  name: string,
  request: XlsxInternalHyperlinkPatch,
  part: string,
): string {
  return `<${name} ref="${request.cell}" location="${escapeAttribute(
    request.target!.location,
    part,
    request,
  )}"/>`;
}

export function patchXlsxInternalHyperlinks(
  bytes: Uint8Array,
  requested: readonly XlsxInternalHyperlinkPatch[],
  limits: ResolvedXlsxWriteLimits,
  part: string,
): XlsxHyperlinkPatchResult {
  const requests = new Map<string, XlsxInternalHyperlinkPatch>();
  for (const request of requested) {
    if (requests.has(request.cell)) {
      failure('XLSX hyperlink patch cells must be unique', part, request);
    }
    requests.set(request.cell, request);
  }
  const decoded = decodeXlsxXml(bytes, part);
  const tokens = tokenizeXlsxXml(decoded.text, part);
  const root = tokens.find(
    (token) =>
      token.depth === 0 && xlsxXmlLocalName(token.name) === 'worksheet',
  );
  if (!root) failure('XLSX worksheet root cannot patch hyperlinks', part);
  const prefix = root.name.slice(0, -'worksheet'.length);
  const collection = directCollection(tokens, root, 'hyperlinks');
  const hyperlinkName = `${prefix}hyperlink`;
  const found = new Set<string>();
  const patches: TextPatch[] = [];
  let sourceCount = 0;
  if (collection) {
    const collectionIndex = tokens.indexOf(collection);
    const close = matchingClose(tokens, collectionIndex);
    const closeIndex = tokens.indexOf(close);
    for (const [relativeIndex, token] of tokens
      .slice(collectionIndex + 1, closeIndex)
      .entries()) {
      if (
        token.name !== hyperlinkName ||
        token.depth !== collection.depth + 1
      ) {
        continue;
      }
      const reference = token.attributes.find(
        (attribute) => attribute.name === 'ref',
      )?.value;
      if (reference === undefined) continue;
      sourceCount += 1;
      const request = requests.get(reference);
      if (!request) continue;
      if (found.has(reference)) {
        failure('XLSX hyperlink reference is ambiguous', part, request);
      }
      found.add(reference);
      const end = matchingClose(
        tokens,
        collectionIndex + 1 + relativeIndex,
      ).end;
      patches.push({
        end,
        replacement: hyperlinkReplacement(decoded.text, token, request, part),
        start: token.start,
      });
    }
    const additions = requested.filter(
      (request) => !found.has(request.cell) && request.target !== null,
    );
    const removals = requested.filter(
      (request) => found.has(request.cell) && request.target === null,
    ).length;
    const finalCount = sourceCount - removals + additions.length;
    if (finalCount === 0) {
      patches.length = 0;
      patches.push({
        end: close.end,
        replacement: '',
        start: collection.start,
      });
    } else if (additions.length !== 0) {
      patches.push({
        end: close.start,
        replacement: additions
          .map((request) => newHyperlink(hyperlinkName, request, part))
          .join(''),
        start: close.start,
      });
    }
  } else {
    const additions = requested.filter((request) => request.target !== null);
    if (additions.length !== 0) {
      const close = matchingClose(tokens, tokens.indexOf(root));
      const name = `${prefix}hyperlinks`;
      patches.push({
        end: close.start,
        replacement: `<${name}>${additions
          .map((request) => newHyperlink(hyperlinkName, request, part))
          .join('')}</${name}>`,
        start: close.start,
      });
    }
  }
  let patchBytes = 0;
  for (const patch of patches) {
    patchBytes += encodeXlsxXml({
      bom: false,
      encoding: decoded.encoding,
      text: patch.replacement,
    }).byteLength;
    if (patchBytes > limits.maxPatchBytes) {
      writeLimitFailure(
        'maxPatchBytes',
        patchBytes,
        limits.maxPatchBytes,
        part,
      );
    }
  }
  patches.sort((left, right) => right.start - left.start);
  let output = decoded.text;
  for (const patch of patches) {
    output = `${output.slice(0, patch.start)}${patch.replacement}${output.slice(patch.end)}`;
  }
  const data = encodeXlsxXml({ ...decoded, text: output });
  if (data.byteLength > limits.maxGeneratedXmlBytes) {
    writeLimitFailure(
      'maxGeneratedXmlBytes',
      data.byteLength,
      limits.maxGeneratedXmlBytes,
      part,
    );
  }
  return { data, patchBytes, patchCount: patches.length };
}
