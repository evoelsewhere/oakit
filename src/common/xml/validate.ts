import { SaxesParser } from 'saxes';

import {
  XmlComplexityLimitError,
  type XmlReadLimits,
  XmlStructureError,
} from './types';

const XML_QUALIFIED_NAME =
  /^(?:[A-Za-z_][A-Za-z\d_.-]*:)?[A-Za-z_][A-Za-z\d_.-]*$/;

export function decodeXmlBytes(bytes: Uint8Array): string {
  let encoding = 'utf-8';
  if (bytes[0] === 0xff && bytes[1] === 0xfe) encoding = 'utf-16le';
  else if (bytes[0] === 0xfe && bytes[1] === 0xff) encoding = 'utf-16be';

  try {
    return new TextDecoder(encoding, { fatal: true }).decode(bytes);
  } catch (cause) {
    throw new XmlStructureError(`Invalid ${encoding.toUpperCase()} XML`, {
      cause,
    });
  }
}

/** Reject malformed or pathological XML before the recursive object parser sees it. */
export function assertXmlComplexity(
  xml: string,
  limits: Pick<XmlReadLimits, 'maxDepth' | 'maxNodes'>,
): number {
  let depth = 0;
  let nodes = 0;
  const parser = new SaxesParser({
    defaultXMLVersion: '1.0',
    forceXMLVersion: true,
    xmlns: false,
  });

  parser.on('doctype', () => {
    throw new XmlStructureError(
      'XML document type declarations are not allowed',
    );
  });
  parser.on('opentagstart', (tag) => {
    if (!XML_QUALIFIED_NAME.test(tag.name)) {
      throw new XmlStructureError(`Invalid XML element name ${tag.name}`);
    }
  });
  parser.on('attribute', (attribute) => {
    if (!XML_QUALIFIED_NAME.test(attribute.name)) {
      throw new XmlStructureError(
        `Invalid XML attribute name ${attribute.name}`,
      );
    }
  });
  parser.on('opentag', () => {
    nodes += 1;
    depth += 1;
    if (limits.maxNodes !== undefined && nodes > limits.maxNodes) {
      throw new XmlComplexityLimitError('maxXmlNodes', nodes, limits.maxNodes);
    }
    if (limits.maxDepth !== undefined && depth > limits.maxDepth) {
      throw new XmlComplexityLimitError('maxXmlDepth', depth, limits.maxDepth);
    }
  });
  parser.on('closetag', () => {
    depth -= 1;
  });

  try {
    parser.write(xml).close();
  } catch (cause) {
    if (
      cause instanceof XmlComplexityLimitError ||
      cause instanceof XmlStructureError
    ) {
      throw cause;
    }
    throw new XmlStructureError('Invalid XML structure', { cause });
  }
  return nodes;
}
