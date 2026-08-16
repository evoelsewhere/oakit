import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';

import {
  assertXmlComplexity,
  getXmlNodeOrder,
  readXmlFileResult,
  simplifyLossless,
  XmlComplexityLimitError,
  XmlStructureError,
} from '../../src/common/xml/read-xml';

const DRAWING_NAMESPACE =
  'http://schemas.openxmlformats.org/drawingml/2006/main';

const CANONICAL_NAMESPACES = [
  ['http://purl.oclc.org/ooxml/drawingml/chart', 'c'],
  ['http://purl.oclc.org/ooxml/drawingml/diagram', 'dgm'],
  ['http://purl.oclc.org/ooxml/drawingml/main', 'a'],
  ['http://purl.oclc.org/ooxml/officeDocument/math', 'm'],
  ['http://purl.oclc.org/ooxml/officeDocument/relationships', 'r'],
  ['http://purl.oclc.org/ooxml/presentationml/main', 'p'],
  ['http://schemas.microsoft.com/office/drawing/2008/diagram', 'dsp'],
  ['http://schemas.microsoft.com/office/drawing/2010/main', 'a14'],
  ['http://schemas.openxmlformats.org/drawingml/2006/chart', 'c'],
  ['http://schemas.openxmlformats.org/drawingml/2006/diagram', 'dgm'],
  ['http://schemas.openxmlformats.org/drawingml/2006/main', 'a'],
  ['http://schemas.openxmlformats.org/markup-compatibility/2006', 'mc'],
  ['http://schemas.openxmlformats.org/officeDocument/2006/math', 'm'],
  ['http://schemas.openxmlformats.org/officeDocument/2006/relationships', 'r'],
  ['http://schemas.openxmlformats.org/package/2006/content-types', ''],
  ['http://schemas.openxmlformats.org/package/2006/relationships', ''],
  ['http://schemas.openxmlformats.org/presentationml/2006/main', 'p'],
] as const;

async function readMalformedBytes(
  bytes: Uint8Array,
): Promise<XmlStructureError> {
  const zip = new JSZip();
  zip.file('document.xml', bytes);
  const result = await readXmlFileResult(zip, 'document.xml');

  expect(result.status).toBe('error');
  if (result.status !== 'error') throw new Error('Expected malformed XML');
  expect(result.phase).toBe('parse');
  expect(result.error).toBeInstanceOf(XmlStructureError);
  if (!(result.error instanceof XmlStructureError)) throw result.error;
  return result.error;
}

describe('XML error contracts', () => {
  it('retains typed complexity details and the exact diagnostic message', () => {
    const error = new XmlComplexityLimitError('maxXmlDepth', 3, 2);

    expect(error).toMatchObject({
      actual: 3,
      limit: 2,
      limitName: 'maxXmlDepth',
      message: 'XML resource limit maxXmlDepth exceeded: 3 > 2',
      name: 'XmlComplexityLimitError',
    });
  });

  it('retains the structure error name and cause', () => {
    const cause = new Error('parser detail');
    const error = new XmlStructureError('Invalid XML structure', { cause });

    expect(error.name).toBe('XmlStructureError');
    expect(error.message).toBe('Invalid XML structure');
    expect(error.cause).toBe(cause);
  });
});

describe('lossless XML normalization boundaries', () => {
  it.each([
    [0, 0],
    [12, 12],
    ['0', 0],
    ['12', 12],
  ])('accepts canonical legacy document order %j', (order, expected) => {
    expect(getXmlNodeOrder({ attrs: { order } })).toBe(expected);
  });

  it.each([
    -1,
    1.5,
    Number.POSITIVE_INFINITY,
    true,
    '+1',
    '01',
    '1.0',
    '9007199254740992',
  ])('rejects non-canonical legacy document order %j', (order) => {
    expect(getXmlNodeOrder({ attrs: { order } })).toBeUndefined();
  });

  it.each(CANONICAL_NAMESPACES)(
    'normalizes namespace %s to prefix %s',
    (namespace, prefix) => {
      const tagName = prefix ? `${prefix}:node` : 'node';
      const normalized = simplifyLossless([
        {
          attributes: { 'xmlns:source': namespace },
          children: [],
          tagName: 'source:node',
        },
      ]) as Record<string, unknown>;

      expect(normalized).toEqual({
        [tagName]: {
          attrs: { 'xmlns:source': namespace },
        },
      });
      expect(getXmlNodeOrder(normalized[tagName])).toBe(0);
    },
  );

  it.each(['a', 'a14', 'c', 'dgm', 'dsp', 'm', 'mc', 'p', 'r'])(
    'protects the reserved prefix %s from an unknown namespace',
    (prefix) => {
      expect(
        simplifyLossless([
          {
            attributes: { [`xmlns:${prefix}`]: 'urn:unknown' },
            children: [],
            tagName: `${prefix}:node`,
          },
        ]),
      ).toHaveProperty(`ns_${prefix}:node`);
    },
  );

  it('preserves text values together with parent attributes and order', () => {
    expect(simplifyLossless(['content'])).toBe('content');
    const normalized = simplifyLossless(['content'], {
      lang: 'en-US',
    });
    expect(normalized).toEqual({
      attrs: { lang: 'en-US' },
      value: 'content',
    });
    expect(getXmlNodeOrder(normalized)).toBe(0);
  });

  it('advances document order after an attributed text element', () => {
    const normalized = simplifyLossless([
      {
        attributes: { lang: 'en-US' },
        children: ['first'],
        tagName: 'text',
      },
      { children: [], tagName: 'next' },
    ]) as Record<string, unknown>;

    expect(normalized).toEqual({
      next: { attrs: {} },
      text: {
        attrs: { lang: 'en-US' },
        value: 'first',
      },
    });
    expect(getXmlNodeOrder(normalized.text)).toBe(0);
    expect(getXmlNodeOrder(normalized.next)).toBe(1);
  });

  it('preserves repeated empty elements and their document order', () => {
    const normalized = simplifyLossless([
      { children: [], tagName: 'item' },
      { children: [], tagName: 'item' },
    ]) as { item: unknown[] };

    expect(normalized).toEqual({
      item: [{ attrs: {} }, { attrs: {} }],
    });
    expect(normalized.item.map(getXmlNodeOrder)).toEqual([0, 1]);
  });

  it('treats absent children as empty and ignores the XML declaration node', () => {
    const normalized = simplifyLossless([
      { children: [], tagName: '?xml' },
      { tagName: 'leaf' },
    ]) as { leaf: unknown };

    expect(normalized).toEqual({ leaf: { attrs: {} } });
    expect(getXmlNodeOrder(normalized.leaf)).toBe(0);
  });

  it('does not silently discard visible mixed text', () => {
    expect(
      simplifyLossless(['visible', { children: [], tagName: 'child' }]),
    ).toBeUndefined();
  });

  it('ignores formatting whitespace around child elements', () => {
    const normalized = simplifyLossless([
      ' \n\t ',
      { children: [], tagName: 'child' },
    ]) as { child: unknown };

    expect(normalized).toEqual({ child: { attrs: {} } });
    expect(getXmlNodeOrder(normalized.child)).toBe(0);
  });

  it('applies a canonical default namespace only to element names', () => {
    const normalized = simplifyLossless([
      {
        attributes: { id: '7', xmlns: DRAWING_NAMESPACE },
        children: [],
        tagName: 'shape',
      },
    ]) as Record<string, unknown>;

    expect(normalized).toEqual({
      'a:shape': {
        attrs: { id: '7', xmlns: DRAWING_NAMESPACE },
      },
    });
    expect(getXmlNodeOrder(normalized['a:shape'])).toBe(0);
  });

  it('does not interpret ordinary attributes as namespace bindings', () => {
    const normalized = simplifyLossless([
      {
        attributes: {
          source: DRAWING_NAMESPACE,
          'xmlns:roo': DRAWING_NAMESPACE,
        },
        children: [],
        tagName: 'root',
      },
    ]) as { root: unknown };

    expect(normalized).toEqual({
      root: {
        attrs: {
          source: DRAWING_NAMESPACE,
          'xmlns:roo': DRAWING_NAMESPACE,
        },
      },
    });
    expect(getXmlNodeOrder(normalized.root)).toBe(0);
  });

  it('protects reserved canonical prefixes bound to unknown namespaces', () => {
    const normalized = simplifyLossless([
      {
        attributes: { 'xmlns:a': 'urn:custom-drawing' },
        children: [],
        tagName: 'a:shape',
      },
      {
        attributes: { 'xmlns:q': 'urn:custom-drawing' },
        children: [],
        tagName: 'q:shape',
      },
    ]) as Record<string, unknown>;

    expect(normalized).toEqual({
      'ns_a:shape': {
        attrs: { 'xmlns:a': 'urn:custom-drawing' },
      },
      'q:shape': {
        attrs: { 'xmlns:q': 'urn:custom-drawing' },
      },
    });
    expect(getXmlNodeOrder(normalized['ns_a:shape'])).toBe(0);
    expect(getXmlNodeOrder(normalized['q:shape'])).toBe(1);
  });

  it('keeps an authored order attribute separate from document order', () => {
    const normalized = simplifyLossless([
      {
        attributes: { order: '99' },
        children: [],
        tagName: 'first',
      },
      { children: [], tagName: 'second' },
    ]) as {
      first: { attrs: Record<string, string> };
      second: { attrs: Record<string, string> };
    };

    expect(normalized.first.attrs.order).toBe('99');
    expect(normalized.second.attrs.order).toBeUndefined();
    expect(getXmlNodeOrder(normalized.first)).toBe(0);
    expect(getXmlNodeOrder(normalized.second)).toBe(1);
    expect(JSON.stringify(normalized.second)).not.toContain('order');
  });

  it('reports the exact duplicate expanded attribute', () => {
    expect(() =>
      simplifyLossless([
        {
          attributes: {
            'xmlns:a': DRAWING_NAMESPACE,
            'xmlns:q': DRAWING_NAMESPACE,
            'a:id': 'first',
            'q:id': 'second',
          },
          children: [],
          tagName: 'root',
        },
      ]),
    ).toThrow('XML element has duplicate expanded attribute a:id');
  });
});

describe('strict XML validation boundaries', () => {
  it('accepts node and depth values exactly at their configured limits', () => {
    expect(
      assertXmlComplexity('<root><first/><second/></root>', {
        maxDepth: 2,
        maxNodes: 3,
      }),
    ).toBe(3);
  });

  it('decrements depth between sibling elements', () => {
    expect(() =>
      assertXmlComplexity('<root><first/><second/></root>', { maxDepth: 2 }),
    ).not.toThrow();
  });

  it('rejects XML 1.1 semantics even when requested by the declaration', () => {
    expect(() =>
      assertXmlComplexity('<?xml version="1.1"?><root>&#x1;</root>', {}),
    ).toThrow(XmlStructureError);
  });

  it('rejects document types with a stable policy error', () => {
    expect(() => assertXmlComplexity('<!DOCTYPE root><root/>', {})).toThrow(
      'XML document type declarations are not allowed',
    );
  });

  it('rejects invalid element and attribute qualified names precisely', () => {
    expect(() => assertXmlComplexity('<bad::root/>', {})).toThrow(
      'Invalid XML element name bad::root',
    );
    expect(() =>
      assertXmlComplexity('<root bad::attribute="value"/>', {}),
    ).toThrow('Invalid XML attribute name bad::attribute');
  });

  it('wraps parser failures with their original cause', () => {
    try {
      assertXmlComplexity('<root>&unknown;</root>', {});
      throw new Error('Expected invalid XML to fail');
    } catch (error) {
      expect(error).toBeInstanceOf(XmlStructureError);
      if (!(error instanceof XmlStructureError)) throw error;
      expect(error.message).toBe('Invalid XML structure');
      expect(error.cause).toBeInstanceOf(Error);
    }
  });

  it.each([
    Uint8Array.of(0xff, 0x3c, 0x00),
    Uint8Array.of(0x3c, 0xfe, 0x00),
    Uint8Array.of(0xfe, 0x3c, 0x00),
    Uint8Array.of(0x3c, 0xff, 0x00),
  ])('does not infer a BOM from only one matching byte: %j', async (bytes) => {
    const error = await readMalformedBytes(bytes);

    expect(error.message).toBe('Invalid UTF-8 XML');
    expect(error.cause).toBeInstanceOf(TypeError);
  });

  it('treats an empty filename exactly like a missing archive part', async () => {
    const zip = new JSZip();
    await expect(readXmlFileResult(zip, '')).resolves.toEqual({
      status: 'missing',
    });
  });

  it('distinguishes stream failures from resource-limit failures', async () => {
    const zip = new JSZip();
    zip.file('broken.xml', Promise.reject(new Error('stream failed')));

    const result = await readXmlFileResult(zip, 'broken.xml');

    expect(result).toMatchObject({ status: 'error', phase: 'read' });
    if (result.status !== 'error') throw new Error('Expected a read failure');
    expect(result.error).toBeInstanceOf(Error);
    if (!(result.error instanceof Error)) throw result.error;
    expect(result.error.message).toBe('stream failed');
  });
});
