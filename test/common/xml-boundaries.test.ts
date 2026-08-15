import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';

import {
  assertXmlComplexity,
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
  it.each(CANONICAL_NAMESPACES)(
    'normalizes namespace %s to prefix %s',
    (namespace, prefix) => {
      const tagName = prefix ? `${prefix}:node` : 'node';

      expect(
        simplifyLossless([
          {
            attributes: { 'xmlns:source': namespace },
            children: [],
            tagName: 'source:node',
          },
        ]),
      ).toEqual({
        [tagName]: {
          attrs: { order: 0, 'xmlns:source': namespace },
        },
      });
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
    expect(simplifyLossless(['content'], { lang: 'en-US' })).toEqual({
      attrs: { lang: 'en-US', order: 0 },
      value: 'content',
    });
  });

  it('advances document order after an attributed text element', () => {
    expect(
      simplifyLossless([
        {
          attributes: { lang: 'en-US' },
          children: ['first'],
          tagName: 'text',
        },
        { children: [], tagName: 'next' },
      ]),
    ).toEqual({
      next: { attrs: { order: 2 } },
      text: {
        attrs: { lang: 'en-US', order: 0 },
        value: 'first',
      },
    });
  });

  it('preserves repeated empty elements and their document order', () => {
    expect(
      simplifyLossless([
        { children: [], tagName: 'item' },
        { children: [], tagName: 'item' },
      ]),
    ).toEqual({
      item: [{ attrs: { order: 0 } }, { attrs: { order: 1 } }],
    });
  });

  it('treats absent children as empty and ignores the XML declaration node', () => {
    expect(
      simplifyLossless([
        { children: [], tagName: '?xml' },
        { tagName: 'leaf' },
      ]),
    ).toEqual({ leaf: { attrs: { order: 0 } } });
  });

  it('does not silently discard visible mixed text', () => {
    expect(
      simplifyLossless(['visible', { children: [], tagName: 'child' }]),
    ).toBeUndefined();
  });

  it('ignores formatting whitespace around child elements', () => {
    expect(
      simplifyLossless([' \n\t ', { children: [], tagName: 'child' }]),
    ).toEqual({ child: { attrs: { order: 0 } } });
  });

  it('applies a canonical default namespace only to element names', () => {
    expect(
      simplifyLossless([
        {
          attributes: { id: '7', xmlns: DRAWING_NAMESPACE },
          children: [],
          tagName: 'shape',
        },
      ]),
    ).toEqual({
      'a:shape': {
        attrs: { id: '7', order: 0, xmlns: DRAWING_NAMESPACE },
      },
    });
  });

  it('does not interpret ordinary attributes as namespace bindings', () => {
    expect(
      simplifyLossless([
        {
          attributes: {
            source: DRAWING_NAMESPACE,
            'xmlns:roo': DRAWING_NAMESPACE,
          },
          children: [],
          tagName: 'root',
        },
      ]),
    ).toEqual({
      root: {
        attrs: {
          order: 0,
          source: DRAWING_NAMESPACE,
          'xmlns:roo': DRAWING_NAMESPACE,
        },
      },
    });
  });

  it('protects reserved canonical prefixes bound to unknown namespaces', () => {
    expect(
      simplifyLossless([
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
      ]),
    ).toEqual({
      'ns_a:shape': {
        attrs: { order: 0, 'xmlns:a': 'urn:custom-drawing' },
      },
      'q:shape': {
        attrs: { order: 1, 'xmlns:q': 'urn:custom-drawing' },
      },
    });
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
