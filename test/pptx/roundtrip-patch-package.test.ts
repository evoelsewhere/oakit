import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';

import { resolvePptxResourceLimits } from '../../src/formats/pptx/internal/resource-limits';
import {
  assertSafeEditablePptxPackage,
  decodeEditablePptxXml,
  generatePptxPatchedArchive,
  hasPptxUtf16Bom,
  readPptxPartPayloads,
  verifyPptxPatchedPayloads,
} from '../../src/formats/pptx/roundtrip/package-preservation';

function bytes(...values: number[]): Uint8Array {
  return new Uint8Array(values);
}

describe('PowerPoint patched package helpers', () => {
  it.each([
    '_xmlsignatures/sig1.xml',
    '_XMLSIGNATURES/sig1.xml',
    'ppt/vbaProject.bin',
    'ppt/nested/vbaProject.bin',
    'PPT/VBAPROJECT.BIN',
  ])('rejects protected package part %s', (name) => {
    const archive = new JSZip();
    archive.file(name, bytes(1), { createFolders: false });

    expect(() => assertSafeEditablePptxPackage(archive)).toThrow(
      'PowerPoint text edit does not modify signed or macro-enabled packages',
    );
  });

  it('accepts an ordinary package part', () => {
    const archive = new JSZip();
    archive.file('ppt/slides/slide1.xml', '<p:sld/>');

    expect(() => assertSafeEditablePptxPackage(archive)).not.toThrow();
  });

  it.each([
    [bytes(0xff, 0xfe, 0x3c, 0x00), 'little-endian BOM'],
    [bytes(0xfe, 0xff, 0x00, 0x3c), 'big-endian BOM'],
  ])('rejects UTF-16 %s', (input) => {
    expect(() =>
      decodeEditablePptxXml(input, resolvePptxResourceLimits()),
    ).toThrow('PowerPoint text edit requires UTF-8 slide XML');
  });

  it.each([
    [bytes(0xff, 0xfe), true],
    [bytes(0xfe, 0xff), true],
    [bytes(0xff, 0xff), false],
    [bytes(0xff, 0x00), false],
    [bytes(0x00, 0xfe), false],
    [bytes(0xfe, 0xfe), false],
  ])('classifies BOM %j as UTF-16=%s', (input, expected) => {
    expect(hasPptxUtf16Bom(input)).toBe(expected);
  });

  it.each([
    '<?xml version="1.0" encoding="UTF-16"?><p:sld/>',
    "<?xml version='1.0' encoding = 'utf-16'?><p:sld/>",
    '<?xml version="1.0" encoding=\t"UTF-16"?><p:sld/>',
  ])('rejects declared UTF-16 source %j', (xml) => {
    expect(() =>
      decodeEditablePptxXml(
        new TextEncoder().encode(xml),
        resolvePptxResourceLimits(),
      ),
    ).toThrow('PowerPoint text edit requires UTF-8 slide XML');
  });

  it('decodes ordinary UTF-8 XML and forwards complexity limits', () => {
    const xml = '<?xml version="1.0" encoding="UTF-8"?><p:sld/>\n';
    expect(
      decodeEditablePptxXml(
        new TextEncoder().encode(xml),
        resolvePptxResourceLimits(),
      ),
    ).toBe(xml);
    const limits = resolvePptxResourceLimits();
    limits.maxXmlNodes = 1;
    expect(() =>
      decodeEditablePptxXml(
        new TextEncoder().encode('<root><child/></root>'),
        limits,
      ),
    ).toThrow('XML resource limit maxXmlNodes exceeded: 2 > 1');
  });

  it('omits directory entries while reading part payloads', async () => {
    const archive = new JSZip();
    archive.folder('ppt/slides');
    archive.file('ppt/slides/slide1.xml', '<p:sld/>');

    const payloads = await readPptxPartPayloads(
      archive,
      resolvePptxResourceLimits(),
    );

    expect([...payloads.keys()]).toEqual(['ppt/slides/slide1.xml']);
  });

  it('uses deterministic archive generation options', async () => {
    let received: unknown;
    const archive = {
      generateAsync(options: unknown) {
        received = options;
        return Promise.resolve(bytes(1, 2, 3));
      },
    } as unknown as JSZip;

    await expect(generatePptxPatchedArchive(archive)).resolves.toEqual(
      bytes(1, 2, 3),
    );
    expect(received).toEqual({
      compression: 'DEFLATE',
      compressionOptions: { level: 9 },
      platform: 'DOS',
      streamFiles: false,
      type: 'uint8array',
    });
  });

  it('accepts changed dirty payloads and exact untouched payloads', () => {
    const source = new Map([
      ['dirty.xml', bytes(1, 2)],
      ['exact.xml', bytes(3, 4)],
    ]);
    const output = new Map([
      ['dirty.xml', bytes(1, 3)],
      ['exact.xml', bytes(3, 4)],
    ]);

    expect(() =>
      verifyPptxPatchedPayloads(source, output, new Set(['dirty.xml'])),
    ).not.toThrow();
  });

  it.each([
    [
      'extra part',
      new Map([
        ['dirty.xml', bytes(1, 3)],
        ['exact.xml', bytes(3, 4)],
        ['extra.xml', bytes(5)],
      ]),
      'PowerPoint text edit changed the package part inventory',
    ],
    [
      'missing part',
      new Map([['dirty.xml', bytes(1, 3)]]),
      'PowerPoint text edit changed the package part inventory',
    ],
    [
      'substituted part',
      new Map([
        ['dirty.xml', bytes(1, 3)],
        ['other.xml', bytes(3, 4)],
      ]),
      'PowerPoint text edit changed the package part inventory',
    ],
    [
      'unchanged dirty part',
      new Map([
        ['dirty.xml', bytes(1, 2)],
        ['exact.xml', bytes(3, 4)],
      ]),
      'PowerPoint text edit did not change dirty part dirty.xml',
    ],
    [
      'changed untouched part',
      new Map([
        ['dirty.xml', bytes(1, 3)],
        ['exact.xml', bytes(3, 5)],
      ]),
      'PowerPoint text edit changed untouched part exact.xml',
    ],
    [
      'shortened untouched part',
      new Map([
        ['dirty.xml', bytes(1, 3)],
        ['exact.xml', bytes(3)],
      ]),
      'PowerPoint text edit changed untouched part exact.xml',
    ],
    [
      'extended untouched part',
      new Map([
        ['dirty.xml', bytes(1, 3)],
        ['exact.xml', bytes(3, 4, 5)],
      ]),
      'PowerPoint text edit changed untouched part exact.xml',
    ],
  ])('rejects %s', (_name, output, message) => {
    const source = new Map([
      ['dirty.xml', bytes(1, 2)],
      ['exact.xml', bytes(3, 4)],
    ]);

    expect(() =>
      verifyPptxPatchedPayloads(source, output, new Set(['dirty.xml'])),
    ).toThrow(message);
  });
});
