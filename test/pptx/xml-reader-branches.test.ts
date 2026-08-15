import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';

import type {
  StreamableZipObject,
  ZipEntryStream,
} from '../../src/common/archive/read-entry';
import { PptxParseError } from '../../src/formats/pptx/errors';
import { defaultPptxResourceLimits } from '../../src/formats/pptx/internal/resource-limits';
import { PptxXmlReader } from '../../src/formats/pptx/internal/xml-reader';
import type { PptxDiagnostic } from '../../src/formats/pptx/types';

function createArchive(): JSZip {
  const zip = new JSZip();
  zip.file('first.xml', '<root><child id="first"/></root>');
  zip.file('second.xml', '<root><child id="second"/></root>');
  zip.file('broken.xml', '<root><child></root>');
  zip.file('broken-again.xml', '<root><other></root>');
  zip.file('media.bin', Uint8Array.from([0, 127, 128, 255]));
  return zip;
}

function createReadFailureArchive(error: unknown): JSZip {
  const listeners = new Map<string, (value?: unknown) => void>();
  const streamImplementation = {
    on(event: string, listener: (value?: unknown) => void) {
      listeners.set(event, listener);
      return streamImplementation;
    },
    pause() {
      return streamImplementation;
    },
    resume() {
      listeners.get('error')?.(error);
      return streamImplementation;
    },
  };
  const stream = streamImplementation as unknown as ZipEntryStream;
  const entry: StreamableZipObject = {
    internalStream() {
      return stream;
    },
    name: 'unreadable.xml',
  };
  return {
    file(filename: string) {
      return filename === entry.name ? entry : null;
    },
  } as unknown as JSZip;
}

function limits(
  overrides: Partial<ReturnType<typeof defaultPptxResourceLimits>>,
): ReturnType<typeof defaultPptxResourceLimits> {
  return { ...defaultPptxResourceLimits(), ...overrides };
}

describe('PptxXmlReader branch contracts', () => {
  it('treats an absent optional part as empty without a diagnostic', async () => {
    const diagnostics: PptxDiagnostic[] = [];
    const reader = new PptxXmlReader(createArchive(), 'tolerant', diagnostics);

    await expect(reader.read('missing.xml')).resolves.toEqual({});
    await expect(reader.read('')).resolves.toEqual({});
    expect(diagnostics).toEqual([]);
  });

  it('reports an empty required part name without inventing a part field', async () => {
    const diagnostics: PptxDiagnostic[] = [];
    const reader = new PptxXmlReader(createArchive(), 'tolerant', diagnostics);

    await expect(reader.read('', { required: true })).resolves.toEqual({});
    expect(diagnostics).toEqual([
      {
        code: 'missing-required-part',
        message: 'Required OOXML part name is empty',
        severity: 'error',
      },
    ]);
  });

  it('reports distinct missing required parts with exact messages', async () => {
    const diagnostics: PptxDiagnostic[] = [];
    const reader = new PptxXmlReader(createArchive(), 'tolerant', diagnostics);

    await reader.read('missing-one.xml', { required: true });
    await reader.read('missing-two.xml', { required: true });

    expect(diagnostics).toEqual([
      {
        code: 'missing-required-part',
        message: 'Required OOXML part is missing: missing-one.xml',
        part: 'missing-one.xml',
        severity: 'error',
      },
      {
        code: 'missing-required-part',
        message: 'Required OOXML part is missing: missing-two.xml',
        part: 'missing-two.xml',
        severity: 'error',
      },
    ]);
  });

  it('makes an invalid required part an error diagnostic in tolerant mode', async () => {
    const diagnostics: PptxDiagnostic[] = [];
    const reader = new PptxXmlReader(createArchive(), 'tolerant', diagnostics);

    await expect(
      reader.read('broken.xml', { required: true }),
    ).resolves.toEqual({});
    expect(diagnostics).toHaveLength(1);
    expect(diagnostics[0]).toMatchObject({
      code: 'xml-parse-failed',
      part: 'broken.xml',
      severity: 'error',
    });
    expect(diagnostics[0]?.message).toContain(
      'Failed to parse OOXML part broken.xml:',
    );
  });

  it('keeps failures for distinct malformed XML parts separate', async () => {
    const diagnostics: PptxDiagnostic[] = [];
    const reader = new PptxXmlReader(createArchive(), 'tolerant', diagnostics);

    await reader.read('broken.xml');
    await reader.read('broken-again.xml');

    expect(diagnostics).toHaveLength(2);
    expect(diagnostics.map(({ code, part }) => ({ code, part }))).toEqual([
      { code: 'xml-parse-failed', part: 'broken.xml' },
      { code: 'xml-parse-failed', part: 'broken-again.xml' },
    ]);
  });

  it('distinguishes an archive read failure from an XML parse failure', async () => {
    const diagnostics: PptxDiagnostic[] = [];
    const reader = new PptxXmlReader(
      createReadFailureArchive(new Error('storage unavailable')),
      'tolerant',
      diagnostics,
    );

    await expect(
      reader.read('unreadable.xml', { required: true }),
    ).resolves.toEqual({});
    expect(diagnostics).toEqual([
      {
        code: 'xml-read-failed',
        message:
          'Failed to read OOXML part unreadable.xml: storage unavailable',
        part: 'unreadable.xml',
        severity: 'error',
      },
    ]);
  });

  it('caches successful XML without charging its budgets twice', async () => {
    const diagnostics: PptxDiagnostic[] = [];
    const reader = new PptxXmlReader(
      createArchive(),
      'tolerant',
      diagnostics,
      limits({ maxTotalXmlNodes: 2 }),
    );

    const first = await reader.read('first.xml', { required: true });
    const cached = await reader.read('first.xml', { required: true });

    expect(cached).toBe(first);
    expect(diagnostics).toEqual([]);
  });

  it('charges distinct XML parts to the cumulative node budget', async () => {
    const diagnostics: PptxDiagnostic[] = [];
    const reader = new PptxXmlReader(
      createArchive(),
      'tolerant',
      diagnostics,
      limits({ maxTotalXmlNodes: 3 }),
    );

    await expect(reader.read('first.xml')).resolves.not.toEqual({});
    await expect(reader.read('second.xml')).rejects.toMatchObject({
      diagnostic: {
        code: 'resource-limit-exceeded',
        part: 'second.xml',
        severity: 'error',
      },
      name: 'PptxParseError',
    });
    expect(diagnostics).toHaveLength(1);
  });

  it('returns exact media bytes and null for an absent media part', async () => {
    const reader = new PptxXmlReader(createArchive(), 'tolerant', []);

    await expect(reader.readMedia('media.bin')).resolves.toEqual(
      Uint8Array.from([0, 127, 128, 255]),
    );
    await expect(reader.readMedia('missing.bin')).resolves.toBeNull();
  });

  it('reports the media part that exceeds its individual byte limit', async () => {
    const diagnostics: PptxDiagnostic[] = [];
    const reader = new PptxXmlReader(
      createArchive(),
      'tolerant',
      diagnostics,
      limits({ maxMediaBytes: 3 }),
    );

    let thrown: unknown;
    try {
      await reader.readMedia('media.bin');
    } catch (error) {
      thrown = error;
    }
    expect(thrown).toBeInstanceOf(PptxParseError);
    if (!(thrown instanceof PptxParseError)) throw new Error('Expected error');
    expect(thrown.cause).toMatchObject({
      actual: 4,
      limit: 3,
      limitName: 'maxMediaBytes',
      part: 'media.bin',
    });
    expect(diagnostics).toEqual([
      {
        code: 'resource-limit-exceeded',
        message:
          'PPTX resource limit maxMediaBytes exceeded for media.bin: 4 > 3',
        part: 'media.bin',
        severity: 'error',
      },
    ]);
  });

  it('allows media exactly at the total expansion boundary', async () => {
    const reader = new PptxXmlReader(
      createArchive(),
      'tolerant',
      [],
      limits({ maxTotalUncompressedBytes: 4 }),
    );

    await expect(reader.readMedia('media.bin')).resolves.toEqual(
      Uint8Array.from([0, 127, 128, 255]),
    );
  });

  it('reports XML that exceeds its individual expanded byte limit', async () => {
    const diagnostics: PptxDiagnostic[] = [];
    const reader = new PptxXmlReader(
      createArchive(),
      'tolerant',
      diagnostics,
      limits({ maxXmlBytes: 3 }),
    );

    await expect(reader.read('first.xml')).rejects.toBeInstanceOf(
      PptxParseError,
    );
    expect(diagnostics).toHaveLength(1);
    expect(diagnostics[0]).toMatchObject({
      code: 'resource-limit-exceeded',
      part: 'first.xml',
      severity: 'error',
    });
    expect(diagnostics[0]?.message).toContain(
      'PPTX resource limit maxXmlBytes exceeded for first.xml:',
    );
  });

  it('reports media that exhausts the total expansion budget', async () => {
    const diagnostics: PptxDiagnostic[] = [];
    const reader = new PptxXmlReader(
      createArchive(),
      'tolerant',
      diagnostics,
      limits({ maxTotalUncompressedBytes: 3 }),
    );

    await expect(reader.readMedia('media.bin')).rejects.toBeInstanceOf(
      PptxParseError,
    );
    expect(diagnostics).toEqual([
      {
        code: 'resource-limit-exceeded',
        message:
          'PPTX resource limit maxTotalUncompressedBytes exceeded for media.bin: 4 > 3',
        part: 'media.bin',
        severity: 'error',
      },
    ]);
  });

  it('preserves an external relationship without warning', () => {
    const diagnostics: PptxDiagnostic[] = [];
    const reader = new PptxXmlReader(createArchive(), 'tolerant', diagnostics);

    expect(
      reader.resolveRelationshipTarget(
        'ppt/slides/slide1.xml',
        'https://example.com/video.mp4',
        'External',
      ),
    ).toBe('https://example.com/video.mp4');
    expect(diagnostics).toEqual([]);
  });

  it('deduplicates invalid relationship diagnostics in tolerant mode', () => {
    const diagnostics: PptxDiagnostic[] = [];
    const reader = new PptxXmlReader(createArchive(), 'tolerant', diagnostics);

    for (let attempt = 0; attempt < 2; attempt += 1) {
      expect(
        reader.resolveRelationshipTarget(
          'ppt/slides/slide1.xml',
          '..&#47;..&#47;..&#47;secret.xml',
        ),
      ).toBeNull();
    }

    expect(diagnostics).toHaveLength(1);
    expect(diagnostics[0]?.message).toContain(
      'Invalid relationship target ..&#47;..&#47;..&#47;secret.xml in ppt/slides/slide1.xml:',
    );
  });

  it('reports distinct invalid relationship targets independently', () => {
    const diagnostics: PptxDiagnostic[] = [];
    const reader = new PptxXmlReader(createArchive(), 'tolerant', diagnostics);

    expect(
      reader.resolveRelationshipTarget(
        'ppt/slides/slide1.xml',
        '../../../secret-one.xml',
      ),
    ).toBeNull();
    expect(
      reader.resolveRelationshipTarget(
        'ppt/slides/slide1.xml',
        '../../../secret-two.xml',
      ),
    ).toBeNull();

    expect(diagnostics).toHaveLength(2);
    expect(diagnostics.map(({ message }) => message)).toEqual([
      expect.stringContaining('../../../secret-one.xml'),
      expect.stringContaining('../../../secret-two.xml'),
    ]);
  });

  it('throws the typed relationship diagnostic with its cause in strict mode', () => {
    const diagnostics: PptxDiagnostic[] = [];
    const reader = new PptxXmlReader(createArchive(), 'strict', diagnostics);

    let thrown: unknown;
    try {
      reader.resolveRelationshipTarget(
        'ppt/slides/slide1.xml',
        '../../../secret.xml',
      );
    } catch (error) {
      thrown = error;
    }

    expect(thrown).toBeInstanceOf(PptxParseError);
    if (!(thrown instanceof PptxParseError)) throw new Error('Expected error');
    expect(thrown.cause).toBeInstanceOf(Error);
    expect(diagnostics).toHaveLength(1);
    expect(diagnostics[0]).toMatchObject({
      code: 'invalid-relationship-target',
      part: 'ppt/slides/slide1.xml',
      severity: 'warning',
    });
  });
});
