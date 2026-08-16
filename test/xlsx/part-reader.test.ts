import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';

import type {
  StreamableZipObject,
  ZipEntryStream,
} from '../../src/common/archive/read-entry';
import { XlsxParseError } from '../../src/formats/xlsx/errors';
import { XlsxPartReader } from '../../src/formats/xlsx/internal/part-reader';
import {
  defaultXlsxResourceLimits,
  type ResolvedXlsxResourceLimits,
} from '../../src/formats/xlsx/internal/resource-limits';
import type { XlsxDiagnostic } from '../../src/formats/xlsx/types';

function archive(): JSZip {
  const zip = new JSZip();
  zip.file('first.xml', '<root><child id="first"/></root>');
  zip.file('second.xml', '<root><child id="second"/></root>');
  zip.file('invalid.xml', '<root><child></root>');
  zip.file(
    'doctype.xml',
    '<!DOCTYPE root [<!ENTITY payload "secret">]><root>&payload;</root>',
  );
  zip.file('invalid-utf8.xml', Uint8Array.from([0x3c, 0x72, 0x3e, 0xff]));
  zip.file('media.bin', Uint8Array.from([0, 127, 128, 255]));
  return zip;
}

function limits(
  overrides: Partial<ResolvedXlsxResourceLimits>,
): ResolvedXlsxResourceLimits {
  return { ...defaultXlsxResourceLimits(), ...overrides };
}

function readFailureArchive(error: unknown): JSZip {
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
  const entry: StreamableZipObject = {
    internalStream() {
      return streamImplementation as unknown as ZipEntryStream;
    },
    name: 'unreadable.xml',
  };
  return {
    file(part: string) {
      return part === entry.name ? entry : null;
    },
  } as unknown as JSZip;
}

function captureParseError(action: Promise<unknown>): Promise<XlsxParseError> {
  return action.then(
    () => {
      throw new Error('Expected XLSX part read to fail');
    },
    (error: unknown) => {
      expect(error).toBeInstanceOf(XlsxParseError);
      return error as XlsxParseError;
    },
  );
}

describe('XLSX bounded part reader', () => {
  it('returns a cached XML tree without charging budgets twice', async () => {
    const diagnostics: XlsxDiagnostic[] = [];
    const reader = new XlsxPartReader(
      archive(),
      diagnostics,
      limits({ maxTotalXmlNodes: 2 }),
    );

    const first = await reader.readXml('first.xml', { required: true });
    const cached = await reader.readXml('first.xml', { required: true });

    expect(cached).toBe(first);
    expect(first).toMatchObject({
      root: { child: { attrs: { id: 'first' } } },
    });
    expect(diagnostics).toEqual([]);
  });

  it('accepts expanded XML exactly at the cumulative byte limit', async () => {
    const reader = new XlsxPartReader(
      archive(),
      [],
      limits({ maxTotalUncompressedBytes: 32 }),
    );

    await expect(reader.readXml('first.xml')).resolves.toMatchObject({
      root: { child: {} },
    });
  });

  it('returns null for an absent optional part', async () => {
    const diagnostics: XlsxDiagnostic[] = [];
    const reader = new XlsxPartReader(archive(), diagnostics);

    await expect(reader.readXml('missing.xml')).resolves.toBeNull();
    expect(diagnostics).toEqual([]);
  });

  it.each(['missing.xml', ''])(
    'throws a structured missing-required-part error for %s',
    async (part) => {
      const diagnostics: XlsxDiagnostic[] = [];
      const reader = new XlsxPartReader(archive(), diagnostics);
      const error = await captureParseError(
        reader.readXml(part, { required: true }),
      );

      expect(error.diagnostic).toEqual({
        code: 'missing-required-part',
        message: part
          ? `Required XLSX part is missing: ${part}`
          : 'Required XLSX part name is empty',
        ...(part ? { part } : {}),
        severity: 'error',
      });
      expect(diagnostics).toEqual([error.diagnostic]);
    },
  );

  it.each([
    ['invalid.xml', 'xml-parse-failed'],
    ['doctype.xml', 'xml-parse-failed'],
    ['invalid-utf8.xml', 'xml-parse-failed'],
  ] as const)('rejects unsafe XML for %s', async (part, code) => {
    const diagnostics: XlsxDiagnostic[] = [];
    const reader = new XlsxPartReader(archive(), diagnostics);
    const error = await captureParseError(reader.readXml(part));

    expect(error.diagnostic).toMatchObject({
      code,
      part,
      severity: 'error',
    });
    expect(error.cause).toBeInstanceOf(Error);
    expect(diagnostics).toEqual([error.diagnostic]);
  });

  it('distinguishes archive read failures from XML failures', async () => {
    const diagnostics: XlsxDiagnostic[] = [];
    const reader = new XlsxPartReader(
      readFailureArchive(new Error('storage unavailable')),
      diagnostics,
    );
    const error = await captureParseError(reader.readXml('unreadable.xml'));

    expect(error.diagnostic).toEqual({
      code: 'xml-read-failed',
      message: 'Failed to read XLSX part unreadable.xml',
      part: 'unreadable.xml',
      severity: 'error',
    });
    expect((error.cause as Error).message).toBe('storage unavailable');
  });

  it.each([
    [{ maxXmlBytes: 31 }, 'first.xml', 'maxXmlBytes', 32, 31],
    [{ maxXmlDepth: 1 }, 'first.xml', 'maxXmlDepth', 2, 1],
    [{ maxXmlNodes: 1 }, 'first.xml', 'maxXmlNodes', 2, 1],
    [{ maxTotalXmlNodes: 3 }, 'second.xml', 'maxTotalXmlNodes', 4, 3],
    [
      { maxTotalUncompressedBytes: 63 },
      'second.xml',
      'maxTotalUncompressedBytes',
      65,
      63,
    ],
  ] as const)(
    'maps XML budget failures to %s',
    async (overrides, failingPart, limitName, actual, limit) => {
      const diagnostics: XlsxDiagnostic[] = [];
      const reader = new XlsxPartReader(
        archive(),
        diagnostics,
        limits(overrides),
      );
      if (failingPart === 'second.xml') {
        await reader.readXml('first.xml');
      }
      const error = await captureParseError(reader.readXml(failingPart));

      expect(error.diagnostic).toMatchObject({
        actual,
        code: 'resource-limit-exceeded',
        limit,
        limitName,
        part: failingPart,
        severity: 'error',
      });
      expect(error.cause).toMatchObject({
        actual,
        limit,
        limitName,
        name: 'XlsxResourceLimitError',
        part: failingPart,
      });
      expect(diagnostics).toEqual([error.diagnostic]);
    },
  );

  it('returns exact binary bytes and null for a missing part', async () => {
    const reader = new XlsxPartReader(archive(), []);

    await expect(
      reader.readBytes('media.bin', 'maxMediaBytes'),
    ).resolves.toEqual(Uint8Array.from([0, 127, 128, 255]));
    await expect(
      reader.readBytes('missing.bin', 'maxMediaBytes'),
    ).resolves.toBeNull();
  });

  it.each([
    [{ maxMediaBytes: 3 }, 'maxMediaBytes', 4, 3],
    [{ maxTotalUncompressedBytes: 3 }, 'maxTotalUncompressedBytes', 4, 3],
  ] as const)(
    'enforces binary byte budget %s',
    async (overrides, limitName, actual, limit) => {
      const diagnostics: XlsxDiagnostic[] = [];
      const reader = new XlsxPartReader(
        archive(),
        diagnostics,
        limits(overrides),
      );
      const error = await captureParseError(
        reader.readBytes('media.bin', 'maxMediaBytes'),
      );

      expect(error.diagnostic).toMatchObject({
        actual,
        code: 'resource-limit-exceeded',
        limit,
        limitName,
        part: 'media.bin',
        severity: 'error',
      });
      expect(diagnostics).toEqual([error.diagnostic]);
    },
  );
});
