import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';

import { serializePowerPointArchive } from '../../src/formats/pptx/writer/archive';
import type { PptxSerializedPart } from '../../src/formats/pptx/writer/parts';

const PARTS: readonly PptxSerializedPart[] = [
  {
    path: '[Content_Types].xml',
    xml: '<?xml version="1.0"?><Types><Repeated>aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa</Repeated></Types>',
  },
  {
    path: '_rels/.rels',
    xml: '<?xml version="1.0"?><Relationships/>',
  },
];

describe('PowerPoint archive serialization', () => {
  it('generates byte-identical archives sequentially', async () => {
    const first = await serializePowerPointArchive(PARTS);
    const second = await serializePowerPointArchive(PARTS);

    expect(second).toEqual(first);
  });

  it('isolates concurrent archive writers', async () => {
    const [first, second, third] = await Promise.all([
      serializePowerPointArchive(PARTS),
      serializePowerPointArchive(PARTS),
      serializePowerPointArchive(PARTS),
    ]);

    expect(second).toEqual(first);
    expect(third).toEqual(first);
  });

  it('preserves canonical part order without synthetic folder entries', async () => {
    const bytes = await serializePowerPointArchive(PARTS);
    const archive = await JSZip.loadAsync(bytes);

    expect(Object.keys(archive.files)).toEqual([
      '[Content_Types].xml',
      '_rels/.rels',
    ]);
    expect(Object.values(archive.files).every((entry) => !entry.dir)).toBe(
      true,
    );
  });

  it('round-trips every part without changing its UTF-8 XML', async () => {
    const archive = await JSZip.loadAsync(
      await serializePowerPointArchive(PARTS),
    );

    await expect(archive.file('[Content_Types].xml')?.async('string')).resolves
      .toBe(PARTS[0]?.xml);
    await expect(archive.file('_rels/.rels')?.async('string')).resolves.toBe(
      PARTS[1]?.xml,
    );
  });

  it('writes the fixed minimum DOS date and DEFLATE method', async () => {
    const bytes = await serializePowerPointArchive(PARTS);

    expect(Array.from(bytes.slice(0, 4))).toEqual([0x50, 0x4b, 0x03, 0x04]);
    expect(Array.from(bytes.slice(6, 8))).toEqual([0x00, 0x00]);
    expect(Array.from(bytes.slice(8, 10))).toEqual([0x08, 0x00]);
    expect(Array.from(bytes.slice(10, 14))).toEqual([0x00, 0x00, 0x21, 0x00]);

    const archive = await JSZip.loadAsync(bytes);
    expect(archive.file('[Content_Types].xml')?.date.toISOString()).toBe(
      '1980-01-01T00:00:00.000Z',
    );
    expect(archive.file('_rels/.rels')?.date.toISOString()).toBe(
      '1980-01-01T00:00:00.000Z',
    );
  });

  it('does not mutate caller-owned part records', async () => {
    const parts = PARTS.map((part) => ({ ...part }));
    const before = structuredClone(parts);

    await serializePowerPointArchive(parts);

    expect(parts).toEqual(before);
  });

  it('emits a valid empty ZIP for an empty internal inventory', async () => {
    const bytes = await serializePowerPointArchive([]);
    const archive = await JSZip.loadAsync(bytes);

    expect(Object.keys(archive.files)).toEqual([]);
    expect(Array.from(bytes.slice(0, 4))).toEqual([0x50, 0x4b, 0x05, 0x06]);
  });
});
