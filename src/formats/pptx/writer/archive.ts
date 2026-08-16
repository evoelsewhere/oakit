import JSZip from 'jszip';

import type { PptxSerializedPart } from './parts';

const POWERPOINT_ARCHIVE_DATE = new Date(Date.UTC(1980, 0, 1));

export async function serializePowerPointArchive(
  parts: readonly PptxSerializedPart[],
): Promise<Uint8Array> {
  const archive = new JSZip();
  for (const part of parts) {
    archive.file(part.path, part.xml, {
      createFolders: false,
      date: POWERPOINT_ARCHIVE_DATE,
    });
  }
  return archive.generateAsync({
    compression: 'DEFLATE',
    compressionOptions: { level: 9 },
    platform: 'DOS',
    streamFiles: false,
    type: 'uint8array',
  });
}
