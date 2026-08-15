import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';

import { readXmlFile, readXmlFileResult } from '../../src/common/xml/read-xml';

function createXmlArchive(): JSZip {
  const zip = new JSZip();
  zip.file(
    'document.xml',
    '<root><first value="1"/><second>content</second></root>',
  );
  return zip;
}

describe('readXmlFile', () => {
  it('assigns deterministic order values across sequential reads', async () => {
    const zip = createXmlArchive();

    const first = await readXmlFile(zip, 'document.xml');
    const second = await readXmlFile(zip, 'document.xml');

    expect(second).toEqual(first);
  });

  it('isolates order state across concurrent reads', async () => {
    const zip = createXmlArchive();

    const [first, second] = await Promise.all([
      readXmlFile(zip, 'document.xml'),
      readXmlFile(zip, 'document.xml'),
    ]);

    expect(second).toEqual(first);
  });

  it('distinguishes missing parts from invalid XML', async () => {
    const zip = createXmlArchive();
    zip.file('invalid.xml', '<root><child></root>');

    await expect(readXmlFileResult(zip, 'missing.xml')).resolves.toEqual({
      status: 'missing',
    });
    await expect(readXmlFileResult(zip, 'invalid.xml')).resolves.toMatchObject({
      status: 'error',
      phase: 'parse',
    });
  });
});
