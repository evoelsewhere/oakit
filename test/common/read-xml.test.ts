import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';

import { readXmlFile } from '../../src/common/xml/read-xml';

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
});
