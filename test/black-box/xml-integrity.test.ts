import { describe, expect, it } from 'vitest';

import { parsePptx } from '../../src';
import {
  createIndependentPptx,
  OFFICE_REL_NS,
  PRESENTATION_NS,
} from './pptx-package';

function validPresentation(size = '9144000'): string {
  return `<p:presentation xmlns:p="${PRESENTATION_NS}" xmlns:r="${OFFICE_REL_NS}">
    <p:sldIdLst><p:sldId id="256" r:id="rIdSlide1"/></p:sldIdLst>
    <p:sldSz cx="${size}" cy="5143500"/>
  </p:presentation>`;
}

function encodeUtf16Le(value: string): Uint8Array {
  const bytes = new Uint8Array(2 + value.length * 2);
  bytes[0] = 0xff;
  bytes[1] = 0xfe;
  for (let index = 0; index < value.length; index++) {
    const codeUnit = value.charCodeAt(index);
    bytes[2 + index * 2] = codeUnit & 0xff;
    bytes[3 + index * 2] = codeUnit >> 8;
  }
  return bytes;
}

async function expectStrictXmlFailure(
  presentationXml: string | Uint8Array,
): Promise<void> {
  const input = await createIndependentPptx({
    'ppt/presentation.xml': presentationXml,
  });

  await expect(parsePptx(input, { errorMode: 'strict' })).rejects.toMatchObject(
    {
      diagnostic: {
        code: 'xml-parse-failed',
        part: 'ppt/presentation.xml',
      },
    },
  );
}

describe('PPTX XML integrity black-box cases', () => {
  it('rejects multiple document roots', async () => {
    await expectStrictXmlFailure(
      `${validPresentation()}<p:presentation xmlns:p="${PRESENTATION_NS}"/>`,
    );
  });

  it('rejects document type declarations in OOXML parts', async () => {
    await expectStrictXmlFailure(
      `<!DOCTYPE p:presentation [<!ENTITY injected "text">]>${validPresentation()}`,
    );
  });

  it('rejects undeclared XML entity references', async () => {
    await expectStrictXmlFailure(validPresentation('9144&unknown;000'));
  });

  it('rejects unquoted XML attribute values', async () => {
    await expectStrictXmlFailure(
      validPresentation().replace('cx="9144000"', 'cx=9144000'),
    );
  });

  it('rejects duplicate XML attributes', async () => {
    await expectStrictXmlFailure(
      validPresentation().replace('cx="9144000"', 'cx="9144000" cx="18288000"'),
    );
  });

  it('rejects duplicate expanded attribute names behind namespace aliases', async () => {
    const presentation = validPresentation()
      .replace(
        `xmlns:r="${OFFICE_REL_NS}"`,
        `xmlns:r="${OFFICE_REL_NS}" xmlns:rel="${OFFICE_REL_NS}"`,
      )
      .replace('r:id="rIdSlide1"', 'r:id="rIdSlide1" rel:id="rIdMissing"');

    await expectStrictXmlFailure(presentation);
  });

  it('rejects an XML declaration inside the document root', async () => {
    await expectStrictXmlFailure(
      validPresentation().replace(
        '<p:sldIdLst>',
        '<?xml version="1.0"?><p:sldIdLst>',
      ),
    );
  });

  it('rejects double hyphens inside XML comments', async () => {
    await expectStrictXmlFailure(
      validPresentation().replace(
        '<p:sldIdLst>',
        '<!-- invalid -- comment --><p:sldIdLst>',
      ),
    );
  });

  it('rejects a CDATA terminator in regular character data', async () => {
    await expectStrictXmlFailure(
      validPresentation().replace('<p:sldIdLst>', 'invalid]]><p:sldIdLst>'),
    );
  });

  it('rejects malformed UTF-8 instead of replacing bytes silently', async () => {
    const prefix = new TextEncoder().encode(
      validPresentation().replace('9144000', '9144'),
    );
    const marker = new TextEncoder().encode('9144');
    const markerIndex = prefix.findIndex((value, index) =>
      marker.every((item, offset) => prefix[index + offset] === item),
    );
    const bytes = new Uint8Array(prefix.byteLength + 1);
    bytes.set(prefix.slice(0, markerIndex + marker.byteLength), 0);
    bytes[markerIndex + marker.byteLength] = 0xff;
    bytes.set(
      prefix.slice(markerIndex + marker.byteLength),
      markerIndex + marker.byteLength + 1,
    );

    await expectStrictXmlFailure(bytes);
  });

  it('accepts a valid UTF-16LE OOXML part with a byte-order mark', async () => {
    const input = await createIndependentPptx({
      'ppt/presentation.xml': encodeUtf16Le(
        `<?xml version="1.0" encoding="UTF-16"?>${validPresentation()}`,
      ),
    });

    const result = await parsePptx(input, { errorMode: 'strict' });

    expect(result.size).toEqual({ width: 720, height: 405 });
    expect(result.slides).toHaveLength(1);
  });
});
