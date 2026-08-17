import { TextEncoder } from 'node:util';

import { describe, expect, it } from 'vitest';

import {
  fetchSlidesManiaTemplate,
  powerPointDownloadUrlFromSlidesMania,
  SLIDESMANIA_MAX_PAGE_BYTES,
  SLIDESMANIA_MAX_PPTX_BYTES,
} from '../../scripts/reliability/fetch-slidesmania-template.mjs';
import {
  slidesManiaCorpus,
  slidesManiaCorpusProvenance,
} from '../../scripts/reliability/slidesmania-corpus.mjs';

const SOURCE_PAGE = 'https://slidesmania.com/example-template/';
const DOWNLOAD =
  'https://docs.google.com/presentation/d/presentation-id/export/pptx';

function page(download = DOWNLOAD) {
  return `<html><a class="download" href="${download}"><strong> Download   PowerPoint </strong></a></html>`;
}

function response(body, options = {}) {
  const bytes =
    typeof body === 'string' ? new TextEncoder().encode(body) : body;
  const headerValues = new Map(
    Object.entries(options.headers ?? {}).map(([name, value]) => [
      name.toLowerCase(),
      String(value),
    ]),
  );
  const status = options.status ?? 200;
  return {
    arrayBuffer: () => Promise.resolve(Uint8Array.from(bytes).buffer),
    headers: { get: (name) => headerValues.get(name.toLowerCase()) ?? null },
    ok: status >= 200 && status < 300,
    status,
  };
}

describe('SlidesMania reliability corpus', () => {
  it('selects 30 unique source pages without committing direct download URLs', () => {
    expect(slidesManiaCorpus).toHaveLength(30);
    expect(
      new Set(slidesManiaCorpus.map(({ sourcePage }) => sourcePage)).size,
    ).toBe(30);
    expect(
      slidesManiaCorpus.every(({ sourcePage }) =>
        sourcePage.startsWith('https://slidesmania.com/'),
      ),
    ).toBe(true);
    expect(JSON.stringify(slidesManiaCorpus)).not.toContain('/export/pptx');
    expect(slidesManiaCorpusProvenance).toMatchObject({
      homepage: 'https://slidesmania.com/',
    });
  });

  it('extracts only an exact trusted PowerPoint export link', () => {
    expect(powerPointDownloadUrlFromSlidesMania(page())).toBe(DOWNLOAD);
    expect(() =>
      powerPointDownloadUrlFromSlidesMania(
        page('https://evil.example/presentation/d/id/export/pptx'),
      ),
    ).toThrow('download URL is not trusted');
    expect(() => powerPointDownloadUrlFromSlidesMania('<p>none</p>')).toThrow(
      'exposes no PowerPoint download',
    );
  });

  it('downloads a bounded OOXML package through the declared source page', async () => {
    const calls = [];
    const fetchImplementation = async (url) => {
      calls.push(url);
      return calls.length === 1
        ? response(page())
        : response(Uint8Array.of(0x50, 0x4b, 0x03, 0x04, 1));
    };

    await expect(
      fetchSlidesManiaTemplate(
        { sourcePage: SOURCE_PAGE, title: 'Example' },
        fetchImplementation,
      ),
    ).resolves.toEqual({
      bytes: Uint8Array.of(0x50, 0x4b, 0x03, 0x04, 1),
      sourcePage: SOURCE_PAGE,
      title: 'Example',
    });
    expect(calls).toEqual([SOURCE_PAGE, DOWNLOAD]);
  });

  it.each([
    [
      'invalid entry',
      null,
      async () => response(page()),
      'corpus entry is invalid',
    ],
    [
      'invalid source',
      { sourcePage: 'https://evil.example/template/', title: 'Example' },
      async () => response(page()),
      'source page URL is invalid',
    ],
    [
      'source status',
      { sourcePage: SOURCE_PAGE, title: 'Example' },
      async () => response('no', { status: 503 }),
      'source page failed with status 503',
    ],
    [
      'empty source',
      { sourcePage: SOURCE_PAGE, title: 'Example' },
      async () => response(''),
      'source page returned empty content',
    ],
    [
      'invalid package',
      { sourcePage: SOURCE_PAGE, title: 'Example' },
      async (url) =>
        response(url === SOURCE_PAGE ? page() : Uint8Array.of(1, 2, 3, 4)),
      'download is not an OOXML package',
    ],
  ])('rejects %s', async (_name, entry, fetchImplementation, expected) => {
    await expect(
      fetchSlidesManiaTemplate(entry, fetchImplementation),
    ).rejects.toThrow(expected);
  });

  it.each([
    ['source', SLIDESMANIA_MAX_PAGE_BYTES, 'source page'],
    ['PowerPoint', SLIDESMANIA_MAX_PPTX_BYTES, 'PowerPoint download'],
  ])(
    'rejects an oversized %s response before reading it',
    async (kind, limit, label) => {
      let call = 0;
      const fetchImplementation = async () => {
        call += 1;
        if (kind === 'PowerPoint' && call === 1) return response(page());
        return response('x', {
          headers: { 'content-length': String(limit + 1) },
        });
      };
      await expect(
        fetchSlidesManiaTemplate(
          { sourcePage: SOURCE_PAGE, title: 'Example' },
          fetchImplementation,
        ),
      ).rejects.toThrow(`${label} exceeds its byte limit`);
    },
  );
});
