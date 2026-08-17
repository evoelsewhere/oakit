import { URL } from 'node:url';
import { TextDecoder } from 'node:util';

export const SLIDESMANIA_MAX_PAGE_BYTES = 2 * 1024 * 1024;
export const SLIDESMANIA_MAX_PPTX_BYTES = 50 * 1024 * 1024;

function declaredLength(response) {
  const value = Number(response.headers.get('content-length'));
  return Number.isFinite(value) ? value : null;
}

async function responseBytes(response, maximumBytes, label) {
  if (!response.ok) {
    throw new Error(`${label} failed with status ${response.status}`);
  }
  const length = declaredLength(response);
  if (length !== null && length > maximumBytes) {
    throw new RangeError(`${label} exceeds its byte limit`);
  }
  const bytes = new Uint8Array(await response.arrayBuffer());
  if (bytes.byteLength === 0)
    throw new Error(`${label} returned empty content`);
  if (bytes.byteLength > maximumBytes) {
    throw new RangeError(`${label} exceeds its byte limit`);
  }
  return bytes;
}

function decodeHtmlAttribute(value) {
  return value.replaceAll('&amp;', '&');
}

export function powerPointDownloadUrlFromSlidesMania(html) {
  const anchors = html.matchAll(/<a\b([^>]*)>([\s\S]*?)<\/a\s*>/gi);
  for (const anchor of anchors) {
    const label = (anchor[2] ?? '')
      .replace(/<[^>]*>/g, '')
      .replace(/\s+/g, ' ')
      .trim()
      .toUpperCase();
    if (label !== 'DOWNLOAD POWERPOINT') continue;
    const attributes = anchor[1] ?? '';
    const href = /\bhref\s*=\s*(?:"([^"]*)"|'([^']*)')/i.exec(attributes);
    const value = href?.[1] ?? href?.[2];
    if (value === undefined) break;
    const url = new URL(decodeHtmlAttribute(value));
    if (
      url.protocol !== 'https:' ||
      url.hostname !== 'docs.google.com' ||
      !/^\/presentation\/d\/[\w-]+\/export\/pptx$/.test(url.pathname)
    ) {
      throw new Error('SlidesMania PowerPoint download URL is not trusted');
    }
    return url.toString();
  }
  throw new Error('SlidesMania page exposes no PowerPoint download');
}

function validateSourcePage(sourcePage) {
  const url = new URL(sourcePage);
  if (
    url.protocol !== 'https:' ||
    url.hostname !== 'slidesmania.com' ||
    url.pathname === '/'
  ) {
    throw new TypeError('SlidesMania source page URL is invalid');
  }
  return url.toString();
}

export async function fetchSlidesManiaTemplate(
  entry,
  fetchImplementation = globalThis.fetch,
) {
  if (
    entry === null ||
    typeof entry !== 'object' ||
    Array.isArray(entry) ||
    typeof entry.title !== 'string' ||
    entry.title.length === 0 ||
    typeof entry.sourcePage !== 'string'
  ) {
    throw new TypeError('SlidesMania corpus entry is invalid');
  }
  const sourcePage = validateSourcePage(entry.sourcePage);
  const pageResponse = await fetchImplementation(sourcePage);
  const pageBytes = await responseBytes(
    pageResponse,
    SLIDESMANIA_MAX_PAGE_BYTES,
    'SlidesMania source page',
  );
  const downloadUrl = powerPointDownloadUrlFromSlidesMania(
    new TextDecoder().decode(pageBytes),
  );
  const pptxResponse = await fetchImplementation(downloadUrl);
  const bytes = await responseBytes(
    pptxResponse,
    SLIDESMANIA_MAX_PPTX_BYTES,
    'SlidesMania PowerPoint download',
  );
  if (
    bytes[0] !== 0x50 ||
    bytes[1] !== 0x4b ||
    bytes[2] !== 0x03 ||
    bytes[3] !== 0x04
  ) {
    throw new Error('SlidesMania PowerPoint download is not an OOXML package');
  }
  return { bytes, sourcePage, title: entry.title };
}
