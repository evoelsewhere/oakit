import { describe, expect, it, vi } from 'vitest';

import {
  GOOGLE_SLIDES_MAX_EXPORT_BYTES,
  GOOGLE_SLIDES_MAX_SOURCE_BYTES,
  importGoogleSlidesPresentation,
  roundTripGoogleSlidesPresentation,
} from '../../scripts/reliability/google-slides-drive.mjs';

function jsonResponse(value, status = 200) {
  return new globalThis.Response(JSON.stringify(value), {
    headers: { 'content-type': 'application/json' },
    status,
  });
}

describe('controlled Google Slides transport', () => {
  it('imports PPTX bytes as a temporary Google Slides presentation', async () => {
    const fetchImplementation = vi.fn(async () =>
      jsonResponse({ id: 'temporary-id' }),
    );
    const source = new Uint8Array([1, 2, 3]);

    await expect(
      importGoogleSlidesPresentation(
        source,
        'secret-token',
        'temporary-deck',
        fetchImplementation,
      ),
    ).resolves.toBe('temporary-id');
    const [url, options] = fetchImplementation.mock.calls[0];
    expect(url).toContain('uploadType=multipart');
    expect(options.method).toBe('POST');
    expect(options.headers).toEqual({ Authorization: 'Bearer secret-token' });
    expect(options.body).toBeInstanceOf(globalThis.FormData);
    const entries = [...options.body.entries()];
    expect(entries.map(([name]) => name)).toEqual(['metadata', 'file']);
    const metadata = entries[0]?.[1];
    const file = entries[1]?.[1];
    expect(metadata).toBeInstanceOf(globalThis.Blob);
    expect(file).toBeInstanceOf(globalThis.Blob);
    if (
      !(metadata instanceof globalThis.Blob) ||
      !(file instanceof globalThis.Blob)
    ) {
      throw new Error('Expected multipart blobs');
    }
    expect(JSON.parse(await metadata.text())).toEqual({
      mimeType: 'application/vnd.google-apps.presentation',
      name: 'temporary-deck',
    });
    expect(new Uint8Array(await file.arrayBuffer())).toEqual(source);
  });

  it('exports PPTX bytes and always deletes the temporary presentation', async () => {
    const fetchImplementation = vi
      .fn()
      .mockResolvedValueOnce(jsonResponse({ id: 'temporary/id' }))
      .mockResolvedValueOnce(new globalThis.Response(new Uint8Array([4, 5, 6])))
      .mockResolvedValueOnce(new globalThis.Response(null, { status: 204 }));

    await expect(
      roundTripGoogleSlidesPresentation(
        new Uint8Array([1]),
        'token',
        'deck',
        fetchImplementation,
      ),
    ).resolves.toEqual(new Uint8Array([4, 5, 6]));
    expect(fetchImplementation).toHaveBeenCalledTimes(3);
    expect(fetchImplementation.mock.calls[1]?.[0]).toContain(
      'temporary%2Fid/export',
    );
    expect(fetchImplementation.mock.calls[2]?.[1]?.method).toBe('DELETE');
  });

  it('deletes the temporary presentation after an export failure', async () => {
    const fetchImplementation = vi
      .fn()
      .mockResolvedValueOnce(jsonResponse({ id: 'temporary-id' }))
      .mockResolvedValueOnce(new globalThis.Response(null, { status: 503 }))
      .mockResolvedValueOnce(new globalThis.Response(null, { status: 204 }));

    await expect(
      roundTripGoogleSlidesPresentation(
        new Uint8Array([1]),
        'token',
        'deck',
        fetchImplementation,
      ),
    ).rejects.toThrow('Google Drive export failed with status 503');
    expect(fetchImplementation.mock.calls[2]?.[1]?.method).toBe('DELETE');
  });

  it('retains a bounded Google error reason without exposing response details', async () => {
    const fetchImplementation = vi
      .fn()
      .mockResolvedValueOnce(jsonResponse({ id: 'temporary-id' }))
      .mockResolvedValueOnce(
        jsonResponse(
          {
            error: {
              errors: [
                {
                  message: 'secret server detail',
                  reason: 'exportSizeLimitExceeded',
                },
              ],
            },
          },
          403,
        ),
      )
      .mockResolvedValueOnce(new globalThis.Response(null, { status: 204 }));

    const result = roundTripGoogleSlidesPresentation(
      new Uint8Array([1]),
      'token',
      'deck',
      fetchImplementation,
    );
    await expect(result).rejects.toThrow(
      'Google Drive export failed with status 403 (exportSizeLimitExceeded)',
    );
    await expect(result).rejects.not.toThrow('secret server detail');
  });

  it('fails closed when export and cleanup both fail without exposing a token', async () => {
    const fetchImplementation = vi
      .fn()
      .mockResolvedValueOnce(jsonResponse({ id: 'temporary-id' }))
      .mockResolvedValueOnce(new globalThis.Response(null, { status: 503 }))
      .mockResolvedValueOnce(new globalThis.Response(null, { status: 403 }));

    const result = roundTripGoogleSlidesPresentation(
      new Uint8Array([1]),
      'do-not-log-this-token',
      'deck',
      fetchImplementation,
    );
    await expect(result).rejects.toThrow(
      'Google Slides export and cleanup both failed',
    );
    await expect(result).rejects.not.toThrow('do-not-log-this-token');
  });

  it('rejects missing ids, invalid sources, and bounded exports', async () => {
    await expect(
      importGoogleSlidesPresentation(
        new Uint8Array([1]),
        'token',
        'deck',
        async () => jsonResponse({}),
      ),
    ).rejects.toThrow('returned no presentation id');
    await expect(
      importGoogleSlidesPresentation(
        new Uint8Array(),
        'token',
        'deck',
        vi.fn(),
      ),
    ).rejects.toThrow('source must be non-empty bytes');
    class OversizedBytes extends Uint8Array {
      get byteLength() {
        return GOOGLE_SLIDES_MAX_SOURCE_BYTES + 1;
      }
    }
    await expect(
      importGoogleSlidesPresentation(
        new OversizedBytes(),
        'token',
        'deck',
        vi.fn(),
      ),
    ).rejects.toThrow('source exceeds the upload byte limit');

    const fetchImplementation = vi
      .fn()
      .mockResolvedValueOnce(jsonResponse({ id: 'temporary-id' }))
      .mockResolvedValueOnce(
        new globalThis.Response(null, {
          headers: {
            'content-length': String(GOOGLE_SLIDES_MAX_EXPORT_BYTES + 1),
          },
        }),
      )
      .mockResolvedValueOnce(new globalThis.Response(null, { status: 204 }));
    await expect(
      roundTripGoogleSlidesPresentation(
        new Uint8Array([1]),
        'token',
        'deck',
        fetchImplementation,
      ),
    ).rejects.toThrow('export exceeds the download byte limit');
    expect(fetchImplementation.mock.calls[2]?.[1]?.method).toBe('DELETE');
  });
});
