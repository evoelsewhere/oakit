const DRIVE_API = 'https://www.googleapis.com/drive/v3';
const DRIVE_UPLOAD_API = 'https://www.googleapis.com/upload/drive/v3';
const GOOGLE_SLIDES_MIME = 'application/vnd.google-apps.presentation';
const PPTX_MIME =
  'application/vnd.openxmlformats-officedocument.presentationml.presentation';

export const GOOGLE_SLIDES_MAX_EXPORT_BYTES = 10 * 1024 * 1024;
export const GOOGLE_SLIDES_MAX_SOURCE_BYTES = 50 * 1024 * 1024;

function authorization(accessToken) {
  if (typeof accessToken !== 'string' || accessToken.trim().length === 0) {
    throw new TypeError('Google Drive access token must be non-empty');
  }
  return { Authorization: `Bearer ${accessToken}` };
}

async function assertDriveResponse(response, operation) {
  if (!response.ok) {
    let reason;
    try {
      const value = await response.json();
      const candidate = value?.error?.errors?.[0]?.reason;
      if (typeof candidate === 'string' && /^[\w-]{1,80}$/.test(candidate)) {
        reason = candidate;
      }
    } catch {
      reason = undefined;
    }
    throw new Error(
      `Google Drive ${operation} failed with status ${response.status}${reason === undefined ? '' : ` (${reason})`}`,
    );
  }
}

export async function importGoogleSlidesPresentation(
  source,
  accessToken,
  name,
  fetchImplementation = globalThis.fetch,
) {
  if (!(source instanceof Uint8Array) || source.byteLength === 0) {
    throw new TypeError('Google Slides source must be non-empty bytes');
  }
  if (source.byteLength > GOOGLE_SLIDES_MAX_SOURCE_BYTES) {
    throw new RangeError('Google Slides source exceeds the upload byte limit');
  }
  if (typeof name !== 'string' || name.length === 0) {
    throw new TypeError('Google Slides temporary name must be non-empty');
  }
  const body = new globalThis.FormData();
  body.append(
    'metadata',
    new globalThis.Blob(
      [JSON.stringify({ mimeType: GOOGLE_SLIDES_MIME, name })],
      {
        type: 'application/json',
      },
    ),
  );
  body.append('file', new globalThis.Blob([source], { type: PPTX_MIME }));
  const response = await fetchImplementation(
    `${DRIVE_UPLOAD_API}/files?uploadType=multipart&fields=id`,
    {
      body,
      headers: authorization(accessToken),
      method: 'POST',
    },
  );
  await assertDriveResponse(response, 'import');
  const value = await response.json();
  if (
    value === null ||
    typeof value !== 'object' ||
    Array.isArray(value) ||
    typeof value.id !== 'string' ||
    value.id.length === 0
  ) {
    throw new Error('Google Drive import returned no presentation id');
  }
  return value.id;
}

export async function exportGoogleSlidesPresentation(
  fileId,
  accessToken,
  fetchImplementation = globalThis.fetch,
) {
  const response = await fetchImplementation(
    `${DRIVE_API}/files/${encodeURIComponent(fileId)}/export?mimeType=${encodeURIComponent(PPTX_MIME)}`,
    { headers: authorization(accessToken) },
  );
  await assertDriveResponse(response, 'export');
  const declaredLength = Number(response.headers.get('content-length'));
  if (
    Number.isFinite(declaredLength) &&
    declaredLength > GOOGLE_SLIDES_MAX_EXPORT_BYTES
  ) {
    throw new RangeError(
      'Google Slides export exceeds the download byte limit',
    );
  }
  const output = new Uint8Array(await response.arrayBuffer());
  if (output.byteLength === 0) {
    throw new Error('Google Slides export returned an empty presentation');
  }
  if (output.byteLength > GOOGLE_SLIDES_MAX_EXPORT_BYTES) {
    throw new RangeError(
      'Google Slides export exceeds the download byte limit',
    );
  }
  return output;
}

export async function deleteGoogleSlidesPresentation(
  fileId,
  accessToken,
  fetchImplementation = globalThis.fetch,
) {
  const response = await fetchImplementation(
    `${DRIVE_API}/files/${encodeURIComponent(fileId)}`,
    { headers: authorization(accessToken), method: 'DELETE' },
  );
  await assertDriveResponse(response, 'cleanup');
}

export async function roundTripGoogleSlidesPresentation(
  source,
  accessToken,
  name,
  fetchImplementation = globalThis.fetch,
) {
  const fileId = await importGoogleSlidesPresentation(
    source,
    accessToken,
    name,
    fetchImplementation,
  );
  let operationError;
  let output;
  try {
    output = await exportGoogleSlidesPresentation(
      fileId,
      accessToken,
      fetchImplementation,
    );
  } catch (error) {
    operationError = error;
  }
  let cleanupError;
  try {
    await deleteGoogleSlidesPresentation(
      fileId,
      accessToken,
      fetchImplementation,
    );
  } catch (error) {
    cleanupError = error;
  }
  if (operationError !== undefined && cleanupError !== undefined) {
    throw new AggregateError(
      [operationError, cleanupError],
      'Google Slides export and cleanup both failed',
    );
  }
  if (operationError !== undefined) throw operationError;
  if (cleanupError !== undefined) throw cleanupError;
  return output;
}
