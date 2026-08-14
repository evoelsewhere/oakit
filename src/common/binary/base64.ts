/** Encode binary data without relying on Node.js Buffer. */
export function encodeBase64(input: ArrayBuffer | ArrayBufferView): string {
  const encodings =
    'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/';
  const bytes =
    input instanceof ArrayBuffer
      ? new Uint8Array(input)
      : new Uint8Array(input.buffer, input.byteOffset, input.byteLength);
  const mainLength = bytes.byteLength - (bytes.byteLength % 3);
  let base64 = '';

  for (let index = 0; index < mainLength; index += 3) {
    const chunk =
      (bytes[index]! << 16) | (bytes[index + 1]! << 8) | bytes[index + 2]!;
    base64 += encodings[(chunk & 16_515_072) >> 18];
    base64 += encodings[(chunk & 258_048) >> 12];
    base64 += encodings[(chunk & 4_032) >> 6];
    base64 += encodings[chunk & 63];
  }

  const remainder = bytes.byteLength % 3;
  if (remainder === 1) {
    const chunk = bytes[mainLength]!;
    base64 += encodings[(chunk & 252) >> 2];
    base64 += `${encodings[(chunk & 3) << 4]}==`;
  } else if (remainder === 2) {
    const chunk = (bytes[mainLength]! << 8) | bytes[mainLength + 1]!;
    base64 += encodings[(chunk & 64_512) >> 10];
    base64 += encodings[(chunk & 1_008) >> 4];
    base64 += `${encodings[(chunk & 15) << 2]}=`;
  }

  return base64;
}

/** @internal Compatibility name used by the upstream PPTX parser. */
export const base64ArrayBuffer = encodeBase64;
