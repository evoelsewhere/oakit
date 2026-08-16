function base64Alphabet(): string {
  return 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/';
}

function isCanonicalBase64(input: string): boolean {
  return /^(?:[A-Za-z0-9+/]{4})*(?:[A-Za-z0-9+/]{2}==|[A-Za-z0-9+/]{3}=)?$/.test(
    input,
  );
}

function invalidBase64(): never {
  throw new RangeError('Invalid canonical Base64 encoding');
}

/** Return the decoded size after validating canonical RFC 4648 Base64. */
export function decodedBase64ByteLength(input: string): number {
  if (!isCanonicalBase64(input)) invalidBase64();
  const alphabet = base64Alphabet();
  const padding = input.endsWith('==') ? 2 : input.endsWith('=') ? 1 : 0;
  if (padding === 2) {
    const value = alphabet.indexOf(input.charAt(input.length - 3));
    if ((value & 15) !== 0) invalidBase64();
  } else if (padding === 1) {
    const value = alphabet.indexOf(input.charAt(input.length - 2));
    if ((value & 3) !== 0) invalidBase64();
  }
  return (input.length / 4) * 3 - padding;
}

/** Decode canonical Base64 without relying on Node.js Buffer or atob. */
export function decodeBase64(input: string): Uint8Array {
  const byteLength = decodedBase64ByteLength(input);
  const alphabet = base64Alphabet();
  const output = new Uint8Array(byteLength);
  for (const match of input.matchAll(/.{4}/g)) {
    const inputIndex = match.index;
    const outputIndex = (inputIndex / 4) * 3;
    const chunk = match[0];
    const first = alphabet.indexOf(chunk.charAt(0));
    const second = alphabet.indexOf(chunk.charAt(1));
    const thirdCharacter = chunk.charAt(2);
    const fourthCharacter = chunk.charAt(3);
    const third = alphabet.indexOf(thirdCharacter);
    const fourth = alphabet.indexOf(fourthCharacter);

    output[outputIndex] = (first << 2) | (second >> 4);
    output[outputIndex + 1] = ((second & 15) << 4) | (third >> 2);
    output[outputIndex + 2] = ((third & 3) << 6) | fourth;
  }
  return output;
}

/** Encode binary data without relying on Node.js Buffer. */
export function encodeBase64(input: ArrayBuffer | ArrayBufferView): string {
  const alphabet = base64Alphabet();
  const bytes =
    input instanceof ArrayBuffer
      ? new Uint8Array(input)
      : new Uint8Array(input.buffer, input.byteOffset, input.byteLength);
  const mainLength = bytes.byteLength - (bytes.byteLength % 3);
  let base64 = '';

  for (const [index] of bytes.subarray(0, mainLength).entries()) {
    if (index % 3 !== 0) continue;
    const chunk =
      (bytes[index]! << 16) | (bytes[index + 1]! << 8) | bytes[index + 2]!;
    base64 += alphabet[(chunk & 16_515_072) >> 18];
    base64 += alphabet[(chunk & 258_048) >> 12];
    base64 += alphabet[(chunk & 4_032) >> 6];
    base64 += alphabet[chunk & 63];
  }

  const remainder = bytes.byteLength % 3;
  if (remainder === 1) {
    const chunk = bytes[mainLength]!;
    base64 += alphabet[(chunk & 252) >> 2];
    base64 += `${alphabet[(chunk & 3) << 4]}==`;
  } else if (remainder === 2) {
    const chunk = (bytes[mainLength]! << 8) | bytes[mainLength + 1]!;
    base64 += alphabet[(chunk & 64_512) >> 10];
    base64 += alphabet[(chunk & 1_008) >> 4];
    base64 += `${alphabet[(chunk & 15) << 2]}=`;
  }

  return base64;
}

/** @internal Compatibility name used by the upstream PPTX parser. */
export const base64ArrayBuffer = encodeBase64;
