import { canonicalJson } from './canonical-json';

function hexadecimal(bytes: Uint8Array): string {
  return Array.from(bytes, (byte) => byte.toString(16).padStart(2, '0')).join(
    '',
  );
}

export async function sha256Bytes(bytes: Uint8Array): Promise<string> {
  const owned = new Uint8Array(bytes.byteLength);
  owned.set(bytes);
  const digest = await globalThis.crypto.subtle.digest('SHA-256', owned);
  return hexadecimal(new Uint8Array(digest));
}

export function sha256Text(value: string): Promise<string> {
  return sha256Bytes(new TextEncoder().encode(value));
}

export function canonicalSha256(value: unknown): Promise<string> {
  return sha256Text(canonicalJson(value));
}
