import { decodeXmlEntities } from './html';

/** Serialize one OOXML typeface as a quoted CSS font-family value. */
export function serializeCssFontFamily(value: string): string | null {
  const family = decodeXmlEntities(value).trim();
  if (!family) return null;

  let serialized = '"';
  for (const character of family) {
    const codePoint = character.codePointAt(0)!;
    if (character === '"' || character === '\\') {
      serialized += `\\${character}`;
    } else if (codePoint === 0) {
      serialized += '\uFFFD';
    } else if (codePoint <= 0x1f || codePoint === 0x7f) {
      serialized += `\\${codePoint.toString(16)} `;
    } else {
      serialized += character;
    }
  }
  return `${serialized}"`;
}

/** Accept only the RGB/RGBA hexadecimal colors emitted by DrawingML. */
export function normalizeHexColor(value: string): string | null {
  const hexadecimal = value.trim().replace(/^#/, '');
  return /^(?:[\da-f]{6}|[\da-f]{8})$/i.test(hexadecimal)
    ? `#${hexadecimal}`
    : null;
}
