const OFFICE_TEXT_ESCAPE_PATTERN = /_x([0-9a-f]{4})_/gi;

/** Decode the UTF-16 code-unit escape syntax used by Office Open XML text. */
export function decodeOfficeTextEscapes(value: string): string {
  return value.replace(OFFICE_TEXT_ESCAPE_PATTERN, (_match, hexadecimal) =>
    String.fromCharCode(Number.parseInt(String(hexadecimal), 16)),
  );
}
