const OFFICE_ESCAPE_PATTERN = /_x[0-9a-f]{4}_/gi;

function escapeOfficeSequences(value: string): string {
  return value.replace(
    OFFICE_ESCAPE_PATTERN,
    (match) => `_x005F_${match.slice(1)}`,
  );
}

export function escapeXmlText(value: string): string {
  return escapeOfficeSequences(value)
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;');
}

export function escapeXmlAttribute(value: string): string {
  return escapeXmlText(value)
    .replaceAll('"', '&quot;')
    .replaceAll("'", '&apos;');
}

function needsPreservedSpace(value: string): boolean {
  return (
    /^\s/.test(value) ||
    /\s$/.test(value) ||
    / {2}/.test(value) ||
    /[\t\r\n]/.test(value)
  );
}

export function serializeDrawingText(value: string): string {
  const preserve = needsPreservedSpace(value) ? ' xml:space="preserve"' : '';
  return `<a:t${preserve}>${escapeXmlText(value)}</a:t>`;
}
