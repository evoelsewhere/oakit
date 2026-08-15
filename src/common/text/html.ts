const HTML_ENTITIES: Readonly<Record<string, string>> = {
  '&': '&amp;',
  '<': '&lt;',
  '>': '&gt;',
  '"': '&quot;',
  "'": '&#039;',
};

const SAFE_HYPERLINK_PROTOCOLS = new Set(['http:', 'https:', 'mailto:']);

export function escapeHtml(text: string): string {
  return text.replace(/[&<>"']/g, (character) => HTML_ENTITIES[character]!);
}

export function sanitizeHyperlink(value: string): string | null {
  const candidate = value.trim();
  if (!candidate) return null;

  try {
    const url = new URL(candidate);
    return SAFE_HYPERLINK_PROTOCOLS.has(url.protocol.toLowerCase())
      ? candidate
      : null;
  } catch {
    return null;
  }
}

export function hasValidText(html: string): boolean {
  if (typeof DOMParser === 'undefined') {
    return (
      html
        .replace(/<[^>]+>/g, '')
        .replace(/\s+/g, ' ')
        .trim() !== ''
    );
  }

  const document = new DOMParser().parseFromString(html, 'text/html');
  return (document.body.textContent ?? '').trim() !== '';
}
