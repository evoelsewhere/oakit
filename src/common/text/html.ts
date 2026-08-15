function decodeNamedXmlEntity(name: string): string | undefined {
  switch (name) {
    case 'amp':
      return '&';
    case 'apos':
      return "'";
    case 'gt':
      return '>';
    case 'lt':
      return '<';
    case 'quot':
      return '"';
  }
}

function escapeHtmlCharacter(character: string): string {
  switch (character) {
    case '&':
      return '&amp;';
    case '<':
      return '&lt;';
    case '>':
      return '&gt;';
    case '"':
      return '&quot;';
    case "'":
      return '&#039;';
    default:
      return character;
  }
}

function isSafeHyperlinkProtocol(protocol: string): boolean {
  switch (protocol) {
    case 'http:':
    case 'https:':
    case 'mailto:':
      return true;
    default:
      return false;
  }
}

export function decodeXmlEntities(text: string): string {
  return text.replace(
    /&(?:#(\d+)|#x([\da-f]+)|(amp|apos|gt|lt|quot));/gi,
    (entity, decimal: string, hexadecimal: string, named: string) => {
      if (named) return decodeNamedXmlEntity(named.toLowerCase()) ?? entity;
      const codePoint = Number.parseInt(
        decimal || hexadecimal,
        decimal ? 10 : 16,
      );
      if (!Number.isInteger(codePoint) || codePoint > 0x10ffff) {
        return entity;
      }
      return String.fromCodePoint(codePoint);
    },
  );
}

export function escapeHtml(text: string): string {
  return text.replace(/[&<>"']/g, escapeHtmlCharacter);
}

export function sanitizeHyperlink(value: string): string | null {
  const candidate = value.trim();

  try {
    const url = new URL(candidate);
    return isSafeHyperlinkProtocol(url.protocol.toLowerCase())
      ? candidate
      : null;
  } catch {
    return null;
  }
}

export function hasValidText(html: string): boolean {
  if (typeof DOMParser === 'undefined') {
    return /\S/.test(html.replace(/<[^>]+>/g, ''));
  }

  const document = new DOMParser().parseFromString(html, 'text/html');
  return (document.body.textContent ?? '').trim() !== '';
}
