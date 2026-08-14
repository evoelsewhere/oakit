const HTML_ENTITIES: Readonly<Record<string, string>> = {
  '&': '&amp;',
  '<': '&lt;',
  '>': '&gt;',
  '"': '&quot;',
  "'": '&#039;',
};

export function escapeHtml(text: string): string {
  return text.replace(/[&<>"']/g, (character) => HTML_ENTITIES[character]!);
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
