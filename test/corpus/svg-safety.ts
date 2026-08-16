import { SaxesParser } from 'saxes';

const EMBEDDED_BITMAP =
  /^data:image\/(?:gif|jpeg|png|webp);base64,(?:[A-Za-z\d+/]{4})*(?:[A-Za-z\d+/]{2}==|[A-Za-z\d+/]{3}=)?$/i;
const INTERNAL_FRAGMENT = /^#[^\s]+$/;

function isSelfContainedReference(value: string): boolean {
  return INTERNAL_FRAGMENT.test(value) || EMBEDDED_BITMAP.test(value);
}

function inspectCssReferences(value: string, features: Set<string>): void {
  if (/@import\b/i.test(value)) features.add('stylesheet import');

  const cssUrl = /url\(\s*(?:"([^"]*)"|'([^']*)'|([^)]*?))\s*\)/giu;
  for (const match of value.matchAll(cssUrl)) {
    const reference = match[1] ?? match[2] ?? match[3]?.trim() ?? '';
    if (!isSelfContainedReference(reference)) {
      features.add(`external CSS reference: ${reference}`);
    }
  }
  if (/url\s*\(/i.test(value.replace(cssUrl, ''))) {
    features.add('malformed CSS reference');
  }
}

export function findUnsafeSvgFeatures(source: string): string[] {
  const features = new Set<string>();
  let styleDepth = 0;
  const parser = new SaxesParser({ xmlns: true });

  parser.on('doctype', () => {
    features.add('document type declaration');
  });
  parser.on('processinginstruction', ({ target }) => {
    if (target.toLowerCase() === 'xml-stylesheet') {
      features.add('XML stylesheet instruction');
    }
  });
  parser.on('opentag', (tag) => {
    if (tag.local === 'script' || tag.local === 'foreignObject') {
      features.add(`forbidden element: ${tag.local}`);
    }
    if (tag.local === 'style') styleDepth += 1;

    for (const attribute of Object.values(tag.attributes)) {
      if (
        (attribute.local === 'href' || attribute.local === 'src') &&
        !isSelfContainedReference(attribute.value)
      ) {
        features.add(`external attribute reference: ${attribute.value}`);
      }
      inspectCssReferences(attribute.value, features);
    }
  });
  parser.on('text', (value) => {
    if (styleDepth > 0) inspectCssReferences(value, features);
  });
  parser.on('cdata', (value) => {
    if (styleDepth > 0) inspectCssReferences(value, features);
  });
  parser.on('closetag', (tag) => {
    if (tag.local === 'style') styleDepth -= 1;
  });
  parser.on('error', () => {
    features.add('malformed XML');
  });
  parser.write(source).close();

  return [...features];
}
