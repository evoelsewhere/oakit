const URI_SCHEME = /^[a-z][a-z\d+.-]*:/i;

function stripQueryAndFragment(value: string): string {
  const suffixIndex = value.search(/[?#]/);
  return suffixIndex < 0 ? value : value.slice(0, suffixIndex);
}

function normalizedSegments(value: string): string[] {
  const segments: string[] = [];
  for (const segment of value.replace(/\\/g, '/').split('/')) {
    if (!segment || segment === '.') continue;
    if (segment === '..') {
      if (segments.length === 0) {
        throw new RangeError(
          'OPC relationship target escapes the package root',
        );
      }
      segments.pop();
      continue;
    }
    segments.push(segment);
  }
  return segments;
}

/** Return the `.rels` part owned by an OPC package part. */
export function getRelationshipPartUri(ownerPart: string): string {
  const normalizedOwner = normalizedSegments(ownerPart).join('/');
  const separatorIndex = normalizedOwner.lastIndexOf('/');
  const directory =
    separatorIndex < 0 ? '' : normalizedOwner.slice(0, separatorIndex + 1);
  const filename = normalizedOwner.slice(separatorIndex + 1);
  if (!filename) throw new TypeError('OPC owner part must name a file');
  return `${directory}_rels/${filename}.rels`;
}

/** Resolve an internal relationship target relative to its owning OPC part. */
export function resolvePartUri(ownerPart: string, target: string): string {
  const normalizedTarget = stripQueryAndFragment(target.trim()).replace(
    /\\/g,
    '/',
  );
  if (!normalizedTarget)
    throw new TypeError('OPC relationship target is empty');
  if (URI_SCHEME.test(normalizedTarget)) {
    throw new TypeError('External URI requires TargetMode="External"');
  }

  if (normalizedTarget.startsWith('/')) {
    return normalizedSegments(normalizedTarget).join('/');
  }

  const ownerSegments = normalizedSegments(ownerPart);
  if (ownerSegments.length === 0) {
    throw new TypeError('OPC owner part must name a file');
  }
  ownerSegments.pop();
  return normalizedSegments(
    [...ownerSegments, normalizedTarget].join('/'),
  ).join('/');
}

/** Preserve external targets and resolve internal targets against their owner. */
export function resolveRelationshipTarget(
  ownerPart: string,
  target: string,
  targetMode?: string,
): string {
  return targetMode?.toLowerCase() === 'external'
    ? target.trim()
    : resolvePartUri(ownerPart, target);
}
