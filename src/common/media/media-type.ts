const MIME_TYPES: Readonly<Record<string, string>> = {
  avi: 'video/avi',
  emf: 'image/x-emf',
  gif: 'image/gif',
  jpeg: 'image/jpeg',
  jpg: 'image/jpeg',
  mp3: 'audio/mpeg',
  mp4: 'video/mp4',
  mpg: 'video/mpg',
  ogg: 'video/ogg',
  png: 'image/png',
  svg: 'image/svg+xml',
  tif: 'image/tiff',
  tiff: 'image/tiff',
  wav: 'audio/wav',
  webm: 'video/webm',
  wmf: 'image/x-wmf',
  wmv: 'video/wmv',
};

export function extractFileExtension(filename: string): string {
  const index = filename.lastIndexOf('.');
  return index < 0 ? '' : filename.slice(index + 1);
}

export function getMimeType(extension: string): string {
  return MIME_TYPES[extension.toLowerCase()] ?? '';
}

export function isVideoLink(value: string): boolean {
  try {
    const protocol = new URL(value).protocol;
    return protocol === 'http:' || protocol === 'https:' || protocol === 'ftp:';
  } catch {
    return false;
  }
}
