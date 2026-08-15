import { describe, expect, it } from 'vitest';

import {
  extractFileExtension,
  getMimeType,
  isVideoLink,
} from '../../src/common/media/media-type';

describe('media type helpers', () => {
  it.each([
    ['avi', 'video/avi'],
    ['emf', 'image/x-emf'],
    ['gif', 'image/gif'],
    ['jpeg', 'image/jpeg'],
    ['jpg', 'image/jpeg'],
    ['mp3', 'audio/mpeg'],
    ['mp4', 'video/mp4'],
    ['mpg', 'video/mpg'],
    ['ogg', 'video/ogg'],
    ['png', 'image/png'],
    ['svg', 'image/svg+xml'],
    ['tif', 'image/tiff'],
    ['tiff', 'image/tiff'],
    ['wav', 'audio/wav'],
    ['webm', 'video/webm'],
    ['wmf', 'image/x-wmf'],
    ['wmv', 'video/wmv'],
  ])('maps .%s to %s', (extension, expected) => {
    expect(getMimeType(extension)).toBe(expected);
    expect(getMimeType(extension.toUpperCase())).toBe(expected);
  });

  it('returns an empty MIME type for missing and unsupported extensions', () => {
    expect(getMimeType('')).toBe('');
    expect(getMimeType('exe')).toBe('');
    expect(getMimeType(' png ')).toBe('');
  });

  it.each([
    ['ppt/media/image.png', 'png'],
    ['ppt/media/photo.final.JPEG', 'JPEG'],
    ['ppt/media/image.png?version=2', 'png'],
    ['ppt/media/image.svg#preview', 'svg'],
    ['?query=photo.png', ''],
    ['#fragment.jpg', ''],
    ['ppt/media/archive', ''],
    ['ppt/media/trailing.', ''],
    ['.hidden', 'hidden'],
  ])('extracts an extension from %s', (filename, expected) => {
    expect(extractFileExtension(filename)).toBe(expected);
  });

  it.each([
    'https://example.com/video.mp4',
    'http://example.com/watch?id=1',
    'ftp://example.com/video.avi',
    'HTTPS://EXAMPLE.COM/VIDEO.MP4',
  ])('accepts an explicit external URL: %s', (value) => {
    expect(isVideoLink(value)).toBe(true);
  });

  it.each([
    '',
    'ppt/media/video.mp4',
    '/absolute/video.mp4',
    'javascript:alert(1)',
    'data:video/mp4;base64,AAAA',
    'file:///tmp/video.mp4',
    '://invalid',
  ])('rejects a non-external or unsafe URL: %s', (value) => {
    expect(isVideoLink(value)).toBe(false);
  });
});
