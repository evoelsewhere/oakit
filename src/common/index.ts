export { base64ArrayBuffer, encodeBase64 } from './binary/base64';
export {
  extractFileExtension,
  getMimeType,
  isVideoLink,
} from './media/media-type';
export { angleToDegrees, numberToFixed, toHex } from './numbers';
export {
  getRelationshipPartUri,
  resolvePartUri,
  resolveRelationshipTarget,
} from './opc/part-uri';
export {
  decodeXmlEntities,
  escapeHtml,
  hasValidText,
  sanitizeHyperlink,
} from './text/html';
export { eachElement, getTextByPathList } from './xml/tree';
export type { XmlLookupValue } from './xml/tree';
