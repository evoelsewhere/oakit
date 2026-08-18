import type { PptxSceneTextElement, PptxSceneTransform } from '../scene-types';
import {
  serializeNativeShapeProperties,
  serializeShapeNonVisualProperties,
} from './shape';
export { serializeShapeTransform } from './shape';
import { serializeTextBody } from './text-body';
import type { PptxTextSerializationContext } from './text-node';

export function serializeTextShape(
  element: PptxSceneTextElement,
  transform: PptxSceneTransform,
  shapeId: number,
  context: PptxTextSerializationContext,
): string {
  const nonVisual = serializeShapeNonVisualProperties(element, shapeId, true);
  const shapeProperties = serializeNativeShapeProperties(element, transform);
  return `<p:sp>${nonVisual}${shapeProperties}${serializeTextBody(element.text, context)}</p:sp>`;
}
