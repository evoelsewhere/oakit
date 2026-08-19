import type { PptxSceneGroupElement, PptxSceneTransform } from '../scene-types';
import type { Element, Group } from '../types';

export interface PptxGroupPreviewDependencies {
  mapChild(
    child: Element,
    childIndex: number,
    key: string,
  ): PptxSceneGroupElement['elements'][number];
  resolveTransform(element: Group): PptxSceneTransform | undefined;
}

export function createPptxRoundTripGroupPreview(
  element: Group,
  slideIndex: number,
  elementIndex: number,
  dependencies: PptxGroupPreviewDependencies,
  keyOverride?: string,
): PptxSceneGroupElement | undefined {
  const transform = dependencies.resolveTransform(element);
  const childSpace = element.childSpace;
  if (
    transform === undefined ||
    childSpace === undefined ||
    !Number.isFinite(childSpace.x) ||
    !Number.isFinite(childSpace.y) ||
    !Number.isFinite(childSpace.width) ||
    childSpace.width <= 0 ||
    !Number.isFinite(childSpace.height) ||
    childSpace.height <= 0
  ) {
    return undefined;
  }
  const key =
    keyOverride ?? `slide-${slideIndex + 1}-element-${elementIndex + 1}`;
  return {
    authored: {},
    elements: element.elements.map((child, childIndex) =>
      dependencies.mapChild(
        child,
        childIndex,
        `${key}-element-${childIndex + 1}`,
      ),
    ),
    key,
    resolved: {
      hidden: false,
      transform: { ...transform, childSpace: { ...childSpace } },
    },
    type: 'group',
  };
}
