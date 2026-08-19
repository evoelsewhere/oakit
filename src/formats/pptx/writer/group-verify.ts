import type { PptxSceneElement, PptxSceneGroupElement } from '../scene-types';
import type { Element, Group } from '../types';

export interface PptxGroupVerificationDependencies {
  expectedPointValue(value: number): number;
  verifyChild(
    generated: Element | undefined,
    expected: PptxSceneElement,
    childIndex: number,
  ): void;
  verifyTransform(
    generated: Group,
    expected: PptxSceneGroupElement,
    location: string,
  ): void;
}

function scaledExpectedChild(
  child: PptxSceneElement,
  groupTransform: NonNullable<PptxSceneGroupElement['authored']['transform']>,
): PptxSceneElement {
  const result = structuredClone(child);
  const transform = result.authored.transform;
  if (transform === undefined) return result;
  const widthScale = groupTransform.width / groupTransform.childSpace.width;
  const heightScale = groupTransform.height / groupTransform.childSpace.height;
  const centerX = transform.x + transform.width / 2;
  const centerY = transform.y + transform.height / 2;
  const rotation = (((transform.rotation ?? 0) % 360) + 360) % 360;
  const swapped = rotation === 90 || rotation === 270;
  const width = transform.width * (swapped ? heightScale : widthScale);
  const height = transform.height * (swapped ? widthScale : heightScale);
  result.authored.transform = {
    ...transform,
    height,
    width,
    x: (centerX - groupTransform.childSpace.x) * widthScale - width / 2,
    y: (centerY - groupTransform.childSpace.y) * heightScale - height / 2,
  };
  return result;
}

export function verifyPowerPointGroupElement(
  generated: Element | undefined,
  expected: PptxSceneGroupElement,
  location: string,
  dependencies: PptxGroupVerificationDependencies,
): void {
  if (generated?.type !== 'group') {
    throw new Error(`Generated PowerPoint group missing at ${location}`);
  }
  dependencies.verifyTransform(generated, expected, location);
  const groupTransform = expected.authored.transform;
  if (groupTransform === undefined || groupTransform.childSpace === undefined) {
    throw new Error(
      `Expected PowerPoint group child space missing at ${location}`,
    );
  }
  const expectedChildSpace = groupTransform.childSpace;
  const childSpace = generated.childSpace;
  const actual =
    childSpace === undefined
      ? undefined
      : {
          height: childSpace.height,
          width: childSpace.width,
          x: childSpace.x,
          y: childSpace.y,
        };
  const wanted = {
    height: dependencies.expectedPointValue(expectedChildSpace.height),
    width: dependencies.expectedPointValue(expectedChildSpace.width),
    x: dependencies.expectedPointValue(expectedChildSpace.x),
    y: dependencies.expectedPointValue(expectedChildSpace.y),
  };
  if (JSON.stringify(actual) !== JSON.stringify(wanted)) {
    throw new Error(
      `Generated PowerPoint group child space mismatch at ${location}`,
    );
  }
  if (generated.elements.length !== expected.elements.length) {
    throw new Error(
      `Generated PowerPoint group child count mismatch at ${location}`,
    );
  }
  expected.elements.forEach((child, index) =>
    dependencies.verifyChild(
      generated.elements[index],
      scaledExpectedChild(child, groupTransform),
      index,
    ),
  );
}
