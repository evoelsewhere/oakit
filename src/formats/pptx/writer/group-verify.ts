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
  const expectedChildSpace = expected.authored.transform?.childSpace;
  if (expectedChildSpace === undefined) {
    throw new Error(
      `Expected PowerPoint group child space missing at ${location}`,
    );
  }
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
    dependencies.verifyChild(generated.elements[index], child, index),
  );
}
