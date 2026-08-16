import { RATIO_EMUs_Points } from '../../../common/ooxml/units';
import type { PptxSceneDocument } from '../scene-types';
import { parse } from '../parser';
import type { PptxDocument, PptxParseOptions } from '../types';
import { pointsToEmu } from './units';

type PptxCreationParser = (
  data: Uint8Array,
  options: PptxParseOptions,
) => Promise<PptxDocument>;

function expectedPointValue(value: number): number {
  return pointsToEmu(value) * RATIO_EMUs_Points;
}

export async function verifyPowerPointCreationWithParser(
  data: Uint8Array,
  scene: PptxSceneDocument,
  parseDocument: PptxCreationParser,
): Promise<void> {
  const document = await parseDocument(data, {
    audioMode: 'none',
    errorMode: 'strict',
    imageMode: 'none',
    limits: {
      maxEntries: scene.slides.length * 2 + 9,
      maxSlides: Math.max(1, scene.slides.length),
    },
    videoMode: 'none',
  });
  if (document.slides.length !== scene.slides.length) {
    throw new Error(
      `Generated PowerPoint slide count mismatch: expected ${scene.slides.length}, received ${document.slides.length}`,
    );
  }
  const expectedWidth = expectedPointValue(scene.size.width);
  const expectedHeight = expectedPointValue(scene.size.height);
  if (
    document.size.width !== expectedWidth ||
    document.size.height !== expectedHeight
  ) {
    throw new Error(
      `Generated PowerPoint size mismatch: expected ${expectedWidth}x${expectedHeight}, received ${document.size.width}x${document.size.height}`,
    );
  }
  scene.slides.forEach((slide, index) => {
    const generated = document.slides[index];
    if (!generated || generated.elements.length !== slide.elements.length) {
      throw new Error(
        `Generated PowerPoint element count mismatch on slide ${index + 1}: expected ${slide.elements.length}, received ${generated?.elements.length ?? 0}`,
      );
    }
  });
}

export function verifyPowerPointCreation(
  data: Uint8Array,
  scene: PptxSceneDocument,
): Promise<void> {
  return verifyPowerPointCreationWithParser(data, scene, parse);
}
