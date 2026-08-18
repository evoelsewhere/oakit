import type { PptxSceneDocument } from './scene-types';
import { validatePptxScene } from './scene-validation';
import { PptxWriteError } from './write-error';
import type { PptxWriteReport, PptxWriteResult } from './write-types';
import { serializePowerPointArchive } from './writer/archive';
import {
  type PptxSerializedPart,
  serializePowerPointParts,
} from './writer/parts';
import { verifyPowerPointCreation } from './writer/verify';

interface PptxCreatorDependencies {
  serializeArchive(parts: readonly PptxSerializedPart[]): Promise<Uint8Array>;
  verify(data: Uint8Array, scene: PptxSceneDocument): Promise<void>;
}

const DEFAULT_DEPENDENCIES: PptxCreatorDependencies = {
  serializeArchive: serializePowerPointArchive,
  verify: verifyPowerPointCreation,
};

type PptxCreationProfile = 'create-native-v1' | 'create-text-v1';

function creationProfile(scene: PptxSceneDocument): PptxCreationProfile {
  return scene.slides.some((slide) =>
    slide.elements.some(
      (element) =>
        element.type === 'image' ||
        element.type === 'shape' ||
        element.type === 'table',
    ),
  )
    ? 'create-native-v1'
    : 'create-text-v1';
}

function creationReport(
  addedPartCount: number,
  profile: PptxCreationProfile,
): PptxWriteReport {
  return {
    addedPartCount,
    copiedPartCount: 0,
    diagnostics: [],
    level: 'C2',
    operations: [],
    patchedPartCount: 0,
    producerEvidence: [],
    rebuiltPartCount: 0,
    removedPartCount: 0,
    supportProfile: {
      effectiveLevel: 'C2',
      id:
        profile === 'create-native-v1'
          ? 'pptx-create-native-v1'
          : 'pptx-create-text-v1',
      producerMatrix: [],
      version: '1',
    },
  };
}

export async function createPptxWithDependencies(
  scene: PptxSceneDocument,
  dependencies: PptxCreatorDependencies,
): Promise<PptxWriteResult> {
  const ownedScene = structuredClone(scene);
  const profile = creationProfile(ownedScene);
  const validation = validatePptxScene(ownedScene, { profile });
  if (!validation.valid) {
    throw new PptxWriteError(
      'invalid-scene',
      'PowerPoint scene is not valid for creation',
      { issues: validation.issues },
    );
  }

  let parts: PptxSerializedPart[];
  let data: Uint8Array;
  try {
    parts = serializePowerPointParts(ownedScene);
    data = await dependencies.serializeArchive(parts);
  } catch (cause) {
    throw new PptxWriteError(
      'package-build-failed',
      'Failed to build PowerPoint package',
      { cause },
    );
  }

  try {
    await dependencies.verify(data, ownedScene);
  } catch (cause) {
    throw new PptxWriteError(
      'verification-failed',
      'Generated PowerPoint package failed strict verification',
      { cause },
    );
  }

  return { data, report: creationReport(parts.length, profile) };
}

export function createPptx(scene: PptxSceneDocument): Promise<PptxWriteResult> {
  return createPptxWithDependencies(scene, DEFAULT_DEPENDENCIES);
}
