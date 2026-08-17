import { describe, expect, it } from 'vitest';

import { createPptxProducerMatrixEvidence } from '../../scripts/reliability/pptx-producer-evidence.mjs';

const hash = (digit) => digit.repeat(64);
const artifact = (digit) => ({ byteLength: 10, sha256: hash(digit) });

function evidence() {
  return {
    googleSlides: {
      artifacts: {
        creation: { output: artifact('2'), source: artifact('1') },
        edit: { output: artifact('4'), source: artifact('3') },
      },
      capabilities: {
        creation: {
          internalLevel: 'C2',
          producerVerifiedLevel: 'C3',
          profileId: 'pptx-create-text-v1',
        },
        edit: {
          internalLevel: 'R2',
          operationCount: 2,
          producerVerifiedLevel: 'R3',
          profileId: 'pptx-roundtrip-text-v1',
        },
      },
      producer: {
        application: 'Google Slides',
        transport: 'Google Drive API v3 controlled import/export',
      },
      schemaVersion: 1,
      validation: {
        openWithoutRepair: true,
        semanticTextPreserved: true,
        semanticTransformPreserved: true,
        strictParse: true,
        temporaryPresentationsDeleted: true,
      },
    },
    libreOfficeCreation: {
      artifacts: {
        output: artifact('6'),
        pdf: artifact('7'),
        source: artifact('5'),
      },
      capability: {
        internalLevel: 'C2',
        producerVerifiedLevel: 'C3',
        profileId: 'pptx-create-text-v1',
      },
      producer: {
        application: 'LibreOffice Impress',
        version: 'LibreOffice 26',
      },
      schemaVersion: 1,
      validation: {
        openWithoutRepair: true,
        pdfExport: true,
        semanticTextPreserved: true,
        strictParse: true,
      },
    },
    libreOfficeEdit: {
      artifacts: {
        output: artifact('9'),
        pdf: artifact('a'),
        source: artifact('8'),
      },
      capability: {
        internalLevel: 'R2',
        operationCount: 2,
        producerVerifiedLevel: 'R3',
        profileId: 'pptx-roundtrip-text-v1',
      },
      producer: {
        application: 'LibreOffice Impress',
        version: 'LibreOffice 26',
      },
      schemaVersion: 1,
      validation: {
        openWithoutRepair: true,
        pdfExport: true,
        semanticTextPreserved: true,
        semanticTransformPreserved: true,
        strictParse: true,
      },
    },
    powerPoint: {
      creation: {
        internalLevel: 'C2',
        nativeReserialized: true,
        openWithoutRepair: true,
        outputSha256: hash('b'),
        saveReopen: true,
        semanticTextPreserved: true,
        sourceSha256: hash('c'),
      },
      edit: {
        internalLevel: 'R2',
        nativeReserialized: true,
        openWithoutRepair: true,
        operationCount: 2,
        outputSha256: hash('d'),
        saveReopen: true,
        semanticTextPreserved: true,
        semanticTransformPreserved: true,
        sourceSha256: hash('e'),
      },
      producer: {
        application: 'Microsoft PowerPoint',
        build: '16.112.1',
        version: '16.112',
      },
    },
  };
}

describe('PowerPoint producer matrix evidence', () => {
  it('classifies creation and editing only after all producer rows pass', () => {
    expect(createPptxProducerMatrixEvidence(evidence())).toEqual({
      producers: [
        {
          id: 'powerpoint-macos',
          version: 'Microsoft PowerPoint 16.112 (16.112.1)',
        },
        { id: 'libreoffice-impress', version: 'LibreOffice 26' },
        {
          id: 'google-slides',
          version: 'Google Drive API v3 controlled import/export',
        },
      ],
      profiles: {
        creation: {
          level: 'C3',
          producerMatrix: [
            'powerpoint-macos',
            'libreoffice-impress',
            'google-slides',
          ],
          profileId: 'pptx-create-text-v1',
        },
        edit: {
          level: 'R3',
          producerMatrix: [
            'powerpoint-macos',
            'libreoffice-impress',
            'google-slides',
          ],
          profileId: 'pptx-roundtrip-text-v1',
        },
      },
      schemaVersion: 1,
    });
  });

  it.each([
    [
      'Google cleanup',
      (value) =>
        (value.googleSlides.validation.temporaryPresentationsDeleted = false),
      'temporary cleanup',
    ],
    [
      'Google profile',
      (value) => (value.googleSlides.capabilities.edit.profileId = 'wrong'),
      'edit profile',
    ],
    [
      'LibreOffice version',
      (value) => (value.libreOfficeEdit.producer.version = 'other'),
      'producer versions',
    ],
    [
      'PowerPoint hash',
      (value) =>
        (value.powerPoint.creation.outputSha256 =
          value.powerPoint.creation.sourceSha256),
      'reserialized',
    ],
  ])('rejects invalid %s evidence', (_name, mutate, message) => {
    const value = evidence();
    mutate(value);
    expect(() => createPptxProducerMatrixEvidence(value)).toThrow(message);
  });
});
