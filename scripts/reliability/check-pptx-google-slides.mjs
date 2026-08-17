import assert from 'node:assert/strict';
import { createHash, randomUUID } from 'node:crypto';
import { mkdir, writeFile } from 'node:fs/promises';
import path from 'node:path';

import {
  createPptx,
  parsePptx,
  readPptxRoundTrip,
  replacePptxRoundTripText,
  setPptxRoundTripTextTransform,
  writePptxRoundTrip,
} from '../../dist/index.js';
import { roundTripGoogleSlidesPresentation } from './google-slides-drive.mjs';

const accessToken = process.env.GOOGLE_DRIVE_ACCESS_TOKEN;
if (accessToken === undefined || accessToken.trim().length === 0) {
  throw new Error(
    'Google Slides producer verification requires GOOGLE_DRIVE_ACCESS_TOKEN',
  );
}

const reportDirectory = path.resolve(
  'reports',
  'reliability',
  'pptx-google-slides',
);
const geometryTolerancePoints = 0.5;

function sha256(bytes) {
  return createHash('sha256').update(bytes).digest('hex');
}

function artifactEvidence(bytes) {
  return { byteLength: bytes.byteLength, sha256: sha256(bytes) };
}

function textScene() {
  return {
    layouts: [],
    masters: [],
    media: [],
    schemaVersion: 2,
    size: { height: 540, width: 960 },
    slides: [
      {
        elements: [
          {
            authored: {
              transform: { height: 80, width: 320, x: 24, y: 32 },
            },
            key: 'google-text',
            resolved: { hidden: false },
            text: {
              body: {},
              paragraphs: [
                {
                  children: [
                    {
                      key: 'google-run',
                      text: 'Before Google producer',
                      type: 'run',
                    },
                  ],
                  key: 'google-paragraph',
                },
              ],
            },
            type: 'text',
          },
        ],
        key: 'google-slide',
      },
    ],
    themes: [],
  };
}

function assertText(document, expected, rejected) {
  const json = JSON.stringify(document);
  assert.match(json, new RegExp(expected.replaceAll(' ', '&nbsp;')));
  assert.doesNotMatch(json, new RegExp(rejected.replaceAll(' ', '&nbsp;')));
}

function assertPoint(actual, expected, description) {
  assert.equal(typeof actual, 'number', `${description} must be numeric`);
  assert.ok(
    Math.abs(actual - expected) <= geometryTolerancePoints,
    `${description} differs by more than ${geometryTolerancePoints} points`,
  );
}

const created = await createPptx(textScene());
const snapshot = await readPptxRoundTrip(created.data);
const element = snapshot.document.slides[0]?.elements[0];
assert.equal(element?.type, 'text');
const run = element.text.paragraphs[0]?.children[0];
assert.equal(run?.type, 'run');
const textOperation = await replacePptxRoundTripText(snapshot, {
  targetKey: run.key,
  value: 'After Google producer',
});
const transformOperation = await setPptxRoundTripTextTransform(textOperation, {
  targetKey: element.key,
  value: {
    flipHorizontal: true,
    flipVertical: true,
    height: 100,
    rotation: 45,
    width: 400,
    x: 50,
    y: 60,
  },
});
const edited = await writePptxRoundTrip(transformOperation);
assert.equal(created.report.level, 'C2');
assert.equal(edited.report.level, 'R2');

const [createdExport, editedExport] = await Promise.all([
  roundTripGoogleSlidesPresentation(
    created.data,
    accessToken,
    `oakit-c3-${randomUUID()}`,
  ),
  roundTripGoogleSlidesPresentation(
    edited.data,
    accessToken,
    `oakit-r3-${randomUUID()}`,
  ),
]);
const [createdDocument, editedDocument] = await Promise.all([
  parsePptx(createdExport, {
    audioMode: 'none',
    errorMode: 'strict',
    imageMode: 'none',
    videoMode: 'none',
  }),
  parsePptx(editedExport, {
    audioMode: 'none',
    errorMode: 'strict',
    imageMode: 'none',
    videoMode: 'none',
  }),
]);
const editedElement = editedDocument.slides[0]?.elements.find(
  (candidate) =>
    candidate.type === 'text' &&
    JSON.stringify(candidate).includes('After&nbsp;Google&nbsp;producer'),
);
await mkdir(reportDirectory, { recursive: true });
await Promise.all([
  writeFile(path.join(reportDirectory, 'created-export.pptx'), createdExport),
  writeFile(path.join(reportDirectory, 'edited-export.pptx'), editedExport),
  writeFile(
    path.join(reportDirectory, 'diagnostic.json'),
    `${JSON.stringify(
      {
        artifacts: {
          creation: {
            output: artifactEvidence(createdExport),
            source: artifactEvidence(created.data),
          },
          edit: {
            output: artifactEvidence(editedExport),
            source: artifactEvidence(edited.data),
          },
        },
        editedTransform:
          editedElement?.type === 'text'
            ? {
                flipHorizontal: editedElement.isFlipH,
                flipVertical: editedElement.isFlipV,
                height: editedElement.height,
                rotation: editedElement.rotate,
                width: editedElement.width,
                x: editedElement.left,
                y: editedElement.top,
              }
            : null,
        schemaVersion: 1,
        temporaryPresentationsDeleted: true,
      },
      null,
      2,
    )}\n`,
  ),
]);
assert.equal(createdDocument.slides.length, 1);
assert.equal(editedDocument.slides.length, 1);
assertText(createdDocument, 'Before Google producer', 'After Google producer');
assertText(editedDocument, 'After Google producer', 'Before Google producer');
assert.equal(editedElement?.type, 'text');
assert.equal(editedElement.isFlipH, true);
assert.equal(editedElement.isFlipV, true);
assert.equal(editedElement.rotate, 45);
assertPoint(editedElement.left, 50, 'edited x');
assertPoint(editedElement.top, 60, 'edited y');
assertPoint(editedElement.width, 400, 'edited width');
assertPoint(editedElement.height, 100, 'edited height');
assert.notEqual(sha256(createdExport), sha256(created.data));
assert.notEqual(sha256(editedExport), sha256(edited.data));

const evidence = {
  artifacts: {
    creation: {
      output: artifactEvidence(createdExport),
      source: artifactEvidence(created.data),
    },
    edit: {
      output: artifactEvidence(editedExport),
      source: artifactEvidence(edited.data),
    },
  },
  capabilities: {
    creation: {
      internalLevel: created.report.level,
      producerVerifiedLevel: 'C3',
      profileId: created.report.supportProfile.id,
    },
    edit: {
      internalLevel: edited.report.level,
      operationCount: edited.report.operations.length,
      producerVerifiedLevel: 'R3',
      profileId: edited.report.supportProfile.id,
    },
  },
  platform: {
    architecture: process.arch,
    node: process.version,
    os: process.platform,
  },
  producer: {
    application: 'Google Slides',
    transport: 'Google Drive API v3 controlled import/export',
  },
  schemaVersion: 1,
  validation: {
    geometryTolerancePoints,
    openWithoutRepair: true,
    semanticTextPreserved: true,
    semanticTransformPreserved: true,
    strictParse: true,
    temporaryPresentationsDeleted: true,
  },
  verifiedAt: new Date().toISOString(),
};
await writeFile(
  path.join(reportDirectory, 'evidence.json'),
  `${JSON.stringify(evidence, null, 2)}\n`,
);
console.log(JSON.stringify(evidence, null, 2));
