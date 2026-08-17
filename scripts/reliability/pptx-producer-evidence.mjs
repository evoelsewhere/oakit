import { createHash } from 'node:crypto';
import { mkdir, readFile, writeFile } from 'node:fs/promises';
import path from 'node:path';
import process from 'node:process';
import { fileURLToPath } from 'node:url';

const CREATE_PROFILE = 'pptx-create-text-v1';
const EDIT_PROFILE = 'pptx-roundtrip-text-v1';
const SHA256_PATTERN = /^[0-9a-f]{64}$/;

function record(value, description) {
  if (value === null || typeof value !== 'object' || Array.isArray(value)) {
    throw new TypeError(`${description} must be an object`);
  }
  return value;
}

function string(value, description) {
  if (typeof value !== 'string' || value.length === 0) {
    throw new TypeError(`${description} must be a non-empty string`);
  }
  return value;
}

function exact(value, expected, description) {
  if (value !== expected) {
    throw new Error(`${description} must equal ${expected}`);
  }
  return value;
}

function truth(value, description) {
  if (value !== true) throw new Error(`${description} must be true`);
  return true;
}

function artifact(value, description) {
  const candidate = record(value, description);
  if (
    !Number.isSafeInteger(candidate.byteLength) ||
    candidate.byteLength <= 0 ||
    typeof candidate.sha256 !== 'string' ||
    !SHA256_PATTERN.test(candidate.sha256)
  ) {
    throw new Error(`${description} must contain bounded SHA-256 evidence`);
  }
  return candidate;
}

function hashPair(sourceValue, outputValue, description) {
  const source = artifact(sourceValue, `${description} source`);
  const output = artifact(outputValue, `${description} output`);
  if (source.sha256 === output.sha256) {
    throw new Error(`${description} must be reserialized by its producer`);
  }
}

function legacyHashPair(value, description) {
  const candidate = record(value, description);
  const source = string(candidate.sourceSha256, `${description} source hash`);
  const output = string(candidate.outputSha256, `${description} output hash`);
  if (!SHA256_PATTERN.test(source) || !SHA256_PATTERN.test(output)) {
    throw new Error(`${description} hashes must be SHA-256`);
  }
  if (source === output) {
    throw new Error(`${description} must be reserialized by its producer`);
  }
}

function validatePowerPoint(value) {
  const evidence = record(value, 'PowerPoint evidence');
  const producer = record(evidence.producer, 'PowerPoint producer');
  exact(producer.application, 'Microsoft PowerPoint', 'PowerPoint application');
  const creation = record(evidence.creation, 'PowerPoint creation evidence');
  exact(creation.internalLevel, 'C2', 'PowerPoint creation internal level');
  truth(creation.nativeReserialized, 'PowerPoint creation reserialization');
  truth(creation.openWithoutRepair, 'PowerPoint creation repair state');
  truth(creation.saveReopen, 'PowerPoint creation save/reopen');
  truth(creation.semanticTextPreserved, 'PowerPoint creation semantics');
  legacyHashPair(creation, 'PowerPoint creation');
  const edit = record(evidence.edit, 'PowerPoint edit evidence');
  exact(edit.internalLevel, 'R2', 'PowerPoint edit internal level');
  exact(edit.operationCount, 2, 'PowerPoint edit operation count');
  truth(edit.nativeReserialized, 'PowerPoint edit reserialization');
  truth(edit.openWithoutRepair, 'PowerPoint edit repair state');
  truth(edit.saveReopen, 'PowerPoint edit save/reopen');
  truth(edit.semanticTextPreserved, 'PowerPoint edit text semantics');
  truth(edit.semanticTransformPreserved, 'PowerPoint edit transform semantics');
  legacyHashPair(edit, 'PowerPoint edit');
  return {
    id: 'powerpoint-macos',
    producer: `${string(producer.application, 'PowerPoint application')} ${string(producer.version, 'PowerPoint version')} (${string(producer.build, 'PowerPoint build')})`,
  };
}

function validateLibreOffice(value, kind) {
  const evidence = record(value, `LibreOffice ${kind} evidence`);
  exact(evidence.schemaVersion, 1, `LibreOffice ${kind} schema version`);
  const capability = record(
    evidence.capability,
    `LibreOffice ${kind} capability`,
  );
  const creation = kind === 'creation';
  exact(
    capability.internalLevel,
    creation ? 'C2' : 'R2',
    `LibreOffice ${kind} internal level`,
  );
  exact(
    capability.producerVerifiedLevel,
    creation ? 'C3' : 'R3',
    `LibreOffice ${kind} producer level`,
  );
  exact(
    capability.profileId,
    creation ? CREATE_PROFILE : EDIT_PROFILE,
    `LibreOffice ${kind} profile`,
  );
  if (!creation) {
    exact(capability.operationCount, 2, 'LibreOffice edit operation count');
  }
  const artifacts = record(evidence.artifacts, `LibreOffice ${kind} artifacts`);
  hashPair(artifacts.source, artifacts.output, `LibreOffice ${kind}`);
  artifact(artifacts.pdf, `LibreOffice ${kind} PDF`);
  const validation = record(
    evidence.validation,
    `LibreOffice ${kind} validation`,
  );
  truth(validation.openWithoutRepair, `LibreOffice ${kind} repair state`);
  truth(validation.pdfExport, `LibreOffice ${kind} PDF export`);
  truth(validation.strictParse, `LibreOffice ${kind} strict parse`);
  truth(validation.semanticTextPreserved, `LibreOffice ${kind} text semantics`);
  if (!creation) {
    truth(
      validation.semanticTransformPreserved,
      'LibreOffice edit transform semantics',
    );
  }
  const producer = record(evidence.producer, `LibreOffice ${kind} producer`);
  exact(
    producer.application,
    'LibreOffice Impress',
    `LibreOffice ${kind} application`,
  );
  return string(producer.version, `LibreOffice ${kind} version`);
}

function validateGoogleSlides(value) {
  const evidence = record(value, 'Google Slides evidence');
  exact(evidence.schemaVersion, 1, 'Google Slides schema version');
  const capabilities = record(
    evidence.capabilities,
    'Google Slides capabilities',
  );
  const creation = record(capabilities.creation, 'Google Slides creation');
  exact(creation.internalLevel, 'C2', 'Google Slides creation internal level');
  exact(
    creation.producerVerifiedLevel,
    'C3',
    'Google Slides creation producer level',
  );
  exact(creation.profileId, CREATE_PROFILE, 'Google Slides creation profile');
  const edit = record(capabilities.edit, 'Google Slides edit');
  exact(edit.internalLevel, 'R2', 'Google Slides edit internal level');
  exact(edit.producerVerifiedLevel, 'R3', 'Google Slides edit producer level');
  exact(edit.profileId, EDIT_PROFILE, 'Google Slides edit profile');
  exact(edit.operationCount, 2, 'Google Slides edit operation count');
  const artifacts = record(evidence.artifacts, 'Google Slides artifacts');
  const creationArtifacts = record(
    artifacts.creation,
    'Google Slides creation artifacts',
  );
  hashPair(
    creationArtifacts.source,
    creationArtifacts.output,
    'Google Slides creation',
  );
  const editArtifacts = record(artifacts.edit, 'Google Slides edit artifacts');
  hashPair(editArtifacts.source, editArtifacts.output, 'Google Slides edit');
  const validation = record(evidence.validation, 'Google Slides validation');
  truth(validation.openWithoutRepair, 'Google Slides repair state');
  truth(validation.strictParse, 'Google Slides strict parse');
  truth(validation.semanticTextPreserved, 'Google Slides text semantics');
  truth(
    validation.semanticTransformPreserved,
    'Google Slides transform semantics',
  );
  truth(
    validation.temporaryPresentationsDeleted,
    'Google Slides temporary cleanup',
  );
  const producer = record(evidence.producer, 'Google Slides producer');
  exact(producer.application, 'Google Slides', 'Google Slides application');
  return string(producer.transport, 'Google Slides transport');
}

export function createPptxProducerMatrixEvidence(input) {
  const powerPoint = validatePowerPoint(input.powerPoint);
  const libreOfficeCreationVersion = validateLibreOffice(
    input.libreOfficeCreation,
    'creation',
  );
  const libreOfficeEditVersion = validateLibreOffice(
    input.libreOfficeEdit,
    'edit',
  );
  exact(
    libreOfficeEditVersion,
    libreOfficeCreationVersion,
    'LibreOffice producer versions',
  );
  const googleTransport = validateGoogleSlides(input.googleSlides);
  return {
    profiles: {
      creation: {
        level: 'C3',
        producerMatrix: [powerPoint.id, 'libreoffice-impress', 'google-slides'],
        profileId: CREATE_PROFILE,
      },
      edit: {
        level: 'R3',
        producerMatrix: [powerPoint.id, 'libreoffice-impress', 'google-slides'],
        profileId: EDIT_PROFILE,
      },
    },
    producers: [
      { id: powerPoint.id, version: powerPoint.producer },
      { id: 'libreoffice-impress', version: libreOfficeCreationVersion },
      { id: 'google-slides', version: googleTransport },
    ],
    schemaVersion: 1,
  };
}

function sha256(text) {
  return createHash('sha256').update(text).digest('hex');
}

async function readEvidence(filename) {
  let text;
  try {
    text = await readFile(filename, 'utf8');
  } catch (cause) {
    throw new Error(`PPTX producer evidence is missing: ${filename}`, {
      cause,
    });
  }
  return { sha256: sha256(text), value: JSON.parse(text) };
}

export async function mergePptxProducerEvidence(reportRoot) {
  const paths = {
    googleSlides: path.join(reportRoot, 'pptx-google-slides', 'evidence.json'),
    libreOfficeCreation: path.join(
      reportRoot,
      'pptx-libreoffice',
      'creation.json',
    ),
    libreOfficeEdit: path.join(reportRoot, 'pptx-libreoffice', 'edit.json'),
    powerPoint: path.join(reportRoot, 'pptx-powerpoint-macos', 'evidence.json'),
  };
  const entries = await Promise.all(
    Object.entries(paths).map(async ([key, filename]) => [
      key,
      await readEvidence(filename),
    ]),
  );
  const evidence = Object.fromEntries(entries);
  const matrix = createPptxProducerMatrixEvidence(
    Object.fromEntries(
      Object.entries(evidence).map(([key, entry]) => [key, entry.value]),
    ),
  );
  return {
    ...matrix,
    evidenceSha256: Object.fromEntries(
      Object.entries(evidence).map(([key, entry]) => [key, entry.sha256]),
    ),
  };
}

if (process.argv[1] === fileURLToPath(import.meta.url)) {
  const reportRoot = path.resolve(process.argv[2] ?? 'reports/reliability');
  const output = path.join(reportRoot, 'pptx-producer-matrix', 'evidence.json');
  const matrix = await mergePptxProducerEvidence(reportRoot);
  await mkdir(path.dirname(output), { recursive: true });
  await writeFile(output, `${JSON.stringify(matrix, null, 2)}\n`);
  console.log(JSON.stringify(matrix, null, 2));
}
