import assert from 'node:assert/strict';
import { createHash } from 'node:crypto';
import { execFile } from 'node:child_process';
import { mkdir, mkdtemp, readFile, rm, writeFile } from 'node:fs/promises';
import path from 'node:path';
import { promisify } from 'node:util';

import {
  createPptx,
  parsePptx,
  readPptxRoundTrip,
  replacePptxRoundTripText,
  setPptxRoundTripTextTransform,
  writePptxRoundTrip,
} from '../../dist/index.js';

const execFileAsync = promisify(execFile);
const projectRoot = path.resolve(import.meta.dirname, '../..');
const reportDirectory = path.join(
  projectRoot,
  'reports',
  'reliability',
  'pptx-powerpoint-macos',
);
const powerPointApp =
  process.env.POWERPOINT_APP_PATH ?? '/Applications/Microsoft PowerPoint.app';
const powerPointName =
  process.env.POWERPOINT_APP_NAME ?? 'Microsoft PowerPoint';

if (process.platform !== 'darwin') {
  throw new Error('PowerPoint native producer verification requires macOS');
}

function sha256(bytes) {
  return createHash('sha256').update(bytes).digest('hex');
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
            key: 'native-text',
            resolved: { hidden: false },
            text: {
              body: {},
              paragraphs: [
                {
                  children: [
                    {
                      key: 'native-run',
                      text: 'Before native producer',
                      type: 'run',
                    },
                  ],
                  key: 'native-paragraph',
                },
              ],
            },
            type: 'text',
          },
        ],
        key: 'native-slide',
      },
    ],
    themes: [],
  };
}

async function powerPointVersion() {
  const plist = path.join(powerPointApp, 'Contents', 'Info.plist');
  const [shortVersion, buildVersion] = await Promise.all([
    execFileAsync('/usr/libexec/PlistBuddy', [
      '-c',
      'Print :CFBundleShortVersionString',
      plist,
    ]),
    execFileAsync('/usr/libexec/PlistBuddy', [
      '-c',
      'Print :CFBundleVersion',
      plist,
    ]),
  ]);
  return {
    build: buildVersion.stdout.trim(),
    version: shortVersion.stdout.trim(),
  };
}

async function runPowerPointScript(lines, values) {
  const scriptArguments = lines.flatMap((line) => ['-e', line]);
  scriptArguments.push(...values);
  await execFileAsync('/usr/bin/osascript', scriptArguments, {
    maxBuffer: 1024 * 1024,
    timeout: 45_000,
  });
}

async function openPresentation(filename) {
  await execFileAsync('/usr/bin/open', ['-a', powerPointName, filename]);
}

async function waitForPresentation(name, actionLines, values = []) {
  await runPowerPointScript(
    [
      'on run argv',
      'set deckName to item 1 of argv',
      `tell application "${powerPointName}"`,
      'repeat with attempt from 1 to 120',
      'if exists presentation deckName then exit repeat',
      'delay 0.25',
      'end repeat',
      'if not (exists presentation deckName) then error "PowerPoint did not open the expected presentation"',
      'set targetDeck to presentation deckName',
      ...actionLines,
      'end tell',
      'end run',
    ],
    [name, ...values],
  );
}

async function saveAndReopen(input) {
  await openPresentation(input);
  await waitForPresentation(path.basename(input), [
    'set probeSlide to make new slide at end of targetDeck with properties {layout:slide layout blank}',
    'delete probeSlide',
    'save targetDeck',
    'close targetDeck saving no',
  ]);
  await openPresentation(input);
  await waitForPresentation(path.basename(input), [
    'close targetDeck saving no',
  ]);
}

async function runNativeCycles(paths) {
  await saveAndReopen(paths.created);
  await saveAndReopen(paths.edited);
}

function assertText(document, expected, rejected) {
  const json = JSON.stringify(document);
  assert.match(json, new RegExp(expected.replaceAll(' ', '&nbsp;')));
  assert.doesNotMatch(json, new RegExp(rejected.replaceAll(' ', '&nbsp;')));
}

const directory = await mkdtemp('/tmp/oakit-powerpoint-macos-');
try {
  const paths = {
    created: path.join(directory, 'oakit-c3-created.pptx'),
    edited: path.join(directory, 'oakit-r3-edited.pptx'),
  };
  const created = await createPptx(textScene());
  const snapshot = await readPptxRoundTrip(created.data);
  const element = snapshot.document.slides[0]?.elements[0];
  assert.equal(element?.type, 'text');
  const run = element.text.paragraphs[0]?.children[0];
  assert.equal(run?.type, 'run');
  const textOperation = await replacePptxRoundTripText(snapshot, {
    targetKey: run.key,
    value: 'After native producer',
  });
  const operation = await setPptxRoundTripTextTransform(textOperation, {
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
  const edited = await writePptxRoundTrip(operation);
  assert.equal(edited.report.level, 'R2');
  await Promise.all([
    writeFile(paths.created, created.data),
    writeFile(paths.edited, edited.data),
  ]);

  await runNativeCycles(paths);

  const [createdResaved, editedResaved] = await Promise.all([
    readFile(paths.created),
    readFile(paths.edited),
  ]);
  const [createdDocument, editedDocument, producer] = await Promise.all([
    parsePptx(new Uint8Array(createdResaved), {
      audioMode: 'none',
      errorMode: 'strict',
      imageMode: 'none',
      videoMode: 'none',
    }),
    parsePptx(new Uint8Array(editedResaved), {
      audioMode: 'none',
      errorMode: 'strict',
      imageMode: 'none',
      videoMode: 'none',
    }),
    powerPointVersion(),
  ]);
  assertText(
    createdDocument,
    'Before native producer',
    'After native producer',
  );
  assertText(editedDocument, 'After native producer', 'Before native producer');
  assert.equal(createdDocument.slides.length, 1);
  assert.equal(editedDocument.slides.length, 1);
  const editedElement = editedDocument.slides[0]?.elements[0];
  assert.equal(editedElement?.type, 'text');
  assert.deepEqual(
    {
      flipHorizontal: editedElement.isFlipH,
      flipVertical: editedElement.isFlipV,
      height: editedElement.height,
      rotation: editedElement.rotate,
      width: editedElement.width,
      x: editedElement.left,
      y: editedElement.top,
    },
    {
      flipHorizontal: true,
      flipVertical: true,
      height: 100,
      rotation: 45,
      width: 400,
      x: 50,
      y: 60,
    },
  );
  assert.notEqual(sha256(createdResaved), sha256(created.data));
  assert.notEqual(sha256(editedResaved), sha256(edited.data));

  const evidence = {
    creation: {
      internalLevel: created.report.level,
      nativeReserialized: true,
      openWithoutRepair: true,
      outputSha256: sha256(createdResaved),
      saveReopen: true,
      semanticTextPreserved: true,
      sourceSha256: sha256(created.data),
    },
    edit: {
      internalLevel: edited.report.level,
      nativeReserialized: true,
      openWithoutRepair: true,
      operationCount: edited.report.operations.length,
      outputSha256: sha256(editedResaved),
      saveReopen: true,
      semanticTextPreserved: true,
      semanticTransformPreserved: true,
      sourceSha256: sha256(edited.data),
    },
    platform: {
      architecture: process.arch,
      node: process.version,
      os: process.platform,
    },
    producer: {
      application: powerPointName,
      ...producer,
    },
  };
  await mkdir(reportDirectory, { recursive: true });
  await writeFile(
    path.join(reportDirectory, 'evidence.json'),
    `${JSON.stringify(evidence, null, 2)}\n`,
  );
  console.log(JSON.stringify(evidence, null, 2));
} finally {
  await rm(directory, { force: true, recursive: true });
}
