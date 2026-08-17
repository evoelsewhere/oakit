import assert from 'node:assert/strict';
import { Buffer } from 'node:buffer';
import { createHash } from 'node:crypto';
import { execFile } from 'node:child_process';
import { mkdir, readFile, rm, writeFile } from 'node:fs/promises';
import path from 'node:path';
import { promisify, TextDecoder } from 'node:util';
import { inflateSync } from 'node:zlib';

import JSZip from 'jszip';

const execFileAsync = promisify(execFile);
const projectRoot = path.resolve(import.meta.dirname, '../..');
const outputDirectory = path.join(
  projectRoot,
  'reports',
  'reliability',
  'pptx-render',
);
const fixturePath = path.join(outputDirectory, 'agent-ready.pptx');
const workerPath = path.join(
  projectRoot,
  'scripts',
  'reliability',
  'render-pptx-worker.mjs',
);

const PRESENTATION_NS =
  'http://schemas.openxmlformats.org/presentationml/2006/main';
const DRAWING_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const OFFICE_REL_NS =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const PACKAGE_REL_NS =
  'http://schemas.openxmlformats.org/package/2006/relationships';

function sha256(bytes) {
  return createHash('sha256').update(bytes).digest('hex');
}

async function partPayloads(bytes) {
  const archive = await JSZip.loadAsync(bytes, { checkCRC32: true });
  const result = new Map();
  for (const part of Object.values(archive.files)) {
    if (!part.dir) result.set(part.name, await part.async('uint8array'));
  }
  return result;
}

async function createFixture() {
  const zip = new JSZip();
  const parts = {
    '[Content_Types].xml': `<?xml version="1.0" encoding="UTF-8"?>
      <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
        <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
        <Default Extension="xml" ContentType="application/xml"/>
        <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
        <Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
        <Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>
        <Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>
        <Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
      </Types>`,
    '_rels/.rels': `<?xml version="1.0" encoding="UTF-8"?>
      <Relationships xmlns="${PACKAGE_REL_NS}"><Relationship Id="rId1" Type="${OFFICE_REL_NS}/officeDocument" Target="ppt/presentation.xml"/></Relationships>`,
    'customXml/agent-evidence.xml':
      '<?xml version="1.0"?><agent xmlns="urn:oakit:e2e">opaque source bytes</agent>',
    'ppt/presentation.xml': `<?xml version="1.0" encoding="UTF-8"?>
      <p:presentation xmlns:p="${PRESENTATION_NS}" xmlns:r="${OFFICE_REL_NS}">
        <p:sldIdLst><p:sldId id="256" r:id="rIdSlide1"/></p:sldIdLst>
        <p:sldSz cx="9144000" cy="5143500"/>
      </p:presentation>`,
    'ppt/_rels/presentation.xml.rels': `<?xml version="1.0" encoding="UTF-8"?>
      <Relationships xmlns="${PACKAGE_REL_NS}">
        <Relationship Id="rIdTheme" Type="${OFFICE_REL_NS}/theme" Target="theme/theme1.xml"/>
        <Relationship Id="rIdSlide1" Type="${OFFICE_REL_NS}/slide" Target="slides/slide1.xml"/>
      </Relationships>`,
    'ppt/theme/theme1.xml': `<?xml version="1.0" encoding="UTF-8"?>
      <a:theme xmlns:a="${DRAWING_NS}" name="Agent Ready">
        <a:themeElements><a:clrScheme name="Agent Ready"><a:dk1><a:srgbClr val="000000"/></a:dk1><a:lt1><a:srgbClr val="FFFFFF"/></a:lt1><a:accent1><a:srgbClr val="4472C4"/></a:accent1></a:clrScheme><a:fontScheme name="Agent Ready"><a:majorFont><a:latin typeface="Arial"/></a:majorFont><a:minorFont><a:latin typeface="Arial"/></a:minorFont></a:fontScheme><a:fmtScheme name="Agent Ready"/></a:themeElements>
      </a:theme>`,
    'ppt/slides/slide1.xml': `<?xml version="1.0" encoding="UTF-8"?>
      <p:sld xmlns:p="${PRESENTATION_NS}" xmlns:a="${DRAWING_NS}">
        <p:cSld>
          <p:bg><p:bgPr><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill><a:effectLst/></p:bgPr></p:bg>
          <p:spTree>
            <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>
            <p:sp><p:nvSpPr><p:cNvPr id="2" name="Agent text"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="3657600" cy="914400"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln><a:noFill/></a:ln></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US"/><a:t>Agent Ready preview</a:t></a:r></a:p></p:txBody></p:sp>
          </p:spTree>
        </p:cSld>
      </p:sld>`,
    'ppt/slides/_rels/slide1.xml.rels': `<?xml version="1.0" encoding="UTF-8"?>
      <Relationships xmlns="${PACKAGE_REL_NS}"><Relationship Id="rIdLayout" Type="${OFFICE_REL_NS}/slideLayout" Target="../slideLayouts/slideLayout1.xml"/></Relationships>`,
    'ppt/slideLayouts/slideLayout1.xml': `<?xml version="1.0" encoding="UTF-8"?><p:sldLayout xmlns:p="${PRESENTATION_NS}" xmlns:a="${DRAWING_NS}"><p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/></p:spTree></p:cSld></p:sldLayout>`,
    'ppt/slideLayouts/_rels/slideLayout1.xml.rels': `<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="${PACKAGE_REL_NS}"><Relationship Id="rIdMaster" Type="${OFFICE_REL_NS}/slideMaster" Target="../slideMasters/slideMaster1.xml"/></Relationships>`,
    'ppt/slideMasters/slideMaster1.xml': `<?xml version="1.0" encoding="UTF-8"?><p:sldMaster xmlns:p="${PRESENTATION_NS}" xmlns:a="${DRAWING_NS}"><p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/></p:spTree></p:cSld><p:clrMap accent1="accent1" bg1="lt1" tx1="dk1"/></p:sldMaster>`,
    'ppt/slideMasters/_rels/slideMaster1.xml.rels': `<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="${PACKAGE_REL_NS}"><Relationship Id="rIdTheme" Type="${OFFICE_REL_NS}/theme" Target="../theme/theme1.xml"/></Relationships>`,
  };
  for (const [name, content] of Object.entries(parts)) zip.file(name, content);
  return zip.generateAsync({ compression: 'STORE', type: 'uint8array' });
}

function paeth(left, above, upperLeft) {
  const estimate = left + above - upperLeft;
  const leftDistance = Math.abs(estimate - left);
  const aboveDistance = Math.abs(estimate - above);
  const upperLeftDistance = Math.abs(estimate - upperLeft);
  if (leftDistance <= aboveDistance && leftDistance <= upperLeftDistance) {
    return left;
  }
  return aboveDistance <= upperLeftDistance ? above : upperLeft;
}

function decodePng(data) {
  assert.deepEqual(
    Array.from(data.subarray(0, 8)),
    [137, 80, 78, 71, 13, 10, 26, 10],
  );
  const view = new DataView(data.buffer, data.byteOffset, data.byteLength);
  const idat = [];
  let width;
  let height;
  let colorType;
  let offset = 8;
  while (offset < data.byteLength) {
    const length = view.getUint32(offset);
    const type = new TextDecoder().decode(
      data.subarray(offset + 4, offset + 8),
    );
    const payload = data.subarray(offset + 8, offset + 8 + length);
    if (type === 'IHDR') {
      width = view.getUint32(offset + 8);
      height = view.getUint32(offset + 12);
      assert.equal(payload[8], 8);
      colorType = payload[9];
      assert.equal(payload[12], 0);
    }
    if (type === 'IDAT') idat.push(payload);
    offset += length + 12;
    if (type === 'IEND') break;
  }
  assert.equal(typeof width, 'number');
  assert.equal(typeof height, 'number');
  assert.ok(colorType === 2 || colorType === 6);
  const bytesPerPixel = colorType === 6 ? 4 : 3;
  const compressed = Buffer.concat(idat.map((chunk) => Buffer.from(chunk)));
  const inflated = inflateSync(compressed);
  const stride = width * bytesPerPixel;
  const pixels = new Uint8Array(height * stride);
  let sourceOffset = 0;
  for (let y = 0; y < height; y += 1) {
    const filter = inflated[sourceOffset];
    sourceOffset += 1;
    for (let x = 0; x < stride; x += 1) {
      const raw = inflated[sourceOffset + x];
      const target = y * stride + x;
      const left = x >= bytesPerPixel ? pixels[target - bytesPerPixel] : 0;
      const above = y > 0 ? pixels[target - stride] : 0;
      const upperLeft =
        y > 0 && x >= bytesPerPixel
          ? pixels[target - stride - bytesPerPixel]
          : 0;
      const predictor =
        filter === 0
          ? 0
          : filter === 1
            ? left
            : filter === 2
              ? above
              : filter === 3
                ? Math.floor((left + above) / 2)
                : paeth(left, above, upperLeft);
      pixels[target] = (raw + predictor) & 0xff;
    }
    sourceOffset += stride;
  }
  return {
    height,
    pixel(x, y) {
      const start = y * stride + x * bytesPerPixel;
      return [
        pixels[start],
        pixels[start + 1],
        pixels[start + 2],
        bytesPerPixel === 4 ? pixels[start + 3] : 255,
      ];
    },
    width,
  };
}

await rm(outputDirectory, { force: true, recursive: true });
await mkdir(outputDirectory, { recursive: true });
const source = await createFixture();
await writeFile(fixturePath, source);

const child = await execFileAsync(
  process.execPath,
  [workerPath, fixturePath, outputDirectory],
  {
    cwd: projectRoot,
    env: { ...process.env, OAKIT_RENDER_E2E_NO_OFFICE: '1', PATH: '' },
    maxBuffer: 1024 * 1024,
  },
);
assert.equal(child.stderr, '');
const workerEvidence = JSON.parse(child.stdout);
assert.equal(workerEvidence.path, '');
assert.equal(workerEvidence.fidelityLevel, 'R2');
assert.equal(workerEvidence.operationCount, 1);
assert.equal(workerEvidence.slideCount, 1);
assert.notEqual(workerEvidence.sourceSha256, workerEvidence.outputSha256);
assert.equal(workerEvidence.sourceStoredInPortableJson, true);
assert.deepEqual(workerEvidence.writeReport, {
  addedPartCount: 0,
  copiedPartCount: 11,
  patchedPartCount: 1,
  rebuiltPartCount: 0,
  removedPartCount: 0,
});

const restored = new Uint8Array(
  await readFile(path.join(outputDirectory, 'restored.pptx')),
);
const [sourceParts, restoredParts] = await Promise.all([
  partPayloads(source),
  partPayloads(restored),
]);
assert.equal(restoredParts.size, sourceParts.size);
for (const [name, sourcePayload] of sourceParts) {
  const restoredPayload = restoredParts.get(name);
  assert.ok(restoredPayload, `Missing restored part ${name}`);
  if (name === 'ppt/slides/slide1.xml') {
    assert.notDeepEqual(restoredPayload, sourcePayload);
  } else {
    assert.deepEqual(restoredPayload, sourcePayload);
  }
}
const svg = await readFile(path.join(outputDirectory, 'slide-1.svg'), 'utf8');
assert.match(svg.replaceAll('\u00a0', ' '), /Agent edited preview/);
assert.doesNotMatch(svg, /<(?:foreignObject|script)\b/i);
assert.doesNotMatch(svg, /(?:blob|file|https):/i);
const png = new Uint8Array(
  await readFile(path.join(outputDirectory, 'slide-1.png')),
);
const decoded = decodePng(png);
assert.deepEqual(
  { height: decoded.height, width: decoded.width },
  { height: 405, width: 720 },
);
assert.deepEqual(decoded.pixel(719, 404), [255, 0, 0, 255]);

const evidence = {
  execution: {
    officeRuntimeRequired: false,
    path: workerEvidence.path,
    runtime: process.version,
  },
  package: {
    fidelityLevel: workerEvidence.fidelityLevel,
    operationCount: workerEvidence.operationCount,
    outputSha256: sha256(restored),
    partInventoryPreserved: true,
    patchedParts: ['ppt/slides/slide1.xml'],
    sourceSha256: sha256(source),
    sourceStoredInPortableJson: workerEvidence.sourceStoredInPortableJson,
    untouchedPartPayloadsPreserved: true,
    writeReport: workerEvidence.writeReport,
  },
  png: {
    ...workerEvidence.png,
    independentlyDecoded: true,
    sampledBottomRightRgba: decoded.pixel(719, 404),
    sha256: sha256(png),
  },
  portableJsonBytes: workerEvidence.portableJsonBytes,
  svg: {
    ...workerEvidence.svg,
    externalReferences: false,
    sha256: sha256(Buffer.from(svg)),
  },
};
await writeFile(
  path.join(outputDirectory, 'evidence.json'),
  `${JSON.stringify(evidence, null, 2)}\n`,
);
console.log(JSON.stringify(evidence, null, 2));
