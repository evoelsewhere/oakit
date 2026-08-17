import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';

import {
  createPptx,
  parsePptxRoundTripJson,
  readPptxRoundTrip,
  replacePptxRoundTripText,
  serializePptxRoundTripJson,
  setPptxRoundTripTextTransform,
  writePptxRoundTrip,
  type PptxSceneDocument,
  type PptxSceneTransform,
} from '../../src';

function scene(): PptxSceneDocument {
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
              transform: { height: 80, width: 300, x: 20, y: 30 },
            },
            key: 'transform-text',
            resolved: { hidden: false },
            text: {
              body: {},
              paragraphs: [
                {
                  children: [
                    { key: 'transform-run', text: 'Before', type: 'run' },
                  ],
                  key: 'transform-paragraph',
                },
              ],
            },
            type: 'text',
          },
        ],
        key: 'transform-slide',
      },
    ],
    themes: [],
  };
}

function changedTransform(): PptxSceneTransform {
  return {
    flipHorizontal: true,
    flipVertical: true,
    height: 100,
    rotation: 45,
    width: 400,
    x: 50,
    y: 60,
  };
}

async function fixture() {
  const created = await createPptx(scene());
  return {
    bytes: created.data,
    snapshot: await readPptxRoundTrip(created.data),
  };
}

async function payloads(data: Uint8Array): Promise<Map<string, Uint8Array>> {
  const archive = await JSZip.loadAsync(data);
  const result = new Map<string, Uint8Array>();
  for (const file of Object.values(archive.files)) {
    if (!file.dir) result.set(file.name, await file.async('uint8array'));
  }
  return result;
}

async function rewriteSlide(
  data: Uint8Array,
  update: (xml: string) => string,
): Promise<Uint8Array> {
  const archive = await JSZip.loadAsync(data);
  const entry = archive.file('ppt/slides/slide1.xml');
  if (entry === null) throw new Error('Missing slide part');
  archive.file('ppt/slides/slide1.xml', update(await entry.async('text')));
  return archive.generateAsync({ type: 'uint8array' });
}

describe('PowerPoint part-preserving text transform editing', () => {
  it('moves, resizes, rotates, and flips a text element through the public API', async () => {
    const { bytes, snapshot } = await fixture();
    const edited = await setPptxRoundTripTextTransform(snapshot, {
      targetKey: 'slide-1-element-1',
      value: changedTransform(),
    });
    const output = await writePptxRoundTrip(edited);
    const [sourceParts, outputParts] = await Promise.all([
      payloads(bytes),
      payloads(output.data),
    ]);
    const [sourceArchive, outputArchive] = await Promise.all([
      JSZip.loadAsync(bytes),
      JSZip.loadAsync(output.data),
    ]);

    expect(output.report).toMatchObject({
      level: 'R2',
      operations: [
        { id: 'set-transform-1', kind: 'set-transform', status: 'verified' },
      ],
      patchedPartCount: 1,
    });
    for (const [name, source] of sourceParts) {
      const result = outputParts.get(name);
      expect(result, name).toBeDefined();
      if (name === 'ppt/slides/slide1.xml') {
        expect(result).not.toEqual(source);
      } else {
        expect(result, name).toEqual(source);
      }
    }
    const slideXml = new TextDecoder().decode(
      outputParts.get('ppt/slides/slide1.xml'),
    );
    expect(slideXml).toContain(
      '<a:xfrm rot="2700000" flipH="1" flipV="1"><a:off x="635000" y="762000"/><a:ext cx="5080000" cy="1270000"/></a:xfrm>',
    );
    expect(outputArchive.file('ppt/slides/slide1.xml')?.date.getTime()).toBe(
      sourceArchive.file('ppt/slides/slide1.xml')?.date.getTime(),
    );
    const verified = await readPptxRoundTrip(output.data);
    expect(verified.document.slides[0]?.elements[0]).toMatchObject({
      key: 'slide-1-element-1',
      resolved: { transform: changedTransform() },
      text: { paragraphs: [{ children: [{ text: 'Before' }] }] },
      type: 'text',
    });
  });

  it('survives portable JSON without replacing the original source bytes', async () => {
    const { bytes, snapshot } = await fixture();
    const edited = await setPptxRoundTripTextTransform(snapshot, {
      targetKey: 'slide-1-element-1',
      value: changedTransform(),
    });
    const portable = await serializePptxRoundTripJson(edited);
    const restored = await parsePptxRoundTripJson(
      JSON.parse(JSON.stringify(portable)) as unknown,
    );
    const output = await writePptxRoundTrip(restored);

    expect(restored.source.data).toEqual(bytes);
    expect(restored.operations).toEqual(edited.operations);
    expect(output.data).not.toEqual(bytes);
    expect(output.report.level).toBe('R2');
  });

  it('applies transform and text operations to one dirty part in order', async () => {
    const { snapshot } = await fixture();
    const transformed = await setPptxRoundTripTextTransform(snapshot, {
      targetKey: 'slide-1-element-1',
      value: changedTransform(),
    });
    const edited = await replacePptxRoundTripText(transformed, {
      targetKey: 'slide-1-element-1-run-1',
      value: 'After',
    });
    const output = await writePptxRoundTrip(edited);
    const verified = await readPptxRoundTrip(output.data);

    expect(output.report.patchedPartCount).toBe(1);
    expect(output.report.operations).toEqual([
      { id: 'set-transform-1', kind: 'set-transform', status: 'verified' },
      { id: 'replace-text-2', kind: 'replace-text', status: 'verified' },
    ]);
    expect(verified.document.slides[0]?.elements[0]).toMatchObject({
      resolved: { transform: changedTransform() },
      text: { paragraphs: [{ children: [{ text: 'After' }] }] },
    });
  });

  it('allows transform after a text operation targeting a different stable key', async () => {
    const { snapshot } = await fixture();
    const textEdited = await replacePptxRoundTripText(snapshot, {
      targetKey: 'slide-1-element-1-run-1',
      value: 'After',
    });
    const transformed = await setPptxRoundTripTextTransform(textEdited, {
      targetKey: 'slide-1-element-1',
      value: changedTransform(),
    });

    expect(transformed.operations.map((operation) => operation.kind)).toEqual([
      'replace-text',
      'set-transform',
    ]);
  });

  it('normalizes omitted rotation and flip fields', async () => {
    const { snapshot } = await fixture();
    const edited = await setPptxRoundTripTextTransform(snapshot, {
      targetKey: 'slide-1-element-1',
      value: { height: 90, width: 310, x: 21, y: 31 },
    });

    expect(edited.operations[0]).toMatchObject({
      kind: 'set-transform',
      value: {
        flipHorizontal: false,
        flipVertical: false,
        height: 90,
        rotation: 0,
        width: 310,
        x: 21,
        y: 31,
      },
    });
  });

  it('patches transform XML through namespace aliases', async () => {
    const { bytes } = await fixture();
    const aliased = await rewriteSlide(bytes, (xml) =>
      xml
        .replaceAll('xmlns:p=', 'xmlns:pres=')
        .replaceAll('<p:', '<pres:')
        .replaceAll('</p:', '</pres:')
        .replaceAll('xmlns:a=', 'xmlns:draw=')
        .replaceAll('<a:', '<draw:')
        .replaceAll('</a:', '</draw:'),
    );
    const edited = await setPptxRoundTripTextTransform(
      await readPptxRoundTrip(aliased),
      { targetKey: 'slide-1-element-1', value: changedTransform() },
    );
    const output = await writePptxRoundTrip(edited);
    const archive = await JSZip.loadAsync(output.data);
    const slide = await archive.file('ppt/slides/slide1.xml')?.async('text');

    expect(slide).toContain('<draw:xfrm rot="2700000"');
    expect(slide).not.toContain('<a:xfrm rot="2700000"');
  });

  it('rejects a transform with extra coordinate-space children', async () => {
    const { bytes } = await fixture();
    const complex = await rewriteSlide(bytes, (xml) =>
      xml.replace(
        /<a:xfrm><a:off x="254000" y="381000"\/><a:ext cx="3810000" cy="1016000"\/><\/a:xfrm>/,
        '<a:xfrm><a:off x="254000" y="381000"/><a:ext cx="3810000" cy="1016000"/><a:chOff x="0" y="0"/></a:xfrm>',
      ),
    );
    const edited = await setPptxRoundTripTextTransform(
      await readPptxRoundTrip(complex),
      { targetKey: 'slide-1-element-1', value: changedTransform() },
    );

    await expect(writePptxRoundTrip(edited)).rejects.toMatchObject({
      code: 'unsupported-edit-operation',
      message:
        'PowerPoint transform target must contain one simple shape transform',
    });
  });

  it.each([
    [
      'missing target',
      'missing',
      changedTransform(),
      'PowerPoint transform target key does not exist',
    ],
    [
      'zero width',
      'slide-1-element-1',
      { ...changedTransform(), width: 0 },
      'PowerPoint transform value is not a valid scene transform',
    ],
    [
      'zero height',
      'slide-1-element-1',
      { ...changedTransform(), height: 0 },
      'PowerPoint transform value is not a valid scene transform',
    ],
    [
      'non-finite x',
      'slide-1-element-1',
      { ...changedTransform(), x: Number.NaN },
      'PowerPoint transform value is not a valid scene transform',
    ],
    [
      'non-numeric y',
      'slide-1-element-1',
      { ...changedTransform(), y: '60' as never },
      'PowerPoint transform value is not a valid scene transform',
    ],
    [
      'missing required x',
      'slide-1-element-1',
      { ...changedTransform(), x: undefined as never },
      'PowerPoint transform value is not a valid scene transform',
    ],
    [
      'non-finite rotation',
      'slide-1-element-1',
      { ...changedTransform(), rotation: Number.NaN },
      'PowerPoint transform value is not a valid scene transform',
    ],
    [
      'non-numeric rotation',
      'slide-1-element-1',
      { ...changedTransform(), rotation: '45' as never },
      'PowerPoint transform value is not a valid scene transform',
    ],
    [
      'non-boolean flip',
      'slide-1-element-1',
      { ...changedTransform(), flipHorizontal: 1 as never },
      'PowerPoint transform value is not a valid scene transform',
    ],
    [
      'non-boolean vertical flip',
      'slide-1-element-1',
      { ...changedTransform(), flipVertical: 'true' as never },
      'PowerPoint transform value is not a valid scene transform',
    ],
    [
      'extra property',
      'slide-1-element-1',
      { ...changedTransform(), extra: true } as PptxSceneTransform,
      'PowerPoint transform value is not a valid scene transform',
    ],
  ])(
    'rejects an invalid request: %s',
    async (_name, targetKey, value, message) => {
      const { snapshot } = await fixture();

      await expect(
        setPptxRoundTripTextTransform(snapshot, { targetKey, value }),
      ).rejects.toMatchObject({ code: 'invalid-edit-operation', message });
    },
  );

  it('rejects a no-op and a duplicate transform target', async () => {
    const { snapshot } = await fixture();
    const element = snapshot.document.slides[0]?.elements[0];
    if (element?.resolved.transform === undefined) {
      throw new Error('Expected resolved transform');
    }
    await expect(
      setPptxRoundTripTextTransform(snapshot, {
        targetKey: element.key,
        value: element.resolved.transform,
      }),
    ).rejects.toMatchObject({
      code: 'invalid-edit-operation',
      message: 'PowerPoint transform edit must change the target value',
    });
    const edited = await setPptxRoundTripTextTransform(snapshot, {
      targetKey: element.key,
      value: changedTransform(),
    });
    await expect(
      setPptxRoundTripTextTransform(edited, {
        targetKey: element.key,
        value: { ...changedTransform(), x: 70 },
      }),
    ).rejects.toMatchObject({
      code: 'invalid-edit-operation',
      message: 'PowerPoint transform target is already scheduled',
    });
  });
});
