import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';

import {
  createPptx,
  parsePptxRoundTripJson,
  readPptxRoundTrip,
  renderPptxToSvg,
  serializePptxRoundTripJson,
  setPptxRoundTripChartTransform,
  setPptxRoundTripShapeTransform,
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
              transform: { height: 220, width: 420, x: 40, y: 60 },
            },
            barDirection: 'col',
            chartType: 'barChart',
            grouping: 'clustered',
            key: 'source-chart',
            resolved: { hidden: false },
            series: [
              {
                categories: ['Q1', 'Q2', 'Q3'],
                color: '#4F46E5',
                key: 'revenue-series',
                name: 'Revenue',
                values: [12, 18, 27],
              },
            ],
            type: 'chart',
          },
        ],
        key: 'source-slide',
      },
    ],
    themes: [],
  };
}

const CHANGED_TRANSFORM: PptxSceneTransform = {
  flipHorizontal: false,
  flipVertical: false,
  height: 260,
  rotation: 0,
  width: 500,
  x: 80,
  y: 90,
};

async function payloads(data: Uint8Array): Promise<Map<string, Uint8Array>> {
  const archive = await JSZip.loadAsync(data);
  const result = new Map<string, Uint8Array>();
  for (const file of Object.values(archive.files)) {
    if (!file.dir) result.set(file.name, await file.async('uint8array'));
  }
  return result;
}

describe('native PowerPoint chart transform editing', () => {
  it('patches only the chart frame while preserving exact ChartML', async () => {
    const created = await createPptx(scene());
    const snapshot = await readPptxRoundTrip(created.data);
    expect(snapshot.document.slides[0]?.elements[0]).toMatchObject({
      barDirection: 'col',
      chartType: 'barChart',
      grouping: 'clustered',
      key: 'slide-1-element-1',
      series: [
        {
          categories: ['Q1', 'Q2', 'Q3'],
          color: '#4F46E5',
          key: 'slide-1-element-1-series-1',
          name: 'Revenue',
          values: [12, 18, 27],
        },
      ],
      type: 'chart',
    });
    const edited = await setPptxRoundTripChartTransform(snapshot, {
      targetKey: 'slide-1-element-1',
      value: CHANGED_TRANSFORM,
    });
    const output = await writePptxRoundTrip(edited);
    const [sourceParts, outputParts, verified, rendered] = await Promise.all([
      payloads(created.data),
      payloads(output.data),
      readPptxRoundTrip(output.data),
      renderPptxToSvg(output.data),
    ]);

    expect(output.report).toMatchObject({
      level: 'R2',
      patchedPartCount: 1,
      supportProfile: { id: 'pptx-roundtrip-native-v1' },
    });
    expect(verified.document.slides[0]?.elements[0]).toMatchObject({
      chartType: 'barChart',
      resolved: { transform: CHANGED_TRANSFORM },
      series: [{ values: [12, 18, 27] }],
      type: 'chart',
    });
    for (const [name, source] of sourceParts) {
      const result = outputParts.get(name);
      expect(result, name).toBeDefined();
      if (name === 'ppt/slides/slide1.xml') expect(result).not.toEqual(source);
      else expect(result, name).toEqual(source);
    }
    expect(outputParts.get('ppt/charts/chart1.xml')).toEqual(
      sourceParts.get('ppt/charts/chart1.xml'),
    );
    expect(new TextDecoder().decode(rendered.slides[0]?.data)).toContain(
      'barChart',
    );
  });

  it('survives portable JSON with native chart capability binding', async () => {
    const snapshot = await readPptxRoundTrip((await createPptx(scene())).data);
    const edited = await setPptxRoundTripChartTransform(snapshot, {
      targetKey: 'slide-1-element-1',
      value: CHANGED_TRANSFORM,
    });
    const restored = await parsePptxRoundTripJson(
      JSON.parse(JSON.stringify(await serializePptxRoundTripJson(edited))),
    );
    const output = await writePptxRoundTrip(restored);

    expect(restored.consistency.capabilityProfileVersion).toBe(
      'pptx-roundtrip-native-v1',
    );
    expect(output.report.operations).toMatchObject([
      { kind: 'set-transform', status: 'verified' },
    ]);
  });

  it('keeps chart and shape transform APIs type-specific', async () => {
    const snapshot = await readPptxRoundTrip((await createPptx(scene())).data);

    await expect(
      setPptxRoundTripShapeTransform(snapshot, {
        targetKey: 'slide-1-element-1',
        value: CHANGED_TRANSFORM,
      }),
    ).rejects.toMatchObject({
      code: 'invalid-edit-operation',
      message: 'PowerPoint shape transform target key does not exist',
    });
  });
});
