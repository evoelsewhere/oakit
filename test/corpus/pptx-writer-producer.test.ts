import {
  mkdtemp,
  mkdir,
  readFile,
  rm,
  stat,
  writeFile,
} from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { pathToFileURL } from 'node:url';
import { spawnSync } from 'node:child_process';

import { describe, expect, it } from 'vitest';

import {
  createPptx,
  parsePptx,
  readPptxRoundTrip,
  replacePptxRoundTripText,
  writePptxRoundTrip,
  type PptxSceneDocument,
} from '../../src';

function producerScene(): PptxSceneDocument {
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
            key: 'producer-text-1',
            resolved: { hidden: false },
            text: {
              body: { anchor: 'center', wrap: true },
              paragraphs: [
                {
                  children: [
                    {
                      key: 'producer-run-1',
                      properties: { bold: true, fontSize: 24 },
                      text: 'OAKit LibreOffice verification',
                      type: 'run',
                    },
                  ],
                  key: 'producer-paragraph-1',
                },
              ],
            },
            type: 'text',
          },
        ],
        key: 'producer-slide-1',
        name: 'Producer slide one',
      },
      {
        elements: [
          {
            authored: {
              transform: { height: 60, width: 280, x: 40, y: 50 },
            },
            key: 'producer-text-2',
            resolved: { hidden: false },
            text: {
              body: {},
              paragraphs: [
                {
                  children: [
                    {
                      key: 'producer-run-2',
                      text: 'Second producer slide',
                      type: 'run',
                    },
                  ],
                  key: 'producer-paragraph-2',
                },
              ],
            },
            type: 'text',
          },
        ],
        key: 'producer-slide-2',
        name: 'Producer slide two',
      },
    ],
    themes: [],
  };
}

function convertWithLibreOffice(
  executable: string,
  profileDirectory: string,
  outputDirectory: string,
  format: 'pdf' | 'pptx',
  input: string,
): void {
  const conversion = spawnSync(
    executable,
    [
      '--headless',
      `-env:UserInstallation=${pathToFileURL(profileDirectory).href}`,
      '--convert-to',
      format,
      '--outdir',
      outputDirectory,
      input,
    ],
    { encoding: 'utf8' },
  );
  if (conversion.error || conversion.status !== 0) {
    throw new Error(
      `LibreOffice ${format} conversion failed: ${conversion.error?.message ?? conversion.stderr ?? 'unknown error'}`,
    );
  }
}

describe('PowerPoint writer producer compatibility', () => {
  it('survives LibreOffice open, PDF export, and PPTX resave', async () => {
    const directory = await mkdtemp(join(tmpdir(), 'oakit-pptx-writer-'));
    try {
      const source = join(directory, 'created.pptx');
      const resavedDirectory = join(directory, 'resaved');
      const pdfDirectory = join(directory, 'pdf');
      await mkdir(resavedDirectory);
      await mkdir(pdfDirectory);

      const created = await createPptx(producerScene());
      await writeFile(source, created.data);
      const executable = process.env.SOFFICE_PATH ?? 'soffice';
      convertWithLibreOffice(
        executable,
        join(directory, 'profile-pptx'),
        resavedDirectory,
        'pptx',
        source,
      );
      const resaved = join(resavedDirectory, 'created.pptx');
      convertWithLibreOffice(
        executable,
        join(directory, 'profile-pdf'),
        pdfDirectory,
        'pdf',
        resaved,
      );

      const resavedBytes = new Uint8Array(await readFile(resaved));
      const parsed = await parsePptx(resavedBytes, {
        errorMode: 'strict',
        imageMode: 'none',
      });
      expect(parsed.size).toEqual({ height: 540, width: 960 });
      expect(parsed.slides).toHaveLength(2);
      expect(parsed.slides.map((slide) => slide.elements.length)).toEqual([
        1, 1,
      ]);
      expect(JSON.stringify(parsed)).toContain(
        'OAKit&nbsp;LibreOffice&nbsp;verification',
      );
      expect(JSON.stringify(parsed)).toContain(
        'Second&nbsp;producer&nbsp;slide',
      );
      expect(
        (await stat(join(pdfDirectory, 'created.pdf'))).size,
      ).toBeGreaterThan(0);
    } finally {
      await rm(directory, { force: true, recursive: true });
    }
  });

  it('survives two LibreOffice save and reopen cycles after a text edit', async () => {
    const directory = await mkdtemp(join(tmpdir(), 'oakit-pptx-edit-'));
    try {
      const source = join(directory, 'edited.pptx');
      const firstDirectory = join(directory, 'first-save');
      const secondDirectory = join(directory, 'second-save');
      const pdfDirectory = join(directory, 'pdf');
      await Promise.all([
        mkdir(firstDirectory),
        mkdir(secondDirectory),
        mkdir(pdfDirectory),
      ]);

      const created = await createPptx(producerScene());
      const snapshot = await readPptxRoundTrip(created.data);
      const run = snapshot.document.slides[0]?.elements[0];
      if (run?.type !== 'text') throw new Error('Expected editable text');
      const target = run.text.paragraphs[0]?.children[0];
      if (target?.type !== 'run') throw new Error('Expected editable run');
      const operation = await replacePptxRoundTripText(snapshot, {
        targetKey: target.key,
        value: 'OAKit edited by an agent',
      });
      const edited = await writePptxRoundTrip(operation);
      expect(edited.report).toMatchObject({
        level: 'R2',
        operations: [{ kind: 'replace-text', status: 'verified' }],
        patchedPartCount: 1,
      });
      await writeFile(source, edited.data);

      const executable = process.env.SOFFICE_PATH ?? 'soffice';
      convertWithLibreOffice(
        executable,
        join(directory, 'profile-first'),
        firstDirectory,
        'pptx',
        source,
      );
      const firstSave = join(firstDirectory, 'edited.pptx');
      convertWithLibreOffice(
        executable,
        join(directory, 'profile-second'),
        secondDirectory,
        'pptx',
        firstSave,
      );
      const secondSave = join(secondDirectory, 'edited.pptx');
      convertWithLibreOffice(
        executable,
        join(directory, 'profile-pdf'),
        pdfDirectory,
        'pdf',
        secondSave,
      );

      const reopened = await parsePptx(
        new Uint8Array(await readFile(secondSave)),
        { errorMode: 'strict', imageMode: 'none' },
      );
      const serialized = JSON.stringify(reopened);
      expect(reopened.size).toEqual({ height: 540, width: 960 });
      expect(reopened.slides).toHaveLength(2);
      expect(serialized).toContain(
        'OAKit&nbsp;edited&nbsp;by&nbsp;an&nbsp;agent',
      );
      expect(serialized).not.toContain(
        'OAKit&nbsp;LibreOffice&nbsp;verification',
      );
      expect(serialized).toContain('Second&nbsp;producer&nbsp;slide');
      expect(
        (await stat(join(pdfDirectory, 'edited.pdf'))).size,
      ).toBeGreaterThan(0);
    } finally {
      await rm(directory, { force: true, recursive: true });
    }
  });
});
