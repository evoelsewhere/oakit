import { describe, expect, it } from 'vitest';

import { parsePptx } from '../../src';
import {
  createIndependentPptx,
  DRAWING_NS,
  OFFICE_REL_NS,
  OFFICE_REL_TYPE,
  PACKAGE_REL_NS,
  PRESENTATION_NS,
} from './pptx-package';

function mediaPicture(id: number, media: string): string {
  return `<p:pic>
    <p:nvPicPr>
      <p:cNvPr id="${id}" name="Media ${id}"/><p:cNvPicPr/>
      <p:nvPr>${media}</p:nvPr>
    </p:nvPicPr>
    <p:blipFill><a:blip r:embed="rIdPoster"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>
    <p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="457200"/></a:xfrm></p:spPr>
  </p:pic>`;
}

describe('PowerPoint media classification through the public API', () => {
  it('distinguishes linked, embedded, unsupported, missing, and overlapping media', async () => {
    const input = await createIndependentPptx({
      'ppt/slides/slide1.xml': `
        <p:sld xmlns:p="${PRESENTATION_NS}" xmlns:a="${DRAWING_NS}" xmlns:r="${OFFICE_REL_NS}">
          <p:cSld><p:spTree>
            <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
            <p:grpSpPr/>
            ${mediaPicture(730, '<a:videoFile r:link="rIdHttp"/>')}
            ${mediaPicture(731, '<a:videoFile r:link="rIdMp4"/>')}
            ${mediaPicture(732, '<a:videoFile r:link="rIdWebm"/>')}
            ${mediaPicture(733, '<a:videoFile r:link="rIdVideoOgg"/>')}
            ${mediaPicture(734, '<a:videoFile r:link="rIdAvi"/>')}
            ${mediaPicture(735, '<a:audioFile r:link="rIdMp3"/>')}
            ${mediaPicture(736, '<a:audioFile r:link="rIdWav"/>')}
            ${mediaPicture(737, '<a:audioFile r:link="rIdAudioOgg"/>')}
            ${mediaPicture(738, '<a:audioFile r:link="rIdM4a"/>')}
            ${mediaPicture(
              739,
              '<a:videoFile r:link="rIdMp4"/><a:audioFile r:link="rIdMp3"/>',
            )}
            ${mediaPicture(740, '<a:videoFile r:link="rIdMissing"/>')}
            ${mediaPicture(741, '<a:audioFile r:link="rIdMissing"/>')}
            ${mediaPicture(742, '')}
          </p:spTree></p:cSld>
        </p:sld>`,
      'ppt/slides/_rels/slide1.xml.rels': `
        <Relationships xmlns="${PACKAGE_REL_NS}">
          <Relationship Id="rIdLayout" Type="${OFFICE_REL_TYPE}slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
          <Relationship Id="rIdPoster" Type="${OFFICE_REL_TYPE}image" Target="../media/poster.png"/>
          <Relationship Id="rIdHttp" Type="${OFFICE_REL_TYPE}video" Target="https://media.example/video" TargetMode="External"/>
          <Relationship Id="rIdMp4" Type="${OFFICE_REL_TYPE}video" Target="../media/movie.MP4"/>
          <Relationship Id="rIdWebm" Type="${OFFICE_REL_TYPE}video" Target="../media/movie.WEBM"/>
          <Relationship Id="rIdVideoOgg" Type="${OFFICE_REL_TYPE}video" Target="../media/movie.ogg"/>
          <Relationship Id="rIdAvi" Type="${OFFICE_REL_TYPE}video" Target="../media/movie.avi"/>
          <Relationship Id="rIdMp3" Type="${OFFICE_REL_TYPE}audio" Target="../media/sound.MP3"/>
          <Relationship Id="rIdWav" Type="${OFFICE_REL_TYPE}audio" Target="../media/sound.wav"/>
          <Relationship Id="rIdAudioOgg" Type="${OFFICE_REL_TYPE}audio" Target="../media/sound.ogg"/>
          <Relationship Id="rIdM4a" Type="${OFFICE_REL_TYPE}audio" Target="../media/sound.m4a"/>
        </Relationships>`,
      'ppt/media/movie.MP4': new Uint8Array([1, 2, 3]),
      'ppt/media/movie.WEBM': new Uint8Array([4, 5, 6]),
      'ppt/media/movie.ogg': new Uint8Array([7, 8, 9]),
      'ppt/media/movie.avi': new Uint8Array([10, 11, 12]),
      'ppt/media/sound.MP3': new Uint8Array([13, 14, 15]),
      'ppt/media/sound.wav': new Uint8Array([16, 17, 18]),
      'ppt/media/sound.ogg': new Uint8Array([19, 20, 21]),
      'ppt/media/sound.m4a': new Uint8Array([22, 23, 24]),
    });

    const result = await parsePptx(input, {
      audioMode: 'blob',
      errorMode: 'strict',
      imageMode: 'none',
      videoMode: 'blob',
    });
    const byId = Object.fromEntries(
      (result.slides[0]?.elements ?? []).map((element) => [
        element.id,
        element,
      ]),
    );

    expect(byId['730']).toMatchObject({
      blob: '',
      ref: 'https://media.example/video',
      type: 'video',
    });
    const blobUrls: string[] = [];
    for (const [id, ref] of [
      ['731', 'ppt/media/movie.MP4'],
      ['732', 'ppt/media/movie.WEBM'],
      ['733', 'ppt/media/movie.ogg'],
      ['739', 'ppt/media/movie.MP4'],
    ] as const) {
      const video = byId[id];
      expect(video?.type).toBe('video');
      if (video?.type !== 'video') throw new Error(`Expected video ${id}`);
      expect(video.ref).toBe(ref);
      expect(video.blob).toMatch(/^blob:/);
      blobUrls.push(video.blob);
    }
    expect(byId['734']).toMatchObject({
      blob: '',
      ref: 'ppt/media/movie.avi',
      type: 'video',
    });
    for (const [id, ref] of [
      ['735', 'ppt/media/sound.MP3'],
      ['736', 'ppt/media/sound.wav'],
      ['737', 'ppt/media/sound.ogg'],
    ] as const) {
      const audio = byId[id];
      expect(audio?.type).toBe('audio');
      if (audio?.type !== 'audio') throw new Error(`Expected audio ${id}`);
      expect(audio.ref).toBe(ref);
      expect(audio.blob).toMatch(/^blob:/);
      blobUrls.push(audio.blob);
    }
    expect(byId['738']).toMatchObject({
      blob: '',
      ref: 'ppt/media/sound.m4a',
      type: 'audio',
    });
    expect(byId['740']).toMatchObject({
      ref: 'ppt/media/poster.png',
      type: 'image',
    });
    expect(byId['741']).toMatchObject({
      ref: 'ppt/media/poster.png',
      type: 'image',
    });
    expect(byId['742']).toMatchObject({
      ref: 'ppt/media/poster.png',
      type: 'image',
    });

    for (const blobUrl of blobUrls) URL.revokeObjectURL(blobUrl);
  });
});
