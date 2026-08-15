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

const DIAGRAM_NS = 'http://schemas.openxmlformats.org/drawingml/2006/diagram';
const DIAGRAM_DRAWING_NS =
  'http://schemas.microsoft.com/office/drawing/2008/diagram';

const MEDIA_DIAGRAM_SLIDE = `
  <p:sld xmlns:p="${PRESENTATION_NS}" xmlns:a="${DRAWING_NS}" xmlns:dgm="${DIAGRAM_NS}" xmlns:r="${OFFICE_REL_NS}">
    <p:cSld><p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>
      <p:graphicFrame>
        <p:nvGraphicFramePr><p:cNvPr id="2" name="Process diagram"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>
        <p:xfrm><a:off x="0" y="0"/><a:ext cx="1828800" cy="914400"/></p:xfrm>
        <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/diagram">
          <dgm:relIds r:dm="rIdDiagramData"/>
        </a:graphicData></a:graphic>
      </p:graphicFrame>
      <p:pic>
        <p:nvPicPr><p:cNvPr id="3" name="Video"/><p:cNvPicPr/><p:nvPr><a:videoFile r:link="rIdVideo"/></p:nvPr></p:nvPicPr>
        <p:blipFill><a:blip r:embed="rIdPoster"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>
        <p:spPr><a:xfrm><a:off x="0" y="914400"/><a:ext cx="914400" cy="914400"/></a:xfrm></p:spPr>
      </p:pic>
      <p:pic>
        <p:nvPicPr><p:cNvPr id="4" name="Audio"/><p:cNvPicPr/><p:nvPr><a:audioFile r:link="rIdAudio"/></p:nvPr></p:nvPicPr>
        <p:blipFill><a:blip r:embed="rIdPoster"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>
        <p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="914400" cy="914400"/></a:xfrm></p:spPr>
      </p:pic>
      <p:grpSp>
        <p:nvGrpSpPr><p:cNvPr id="5" name="Custom group"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
        <p:grpSpPr><a:xfrm><a:off x="1828800" y="0"/><a:ext cx="1828800" cy="1828800"/><a:chOff x="0" y="0"/><a:chExt cx="1828800" cy="1828800"/></a:xfrm></p:grpSpPr>
        <p:sp>
          <p:nvSpPr><p:cNvPr id="6" name="Custom triangle"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
          <p:spPr>
            <a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm>
            <a:custGeom><a:avLst/><a:pathLst><a:path w="100" h="100">
              <a:moveTo><a:pt x="0" y="100"/></a:moveTo>
              <a:lnTo><a:pt x="50" y="0"/></a:lnTo>
              <a:lnTo><a:pt x="100" y="100"/></a:lnTo><a:close/>
            </a:path></a:pathLst></a:custGeom>
          </p:spPr>
        </p:sp>
      </p:grpSp>
    </p:spTree></p:cSld>
  </p:sld>`;

const DIAGRAM_DATA = `
  <dgm:dataModel xmlns:dgm="${DIAGRAM_NS}" xmlns:a="${DRAWING_NS}" xmlns:dsp="${DIAGRAM_DRAWING_NS}">
    <dgm:ptLst>
      <dgm:pt modelId="1"><dgm:t><a:p><a:r><a:t>Plan</a:t></a:r></a:p></dgm:t></dgm:pt>
      <dgm:pt modelId="2"><dgm:t><a:p><a:r><a:t>Build</a:t></a:r></a:p></dgm:t></dgm:pt>
    </dgm:ptLst>
    <dgm:extLst><a:ext uri="diagram-drawing"><dsp:dataModelExt relId="rIdDiagramDrawing"/></a:ext></dgm:extLst>
  </dgm:dataModel>`;

const DIAGRAM_DRAWING = `
  <dsp:drawing xmlns:dsp="${DIAGRAM_DRAWING_NS}" xmlns:a="${DRAWING_NS}">
    <dsp:spTree><dsp:sp>
      <dsp:nvSpPr><dsp:cNvPr id="20" name="Rendered node"/><dsp:cNvSpPr/><dsp:nvPr/></dsp:nvSpPr>
      <dsp:spPr>
        <a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="457200"/></a:xfrm>
        <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
        <a:solidFill><a:srgbClr val="70AD47"/></a:solidFill>
      </dsp:spPr>
    </dsp:sp></dsp:spTree>
  </dsp:drawing>`;

describe('PPTX diagram, media, and grouped custom shapes', () => {
  it('parses diagram text, linked media, and nested custom geometry', async () => {
    const input = await createIndependentPptx({
      'ppt/slides/slide1.xml': MEDIA_DIAGRAM_SLIDE,
      'ppt/slides/_rels/slide1.xml.rels': `
        <Relationships xmlns="${PACKAGE_REL_NS}">
          <Relationship Id="rIdLayout" Type="${OFFICE_REL_TYPE}slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
          <Relationship Id="rIdDiagramData" Type="${OFFICE_REL_TYPE}diagramData" Target="../diagrams/data1.xml"/>
          <Relationship Id="rIdDiagramDrawing" Type="${OFFICE_REL_TYPE}diagramDrawing" Target="../diagrams/drawing1.xml"/>
          <Relationship Id="rIdPoster" Type="${OFFICE_REL_TYPE}image" Target="../media/poster.png"/>
          <Relationship Id="rIdVideo" Type="${OFFICE_REL_TYPE}video" Target="../media/clip.mp4"/>
          <Relationship Id="rIdAudio" Type="${OFFICE_REL_TYPE}audio" Target="../media/sound.mp3"/>
        </Relationships>`,
      'ppt/diagrams/data1.xml': DIAGRAM_DATA,
      'ppt/diagrams/drawing1.xml': DIAGRAM_DRAWING,
    });

    const result = await parsePptx(input, {
      audioMode: 'none',
      errorMode: 'strict',
      imageMode: 'none',
      videoMode: 'none',
    });
    const elements = result.slides[0]?.elements ?? [];
    const diagram = elements.find((element) => element.type === 'diagram');
    const video = elements.find((element) => element.type === 'video');
    const audio = elements.find((element) => element.type === 'audio');
    const group = elements.find((element) => element.type === 'group');

    expect(diagram?.type).toBe('diagram');
    if (diagram?.type !== 'diagram') throw new Error('Expected a diagram');
    expect(diagram.textList).toEqual(['Plan', 'Build']);
    expect(diagram.elements).toHaveLength(1);
    expect(diagram.elements[0]).toMatchObject({
      id: '20',
      shapType: 'rect',
      type: 'shape',
    });
    expect(video).toMatchObject({
      id: '3',
      ref: 'ppt/media/clip.mp4',
      blob: '',
    });
    expect(audio).toMatchObject({
      id: '4',
      ref: 'ppt/media/sound.mp3',
      blob: '',
    });
    expect(group?.type).toBe('group');
    if (group?.type !== 'group') throw new Error('Expected a group');
    expect(group.elements).toHaveLength(1);
    expect(group.elements[0]).toMatchObject({
      id: '6',
      type: 'shape',
      shapType: 'custom',
    });
    if (group.elements[0]?.type !== 'shape') {
      throw new Error('Expected a custom shape');
    }
    expect(group.elements[0].path).toContain('M0,72');
    expect(group.elements[0].path).toContain('L72,72');
  });

  it('loads embedded audio and video bytes as independently readable blobs', async () => {
    const videoBytes = new Uint8Array([0, 0, 0, 24, 102, 116, 121, 112]);
    const audioBytes = new Uint8Array([73, 68, 51, 4, 0, 0, 0, 0]);
    const input = await createIndependentPptx({
      'ppt/slides/slide1.xml': MEDIA_DIAGRAM_SLIDE,
      'ppt/slides/_rels/slide1.xml.rels': `
        <Relationships xmlns="${PACKAGE_REL_NS}">
          <Relationship Id="rIdLayout" Type="${OFFICE_REL_TYPE}slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
          <Relationship Id="rIdPoster" Type="${OFFICE_REL_TYPE}image" Target="../media/poster.png"/>
          <Relationship Id="rIdVideo" Type="${OFFICE_REL_TYPE}video" Target="../media/clip.mp4"/>
          <Relationship Id="rIdAudio" Type="${OFFICE_REL_TYPE}audio" Target="../media/sound.mp3"/>
        </Relationships>`,
      'ppt/media/clip.mp4': videoBytes,
      'ppt/media/sound.mp3': audioBytes,
    });

    const result = await parsePptx(input, {
      audioMode: 'blob',
      errorMode: 'strict',
      imageMode: 'none',
      videoMode: 'blob',
    });
    const elements = result.slides[0]?.elements ?? [];
    const video = elements.find((element) => element.type === 'video');
    const audio = elements.find((element) => element.type === 'audio');
    if (video?.type !== 'video' || audio?.type !== 'audio') {
      throw new Error('Expected both media elements');
    }

    try {
      const loadedVideo = await fetch(video.blob).then((response) =>
        response.arrayBuffer(),
      );
      const loadedAudio = await fetch(audio.blob).then((response) =>
        response.arrayBuffer(),
      );
      expect(new Uint8Array(loadedVideo)).toEqual(videoBytes);
      expect(new Uint8Array(loadedAudio)).toEqual(audioBytes);
    } finally {
      URL.revokeObjectURL(video.blob);
      URL.revokeObjectURL(audio.blob);
    }
  });
});
