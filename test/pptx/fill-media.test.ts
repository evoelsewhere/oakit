import { afterEach, describe, expect, it, vi } from 'vitest';

import {
  getAudioData,
  getImageData,
  getPicFill,
  getPicFillOpacity,
  getPicFilters,
  getVideoData,
  loadAudio,
  loadImage,
  loadVideo,
} from '../../src/formats/pptx/internal/fill';
import { fillContext, xml } from './fill-fixture';

afterEach(() => {
  vi.restoreAllMocks();
});

describe('PPTX embedded media loading', () => {
  it('does not read an empty path or an XML relationship as media', async () => {
    const fixture = fillContext();

    await expect(loadImage('', fixture.context)).resolves.toBe('');
    await expect(
      loadImage('ppt/diagrams/data1.xml', fixture.context),
    ).resolves.toBe('');
    expect(fixture.readMedia).not.toHaveBeenCalled();
  });

  it('encodes an image once and reuses its path-and-mode cache entry', async () => {
    const fixture = fillContext({
      media: { 'ppt/media/image1.png': new Uint8Array([0, 1, 2, 253]) },
    });

    await expect(
      loadImage('ppt/media/image1.png', fixture.context),
    ).resolves.toBe('data:image/png;base64,AAEC/Q==');
    await expect(
      loadImage('ppt/media/image1.png', fixture.context),
    ).resolves.toBe('data:image/png;base64,AAEC/Q==');
    expect(fixture.readMedia).toHaveBeenCalledTimes(1);
    expect(fixture.context.loadedImages['ppt/media/image1.png']).toEqual({
      base64: 'data:image/png;base64,AAEC/Q==',
      blob: '',
      ref: 'ppt/media/image1.png',
    });
  });

  it('creates a typed Blob from only the selected Uint8Array view', async () => {
    const backing = new Uint8Array([9, 1, 2, 9]);
    const fixture = fillContext({
      media: { 'ppt/media/image1.png': backing.subarray(1, 3) },
    });
    let received: Blob | undefined;
    const createObjectUrl = vi
      .spyOn(URL, 'createObjectURL')
      .mockImplementation((blob) => {
        if (!(blob instanceof Blob)) throw new TypeError('Expected a Blob');
        received = blob;
        return 'blob:image-1';
      });

    await expect(
      loadImage('ppt/media/image1.png', fixture.context, 'blob'),
    ).resolves.toBe('blob:image-1');
    await expect(
      loadImage('ppt/media/image1.png', fixture.context, 'blob'),
    ).resolves.toBe('blob:image-1');
    expect(received?.type).toBe('image/png');
    await expect(received?.arrayBuffer()).resolves.toEqual(
      new Uint8Array([1, 2]).buffer,
    );
    expect(fixture.readMedia).toHaveBeenCalledTimes(1);
    expect(createObjectUrl).toHaveBeenCalledTimes(1);
  });

  it('returns an empty value when the relationship media is absent', async () => {
    const fixture = fillContext();

    await expect(
      loadImage('ppt/media/missing.png', fixture.context),
    ).resolves.toBe('');
    expect(fixture.readMedia).toHaveBeenCalledWith('ppt/media/missing.png');
  });

  it('honors all image output modes', async () => {
    vi.spyOn(URL, 'createObjectURL').mockReturnValue('blob:image-1');
    const bytes = new Uint8Array([1, 2]);

    await expect(
      getImageData(
        'ppt/media/image1.png',
        fillContext({ media: { 'ppt/media/image1.png': bytes } }).context,
      ),
    ).resolves.toEqual({
      base64: 'data:image/png;base64,AQI=',
      blob: '',
      ref: 'ppt/media/image1.png',
    });
    await expect(
      getImageData(
        'ppt/media/image1.png',
        fillContext({
          media: { 'ppt/media/image1.png': bytes },
          options: { imageMode: 'blob' },
        }).context,
      ),
    ).resolves.toEqual({
      base64: '',
      blob: 'blob:image-1',
      ref: 'ppt/media/image1.png',
    });
    await expect(
      getImageData(
        'ppt/media/image1.png',
        fillContext({
          media: { 'ppt/media/image1.png': bytes },
          options: { imageMode: 'both' },
        }).context,
      ),
    ).resolves.toEqual({
      base64: 'data:image/png;base64,AQI=',
      blob: 'blob:image-1',
      ref: 'ppt/media/image1.png',
    });
    await expect(
      getImageData(
        'ppt/media/image1.png',
        fillContext({
          media: { 'ppt/media/image1.png': bytes },
          options: { imageMode: 'none' },
        }).context,
      ),
    ).resolves.toEqual({
      base64: '',
      blob: '',
      ref: 'ppt/media/image1.png',
    });
  });

  it('returns neutral image data for an empty path in both mode', async () => {
    const fixture = fillContext({ options: { imageMode: 'both' } });

    await expect(getImageData('', fixture.context)).resolves.toEqual({
      base64: '',
      blob: '',
      ref: '',
    });
    expect(fixture.readMedia).not.toHaveBeenCalled();
  });

  it('only loads audio and video in their supported Blob mode', async () => {
    vi.spyOn(URL, 'createObjectURL')
      .mockReturnValueOnce('blob:video-1')
      .mockReturnValueOnce('blob:audio-1');
    const fixture = fillContext({
      media: {
        'ppt/media/audio1.wav': new Uint8Array([1]),
        'ppt/media/video1.mp4': new Uint8Array([2]),
      },
    });

    await expect(
      loadVideo('ppt/media/video1.mp4', fixture.context, 'base64'),
    ).resolves.toBe('');
    await expect(
      loadAudio('ppt/media/audio1.wav', fixture.context, 'base64'),
    ).resolves.toBe('');
    expect(fixture.readMedia).not.toHaveBeenCalled();
    await expect(
      loadVideo('ppt/media/video1.mp4', fixture.context),
    ).resolves.toBe('blob:video-1');
    await expect(
      loadAudio('ppt/media/audio1.wav', fixture.context),
    ).resolves.toBe('blob:audio-1');
    expect(fixture.readMedia).toHaveBeenCalledTimes(2);
  });

  it('honors enabled and disabled audio/video parser options', async () => {
    vi.spyOn(URL, 'createObjectURL')
      .mockReturnValueOnce('blob:video-1')
      .mockReturnValueOnce('blob:audio-1');
    const disabled = fillContext();
    await expect(
      getVideoData('ppt/media/video1.mp4', disabled.context),
    ).resolves.toEqual({ blob: '', ref: 'ppt/media/video1.mp4' });
    await expect(
      getAudioData('ppt/media/audio1.wav', disabled.context),
    ).resolves.toEqual({ blob: '', ref: 'ppt/media/audio1.wav' });
    expect(disabled.readMedia).not.toHaveBeenCalled();

    const enabled = fillContext({
      media: {
        'ppt/media/audio1.wav': new Uint8Array([1]),
        'ppt/media/video1.mp4': new Uint8Array([2]),
      },
      options: { audioMode: 'blob', videoMode: 'blob' },
    });
    await expect(
      getVideoData('ppt/media/video1.mp4', enabled.context),
    ).resolves.toEqual({
      blob: 'blob:video-1',
      ref: 'ppt/media/video1.mp4',
    });
    await expect(
      getAudioData('ppt/media/audio1.wav', enabled.context),
    ).resolves.toEqual({
      blob: 'blob:audio-1',
      ref: 'ppt/media/audio1.wav',
    });
  });
});

describe('PPTX picture relationships', () => {
  const relationship = (target: string) => ({ target, type: 'image' });

  it.each([
    ['slide', 'slide'],
    ['slideBg', 'slide'],
    ['slideLayoutBg', 'layout'],
    ['slideMasterBg', 'master'],
    ['themeBg', 'theme'],
    ['diagramBg', 'diagram'],
  ] as const)(
    'resolves %s against the %s relationship map',
    async (source, map) => {
      const fixture = fillContext({
        layoutRelationships: { rIdImage: relationship('layout') },
        masterRelationships: { rIdImage: relationship('master') },
        options: { imageMode: 'none' },
        slideRelationships: { rIdImage: relationship('slide') },
        themeRelationships: { rIdImage: relationship('theme') },
      });
      fixture.context.diagramResObj = {
        rIdImage: relationship('diagram'),
      };

      await expect(
        getPicFill(
          source,
          xml({ 'a:blip': { attrs: { 'r:embed': 'rIdImage' } } }),
          fixture.context,
        ),
      ).resolves.toEqual({ base64: '', blob: '', ref: map });
    },
  );

  it('returns neutral data for missing nodes, IDs, maps, and targets', async () => {
    const fixture = fillContext({ options: { imageMode: 'none' } });

    await expect(
      getPicFill('slide', undefined, fixture.context),
    ).resolves.toEqual({ base64: '', blob: '', ref: '' });
    await expect(
      getPicFill('slide', xml({}), fixture.context),
    ).resolves.toEqual({ base64: '', blob: '', ref: '' });
    await expect(
      getPicFill(
        'unknown',
        xml({ 'a:blip': { attrs: { 'r:embed': 'rIdImage' } } }),
        fixture.context,
      ),
    ).resolves.toEqual({ base64: '', blob: '', ref: '' });
    await expect(
      getPicFill(
        'slide',
        xml({ 'a:blip': { attrs: { 'r:embed': 'rIdImage' } } }),
        fixture.context,
      ),
    ).resolves.toEqual({ base64: '', blob: '', ref: '' });
  });
});

describe('PPTX picture numeric effects', () => {
  it.each([
    [undefined, 1],
    ['0', 0],
    ['25%', 0.25],
    ['50000', 0.5],
    ['100000', 1],
    ['200000', 1],
    ['-100000', 0],
    ['50000x', 1],
    ['  ', 1],
    ['5%0', 1],
    ['Infinity', 1],
  ] as const)('normalizes opacity %j', (amount, expected) => {
    expect(
      getPicFillOpacity(
        xml({
          'a:blip': {
            'a:alphaModFix': {
              attrs: { ...(amount === undefined ? {} : { amt: amount }) },
            },
          },
        }),
      ),
    ).toBe(expected);
  });

  it('extracts all finite picture filters across extension effects', () => {
    expect(
      getPicFilters(
        xml({
          'a:blip': {
            'a:extLst': {
              'a:ext': [
                {
                  'a14:imgProps': {
                    'a14:imgLayer': {
                      'a14:imgEffect': [
                        {
                          'a14:brightnessContrast': {
                            attrs: { bright: '-25000', contrast: '50%' },
                          },
                          'a14:colorTemperature': {
                            attrs: { colorTemp: '6500' },
                          },
                          'a14:saturation': { attrs: { sat: '125000' } },
                          'a14:sharpenSoften': {
                            attrs: { amount: '40000' },
                          },
                        },
                        {
                          'a14:sharpenSoften': {
                            attrs: { amount: '-30000' },
                          },
                        },
                      ],
                    },
                  },
                },
              ],
            },
          },
        }),
      ),
    ).toEqual({
      brightness: -0.25,
      colorTemperature: 6500,
      contrast: 0.5,
      saturation: 1.25,
      sharpen: 0.4,
      soften: 0.3,
    });
  });

  it('omits malformed and zero-value effects', () => {
    expect(
      getPicFilters(
        xml({
          'a:blip': {
            'a:extLst': {
              'a:ext': {
                'a14:imgProps': {
                  'a14:imgLayer': {
                    'a14:imgEffect': {
                      'a14:brightnessContrast': {
                        attrs: { bright: '10x', contrast: 'Infinity' },
                      },
                      'a14:colorTemperature': {
                        attrs: { colorTemp: '6500x' },
                      },
                      'a14:saturation': { attrs: { sat: '' } },
                      'a14:sharpenSoften': { attrs: { amount: '0' } },
                    },
                  },
                },
              },
            },
          },
        }),
      ),
    ).toBeNull();
    expect(getPicFilters(undefined)).toBeNull();
  });
});
