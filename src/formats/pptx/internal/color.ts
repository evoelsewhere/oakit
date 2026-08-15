import tinycolor from 'tinycolor2';

function finiteOr(value: number, fallback: number): number {
  return Number.isFinite(value) ? value : fallback;
}

function clampUnit(value: number): number {
  return Math.min(Math.max(value, 0), 1);
}

function wrap(value: number, period: number): number {
  return ((value % period) + period) % period;
}

function formatHsl(
  rgbStr: string,
  hue: number,
  saturation: number,
  lightness: number,
  isAlpha: boolean,
): string {
  const source = tinycolor(rgbStr).toHsl();
  const transformed = tinycolor({
    h: wrap(finiteOr(hue, source.h), 360),
    s: clampUnit(finiteOr(saturation, source.s)),
    l: clampUnit(finiteOr(lightness, source.l)),
    a: source.a,
  });
  return isAlpha ? transformed.toHex8() : transformed.toHex();
}

export function hueToRgb(t1: number, t2: number, hue: number): number {
  const lower = finiteOr(t1, 0);
  const upper = finiteOr(t2, 0);
  const normalizedHue = wrap(finiteOr(hue, 0), 6);
  switch (Math.floor(normalizedHue)) {
    case 0:
      return (upper - lower) * normalizedHue + lower;
    case 1:
    case 2:
      return upper;
    case 3:
      return (upper - lower) * (4 - normalizedHue) + lower;
    default:
      return lower;
  }
}

export function hslToRgb(hue: number, sat: number, light: number) {
  const normalizedHue = wrap(finiteOr(hue, 0), 360) / 60;
  const saturation = clampUnit(finiteOr(sat, 0));
  const lightness = clampUnit(finiteOr(light, 0));
  const t2 = lightness + saturation * Math.min(lightness, 1 - lightness);
  const t1 = lightness * 2 - t2;
  return {
    r: hueToRgb(t1, t2, normalizedHue + 2) * 255,
    g: hueToRgb(t1, t2, normalizedHue) * 255,
    b: hueToRgb(t1, t2, normalizedHue - 2) * 255,
  };
}

export function applyShade(
  rgbStr: string,
  shadeValue: number,
  isAlpha = false,
): string {
  const color = tinycolor(rgbStr).toHsl();
  const multiplier = clampUnit(finiteOr(shadeValue, 1));
  return formatHsl(rgbStr, color.h, color.s, color.l * multiplier, isAlpha);
}

export function applyTint(
  rgbStr: string,
  tintValue: number,
  isAlpha = false,
): string {
  const color = tinycolor(rgbStr).toHsl();
  const multiplier = clampUnit(finiteOr(tintValue, 1));
  return formatHsl(
    rgbStr,
    color.h,
    color.s,
    color.l * multiplier + (1 - multiplier),
    isAlpha,
  );
}

export function applyLumOff(
  rgbStr: string,
  offset: number,
  isAlpha = false,
): string {
  const color = tinycolor(rgbStr).toHsl();
  return formatHsl(
    rgbStr,
    color.h,
    color.s,
    color.l + finiteOr(offset, 0),
    isAlpha,
  );
}

export function applyLumMod(
  rgbStr: string,
  multiplier: number,
  isAlpha = false,
): string {
  const color = tinycolor(rgbStr).toHsl();
  return formatHsl(
    rgbStr,
    color.h,
    color.s,
    color.l * finiteOr(multiplier, 1),
    isAlpha,
  );
}

export function applyHueMod(
  rgbStr: string,
  multiplier: number,
  isAlpha = false,
): string {
  const color = tinycolor(rgbStr).toHsl();
  return formatHsl(
    rgbStr,
    color.h * finiteOr(multiplier, 1),
    color.s,
    color.l,
    isAlpha,
  );
}

export function applySatMod(
  rgbStr: string,
  multiplier: number,
  isAlpha = false,
): string {
  const color = tinycolor(rgbStr).toHsl();
  return formatHsl(
    rgbStr,
    color.h,
    color.s * finiteOr(multiplier, 1),
    color.l,
    isAlpha,
  );
}

function normalizePresetName(name: string): string {
  const normalized = name.trim().toLowerCase();
  if (normalized.startsWith('dk')) return `dark${normalized.slice(2)}`;
  if (normalized.startsWith('lt')) return `light${normalized.slice(2)}`;
  if (normalized.startsWith('med') && !normalized.startsWith('medium')) {
    return `medium${normalized.slice(3)}`;
  }
  return normalized;
}

export function getColorName2Hex(name: string): string | undefined {
  const normalized = normalizePresetName(name);
  if (!Object.hasOwn(tinycolor.names, normalized)) return undefined;
  return tinycolor(normalized).toHex();
}
