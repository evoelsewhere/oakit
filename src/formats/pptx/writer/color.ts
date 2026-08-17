const HEX_COLOR_PATTERN = /^#[0-9a-f]{6}$/i;

export function powerPointRgbValue(color: string): string {
  if (!HEX_COLOR_PATTERN.test(color)) {
    throw new TypeError('PowerPoint color must use #RRGGBB');
  }
  return color.slice(1).toUpperCase();
}

export function serializeSolidColorFill(color: string): string {
  return `<a:solidFill><a:srgbClr val="${powerPointRgbValue(color)}"/></a:solidFill>`;
}
