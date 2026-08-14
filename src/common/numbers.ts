export function angleToDegrees(
  angle: number | string | null | undefined,
): number {
  if (!angle) return 0;
  return Math.round(Number(angle) / 60_000);
}

export function toHex(value: number): string {
  return value.toString(16).padStart(2, '0');
}

export function numberToFixed(value: number, fractionDigits = 4): number {
  return Number.parseFloat(value.toFixed(fractionDigits));
}
