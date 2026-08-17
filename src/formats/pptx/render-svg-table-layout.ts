export function normalizedTableSizes(
  values: unknown,
  count: number,
  total: number,
): number[] {
  if (Array.isArray(values) && values.length >= count) {
    const selected = values.slice(0, count);
    if (
      selected.every((value) => Number.isFinite(value) && (value as number) > 0)
    ) {
      const measured = selected.reduce<number>(
        (sum, value) => sum + (value as number),
        0,
      );
      return selected.map((value) => ((value as number) / measured) * total);
    }
  }
  return Array.from({ length: count }, () => total / count);
}

export function tableOffsets(sizes: number[]): number[] {
  const result = [0];
  for (const size of sizes) result.push((result.at(-1) as number) + size);
  return result;
}

export function boundedTableSpan(value: unknown, available: number): number {
  return typeof value === 'number' && Number.isSafeInteger(value) && value > 0
    ? Math.min(value, available)
    : 1;
}
