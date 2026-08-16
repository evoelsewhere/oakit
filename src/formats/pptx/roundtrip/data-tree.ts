import type { ResolvedPptxResourceLimits } from '../internal/resource-limits';
import { PptxWriteError } from '../write-error';

const MAX_SNAPSHOT_DEPTH = 64;

type SnapshotLocation = 'other' | 'root' | 'source' | 'source-data';

interface PendingValue {
  depth: number;
  location: SnapshotLocation;
  value: unknown;
}

function invalidDataTree(message: string): never {
  throw new PptxWriteError('invalid-snapshot', message);
}

function assertNoSymbolKeys(value: object): void {
  if (Object.getOwnPropertySymbols(value).length !== 0) {
    invalidDataTree('PowerPoint round-trip snapshot contains a symbol key');
  }
}

function dataProperty(
  value: object,
  name: string,
  arrayProperty: boolean,
): unknown {
  const descriptor = Object.getOwnPropertyDescriptor(value, name);
  if (
    descriptor === undefined ||
    !('value' in descriptor) ||
    !descriptor.enumerable
  ) {
    invalidDataTree(
      arrayProperty
        ? 'PowerPoint round-trip snapshot contains a sparse or accessor array'
        : 'PowerPoint round-trip snapshot contains an accessor or hidden property',
    );
  }
  return descriptor.value;
}

function arrayChildren(
  value: unknown[],
  depth: number,
  maxContainers: number,
): PendingValue[] {
  if (Object.getPrototypeOf(value) !== Array.prototype) {
    invalidDataTree('PowerPoint round-trip snapshot requires plain arrays');
  }
  assertNoSymbolKeys(value);
  if (value.length > maxContainers) {
    invalidDataTree(
      'PowerPoint round-trip snapshot exceeds the container budget',
    );
  }
  if (Object.getOwnPropertyNames(value).length !== value.length + 1) {
    invalidDataTree(
      'PowerPoint round-trip snapshot contains a sparse or extended array',
    );
  }

  const children: PendingValue[] = [];
  for (let index = 0; index < value.length; index += 1) {
    children.push({
      depth: depth + 1,
      location: 'other',
      value: dataProperty(value, String(index), true),
    });
  }
  return children;
}

function objectChildren(
  value: object,
  location: SnapshotLocation,
  depth: number,
  consumeString: (length: number) => void,
): PendingValue[] {
  if (Object.getPrototypeOf(value) !== Object.prototype) {
    invalidDataTree('PowerPoint round-trip snapshot requires plain objects');
  }
  assertNoSymbolKeys(value);

  const children: PendingValue[] = [];
  for (const name of Object.getOwnPropertyNames(value)) {
    consumeString(name.length);
    children.push({
      depth: depth + 1,
      location:
        location === 'root' && name === 'source'
          ? 'source'
          : location === 'source' && name === 'data'
            ? 'source-data'
            : 'other',
      value: dataProperty(value, name, false),
    });
  }
  return children;
}

export function assertPptxRoundTripDataTree(
  value: unknown,
  limits: ResolvedPptxResourceLimits,
): void {
  const pending: PendingValue[] = [{ depth: 0, location: 'root', value }];
  const seen = new WeakSet<object>();
  let containers = 0;
  let stringCodeUnits = 0;
  const consumeString = (length: number): void => {
    stringCodeUnits += length;
    if (
      !Number.isSafeInteger(stringCodeUnits) ||
      stringCodeUnits > limits.maxTotalUncompressedBytes
    ) {
      invalidDataTree(
        'PowerPoint round-trip snapshot exceeds the string budget',
      );
    }
  };

  for (const current of pending) {
    if (current.depth > MAX_SNAPSHOT_DEPTH) {
      invalidDataTree(
        'PowerPoint round-trip snapshot exceeds the maximum data depth',
      );
    }
    if (typeof current.value === 'string') {
      consumeString(current.value.length);
      continue;
    }
    if (current.value === null || typeof current.value === 'boolean') {
      continue;
    }
    if (typeof current.value === 'number') {
      if (Number.isFinite(current.value)) continue;
      invalidDataTree(
        'PowerPoint round-trip snapshot contains a non-JSON value',
      );
    }
    if (typeof current.value !== 'object') {
      invalidDataTree(
        'PowerPoint round-trip snapshot contains a non-JSON value',
      );
    }
    if (current.value instanceof Uint8Array || current.value instanceof Blob) {
      if (current.location !== 'source-data') {
        invalidDataTree(
          'PowerPoint round-trip snapshot contains binary data outside its source',
        );
      }
      continue;
    }
    if (seen.has(current.value)) {
      invalidDataTree(
        'PowerPoint round-trip snapshot contains a repeated object reference',
      );
    }
    seen.add(current.value);
    containers += 1;
    if (containers > limits.maxTotalXmlNodes) {
      invalidDataTree(
        'PowerPoint round-trip snapshot exceeds the container budget',
      );
    }

    const children = Array.isArray(current.value)
      ? arrayChildren(current.value, current.depth, limits.maxTotalXmlNodes)
      : objectChildren(
          current.value,
          current.location,
          current.depth,
          consumeString,
        );
    pending.push(...children);
  }
}
