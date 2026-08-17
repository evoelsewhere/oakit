function finiteNumber(value, description) {
  if (typeof value !== 'number' || !Number.isFinite(value)) {
    throw new TypeError(`${description} must be finite`);
  }
  return value;
}

function boolean(value, description) {
  if (typeof value !== 'boolean') {
    throw new TypeError(`${description} must be boolean`);
  }
  return value;
}

export function pptxTransformMatrix(value) {
  if (value === null || typeof value !== 'object' || Array.isArray(value)) {
    throw new TypeError('PowerPoint transform must be an object');
  }
  const radians =
    (finiteNumber(value.rotation, 'PowerPoint rotation') * Math.PI) / 180;
  const scaleX = boolean(value.flipHorizontal, 'PowerPoint horizontal flip')
    ? -1
    : 1;
  const scaleY = boolean(value.flipVertical, 'PowerPoint vertical flip')
    ? -1
    : 1;
  const cosine = Math.cos(radians);
  const sine = Math.sin(radians);
  return [cosine * scaleX, -sine * scaleY, sine * scaleX, cosine * scaleY];
}

export function pptxTransformsAreEquivalent(left, right, tolerance = 1e-10) {
  finiteNumber(tolerance, 'PowerPoint transform tolerance');
  if (tolerance <= 0) {
    throw new RangeError('PowerPoint transform tolerance must be positive');
  }
  const leftMatrix = pptxTransformMatrix(left);
  const rightMatrix = pptxTransformMatrix(right);
  return leftMatrix.every(
    (value, index) => Math.abs(value - rightMatrix[index]) <= tolerance,
  );
}
