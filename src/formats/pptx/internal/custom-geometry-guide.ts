export type DrawingGuideResolver = (token: string) => number | undefined;

const ANGLE_UNITS_PER_DEGREE = 60_000;
const MAX_GUIDE_DEPTH = 256;

function bounded(value: number): number | undefined {
  return Number.isFinite(value) && Math.abs(value) <= Number.MAX_SAFE_INTEGER
    ? value
    : undefined;
}

function builtInValues(
  width: number,
  height: number,
): ReadonlyMap<string, number> {
  const shortSide = Math.min(width, height);
  return new Map<string, number>([
    ['3cd4', 16_200_000],
    ['3cd8', 8_100_000],
    ['5cd8', 13_500_000],
    ['7cd8', 18_900_000],
    ['b', height],
    ['cd2', 10_800_000],
    ['cd4', 5_400_000],
    ['cd8', 2_700_000],
    ['h', height],
    ['hc', width / 2],
    ['hd2', height / 2],
    ['hd3', height / 3],
    ['hd4', height / 4],
    ['hd5', height / 5],
    ['hd6', height / 6],
    ['hd8', height / 8],
    ['l', 0],
    ['ls', Math.max(width, height)],
    ['r', width],
    ['ss', shortSide],
    ['ssd2', shortSide / 2],
    ['ssd4', shortSide / 4],
    ['ssd6', shortSide / 6],
    ['ssd8', shortSide / 8],
    ['ssd16', shortSide / 16],
    ['ssd32', shortSide / 32],
    ['t', 0],
    ['vc', height / 2],
    ['w', width],
    ['wd2', width / 2],
    ['wd3', width / 3],
    ['wd4', width / 4],
    ['wd5', width / 5],
    ['wd6', width / 6],
    ['wd8', width / 8],
    ['wd10', width / 10],
    ['wd32', width / 32],
  ]);
}

function literalValue(token: string): number | undefined {
  if (!/^[+-]?\d+$/.test(token)) return undefined;
  const value = Number(token);
  return Number.isSafeInteger(value) ? value : undefined;
}

function formulaArity(operator: string): 1 | 2 | 3 | undefined {
  switch (operator) {
    case 'abs':
    case 'sqrt':
    case 'val':
      return 1;
    case 'at2':
    case 'cos':
    case 'max':
    case 'min':
    case 'sin':
    case 'tan':
      return 2;
    case '*/':
    case '+-':
    case '+/':
    case '?:':
    case 'cat2':
    case 'mod':
    case 'pin':
    case 'sat2':
      return 3;
  }
  return undefined;
}

function angleToRadians(angle: number): number {
  return (angle * Math.PI) / (180 * ANGLE_UNITS_PER_DEGREE);
}

function safeDivide(numerator: number, denominator: number): number {
  return denominator === 0 ? 0 : numerator / denominator;
}

function evaluateFormula(
  formula: string,
  resolve: DrawingGuideResolver,
): number | undefined {
  const tokens = formula.trim().split(/\s+/);
  const operator = tokens[0] as string;
  const arity = formulaArity(operator);
  if (arity === undefined || tokens.length !== arity + 1) return undefined;

  const operands = tokens.slice(1).map(resolve);
  if (operands.some((value) => value === undefined)) return undefined;
  const [x, y, z] = operands as [
    number,
    number | undefined,
    number | undefined,
  ];
  let result: number;

  switch (operator) {
    case 'abs':
      result = Math.abs(x);
      break;
    case '+/':
      result = safeDivide(x + (y ?? 0), z ?? 0);
      break;
    case '+-':
      result = x + (y ?? 0) - (z ?? 0);
      break;
    case 'at2':
      result = (Math.atan2(y ?? 0, x) * 180 * ANGLE_UNITS_PER_DEGREE) / Math.PI;
      break;
    case 'cos':
      result = x * Math.cos(angleToRadians(y ?? 0));
      break;
    case 'cat2':
      result = x * Math.cos(Math.atan2(z ?? 0, y ?? 0));
      break;
    case '?:':
      result = x > 0 ? (y ?? 0) : (z ?? 0);
      break;
    case 'max':
      result = Math.max(x, y ?? 0);
      break;
    case 'min':
      result = Math.min(x, y ?? 0);
      break;
    case 'mod':
      result = Math.hypot(x, y ?? 0, z ?? 0);
      break;
    case '*/':
      result = safeDivide(x * (y ?? 0), z ?? 0);
      break;
    case 'pin':
      result = Math.max(x, Math.min(y ?? 0, z ?? 0));
      break;
    case 'sat2':
      result = x * Math.sin(Math.atan2(z ?? 0, y ?? 0));
      break;
    case 'sin':
      result = x * Math.sin(angleToRadians(y ?? 0));
      break;
    case 'sqrt':
      result = Math.sqrt(x);
      break;
    case 'tan':
      result = x * Math.tan(angleToRadians(y ?? 0));
      break;
    case 'val':
      result = x;
      break;
    default:
      return undefined;
  }
  return bounded(result);
}

/** Build a lazy, cycle-safe evaluator for DrawingML custom geometry guides. */
export function createDrawingGuideResolver(
  width: number,
  height: number,
  formulas: ReadonlyMap<string, string> = new Map(),
): DrawingGuideResolver {
  const builtIns = builtInValues(width, height);
  const active = new Set<string>();

  const resolve: DrawingGuideResolver = (token) => {
    const literal = literalValue(token);
    if (literal !== undefined) return literal;

    const builtIn = builtIns.get(token);
    if (builtIn !== undefined) return bounded(builtIn);

    const formula = formulas.get(token);
    if (
      formula === undefined ||
      active.has(token) ||
      active.size >= MAX_GUIDE_DEPTH
    ) {
      return undefined;
    }

    active.add(token);
    const value = evaluateFormula(formula, resolve);
    active.delete(token);
    return value;
  };

  return resolve;
}
