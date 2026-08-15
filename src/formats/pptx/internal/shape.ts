import type { XmlLookupValue } from '../../../common';

import { getTextByPathList } from '../../../common';
import {
  createDrawingGuideResolver,
  type DrawingGuideResolver,
} from './custom-geometry-guide';

interface Point {
  x: number;
  y: number;
}

interface Edge {
  dx: number;
  dy: number;
  lengthSquared: number;
}

type Quadrilateral = readonly [Point, Point, Point, Point];

interface CustomPathRenderState {
  currentPoint: Point | undefined;
  subpathStart: Point | undefined;
}

type PathCommand =
  | { points: Point[]; type: 'cubicBezTo' | 'lineTo' | 'moveTo' }
  | { type: 'arcTo' | 'close' | 'quadBezTo' };

interface PathAnalysis {
  arcCount: number;
  curveCount: number;
  isCircular: boolean;
  isClosed: boolean;
  lineCount: number;
  vertices: Point[];
}

type OrderedCustomCommand =
  | { order: number; type: 'close' }
  | { order: number; point: Point; type: 'lineTo' | 'moveTo' }
  | {
      order: number;
      points: readonly [Point, Point, Point];
      type: 'cubicBezTo';
    }
  | { order: number; points: readonly [Point, Point]; type: 'quadBezTo' }
  | {
      hR: number;
      order: number;
      stAng: number;
      swAng: number;
      type: 'arcTo';
      wR: number;
    };

function asArray(value: XmlLookupValue | undefined): XmlLookupValue[] {
  if (!value) return [];
  return Array.isArray(value) ? value : [value];
}

function nodeAt(
  node: unknown,
  path: readonly string[],
): XmlLookupValue | undefined {
  return getTextByPathList<XmlLookupValue>(node, path);
}

function attributes(node: unknown): Record<string, string> {
  return getTextByPathList<Record<string, string>>(node, ['attrs']) ?? {};
}

function pointFromNode(node: XmlLookupValue): Point {
  const attrs = attributes(node);
  return {
    x: Number.parseInt(attrs.x ?? '0'),
    y: Number.parseInt(attrs.y ?? '0'),
  };
}

function pointsFromNode(node: XmlLookupValue | undefined): Point[] {
  return asArray(node).map(pointFromNode);
}

function resolvedPointFromNode(
  node: XmlLookupValue,
  resolve: DrawingGuideResolver,
): Point | undefined {
  const attrs = attributes(node);
  const x = attrs.x === undefined ? undefined : resolve(attrs.x);
  const y = attrs.y === undefined ? undefined : resolve(attrs.y);
  return x === undefined || y === undefined ? undefined : { x, y };
}

function resolvedPointsFromNode(
  node: XmlLookupValue | undefined,
  resolve: DrawingGuideResolver,
): Point[] | undefined {
  const points = asArray(node).map((point) =>
    resolvedPointFromNode(point, resolve),
  );
  return points.some((point) => point === undefined)
    ? undefined
    : (points as Point[]);
}

function commandOrder(node: XmlLookupValue | undefined): number {
  return Number(attributes(node).order ?? 0);
}

function collectPointCommands(
  path: XmlLookupValue,
  elementName: 'a:lnTo' | 'a:moveTo',
  type: 'lineTo' | 'moveTo',
  resolve: DrawingGuideResolver,
): OrderedCustomCommand[] {
  return asArray(nodeAt(path, [elementName])).flatMap((item) => {
    const point = resolvedPointsFromNode(nodeAt(item, ['a:pt']), resolve)?.[0];
    return point ? [{ type, point, order: commandOrder(item) }] : [];
  });
}

function collectBezierCommands(
  path: XmlLookupValue,
  elementName: 'a:cubicBezTo' | 'a:quadBezTo',
  type: 'cubicBezTo' | 'quadBezTo',
  resolve: DrawingGuideResolver,
): OrderedCustomCommand[] {
  return asArray(nodeAt(path, [elementName])).flatMap(
    (item): OrderedCustomCommand[] => {
      const pointNodes = nodeAt(item, ['a:pt']);
      const points = resolvedPointsFromNode(pointNodes, resolve);
      if (!points) return [];
      if (type === 'cubicBezTo') {
        return points.length === 3
          ? [
              {
                type,
                points: points as [Point, Point, Point],
                order: commandOrder(item),
              },
            ]
          : [];
      }
      return points.length === 2
        ? [
            {
              type,
              points: points as [Point, Point],
              order: commandOrder(item),
            },
          ]
        : [];
    },
  );
}

function collectArcCommands(
  path: XmlLookupValue,
  resolve: DrawingGuideResolver,
): OrderedCustomCommand[] {
  return asArray(nodeAt(path, ['a:arcTo'])).flatMap((item) => {
    const attrs = attributes(item);
    const hR = attrs.hR === undefined ? undefined : resolve(attrs.hR);
    const wR = attrs.wR === undefined ? undefined : resolve(attrs.wR);
    const stAng = attrs.stAng === undefined ? undefined : resolve(attrs.stAng);
    const swAng = attrs.swAng === undefined ? undefined : resolve(attrs.swAng);
    return hR === undefined ||
      wR === undefined ||
      stAng === undefined ||
      swAng === undefined
      ? []
      : [
          {
            type: 'arcTo',
            hR,
            wR,
            stAng,
            swAng,
            order: commandOrder(item),
          },
        ];
  });
}

function collectGuideFormulas(
  customGeometry: XmlLookupValue,
): ReadonlyMap<string, string> {
  const formulas = new Map<string, string>();
  const guideLists = [
    nodeAt(customGeometry, ['a:avLst', 'a:gd']),
    nodeAt(customGeometry, ['a:gdLst', 'a:gd']),
  ];
  for (const guideList of guideLists) {
    for (const guide of asArray(guideList)) {
      const attrs = attributes(guide);
      if (!attrs.name) continue;
      if (attrs.fmla) formulas.set(attrs.name, attrs.fmla);
      else formulas.delete(attrs.name);
    }
  }
  return formulas;
}

function collectCloseCommands(path: XmlLookupValue): OrderedCustomCommand[] {
  return asArray(nodeAt(path, ['a:close'])).map((item) => ({
    type: 'close',
    order: commandOrder(item),
  }));
}

function roundedCoordinate(value: number): number {
  return Math.round(value * 1_000_000_000_000) / 1_000_000_000_000;
}

function ellipseRayOffset(
  radiusX: number,
  radiusY: number,
  angleDegrees: number,
): Point {
  const normalizedAngle = ((angleDegrees % 360) + 360) % 360;
  let cosine: number;
  let sine: number;
  switch (normalizedAngle) {
    case 0:
      cosine = 1;
      sine = 0;
      break;
    case 90:
      cosine = 0;
      sine = 1;
      break;
    case 180:
      cosine = -1;
      sine = 0;
      break;
    case 270:
      cosine = 0;
      sine = -1;
      break;
    default: {
      const radians = (normalizedAngle * Math.PI) / 180;
      cosine = Math.cos(radians);
      sine = Math.sin(radians);
    }
  }
  const denominator = Math.hypot(radiusY * cosine, radiusX * sine);
  const scale = (radiusX * radiusY) / denominator;
  return {
    x: scale * cosine,
    y: scale * sine,
  };
}

function renderCustomArc(
  command: Extract<OrderedCustomCommand, { type: 'arcTo' }>,
  currentPoint: Point,
  scaleX: number,
  scaleY: number,
): { data: string; endpoint: Point } | undefined {
  const radiusX = command.wR * scaleX;
  const radiusY = command.hR * scaleY;
  if (
    Math.abs(radiusX) > Number.MAX_SAFE_INTEGER ||
    Math.abs(radiusY) > Number.MAX_SAFE_INTEGER ||
    radiusX <= 0 ||
    radiusY <= 0
  ) {
    return undefined;
  }

  const startDegrees = command.stAng / 60_000;
  const requestedSweep = command.swAng / 60_000;
  const sweepDegrees = Math.max(-360, Math.min(360, requestedSweep));
  if (sweepDegrees === 0) return undefined;

  const startOffset = ellipseRayOffset(radiusX, radiusY, startDegrees);
  const center = {
    x: currentPoint.x - startOffset.x,
    y: currentPoint.y - startOffset.y,
  };
  const segmentSweeps =
    Math.abs(sweepDegrees) === 360
      ? [sweepDegrees / 2, sweepDegrees / 2]
      : [sweepDegrees];
  const sweepFlag = Math.max(0, Math.sign(sweepDegrees));
  let traversed = 0;
  let data = '';
  let endpoint = currentPoint;

  for (const segmentSweep of segmentSweeps) {
    traversed += segmentSweep;
    const endOffset = ellipseRayOffset(
      radiusX,
      radiusY,
      startDegrees + traversed,
    );
    endpoint = {
      x: roundedCoordinate(center.x + endOffset.x),
      y: roundedCoordinate(center.y + endOffset.y),
    };
    const largeArcFlag = Math.abs(segmentSweep) > 180 ? 1 : 0;
    data += ` A${roundedCoordinate(radiusX)},${roundedCoordinate(radiusY)} 0 ${largeArcFlag},${sweepFlag} ${endpoint.x},${endpoint.y}`;
  }

  return { data, endpoint };
}

function renderCustomCommand(
  command: OrderedCustomCommand,
  scaleX: number,
  scaleY: number,
  state: CustomPathRenderState,
): string {
  switch (command.type) {
    case 'moveTo': {
      const point = {
        x: command.point.x * scaleX,
        y: command.point.y * scaleY,
      };
      state.currentPoint = point;
      state.subpathStart = point;
      return ` M${point.x},${point.y}`;
    }
    case 'lineTo': {
      const point = {
        x: command.point.x * scaleX,
        y: command.point.y * scaleY,
      };
      state.currentPoint = point;
      return ` L${point.x},${point.y}`;
    }
    case 'cubicBezTo': {
      const [first, second, third] = command.points;
      state.currentPoint = {
        x: third.x * scaleX,
        y: third.y * scaleY,
      };
      return ` C${first.x * scaleX},${first.y * scaleY} ${second.x * scaleX},${second.y * scaleY} ${third.x * scaleX},${third.y * scaleY}`;
    }
    case 'quadBezTo': {
      const [first, second] = command.points;
      state.currentPoint = {
        x: second.x * scaleX,
        y: second.y * scaleY,
      };
      return ` Q${first.x * scaleX},${first.y * scaleY} ${second.x * scaleX},${second.y * scaleY}`;
    }
    case 'arcTo': {
      if (!state.currentPoint) return '';
      const arc = renderCustomArc(command, state.currentPoint, scaleX, scaleY);
      if (!arc) return '';
      state.currentPoint = arc.endpoint;
      return arc.data;
    }
    case 'close':
      state.currentPoint = state.subpathStart;
      return 'z';
  }
}

function renderCustomPath(
  pathNode: XmlLookupValue,
  width: number,
  height: number,
  formulas: ReadonlyMap<string, string>,
): string {
  const pathAttributes = attributes(pathNode);
  const sourceWidth = Number(pathAttributes.w ?? 0);
  const sourceHeight = Number(pathAttributes.h ?? 0);
  if (
    !Number.isSafeInteger(sourceWidth) ||
    !Number.isSafeInteger(sourceHeight) ||
    !Number.isFinite(width) ||
    !Number.isFinite(height) ||
    Math.abs(width) > Number.MAX_SAFE_INTEGER ||
    Math.abs(height) > Number.MAX_SAFE_INTEGER ||
    sourceWidth <= 0 ||
    sourceHeight <= 0 ||
    width < 0 ||
    height < 0
  ) {
    return '';
  }
  const scaleX = width / sourceWidth;
  const scaleY = height / sourceHeight;
  const resolve = createDrawingGuideResolver(
    sourceWidth,
    sourceHeight,
    formulas,
  );
  const commands: OrderedCustomCommand[] = [
    ...collectPointCommands(pathNode, 'a:moveTo', 'moveTo', resolve),
    ...collectPointCommands(pathNode, 'a:lnTo', 'lineTo', resolve),
    ...collectBezierCommands(pathNode, 'a:cubicBezTo', 'cubicBezTo', resolve),
    ...collectBezierCommands(pathNode, 'a:quadBezTo', 'quadBezTo', resolve),
    ...collectArcCommands(pathNode, resolve),
    ...collectCloseCommands(pathNode),
  ];
  const state: CustomPathRenderState = {
    currentPoint: undefined,
    subpathStart: undefined,
  };

  return commands
    .sort((left, right) => left.order - right.order)
    .map((command) => renderCustomCommand(command, scaleX, scaleY, state))
    .join('');
}

export function getCustomShapePath(
  customGeometry: XmlLookupValue,
  width: number,
  height: number,
): string {
  const formulas = collectGuideFormulas(customGeometry);
  return asArray(nodeAt(customGeometry, ['a:pathLst', 'a:path']))
    .map((path) => renderCustomPath(path, width, height, formulas))
    .join('');
}

export function isStrokeOnlyCustomGeometry(
  customGeometry: XmlLookupValue,
): boolean {
  const paths = asArray(nodeAt(customGeometry, ['a:pathLst', 'a:path']));
  return (
    paths.length > 0 && paths.every((path) => attributes(path).fill === 'none')
  );
}

function extractPathCommands(path: XmlLookupValue): PathCommand[] {
  const commands: PathCommand[] = [];
  const moveTo = nodeAt(path, ['a:moveTo', 'a:pt']);
  if (moveTo)
    commands.push({ type: 'moveTo', points: [pointFromNode(moveTo)] });

  for (const line of asArray(nodeAt(path, ['a:lnTo']))) {
    const point = nodeAt(line, ['a:pt']);
    if (point)
      commands.push({ type: 'lineTo', points: [pointFromNode(point)] });
  }
  for (const cubic of asArray(nodeAt(path, ['a:cubicBezTo']))) {
    const points = pointsFromNode(nodeAt(cubic, ['a:pt']));
    if (points.length === 3) commands.push({ type: 'cubicBezTo', points });
  }
  for (const quadratic of asArray(nodeAt(path, ['a:quadBezTo']))) {
    const points = pointsFromNode(nodeAt(quadratic, ['a:pt']));
    if (points.length === 2) commands.push({ type: 'quadBezTo' });
  }
  asArray(nodeAt(path, ['a:arcTo'])).forEach(() => {
    commands.push({ type: 'arcTo' });
  });
  if (nodeAt(path, ['a:close'])) commands.push({ type: 'close' });
  return commands;
}

function analyzePathCommands(
  commands: PathCommand[],
  pathWidth: number,
  pathHeight: number,
): PathAnalysis {
  const analysis: PathAnalysis = {
    lineCount: 0,
    curveCount: 0,
    arcCount: 0,
    isClosed: false,
    vertices: [],
    isCircular: false,
  };

  for (const command of commands) {
    switch (command.type) {
      case 'moveTo':
      case 'lineTo': {
        if (command.type === 'lineTo') analysis.lineCount += 1;
        const point = command.points[0];
        if (point) analysis.vertices.push(point);
        break;
      }
      case 'cubicBezTo':
      case 'quadBezTo': {
        analysis.curveCount += 1;
        break;
      }
      case 'arcTo':
        analysis.arcCount += 1;
        break;
      case 'close':
        analysis.isClosed = true;
        break;
    }
  }

  if (
    analysis.curveCount === 4 &&
    analysis.lineCount === 0 &&
    analysis.isClosed
  ) {
    analysis.isCircular = checkIfCircular(commands, pathWidth, pathHeight);
  }
  return analysis;
}

function checkIfCircular(
  commands: PathCommand[],
  width: number,
  height: number,
): boolean {
  const endpoints = commands
    .map((command) =>
      command.type === 'cubicBezTo' ? command.points[2] : undefined,
    )
    .filter((point): point is Point => point !== undefined);
  const hasTop = endpoints.some((point) => Math.abs(point.y) < height * 0.1);
  const hasBottom = endpoints.some(
    (point) => Math.abs(point.y - height) < height * 0.1,
  );
  const hasLeft = endpoints.some((point) => Math.abs(point.x) < width * 0.1);
  const hasRight = endpoints.some(
    (point) => Math.abs(point.x - width) < width * 0.1,
  );
  return hasTop && hasBottom && hasLeft && hasRight;
}

function removeDuplicateVertices(vertices: Point[]): Point[] {
  const unique: Point[] = [];
  for (const vertex of vertices) {
    const duplicate = unique.some(
      (item) => item.x === vertex.x && item.y === vertex.y,
    );
    if (!duplicate) unique.push(vertex);
  }
  return unique;
}

function edgeBetween(first: Point, second: Point): Edge {
  const dx = second.x - first.x;
  const dy = second.y - first.y;
  return { dx, dy, lengthSquared: dx * dx + dy * dy };
}

function isRectangle(vertices: Quadrilateral): boolean {
  const [first, second, third, fourth] = vertices;
  return (
    (first.y === second.y &&
      second.x === third.x &&
      third.y === fourth.y &&
      fourth.x === first.x) ||
    (first.x === second.x &&
      second.y === third.y &&
      third.x === fourth.x &&
      fourth.y === first.y)
  );
}

type QuadrilateralEdges = readonly [Edge, Edge, Edge, Edge];

function isRhombus(edges: QuadrilateralEdges): boolean {
  const [first] = edges;
  return edges.every((edge) => edge.lengthSquared === first.lengthSquared);
}

function areParallel(first: Edge, second: Edge): boolean {
  return first.dx * second.dy === first.dy * second.dx;
}

function matchQuadrilateral(vertices: Quadrilateral): string {
  const [firstPoint, secondPoint, thirdPoint, fourthPoint] = vertices;
  const edges: QuadrilateralEdges = [
    edgeBetween(firstPoint, secondPoint),
    edgeBetween(secondPoint, thirdPoint),
    edgeBetween(thirdPoint, fourthPoint),
    edgeBetween(fourthPoint, firstPoint),
  ];
  const [first, second, third, fourth] = edges;
  if (isRectangle(vertices)) return 'rect';
  if (isRhombus(edges)) return 'rhombus';

  const parallel02 = areParallel(first, third);
  const parallel13 = areParallel(second, fourth);
  if (parallel02 && parallel13) return 'parallelogram';
  if (parallel02 !== parallel13) return 'trapezoid';
  return 'custom';
}

function matchPolygon(vertices: Point[]): string {
  const unique = removeDuplicateVertices(vertices);
  switch (unique.length) {
    case 3:
      return 'triangle';
    case 4:
      return matchQuadrilateral(unique as unknown as Quadrilateral);
    case 5:
      return 'pentagon';
    case 6:
      return 'hexagon';
    case 7:
      return 'heptagon';
    case 8:
      return 'octagon';
    case 10:
      return 'decagon';
    case 12:
      return 'dodecagon';
    default:
      return 'custom';
  }
}

function matchPolygonByLineCount(lineCount: number): string {
  return (
    {
      3: 'triangle',
      4: 'rectangle',
      5: 'pentagon',
      6: 'hexagon',
      7: 'heptagon',
      8: 'octagon',
    }[lineCount] ?? 'custom'
  );
}

function matchShape(analysis: PathAnalysis): string {
  if (analysis.isCircular) return 'ellipse';
  if (analysis.arcCount >= 2 && analysis.isClosed && analysis.lineCount === 0)
    return 'ellipse';
  if (analysis.curveCount === 0 && analysis.arcCount === 0 && analysis.isClosed)
    return matchPolygon(analysis.vertices);
  if (
    analysis.curveCount > 0 &&
    analysis.curveCount <= analysis.lineCount &&
    analysis.isClosed
  ) {
    const base = matchPolygonByLineCount(analysis.lineCount);
    return base === 'rectangle' ? 'roundRect' : base;
  }
  return 'custom';
}

export function identifyShape(shapeData: XmlLookupValue): string {
  const path = asArray(nodeAt(shapeData, ['a:pathLst', 'a:path']))[0];
  if (!path) return 'custom';
  const attrs = attributes(path);
  const commands = extractPathCommands(path);
  return matchShape(
    analyzePathCommands(commands, Number(attrs.w), Number(attrs.h)),
  );
}
