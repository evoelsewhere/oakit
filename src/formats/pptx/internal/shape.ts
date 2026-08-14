import type { XmlLookupValue } from '../../../common';

import { getTextByPathList } from '../../../common';

interface Point {
  x: number;
  y: number;
}

interface Edge {
  dx: number;
  dy: number;
  length: number;
}

type PathCommand =
  | { points: Point[]; type: 'cubicBezTo' | 'lineTo' | 'moveTo' | 'quadBezTo' }
  | {
      hR: number;
      stAng: number;
      swAng: number;
      type: 'arcTo';
      wR: number;
    }
  | { type: 'close' };

interface PathAnalysis {
  arcCount: number;
  commands: PathCommand[];
  curveCount: number;
  hasCurves: boolean;
  isCircular: boolean;
  isClosed: boolean;
  lineCount: number;
  pathHeight: number;
  pathWidth: number;
  vertices: Point[];
}

type OrderedCustomCommand =
  | { order: number; type: 'close' }
  | { order: number; point: Point; type: 'lineTo' | 'moveTo' }
  | { order: number; points: Point[]; type: 'cubicBezTo' | 'quadBezTo' }
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

export function shapeArc(
  centerX: number,
  centerY: number,
  radiusX: number,
  radiusY: number,
  startAngle: number,
  endAngle: number,
  close: boolean,
): string {
  let path = '';
  const increment = endAngle >= startAngle ? 1 : -1;
  for (
    let angle = startAngle;
    increment > 0 ? angle <= endAngle : angle > endAngle;
    angle += increment
  ) {
    const radians = angle * (Math.PI / 180);
    const x = centerX + Math.cos(radians) * radiusX;
    const y = centerY + Math.sin(radians) * radiusY;
    path += angle === startAngle ? ` M${x} ${y}` : ` L${x} ${y}`;
  }
  return `${path}${close ? ' z' : ''}`;
}

function commandOrder(node: XmlLookupValue | undefined): number {
  return Number(attributes(node).order ?? 0);
}

function collectPointCommands(
  path: XmlLookupValue,
  elementName: 'a:lnTo' | 'a:moveTo',
  type: 'lineTo' | 'moveTo',
): OrderedCustomCommand[] {
  return asArray(nodeAt(path, [elementName])).flatMap((item) => {
    const points = pointsFromNode(nodeAt(item, ['a:pt']));
    return points.map((point) => ({
      type,
      point,
      order: commandOrder(nodeAt(item, ['a:pt']) ?? item),
    }));
  });
}

function collectBezierCommands(
  path: XmlLookupValue,
  elementName: 'a:cubicBezTo' | 'a:quadBezTo',
  type: 'cubicBezTo' | 'quadBezTo',
): OrderedCustomCommand[] {
  return asArray(nodeAt(path, [elementName])).map((item) => {
    const pointNodes = nodeAt(item, ['a:pt']);
    return {
      type,
      points: pointsFromNode(pointNodes),
      order: commandOrder(asArray(pointNodes)[0] ?? item),
    };
  });
}

function collectArcCommands(path: XmlLookupValue): OrderedCustomCommand[] {
  return asArray(nodeAt(path, ['a:arcTo'])).map((item) => {
    const attrs = attributes(item);
    return {
      type: 'arcTo',
      hR: Number(attrs.hR ?? 0),
      wR: Number(attrs.wR ?? 0),
      stAng: Number(attrs.stAng ?? 0),
      swAng: Number(attrs.swAng ?? 0),
      order: Number(attrs.order ?? 0),
    };
  });
}

function renderCustomCommand(
  command: OrderedCustomCommand,
  scaleX: number,
  scaleY: number,
): string {
  switch (command.type) {
    case 'moveTo':
      return ` M${command.point.x * scaleX},${command.point.y * scaleY}`;
    case 'lineTo':
      return ` L${command.point.x * scaleX},${command.point.y * scaleY}`;
    case 'cubicBezTo': {
      const [first, second, third] = command.points;
      if (!first || !second || !third) return '';
      return ` C${first.x * scaleX},${first.y * scaleY} ${second.x * scaleX},${second.y * scaleY} ${third.x * scaleX},${third.y * scaleY}`;
    }
    case 'quadBezTo': {
      const [first, second] = command.points;
      if (!first || !second) return '';
      return ` Q${first.x * scaleX},${first.y * scaleY} ${second.x * scaleX},${second.y * scaleY}`;
    }
    case 'arcTo': {
      const radiusX = command.wR * scaleX;
      const radiusY = command.hR * scaleY;
      const start = command.stAng / 60_000;
      return shapeArc(
        radiusX,
        radiusY,
        radiusX,
        radiusY,
        start,
        start + command.swAng / 60_000,
        false,
      );
    }
    case 'close':
      return 'z';
  }
}

export function getCustomShapePath(
  customGeometry: XmlLookupValue,
  width: number,
  height: number,
): string {
  const pathNode = asArray(nodeAt(customGeometry, ['a:pathLst', 'a:path']))[0];
  if (!pathNode) return '';

  const pathAttributes = attributes(pathNode);
  const sourceWidth = Number(pathAttributes.w ?? 0);
  const sourceHeight = Number(pathAttributes.h ?? 0);
  const scaleX = sourceWidth === 0 ? 0 : width / sourceWidth;
  const scaleY = sourceHeight === 0 ? 0 : height / sourceHeight;
  const commands: OrderedCustomCommand[] = [
    ...collectPointCommands(pathNode, 'a:moveTo', 'moveTo'),
    ...collectPointCommands(pathNode, 'a:lnTo', 'lineTo'),
    ...collectBezierCommands(pathNode, 'a:cubicBezTo', 'cubicBezTo'),
    ...collectBezierCommands(pathNode, 'a:quadBezTo', 'quadBezTo'),
    ...collectArcCommands(pathNode),
  ];
  if (nodeAt(pathNode, ['a:close'])) {
    commands.push({ type: 'close', order: Number.POSITIVE_INFINITY });
  }

  return commands
    .sort((left, right) => left.order - right.order)
    .map((command) => renderCustomCommand(command, scaleX, scaleY))
    .join('');
}

export function isStrokeOnlyCustomGeometry(
  customGeometry: XmlLookupValue,
): boolean {
  const paths = asArray(nodeAt(customGeometry, ['a:pathLst', 'a:path']));
  return paths.length === 1 && attributes(paths[0]).fill === 'none';
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
    commands.push({
      type: 'quadBezTo',
      points: pointsFromNode(nodeAt(quadratic, ['a:pt'])),
    });
  }
  for (const arc of asArray(nodeAt(path, ['a:arcTo']))) {
    const attrs = attributes(arc);
    commands.push({
      type: 'arcTo',
      wR: Number(attrs.wR ?? 0),
      hR: Number(attrs.hR ?? 0),
      stAng: Number(attrs.stAng ?? 0),
      swAng: Number(attrs.swAng ?? 0),
    });
  }
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
    pathWidth,
    pathHeight,
    hasCurves: false,
    isCircular: false,
    commands,
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
        analysis.hasCurves = true;
        const point = command.points.at(-1);
        if (point) analysis.vertices.push(point);
        break;
      }
      case 'arcTo':
        analysis.arcCount += 1;
        analysis.hasCurves = true;
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
  if (endpoints.length !== 4) return false;

  const hasTop = endpoints.some((point) => Math.abs(point.y) < height * 0.1);
  const hasBottom = endpoints.some(
    (point) => Math.abs(point.y - height) < height * 0.1,
  );
  const hasLeft = endpoints.some((point) => Math.abs(point.x) < width * 0.1);
  const hasRight = endpoints.some(
    (point) => Math.abs(point.x - width) < width * 0.1,
  );
  return (hasTop || hasBottom) && (hasLeft || hasRight);
}

function removeDuplicateVertices(vertices: Point[]): Point[] {
  const unique: Point[] = [];
  for (const vertex of vertices) {
    const duplicate = unique.some(
      (item) =>
        Math.abs(item.x - vertex.x) < 100 && Math.abs(item.y - vertex.y) < 100,
    );
    if (!duplicate) unique.push(vertex);
  }
  return unique;
}

function edgesFromVertices(vertices: Point[]): Edge[] {
  return vertices.map((point, index) => {
    const next = vertices[(index + 1) % vertices.length] ?? point;
    const dx = next.x - point.x;
    const dy = next.y - point.y;
    return { dx, dy, length: Math.hypot(dx, dy) };
  });
}

function isRectangle(edges: Edge[]): boolean {
  const [first, second, third, fourth] = edges;
  if (!first || !second || !third || !fourth) return false;
  const oppositeEdgesMatch =
    Math.abs(first.length - third.length) /
      Math.max(first.length, third.length) <
      0.1 &&
    Math.abs(second.length - fourth.length) /
      Math.max(second.length, fourth.length) <
      0.1;
  if (!oppositeEdgesMatch) return false;

  return edges.every((edge, index) => {
    const next = edges[(index + 1) % edges.length] ?? edge;
    const cosine =
      (edge.dx * next.dx + edge.dy * next.dy) / (edge.length * next.length);
    return Math.abs(cosine) <= 0.1;
  });
}

function isRhombus(edges: Edge[]): boolean {
  const average = edges.reduce((sum, edge) => sum + edge.length, 0) / 4;
  return edges.every((edge) => Math.abs(edge.length - average) / average < 0.1);
}

function slope(edge: Edge): number {
  return edge.dx === 0 ? Number.POSITIVE_INFINITY : edge.dy / edge.dx;
}

function areParallel(first: Edge, second: Edge): boolean {
  const firstSlope = slope(first);
  const secondSlope = slope(second);
  return (
    Math.abs(firstSlope - secondSlope) < 0.15 ||
    (Math.abs(firstSlope) > 1000 && Math.abs(secondSlope) > 1000)
  );
}

function matchQuadrilateral(vertices: Point[]): string {
  const edges = edgesFromVertices(vertices);
  const [first, second, third, fourth] = edges;
  if (!first || !second || !third || !fourth) return 'custom';
  if (isRectangle(edges)) return 'roundRect';
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
      return matchQuadrilateral(unique);
    case 5:
      return 'pentagon';
    case 6:
      return 'hexagon';
    case 7:
      return 'heptagon';
    case 8:
      return 'octagon';
    default:
      return unique.length > 8 ? 'ellipse' : 'custom';
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
  if (!analysis.hasCurves && analysis.isClosed && analysis.vertices.length >= 3)
    return matchPolygon(analysis.vertices);
  if (
    analysis.lineCount === 4 &&
    analysis.curveCount === 4 &&
    analysis.isClosed
  )
    return 'roundRect';
  if (
    analysis.lineCount >= 3 &&
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
  if (commands.length === 0) return 'custom';
  return matchShape(
    analyzePathCommands(
      commands,
      Number.parseInt(attrs.w ?? '0'),
      Number.parseInt(attrs.h ?? '0'),
    ),
  );
}
