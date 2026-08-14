import type { XmlLookupValue } from '../../../common';
import type {
  ChartItem,
  ChartType,
  ChartValue,
  ScatterChartData,
} from '../types';
import type { PptxParserContext } from './context';

import { getTextByPathList } from '../../../common';
import { applyTint } from './color';

interface ParsedChartBase {
  barDir?: 'bar' | 'col';
  colors: string[];
  grouping?: string;
  holeSize?: string;
  marker?: boolean;
  style?: string;
}

interface ParsedScatterChart extends ParsedChartBase {
  data: ScatterChartData;
  type: 'bubbleChart' | 'scatterChart';
}

interface ParsedCommonChart extends ParsedChartBase {
  data: ChartItem[];
  type: Exclude<ChartType, 'bubbleChart' | 'scatterChart'>;
}

export type ParsedChart = ParsedCommonChart | ParsedScatterChart;

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

function textAt(node: unknown, path: readonly string[]): string | undefined {
  return getTextByPathList<string>(node, path);
}

function extractChartColors(
  series: XmlLookupValue | undefined,
  context: PptxParserContext,
): string[] {
  const colors: string[] = [];

  for (const node of asArray(series)) {
    const schemeColorNode =
      nodeAt(node, ['c:spPr', 'a:solidFill', 'a:schemeClr']) ??
      nodeAt(node, ['c:spPr', 'a:ln', 'a:solidFill', 'a:schemeClr']) ??
      nodeAt(node, [
        'c:marker',
        'c:spPr',
        'a:ln',
        'a:solidFill',
        'a:schemeClr',
      ]);

    let color: string | undefined;
    const schemeName = textAt(schemeColorNode, ['attrs', 'val']);
    if (schemeName) {
      color = textAt(context.themeContent, [
        'a:theme',
        'a:themeElements',
        'a:clrScheme',
        `a:${schemeName}`,
        'a:srgbClr',
        'attrs',
        'val',
      ]);
      const tintValue = textAt(schemeColorNode, ['a:tint', 'attrs', 'val']);
      const tint = Number(tintValue) / 100_000;
      if (color && Number.isFinite(tint)) color = applyTint(color, tint);
    } else {
      color = textAt(node, [
        'c:spPr',
        'a:solidFill',
        'a:srgbClr',
        'attrs',
        'val',
      ]);
    }

    colors.push(color ? `#${color}` : '');
  }

  return colors;
}

function extractCategoryLabels(series: XmlLookupValue): Record<string, string> {
  const labels: Record<string, string> = {};
  const points =
    nodeAt(series, ['c:cat', 'c:strRef', 'c:strCache', 'c:pt']) ??
    nodeAt(series, ['c:cat', 'c:numRef', 'c:numCache', 'c:pt']);

  for (const point of asArray(points)) {
    const index = textAt(point, ['attrs', 'idx']);
    const value = textAt(point, ['c:v']);
    if (index !== undefined && value !== undefined) labels[index] = value;
  }
  return labels;
}

function extractChartData(series: XmlLookupValue | undefined): ChartItem[] {
  return asArray(series).map((item, seriesIndex) => {
    const values: ChartValue[] = [];
    const valuePoints = nodeAt(item, [
      'c:val',
      'c:numRef',
      'c:numCache',
      'c:pt',
    ]);
    for (const point of asArray(valuePoints)) {
      values.push({
        x: textAt(point, ['attrs', 'idx']) ?? '',
        y: Number.parseFloat(textAt(point, ['c:v']) ?? '0'),
      });
    }

    return {
      key:
        textAt(item, ['c:tx', 'c:strRef', 'c:strCache', 'c:pt', 'c:v']) ??
        String(seriesIndex),
      values,
      xlabels: extractCategoryLabels(item),
    };
  });
}

function extractNumericPoints(
  series: XmlLookupValue,
  axis: 'c:xVal' | 'c:yVal',
): number[] {
  const points = nodeAt(series, [axis, 'c:numRef', 'c:numCache', 'c:pt']);
  return asArray(points).map((point) =>
    Number.parseFloat(textAt(point, ['c:v']) ?? '0'),
  );
}

function extractScatterChartData(
  series: XmlLookupValue | undefined,
): ScatterChartData {
  const allSeries = asArray(series);
  const firstSeries = allSeries[0];
  if (!firstSeries) return [];

  return [
    extractNumericPoints(firstSeries, 'c:xVal'),
    ...allSeries.map((item) => extractNumericPoints(item, 'c:yVal')),
  ];
}

function chartTypeFromKey(key: string): ChartType | undefined {
  const value = key.startsWith('c:') ? key.slice(2) : key;
  const supported: readonly ChartType[] = [
    'lineChart',
    'line3DChart',
    'barChart',
    'bar3DChart',
    'pieChart',
    'pie3DChart',
    'doughnutChart',
    'areaChart',
    'area3DChart',
    'scatterChart',
    'bubbleChart',
    'radarChart',
    'surfaceChart',
    'surface3DChart',
    'stockChart',
  ];
  return supported.find((type) => type === value);
}

export function getChartInfo(
  plotArea: XmlLookupValue,
  context: PptxParserContext,
): ParsedChart | null {
  for (const key of Object.keys(plotArea)) {
    const type = chartTypeFromKey(key);
    const chartNode = plotArea[key];
    const series = nodeAt(chartNode, ['c:ser']);
    if (!type || !chartNode || !series) continue;

    const isScatter = type === 'scatterChart' || type === 'bubbleChart';
    const colorSeries =
      type === 'pieChart' || type === 'pie3DChart' || type === 'doughnutChart'
        ? nodeAt(series, ['c:dPt'])
        : series;
    const colors =
      type === 'stockChart' ? [] : extractChartColors(colorSeries, context);
    const result: ParsedChart = isScatter
      ? { type, data: extractScatterChartData(series), colors }
      : { type, data: extractChartData(series), colors };

    const grouping = textAt(chartNode, ['c:grouping', 'attrs', 'val']);
    if (grouping) result.grouping = grouping;

    const direction = textAt(chartNode, ['c:barDir', 'attrs', 'val']);
    if (direction === 'bar' || direction === 'col') result.barDir = direction;

    const holeSize = textAt(chartNode, ['c:holeSize', 'attrs', 'val']);
    if (holeSize) result.holeSize = holeSize;

    if (type === 'lineChart') result.marker = Boolean(chartNode['c:marker']);

    const style =
      textAt(chartNode, ['c:scatterStyle', 'attrs', 'val']) ??
      textAt(chartNode, ['c:radarStyle', 'attrs', 'val']);
    if (style) result.style = style;

    return result;
  }

  return null;
}
