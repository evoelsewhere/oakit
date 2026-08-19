import type {
  PptxSceneChartElement,
  PptxSceneChartSeries,
  PptxSceneTransform,
} from '../scene-types';
import { serializeGraphicFrameTransform } from './shape';
import { escapeXmlAttribute, escapeXmlText } from './xml';

const CHART_NAMESPACE =
  'http://schemas.openxmlformats.org/drawingml/2006/chart';
const CHART_GRAPHIC_URI =
  'http://schemas.openxmlformats.org/drawingml/2006/chart';
const DRAWING_NAMESPACE =
  'http://schemas.openxmlformats.org/drawingml/2006/main';
const RELATIONSHIPS_NAMESPACE =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

function booleanAttribute(value: boolean): string {
  return value ? '1' : '0';
}

function spreadsheetColumn(index: number): string {
  let value = index;
  let result = '';
  while (value > 0) {
    value -= 1;
    result = String.fromCharCode(65 + (value % 26)) + result;
    value = Math.floor(value / 26);
  }
  return result;
}

function serializeStringCache(values: readonly string[]): string {
  const points = values
    .map(
      (value, index) =>
        `<c:pt idx="${index}"><c:v>${escapeXmlText(value)}</c:v></c:pt>`,
    )
    .join('');
  return `<c:strCache><c:ptCount val="${values.length}"/>${points}</c:strCache>`;
}

function serializeNumberCache(values: readonly number[]): string {
  const points = values
    .map((value, index) => `<c:pt idx="${index}"><c:v>${value}</c:v></c:pt>`)
    .join('');
  return `<c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="${values.length}"/>${points}</c:numCache>`;
}

function serializeSeriesColor(color: string | undefined): string {
  if (color === undefined) return '';
  return `<c:spPr><a:solidFill><a:srgbClr val="${color.slice(1).toUpperCase()}"/></a:solidFill><a:ln><a:solidFill><a:srgbClr val="${color.slice(1).toUpperCase()}"/></a:solidFill></a:ln></c:spPr>`;
}

function serializeSeries(
  series: PptxSceneChartSeries,
  seriesIndex: number,
  chartType: PptxSceneChartElement['chartType'],
  marker: boolean,
): string {
  const valueColumn = spreadsheetColumn(seriesIndex + 2);
  const lastRow = series.values.length + 1;
  const title =
    `<c:tx><c:strRef><c:f>Sheet1!$${valueColumn}$1</c:f>` +
    `${serializeStringCache([series.name])}</c:strRef></c:tx>`;
  const categories =
    `<c:cat><c:strRef><c:f>Sheet1!$A$2:$A$${lastRow}</c:f>` +
    `${serializeStringCache(series.categories)}</c:strRef></c:cat>`;
  const values =
    `<c:val><c:numRef><c:f>Sheet1!$${valueColumn}$2:$${valueColumn}$${lastRow}</c:f>` +
    `${serializeNumberCache(series.values)}</c:numRef></c:val>`;
  const markerXml =
    chartType === 'lineChart'
      ? `<c:marker><c:symbol val="${marker ? 'circle' : 'none'}"/><c:size val="5"/></c:marker>`
      : '';
  return `<c:ser><c:idx val="${seriesIndex}"/><c:order val="${seriesIndex}"/>${title}${serializeSeriesColor(series.color)}${markerXml}${categories}${values}<c:smooth val="0"/></c:ser>`;
}

function serializeAxes(categoryAxisId: number, valueAxisId: number): string {
  return (
    `<c:catAx><c:axId val="${categoryAxisId}"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/><c:tickLblPos val="nextTo"/><c:crossAx val="${valueAxisId}"/><c:crosses val="autoZero"/><c:auto val="1"/><c:lblAlgn val="ctr"/><c:lblOffset val="100"/></c:catAx>` +
    `<c:valAx><c:axId val="${valueAxisId}"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="l"/><c:majorGridlines/><c:numFmt formatCode="General" sourceLinked="1"/><c:tickLblPos val="nextTo"/><c:crossAx val="${categoryAxisId}"/><c:crosses val="autoZero"/><c:crossBetween val="between"/></c:valAx>`
  );
}

function serializePlot(
  element: PptxSceneChartElement,
  chartNumber: number,
): string {
  const marker = element.marker ?? false;
  const series = element.series
    .map((item, index) =>
      serializeSeries(item, index, element.chartType, marker),
    )
    .join('');
  if (
    element.chartType === 'pieChart' ||
    element.chartType === 'doughnutChart'
  ) {
    const hole =
      element.chartType === 'doughnutChart'
        ? `<c:holeSize val="${element.holeSize ?? 50}"/>`
        : '';
    return `<c:${element.chartType}><c:varyColors val="1"/>${series}<c:firstSliceAng val="0"/>${hole}</c:${element.chartType}>`;
  }
  const categoryAxisId = 10_000_000 + chartNumber * 2;
  const valueAxisId = categoryAxisId + 1;
  if (element.chartType === 'barChart') {
    const direction = element.barDirection ?? 'col';
    const grouping = element.grouping ?? 'clustered';
    return `<c:barChart><c:barDir val="${direction}"/><c:grouping val="${grouping}"/><c:varyColors val="0"/>${series}<c:gapWidth val="150"/><c:axId val="${categoryAxisId}"/><c:axId val="${valueAxisId}"/></c:barChart>${serializeAxes(categoryAxisId, valueAxisId)}`;
  }
  const grouping = element.grouping ?? 'standard';
  return `<c:lineChart><c:grouping val="${grouping}"/><c:varyColors val="0"/>${series}<c:marker val="${booleanAttribute(marker)}"/><c:smooth val="0"/><c:axId val="${categoryAxisId}"/><c:axId val="${valueAxisId}"/></c:lineChart>${serializeAxes(categoryAxisId, valueAxisId)}`;
}

export function serializeChartPart(
  element: PptxSceneChartElement,
  chartNumber: number,
): string {
  const plot = serializePlot(element, chartNumber);
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><c:chartSpace xmlns:c="${CHART_NAMESPACE}" xmlns:a="${DRAWING_NAMESPACE}" xmlns:r="${RELATIONSHIPS_NAMESPACE}"><c:date1904 val="0"/><c:lang val="en-US"/><c:roundedCorners val="0"/><c:chart><c:autoTitleDeleted val="1"/><c:plotArea><c:layout/>${plot}</c:plotArea><c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/><c:showDLblsOverMax val="0"/></c:chart><c:printSettings><c:headerFooter/><c:pageMargins b="0.75" l="0.7" r="0.7" t="0.75" header="0.3" footer="0.3"/><c:pageSetup/></c:printSettings></c:chartSpace>`;
}

export function serializeChartFrame(
  element: PptxSceneChartElement,
  transform: PptxSceneTransform,
  shapeId: number,
  relationshipId: string,
): string {
  const attributes = [
    `id="${shapeId}"`,
    `name="${escapeXmlAttribute(element.name ?? `Chart ${shapeId}`)}"`,
  ];
  if (element.description !== undefined) {
    attributes.push(`descr="${escapeXmlAttribute(element.description)}"`);
  }
  if (element.title !== undefined) {
    attributes.push(`title="${escapeXmlAttribute(element.title)}"`);
  }
  if (element.authored.hidden !== undefined) {
    attributes.push(`hidden="${booleanAttribute(element.authored.hidden)}"`);
  }
  const nonVisual = `<p:nvGraphicFramePr><p:cNvPr ${attributes.join(' ')}/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>`;
  const graphic = `<a:graphic><a:graphicData uri="${CHART_GRAPHIC_URI}"><c:chart xmlns:c="${CHART_NAMESPACE}" xmlns:r="${RELATIONSHIPS_NAMESPACE}" r:id="${escapeXmlAttribute(relationshipId)}"/></a:graphicData></a:graphic>`;
  return `<p:graphicFrame>${nonVisual}${serializeGraphicFrameTransform(transform)}${graphic}</p:graphicFrame>`;
}
