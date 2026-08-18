import type {
  PptxSceneElementBase,
  PptxSceneTableBorder,
  PptxSceneTableBorders,
  PptxSceneTableCell,
  PptxSceneTableElement,
  PptxSceneTransform,
} from '../scene-types';
import { serializeSolidColorFill } from './color';
import { serializeGraphicFrameTransform } from './shape';
import { serializeDrawingTextBody } from './text-body';
import type { PptxTextSerializationContext } from './text-node';
import { pointsToEmu } from './units';
import { escapeXmlAttribute } from './xml';

const TABLE_URI = 'http://schemas.openxmlformats.org/drawingml/2006/table';

const BORDER_TAGS: ReadonlyArray<
  readonly [keyof PptxSceneTableBorders, string]
> = [
  ['left', 'a:lnL'],
  ['right', 'a:lnR'],
  ['top', 'a:lnT'],
  ['bottom', 'a:lnB'],
];

const DASH_VALUES: Record<
  NonNullable<PptxSceneTableBorder['style']>,
  string
> = {
  dashed: 'dash',
  dotted: 'dot',
  solid: 'solid',
};

const CELL_ANCHORS: Record<
  NonNullable<PptxSceneTableCell['text']['body']['anchor']>,
  string
> = {
  bottom: 'b',
  center: 'ctr',
  distributed: 'dist',
  justified: 'just',
  top: 't',
};

function booleanAttribute(value: boolean): string {
  return value ? '1' : '0';
}

function serializeNonVisualProperties(
  element: PptxSceneElementBase,
  shapeId: number,
): string {
  const attributes = [
    `id="${shapeId}"`,
    `name="${escapeXmlAttribute(element.name ?? `Table ${shapeId}`)}"`,
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
  return `<p:nvGraphicFramePr><p:cNvPr ${attributes.join(' ')}/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>`;
}

function serializeBorder(
  tagName: string,
  border: PptxSceneTableBorder,
): string {
  const dash = DASH_VALUES[border.style ?? 'solid'];
  return `<${tagName} w="${pointsToEmu(border.width)}">${serializeSolidColorFill(border.color)}<a:prstDash val="${dash}"/></${tagName}>`;
}

function serializeCellProperties(cell: PptxSceneTableCell): string {
  const anchor = cell.text.body.anchor;
  const attributes =
    anchor === undefined ? '' : ` anchor="${CELL_ANCHORS[anchor]}"`;
  const children: string[] = [];
  for (const [key, tagName] of BORDER_TAGS) {
    const border = cell.borders?.[key];
    if (border !== undefined) children.push(serializeBorder(tagName, border));
  }
  if (cell.fillColor !== undefined) {
    children.push(serializeSolidColorFill(cell.fillColor));
  }
  if (children.length === 0) return `<a:tcPr${attributes}/>`;
  return `<a:tcPr${attributes}>${children.join('')}</a:tcPr>`;
}

function serializeCell(
  cell: PptxSceneTableCell,
  context: PptxTextSerializationContext,
): string {
  const attributes: string[] = [];
  if (cell.rowSpan !== undefined) attributes.push(`rowSpan="${cell.rowSpan}"`);
  if (cell.colSpan !== undefined) attributes.push(`gridSpan="${cell.colSpan}"`);
  if (cell.vMerge !== undefined) {
    attributes.push(`vMerge="${booleanAttribute(cell.vMerge)}"`);
  }
  if (cell.hMerge !== undefined) {
    attributes.push(`hMerge="${booleanAttribute(cell.hMerge)}"`);
  }
  const attributeText =
    attributes.length === 0 ? '' : ` ${attributes.join(' ')}`;
  return `<a:tc${attributeText}>${serializeDrawingTextBody(cell.text, context)}${serializeCellProperties(cell)}</a:tc>`;
}

export function serializeTable(
  element: PptxSceneTableElement,
  transform: PptxSceneTransform,
  shapeId: number,
  context: PptxTextSerializationContext,
): string {
  const columns = element.columns
    .map((width) => `<a:gridCol w="${pointsToEmu(width)}"/>`)
    .join('');
  const rows = element.rows
    .map(
      (row) =>
        `<a:tr h="${pointsToEmu(row.height)}">${row.cells
          .map((cell) => serializeCell(cell, context))
          .join('')}</a:tr>`,
    )
    .join('');
  const table = `<a:tbl><a:tblPr/><a:tblGrid>${columns}</a:tblGrid>${rows}</a:tbl>`;
  return `<p:graphicFrame>${serializeNonVisualProperties(element, shapeId)}${serializeGraphicFrameTransform(transform)}<a:graphic><a:graphicData uri="${TABLE_URI}">${table}</a:graphicData></a:graphic></p:graphicFrame>`;
}
