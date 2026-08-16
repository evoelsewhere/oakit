import { escapeXmlAttribute } from './xml';

const PACKAGE_RELATIONSHIPS_NAMESPACE =
  'http://schemas.openxmlformats.org/package/2006/relationships';

export interface PptxSerializedRelationship {
  id: string;
  target: string;
  targetMode?: 'External';
  type: string;
}

function serializeRelationship(
  relationship: PptxSerializedRelationship,
): string {
  const attributes = [
    `Id="${escapeXmlAttribute(relationship.id)}"`,
    `Type="${escapeXmlAttribute(relationship.type)}"`,
    `Target="${escapeXmlAttribute(relationship.target)}"`,
  ];
  if (relationship.targetMode !== undefined) {
    attributes.push(`TargetMode="${relationship.targetMode}"`);
  }
  return `<Relationship ${attributes.join(' ')}/>`;
}

export function serializeRelationships(
  relationships: readonly PptxSerializedRelationship[],
): string {
  const body = relationships.map(serializeRelationship).join('');
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="${PACKAGE_RELATIONSHIPS_NAMESPACE}">${body}</Relationships>`;
}
