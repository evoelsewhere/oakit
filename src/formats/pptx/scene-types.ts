export interface PptxSceneSize {
  height: number;
  width: number;
}

export interface PptxSceneTransform extends PptxSceneSize {
  flipHorizontal?: boolean;
  flipVertical?: boolean;
  rotation?: number;
  x: number;
  y: number;
}

export interface PptxSceneGroupTransform extends PptxSceneTransform {
  childSpace: PptxSceneTransform;
}

export interface PptxSceneAuthoredElement {
  hidden?: boolean;
  transform?: PptxSceneTransform;
}

export interface PptxSceneResolvedElement {
  hidden: boolean;
  transform?: PptxSceneTransform;
}

export interface PptxScenePlaceholder {
  hasCustomPrompt?: boolean;
  index?: number;
  orientation?: 'horizontal' | 'vertical';
  prompt?: string;
  role: 'layout-definition' | 'master-definition' | 'slide-instance';
  size?: 'full' | 'half' | 'quarter';
  sourceKey?: string;
  type?: string;
}

export interface PptxSceneElementBase {
  authored: PptxSceneAuthoredElement;
  description?: string;
  key: string;
  name?: string;
  placeholder?: PptxScenePlaceholder;
  resolved: PptxSceneResolvedElement;
  title?: string;
}

export interface PptxSceneTextBodyProperties {
  anchor?: 'bottom' | 'center' | 'distributed' | 'justified' | 'top';
  autoFit?: 'none' | 'shape' | 'text';
  vertical?: boolean;
  wrap?: boolean;
}

export interface PptxSceneRunProperties {
  bold?: boolean;
  fontFamily?: string;
  fontSize?: number;
  italic?: boolean;
  language?: string;
}

export interface PptxSceneTextRun {
  key: string;
  preserveSpace?: boolean;
  properties?: PptxSceneRunProperties;
  text: string;
  type: 'run';
}

export interface PptxSceneTextField {
  fieldType: string;
  key: string;
  properties?: PptxSceneRunProperties;
  text: string;
  type: 'field';
}

export interface PptxSceneTextBreak {
  key: string;
  properties?: PptxSceneRunProperties;
  type: 'break';
}

export type PptxSceneTextNode =
  PptxSceneTextBreak | PptxSceneTextField | PptxSceneTextRun;

export interface PptxSceneParagraphProperties {
  alignment?: 'center' | 'distributed' | 'justify' | 'left' | 'right';
  level?: number;
}

export interface PptxSceneParagraph {
  children: PptxSceneTextNode[];
  endProperties?: PptxSceneRunProperties;
  key: string;
  properties?: PptxSceneParagraphProperties;
}

export interface PptxSceneTextBody {
  body: PptxSceneTextBodyProperties;
  paragraphs: PptxSceneParagraph[];
}

export interface PptxSceneTextElement extends PptxSceneElementBase {
  text: PptxSceneTextBody;
  type: 'text';
}

export interface PptxSceneUnsupportedElement extends PptxSceneElementBase {
  feature: string;
  previewText?: string;
  type: 'unsupported';
}

export type PptxSceneElement =
  PptxSceneTextElement | PptxSceneUnsupportedElement;

export interface PptxSceneTheme {
  key: string;
  name?: string;
}

export interface PptxSceneMaster {
  elements: PptxSceneElement[];
  key: string;
  name?: string;
  themeKey: string;
}

export interface PptxSceneLayout {
  elements: PptxSceneElement[];
  key: string;
  masterKey: string;
  name?: string;
}

export interface PptxSceneSlide {
  elements: PptxSceneElement[];
  hidden?: boolean;
  key: string;
  layoutKey?: string;
  name?: string;
}

export interface PptxSceneDocument {
  layouts: PptxSceneLayout[];
  masters: PptxSceneMaster[];
  media: [];
  schemaVersion: 2;
  size: PptxSceneSize;
  slides: PptxSceneSlide[];
  themes: PptxSceneTheme[];
}

export type PptxSceneValidationCode =
  | 'duplicate-public-key'
  | 'invalid-hierarchy-reference'
  | 'invalid-numeric-value'
  | 'invalid-office-text-escape'
  | 'invalid-scene-document'
  | 'unsupported-feature'
  | 'unsupported-schema-version';

export interface PptxSceneValidationIssue {
  code: PptxSceneValidationCode;
  message: string;
  path: string;
}

export interface PptxSceneValidationOptions {
  profile?: 'create-text-v1' | 'scene';
}

export interface PptxSceneValidationResult {
  issues: PptxSceneValidationIssue[];
  valid: boolean;
}
