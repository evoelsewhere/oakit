import type {
  PptxSceneValidationCode,
  PptxSceneValidationIssue,
  PptxSceneValidationOptions,
  PptxSceneValidationResult,
} from './scene-types';

type JsonObject = Record<string, unknown>;

const KEY_PATTERN = /^[A-Za-z0-9][A-Za-z0-9._:-]{0,127}$/;
const EMUS_PER_POINT = 12_700;
const ANGLE_UNITS_PER_DEGREE = 60_000;
const FONT_SIZE_UNITS_PER_POINT = 100;

function isObject(value: unknown): value is JsonObject {
  if (value === null) return false;
  if (value === undefined) return false;
  return Object.getPrototypeOf(value) === Object.prototype;
}

function addIssue(
  issues: PptxSceneValidationIssue[],
  code: PptxSceneValidationCode,
  path: string,
  message: string,
): void {
  issues.push({ code, message, path });
}

function rejectUnknownKeys(
  value: JsonObject,
  allowed: readonly string[],
  path: string,
  issues: PptxSceneValidationIssue[],
): void {
  const allowedKeys = new Set(allowed);
  for (const key of Object.keys(value)) {
    if (!allowedKeys.has(key)) {
      addIssue(
        issues,
        'invalid-scene-document',
        `${path}.${key}`,
        'Unknown property',
      );
    }
  }
}

function requireObject(
  value: unknown,
  path: string,
  issues: PptxSceneValidationIssue[],
): JsonObject | undefined {
  if (isObject(value)) return value;
  addIssue(issues, 'invalid-scene-document', path, 'Expected an object');
  return undefined;
}

function requireArray(
  value: unknown,
  path: string,
  issues: PptxSceneValidationIssue[],
): unknown[] | undefined {
  if (Array.isArray(value)) return value as unknown[];
  addIssue(issues, 'invalid-scene-document', path, 'Expected an array');
  return undefined;
}

function isOneOf(value: unknown, allowed: readonly string[]): value is string {
  return typeof value === 'string' && allowed.includes(value);
}

function optionalString(
  value: JsonObject,
  key: string,
  path: string,
  issues: PptxSceneValidationIssue[],
): void {
  if (value[key] === undefined || typeof value[key] === 'string') return;
  addIssue(
    issues,
    'invalid-scene-document',
    `${path}.${key}`,
    'Expected a string',
  );
}

function optionalBoolean(
  value: JsonObject,
  key: string,
  path: string,
  issues: PptxSceneValidationIssue[],
): void {
  if (value[key] === undefined || typeof value[key] === 'boolean') return;
  addIssue(
    issues,
    'invalid-scene-document',
    `${path}.${key}`,
    'Expected a boolean',
  );
}

function requireFiniteNumber(
  value: unknown,
  path: string,
  issues: PptxSceneValidationIssue[],
  positive: boolean,
): void {
  if (
    typeof value === 'number' &&
    Number.isFinite(value) &&
    (!positive || value > 0)
  ) {
    return;
  }
  addIssue(
    issues,
    'invalid-numeric-value',
    path,
    positive ? 'Expected a positive finite number' : 'Expected a finite number',
  );
}

function requireSerializableInteger(
  value: unknown,
  multiplier: number,
  path: string,
  issues: PptxSceneValidationIssue[],
): void {
  if (typeof value !== 'number' || !Number.isFinite(value)) return;
  if (Number.isSafeInteger(Math.round(value * multiplier))) return;
  addIssue(
    issues,
    'invalid-numeric-value',
    path,
    'Value exceeds the safe OOXML integer range',
  );
}

function validateSize(
  value: unknown,
  path: string,
  issues: PptxSceneValidationIssue[],
  profile: 'create-text-v1' | 'scene',
): void {
  const size = requireObject(value, path, issues);
  if (!size) return;
  rejectUnknownKeys(size, ['height', 'width'], path, issues);
  requireFiniteNumber(size.width, `${path}.width`, issues, true);
  requireFiniteNumber(size.height, `${path}.height`, issues, true);
  if (profile === 'create-text-v1') {
    requireSerializableInteger(
      size.width,
      EMUS_PER_POINT,
      `${path}.width`,
      issues,
    );
    requireSerializableInteger(
      size.height,
      EMUS_PER_POINT,
      `${path}.height`,
      issues,
    );
  }
}

function validateTransform(
  value: unknown,
  path: string,
  issues: PptxSceneValidationIssue[],
  profile: 'create-text-v1' | 'scene',
): void {
  const transform = requireObject(value, path, issues);
  if (!transform) return;
  rejectUnknownKeys(
    transform,
    ['flipHorizontal', 'flipVertical', 'height', 'rotation', 'width', 'x', 'y'],
    path,
    issues,
  );
  requireFiniteNumber(transform.x, `${path}.x`, issues, false);
  requireFiniteNumber(transform.y, `${path}.y`, issues, false);
  requireFiniteNumber(transform.width, `${path}.width`, issues, true);
  requireFiniteNumber(transform.height, `${path}.height`, issues, true);
  if (transform.rotation !== undefined) {
    requireFiniteNumber(transform.rotation, `${path}.rotation`, issues, false);
  }
  optionalBoolean(transform, 'flipHorizontal', path, issues);
  optionalBoolean(transform, 'flipVertical', path, issues);
  if (profile === 'create-text-v1') {
    for (const key of ['height', 'width', 'x', 'y'] as const) {
      requireSerializableInteger(
        transform[key],
        EMUS_PER_POINT,
        `${path}.${key}`,
        issues,
      );
    }
    requireSerializableInteger(
      transform.rotation,
      ANGLE_UNITS_PER_DEGREE,
      `${path}.rotation`,
      issues,
    );
  }
}

function registerKey(
  value: unknown,
  path: string,
  keys: Set<string>,
  issues: PptxSceneValidationIssue[],
): string | undefined {
  if (typeof value !== 'string' || !KEY_PATTERN.test(value)) {
    addIssue(
      issues,
      'invalid-scene-document',
      path,
      'Expected a non-empty portable key of at most 128 characters',
    );
    return undefined;
  }
  if (keys.has(value)) {
    addIssue(
      issues,
      'duplicate-public-key',
      path,
      `Duplicate public key: ${value}`,
    );
    return value;
  }
  keys.add(value);
  return value;
}

function isValidXmlText(value: string): boolean {
  for (const character of value) {
    if (character.length === 2) continue;
    const codeUnit = character.charCodeAt(0);
    if (
      (codeUnit < 0x20 &&
        codeUnit !== 0x09 &&
        codeUnit !== 0x0a &&
        codeUnit !== 0x0d) ||
      (codeUnit >= 0xd800 && codeUnit <= 0xdfff) ||
      codeUnit === 0xfffe ||
      codeUnit === 0xffff
    ) {
      return false;
    }
  }
  return true;
}

function validateTextValue(
  value: unknown,
  path: string,
  issues: PptxSceneValidationIssue[],
): void {
  if (typeof value !== 'string') {
    addIssue(issues, 'invalid-scene-document', path, 'Expected text content');
  } else if (!isValidXmlText(value)) {
    addIssue(
      issues,
      'invalid-office-text-escape',
      path,
      'Text contains a character that cannot be serialized safely',
    );
  }
}

function validateRunProperties(
  value: unknown,
  path: string,
  issues: PptxSceneValidationIssue[],
  profile: 'create-text-v1' | 'scene',
): void {
  const properties = requireObject(value, path, issues);
  if (!properties) return;
  rejectUnknownKeys(
    properties,
    ['bold', 'fontFamily', 'fontSize', 'italic', 'language'],
    path,
    issues,
  );
  optionalBoolean(properties, 'bold', path, issues);
  optionalBoolean(properties, 'italic', path, issues);
  optionalString(properties, 'fontFamily', path, issues);
  optionalString(properties, 'language', path, issues);
  if (properties.fontSize !== undefined) {
    requireFiniteNumber(properties.fontSize, `${path}.fontSize`, issues, true);
    if (profile === 'create-text-v1') {
      requireSerializableInteger(
        properties.fontSize,
        FONT_SIZE_UNITS_PER_POINT,
        `${path}.fontSize`,
        issues,
      );
    }
  }
}

function validateTextNode(
  value: unknown,
  path: string,
  profile: 'create-text-v1' | 'scene',
  keys: Set<string>,
  issues: PptxSceneValidationIssue[],
): void {
  const node = requireObject(value, path, issues);
  if (!node) return;
  registerKey(node.key, `${path}.key`, keys, issues);
  if (node.type === 'run') {
    rejectUnknownKeys(
      node,
      ['key', 'preserveSpace', 'properties', 'text', 'type'],
      path,
      issues,
    );
    validateTextValue(node.text, `${path}.text`, issues);
    optionalBoolean(node, 'preserveSpace', path, issues);
  } else if (node.type === 'field') {
    rejectUnknownKeys(
      node,
      ['fieldType', 'key', 'properties', 'text', 'type'],
      path,
      issues,
    );
    validateTextValue(node.text, `${path}.text`, issues);
    if (typeof node.fieldType !== 'string' || node.fieldType.trim() === '') {
      addIssue(
        issues,
        'invalid-scene-document',
        `${path}.fieldType`,
        'Expected a non-empty field type',
      );
    }
  } else if (node.type === 'break') {
    rejectUnknownKeys(node, ['key', 'properties', 'type'], path, issues);
  } else {
    addIssue(
      issues,
      'invalid-scene-document',
      `${path}.type`,
      'Unknown text node type',
    );
  }
  if (node.properties !== undefined) {
    validateRunProperties(
      node.properties,
      `${path}.properties`,
      issues,
      profile,
    );
  }
}

function validateParagraph(
  value: unknown,
  path: string,
  profile: 'create-text-v1' | 'scene',
  keys: Set<string>,
  issues: PptxSceneValidationIssue[],
): void {
  const paragraph = requireObject(value, path, issues);
  if (!paragraph) return;
  rejectUnknownKeys(
    paragraph,
    ['children', 'endProperties', 'key', 'properties'],
    path,
    issues,
  );
  registerKey(paragraph.key, `${path}.key`, keys, issues);
  const children = requireArray(paragraph.children, `${path}.children`, issues);
  children?.forEach((child, index) =>
    validateTextNode(
      child,
      `${path}.children[${index}]`,
      profile,
      keys,
      issues,
    ),
  );
  if (paragraph.endProperties !== undefined) {
    validateRunProperties(
      paragraph.endProperties,
      `${path}.endProperties`,
      issues,
      profile,
    );
  }
  if (paragraph.properties !== undefined) {
    const properties = requireObject(
      paragraph.properties,
      `${path}.properties`,
      issues,
    );
    if (properties) {
      rejectUnknownKeys(
        properties,
        ['alignment', 'level'],
        `${path}.properties`,
        issues,
      );
      if (
        properties.alignment !== undefined &&
        !isOneOf(properties.alignment, [
          'center',
          'distributed',
          'justify',
          'left',
          'right',
        ])
      ) {
        addIssue(
          issues,
          'invalid-scene-document',
          `${path}.properties.alignment`,
          'Unknown paragraph alignment',
        );
      }
      if (
        properties.level !== undefined &&
        (!Number.isSafeInteger(properties.level) ||
          Number(properties.level) < 0 ||
          Number(properties.level) > 8)
      ) {
        addIssue(
          issues,
          'invalid-numeric-value',
          `${path}.properties.level`,
          'Paragraph level must be an integer from 0 through 8',
        );
      }
    }
  }
}

function validateTextBody(
  value: unknown,
  path: string,
  profile: 'create-text-v1' | 'scene',
  keys: Set<string>,
  issues: PptxSceneValidationIssue[],
): void {
  const text = requireObject(value, path, issues);
  if (!text) return;
  rejectUnknownKeys(text, ['body', 'paragraphs'], path, issues);
  const body = requireObject(text.body, `${path}.body`, issues);
  if (body) {
    rejectUnknownKeys(
      body,
      ['anchor', 'autoFit', 'vertical', 'wrap'],
      `${path}.body`,
      issues,
    );
    optionalBoolean(body, 'vertical', `${path}.body`, issues);
    optionalBoolean(body, 'wrap', `${path}.body`, issues);
    if (
      body.anchor !== undefined &&
      !isOneOf(body.anchor, [
        'bottom',
        'center',
        'distributed',
        'justified',
        'top',
      ])
    ) {
      addIssue(
        issues,
        'invalid-scene-document',
        `${path}.body.anchor`,
        'Unknown text anchor',
      );
    }
    if (
      body.autoFit !== undefined &&
      !isOneOf(body.autoFit, ['none', 'shape', 'text'])
    ) {
      addIssue(
        issues,
        'invalid-scene-document',
        `${path}.body.autoFit`,
        'Unknown text auto-fit mode',
      );
    }
  }
  const paragraphs = requireArray(
    text.paragraphs,
    `${path}.paragraphs`,
    issues,
  );
  if (paragraphs?.length === 0) {
    addIssue(
      issues,
      'invalid-scene-document',
      `${path}.paragraphs`,
      'A text body needs at least one paragraph',
    );
  }
  paragraphs?.forEach((paragraph, index) =>
    validateParagraph(
      paragraph,
      `${path}.paragraphs[${index}]`,
      profile,
      keys,
      issues,
    ),
  );
}

function validatePlaceholder(
  value: unknown,
  path: string,
  referenceKeys: Array<{ path: string; value: string }>,
  issues: PptxSceneValidationIssue[],
): void {
  const placeholder = requireObject(value, path, issues);
  if (!placeholder) return;
  rejectUnknownKeys(
    placeholder,
    [
      'hasCustomPrompt',
      'index',
      'orientation',
      'prompt',
      'role',
      'size',
      'sourceKey',
      'type',
    ],
    path,
    issues,
  );
  optionalBoolean(placeholder, 'hasCustomPrompt', path, issues);
  optionalString(placeholder, 'prompt', path, issues);
  optionalString(placeholder, 'type', path, issues);
  if (
    placeholder.index !== undefined &&
    (!Number.isSafeInteger(placeholder.index) || Number(placeholder.index) < 0)
  ) {
    addIssue(
      issues,
      'invalid-numeric-value',
      `${path}.index`,
      'Placeholder index must be a non-negative integer',
    );
  }
  if (
    !isOneOf(placeholder.role, [
      'layout-definition',
      'master-definition',
      'slide-instance',
    ])
  ) {
    addIssue(
      issues,
      'invalid-scene-document',
      `${path}.role`,
      'Unknown placeholder role',
    );
  }
  if (
    placeholder.orientation !== undefined &&
    !isOneOf(placeholder.orientation, ['horizontal', 'vertical'])
  ) {
    addIssue(
      issues,
      'invalid-scene-document',
      `${path}.orientation`,
      'Unknown placeholder orientation',
    );
  }
  if (
    placeholder.size !== undefined &&
    !isOneOf(placeholder.size, ['full', 'half', 'quarter'])
  ) {
    addIssue(
      issues,
      'invalid-scene-document',
      `${path}.size`,
      'Unknown placeholder size',
    );
  }
  if (placeholder.sourceKey !== undefined) {
    if (typeof placeholder.sourceKey === 'string') {
      referenceKeys.push({
        path: `${path}.sourceKey`,
        value: placeholder.sourceKey,
      });
    } else {
      addIssue(
        issues,
        'invalid-scene-document',
        `${path}.sourceKey`,
        'Expected a public key',
      );
    }
  }
}

function validateElement(
  value: unknown,
  path: string,
  profile: 'create-text-v1' | 'scene',
  keys: Set<string>,
  referenceKeys: Array<{ path: string; value: string }>,
  issues: PptxSceneValidationIssue[],
): void {
  const element = requireObject(value, path, issues);
  if (!element) return;
  registerKey(element.key, `${path}.key`, keys, issues);
  const baseKeys = [
    'authored',
    'description',
    'key',
    'name',
    'placeholder',
    'resolved',
    'title',
    'type',
  ];
  if (element.type === 'text') {
    rejectUnknownKeys(element, [...baseKeys, 'text'], path, issues);
    validateTextBody(element.text, `${path}.text`, profile, keys, issues);
  } else if (element.type === 'unsupported') {
    rejectUnknownKeys(
      element,
      [...baseKeys, 'feature', 'previewText'],
      path,
      issues,
    );
    if (typeof element.feature !== 'string' || element.feature.trim() === '') {
      addIssue(
        issues,
        'invalid-scene-document',
        `${path}.feature`,
        'Expected a non-empty unsupported feature name',
      );
    }
    optionalString(element, 'previewText', path, issues);
    if (profile === 'create-text-v1') {
      addIssue(
        issues,
        'unsupported-feature',
        path,
        'Creation profile create-text-v1 supports text elements only',
      );
    }
  } else {
    addIssue(
      issues,
      'invalid-scene-document',
      `${path}.type`,
      'Unknown scene element type',
    );
  }
  optionalString(element, 'description', path, issues);
  optionalString(element, 'name', path, issues);
  optionalString(element, 'title', path, issues);

  const authored = requireObject(element.authored, `${path}.authored`, issues);
  if (authored) {
    rejectUnknownKeys(
      authored,
      ['hidden', 'transform'],
      `${path}.authored`,
      issues,
    );
    optionalBoolean(authored, 'hidden', `${path}.authored`, issues);
    if (authored.transform !== undefined) {
      validateTransform(
        authored.transform,
        `${path}.authored.transform`,
        issues,
        profile,
      );
    } else if (profile === 'create-text-v1' && element.type === 'text') {
      addIssue(
        issues,
        'unsupported-feature',
        `${path}.authored.transform`,
        'Creation profile create-text-v1 requires an authored text transform',
      );
    }
  }
  const resolved = requireObject(element.resolved, `${path}.resolved`, issues);
  if (resolved) {
    rejectUnknownKeys(
      resolved,
      ['hidden', 'transform'],
      `${path}.resolved`,
      issues,
    );
    if (typeof resolved.hidden !== 'boolean') {
      addIssue(
        issues,
        'invalid-scene-document',
        `${path}.resolved.hidden`,
        'Expected a boolean',
      );
    }
    if (resolved.transform !== undefined) {
      validateTransform(
        resolved.transform,
        `${path}.resolved.transform`,
        issues,
        profile,
      );
    }
  }
  if (element.placeholder !== undefined) {
    validatePlaceholder(
      element.placeholder,
      `${path}.placeholder`,
      referenceKeys,
      issues,
    );
    if (profile === 'create-text-v1') {
      addIssue(
        issues,
        'unsupported-feature',
        `${path}.placeholder`,
        'Creation profile create-text-v1 does not support placeholders yet',
      );
    }
  }
}

function validateElementArray(
  value: unknown,
  path: string,
  profile: 'create-text-v1' | 'scene',
  keys: Set<string>,
  referenceKeys: Array<{ path: string; value: string }>,
  issues: PptxSceneValidationIssue[],
): void {
  const elements = requireArray(value, path, issues);
  elements?.forEach((element, index) =>
    validateElement(
      element,
      `${path}[${index}]`,
      profile,
      keys,
      referenceKeys,
      issues,
    ),
  );
}

export function validatePptxScene(
  value: unknown,
  options: PptxSceneValidationOptions = {},
): PptxSceneValidationResult {
  const issues: PptxSceneValidationIssue[] = [];
  const document = requireObject(value, '$', issues);
  if (!document) return { issues, valid: false };
  rejectUnknownKeys(
    document,
    [
      'layouts',
      'masters',
      'media',
      'schemaVersion',
      'size',
      'slides',
      'themes',
    ],
    '$',
    issues,
  );
  if (document.schemaVersion !== 2) {
    addIssue(
      issues,
      'unsupported-schema-version',
      '$.schemaVersion',
      'Only PowerPoint scene schema version 2 is supported',
    );
  }
  const profile = options.profile ?? 'scene';
  validateSize(document.size, '$.size', issues, profile);
  const keys = new Set<string>();
  const references: Array<{ path: string; value: string }> = [];
  const themeKeys = new Set<string>();
  const masterKeys = new Set<string>();
  const layoutKeys = new Set<string>();

  const themes = requireArray(document.themes, '$.themes', issues);
  themes?.forEach((value, index) => {
    const path = `$.themes[${index}]`;
    const theme = requireObject(value, path, issues);
    if (!theme) return;
    rejectUnknownKeys(theme, ['key', 'name'], path, issues);
    const key = registerKey(theme.key, `${path}.key`, keys, issues);
    if (key) themeKeys.add(key);
    optionalString(theme, 'name', path, issues);
  });

  const masters = requireArray(document.masters, '$.masters', issues);
  masters?.forEach((value, index) => {
    const path = `$.masters[${index}]`;
    const master = requireObject(value, path, issues);
    if (!master) return;
    rejectUnknownKeys(
      master,
      ['elements', 'key', 'name', 'themeKey'],
      path,
      issues,
    );
    const key = registerKey(master.key, `${path}.key`, keys, issues);
    if (key) masterKeys.add(key);
    optionalString(master, 'name', path, issues);
    if (typeof master.themeKey !== 'string') {
      addIssue(
        issues,
        'invalid-scene-document',
        `${path}.themeKey`,
        'Expected a theme key',
      );
    }
    validateElementArray(
      master.elements,
      `${path}.elements`,
      profile,
      keys,
      references,
      issues,
    );
  });

  const layouts = requireArray(document.layouts, '$.layouts', issues);
  layouts?.forEach((value, index) => {
    const path = `$.layouts[${index}]`;
    const layout = requireObject(value, path, issues);
    if (!layout) return;
    rejectUnknownKeys(
      layout,
      ['elements', 'key', 'masterKey', 'name'],
      path,
      issues,
    );
    const key = registerKey(layout.key, `${path}.key`, keys, issues);
    if (key) layoutKeys.add(key);
    optionalString(layout, 'name', path, issues);
    if (typeof layout.masterKey !== 'string') {
      addIssue(
        issues,
        'invalid-scene-document',
        `${path}.masterKey`,
        'Expected a master key',
      );
    }
    validateElementArray(
      layout.elements,
      `${path}.elements`,
      profile,
      keys,
      references,
      issues,
    );
  });

  const slides = requireArray(document.slides, '$.slides', issues);
  slides?.forEach((value, index) => {
    const path = `$.slides[${index}]`;
    const slide = requireObject(value, path, issues);
    if (!slide) return;
    rejectUnknownKeys(
      slide,
      ['elements', 'hidden', 'key', 'layoutKey', 'name'],
      path,
      issues,
    );
    registerKey(slide.key, `${path}.key`, keys, issues);
    optionalBoolean(slide, 'hidden', path, issues);
    optionalString(slide, 'name', path, issues);
    if (slide.layoutKey !== undefined && typeof slide.layoutKey !== 'string') {
      addIssue(
        issues,
        'invalid-scene-document',
        `${path}.layoutKey`,
        'Expected a layout key',
      );
    }
    validateElementArray(
      slide.elements,
      `${path}.elements`,
      profile,
      keys,
      references,
      issues,
    );
  });

  const media = requireArray(document.media, '$.media', issues);
  if (media && media.length > 0) {
    addIssue(
      issues,
      'unsupported-feature',
      '$.media',
      'The first scene contract does not support media resources',
    );
  }

  let hierarchyEmpty: boolean | undefined;
  if (themes && masters && layouts) {
    const emptyHierarchyCollections = [themes, masters, layouts].filter(
      (collection) => collection.length === 0,
    ).length;
    hierarchyEmpty = emptyHierarchyCollections === 3;
    if (profile === 'create-text-v1' && !hierarchyEmpty) {
      addIssue(
        issues,
        'unsupported-feature',
        '$',
        'Creation profile create-text-v1 generates its own minimal hierarchy',
      );
    }
    if (emptyHierarchyCollections > 0 && emptyHierarchyCollections < 3) {
      addIssue(
        issues,
        'invalid-hierarchy-reference',
        '$',
        'A declared hierarchy needs themes, masters, and layouts',
      );
    }
    masters.forEach((value, index) => {
      if (
        isObject(value) &&
        typeof value.themeKey === 'string' &&
        !themeKeys.has(value.themeKey)
      ) {
        addIssue(
          issues,
          'invalid-hierarchy-reference',
          `$.masters[${index}].themeKey`,
          'Master references an unknown theme',
        );
      }
    });
    layouts.forEach((value, index) => {
      if (
        isObject(value) &&
        typeof value.masterKey === 'string' &&
        !masterKeys.has(value.masterKey)
      ) {
        addIssue(
          issues,
          'invalid-hierarchy-reference',
          `$.layouts[${index}].masterKey`,
          'Layout references an unknown master',
        );
      }
    });
  }
  slides?.forEach((value, index) => {
    if (!isObject(value) || hierarchyEmpty === undefined) return;
    if (hierarchyEmpty) {
      if (value.layoutKey !== undefined) {
        addIssue(
          issues,
          'invalid-hierarchy-reference',
          `$.slides[${index}].layoutKey`,
          'A generated minimal hierarchy must not name a layout',
        );
      }
    } else if (
      typeof value.layoutKey !== 'string' ||
      !layoutKeys.has(value.layoutKey)
    ) {
      addIssue(
        issues,
        'invalid-hierarchy-reference',
        `$.slides[${index}].layoutKey`,
        'Slide references an unknown layout',
      );
    }
  });
  for (const reference of references) {
    if (!keys.has(reference.value)) {
      addIssue(
        issues,
        'invalid-hierarchy-reference',
        reference.path,
        'Reference points to an unknown public key',
      );
    }
  }

  return { issues, valid: issues.length === 0 };
}
