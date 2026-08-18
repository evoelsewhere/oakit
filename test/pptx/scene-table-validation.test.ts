import { describe, expect, it } from 'vitest';

import { validatePptxScene } from '../../src/formats/pptx/scene-validation';
import type { PptxSceneValidationOptions } from '../../src/formats/pptx/scene-types';

type Mutable = Record<string, unknown>;
type Profile = NonNullable<PptxSceneValidationOptions['profile']>;

const ELEMENT_PATH = '$.slides[0].elements[0]';

function text(key: string): Mutable {
  return {
    body: { anchor: 'center' },
    paragraphs: [
      {
        children: [{ key: `${key}-run`, text: key, type: 'run' }],
        key: `${key}-paragraph`,
      },
    ],
  };
}

function cell(key: string): Mutable {
  return { text: text(key) };
}

function table(): Mutable {
  return {
    authored: {
      transform: { height: 80, width: 300, x: 40, y: 50 },
    },
    columns: [100, 200],
    key: 'table-1',
    resolved: { hidden: false },
    rows: [
      {
        cells: [cell('cell-1'), cell('cell-2')],
        height: 80,
      },
    ],
    type: 'table',
  };
}

function document(element: Mutable = table()): Mutable {
  return {
    layouts: [],
    masters: [],
    media: [],
    schemaVersion: 2,
    size: { height: 540, width: 960 },
    slides: [{ elements: [element], key: 'slide-1' }],
    themes: [],
  };
}

function validate(element: Mutable, profile: Profile = 'create-native-v1') {
  return validatePptxScene(document(element), { profile });
}

function firstRow(value: Mutable): Mutable {
  return (value.rows as Mutable[])[0] as Mutable;
}

function firstCell(value: Mutable): Mutable {
  return (firstRow(value).cells as Mutable[])[0] as Mutable;
}

function mergedTable(): Mutable {
  const value = table();
  value.rows = [
    {
      cells: [
        { colSpan: 2, rowSpan: 2, text: text('origin') },
        { hMerge: true, text: text('top-right') },
      ],
      height: 40,
    },
    {
      cells: [
        { text: text('bottom-left'), vMerge: true },
        { hMerge: true, text: text('bottom-right'), vMerge: true },
      ],
      height: 40,
    },
  ];
  return value;
}

describe('native PowerPoint table scene validation', () => {
  it('accepts exact create and preservation profiles', () => {
    expect(validate(table())).toEqual({ issues: [], valid: true });
    expect(validate(table(), 'scene')).toEqual({ issues: [], valid: true });
  });

  it.each([
    ['columns', {}, 'Expected an array'],
    ['rows', null, 'Expected an array'],
  ])('rejects primitive %s containers', (key, replacement, message) => {
    const value = table();
    value[key] = replacement;

    expect(validate(value).issues).toContainEqual({
      code: 'invalid-scene-document',
      message,
      path: `${ELEMENT_PATH}.${key}`,
    });
  });

  it.each([
    ['columns', 'A table needs at least one column'],
    ['rows', 'A table needs at least one row'],
  ])('rejects empty %s', (key, message) => {
    const value = table();
    value[key] = [];

    expect(validate(value).issues).toContainEqual({
      code: 'invalid-scene-document',
      message,
      path: `${ELEMENT_PATH}.${key}`,
    });
  });

  it.each([
    ['column', 'columns', 0, `${ELEMENT_PATH}.columns[0]`],
    ['column NaN', 'columns', Number.NaN, `${ELEMENT_PATH}.columns[0]`],
    ['row', 'height', 0, `${ELEMENT_PATH}.rows[0].height`],
    ['row NaN', 'height', Number.NaN, `${ELEMENT_PATH}.rows[0].height`],
  ])('rejects invalid %s dimensions', (_name, kind, replacement, path) => {
    const value = table();
    if (kind === 'columns') {
      (value.columns as unknown[])[0] = replacement;
    } else {
      firstRow(value).height = replacement;
    }

    expect(validate(value).issues).toContainEqual({
      code: 'invalid-numeric-value',
      message: 'Expected a positive finite number',
      path,
    });
  });

  it('enforces serializable column, row, and border EMUs only for creation', () => {
    const huge = Number.MAX_SAFE_INTEGER;
    const value = table();
    value.columns = [huge, 200];
    firstRow(value).height = huge;
    const target = firstCell(value);
    target.borders = {
      top: { color: '#123456', width: huge },
    };
    const creation = validate(value).issues;

    for (const path of [
      `${ELEMENT_PATH}.columns[0]`,
      `${ELEMENT_PATH}.rows[0].height`,
      `${ELEMENT_PATH}.rows[0].cells[0].borders.top.width`,
    ]) {
      expect(creation).toContainEqual({
        code: 'invalid-numeric-value',
        message: 'Value exceeds the safe OOXML integer range',
        path,
      });
    }
    expect(validate(value, 'scene').issues).not.toContainEqual(
      expect.objectContaining({
        message: 'Value exceeds the safe OOXML integer range',
      }),
    );
  });

  it('requires positive serialized EMUs for tiny column and row dimensions', () => {
    const value = table();
    value.columns = [0.000_01, 299.999_99];
    firstRow(value).height = 0.000_01;
    const authored = value.authored as Mutable;
    const transform = authored.transform as Mutable;
    transform.height = 0.000_01;

    expect(validate(value).issues).toEqual(
      expect.arrayContaining([
        {
          code: 'invalid-numeric-value',
          message: 'Value must round to a positive OOXML integer',
          path: `${ELEMENT_PATH}.columns[0]`,
        },
        {
          code: 'invalid-numeric-value',
          message: 'Value must round to a positive OOXML integer',
          path: `${ELEMENT_PATH}.rows[0].height`,
        },
      ]),
    );
  });

  it('rejects malformed rows and exact cell-count mismatches', () => {
    const primitive = table();
    primitive.rows = [null];
    expect(validate(primitive).issues).toContainEqual({
      code: 'invalid-scene-document',
      message: 'Expected an object',
      path: `${ELEMENT_PATH}.rows[0]`,
    });

    const unknown = table();
    firstRow(unknown).extra = true;
    expect(validate(unknown).issues).toContainEqual({
      code: 'invalid-scene-document',
      message: 'Unknown property',
      path: `${ELEMENT_PATH}.rows[0].extra`,
    });

    const cells = table();
    firstRow(cells).cells = [cell('only')];
    expect(validate(cells).issues).toContainEqual({
      code: 'invalid-scene-document',
      message: 'Table row must contain exactly 2 grid cells',
      path: `${ELEMENT_PATH}.rows[0].cells`,
    });

    const primitiveCells = table();
    firstRow(primitiveCells).cells = {};
    expect(validate(primitiveCells).issues).toContainEqual({
      code: 'invalid-scene-document',
      message: 'Expected an array',
      path: `${ELEMENT_PATH}.rows[0].cells`,
    });
  });

  it('rejects primitive cells, unknown fields, colors, and merge booleans', () => {
    const value = table();
    const row = firstRow(value);
    row.cells = [null, cell('second')];
    expect(validate(value).issues).toContainEqual({
      code: 'invalid-scene-document',
      message: 'Expected an object',
      path: `${ELEMENT_PATH}.rows[0].cells[0]`,
    });

    const invalid = table();
    const target = firstCell(invalid);
    target.extra = true;
    target.fillColor = 'blue';
    target.hMerge = 1;
    target.vMerge = 'true';
    const issues = validate(invalid).issues;
    expect(issues).toEqual(
      expect.arrayContaining([
        {
          code: 'invalid-scene-document',
          message: 'Unknown property',
          path: `${ELEMENT_PATH}.rows[0].cells[0].extra`,
        },
        {
          code: 'invalid-scene-document',
          message: 'Expected a #RRGGBB color',
          path: `${ELEMENT_PATH}.rows[0].cells[0].fillColor`,
        },
        {
          code: 'invalid-scene-document',
          message: 'Expected a boolean',
          path: `${ELEMENT_PATH}.rows[0].cells[0].hMerge`,
        },
        {
          code: 'invalid-scene-document',
          message: 'Expected a boolean',
          path: `${ELEMENT_PATH}.rows[0].cells[0].vMerge`,
        },
      ]),
    );
  });

  it.each([
    ['colSpan', 1],
    ['colSpan', 1.5],
    ['colSpan', '2'],
    ['rowSpan', 0],
    ['rowSpan', 1],
    ['rowSpan', 2.5],
  ])('rejects invalid %s value %j', (key, replacement) => {
    const value = table();
    firstCell(value)[key] = replacement;
    const issues = validate(value).issues;

    expect(issues).toContainEqual({
      code: 'invalid-numeric-value',
      message: 'Table span must be an integer greater than one',
      path: `${ELEMENT_PATH}.rows[0].cells[0].${key}`,
    });
    expect(issues.map(({ message }) => message)).not.toEqual(
      expect.arrayContaining([
        'Table span exceeds the grid bounds',
        'Table spans must not overlap',
        'Table merge continuation flags do not match its spans',
      ]),
    );
  });

  it('validates border containers, directions, fields, and every style', () => {
    for (const style of ['dashed', 'dotted', 'solid']) {
      const value = table();
      firstCell(value).borders = {
        bottom: { color: '#123456', style, width: 1 },
        left: { color: '#234567', style, width: 2 },
        right: { color: '#345678', style, width: 3 },
        top: { color: '#456789', style, width: 4 },
      };
      expect(validate(value)).toEqual({ issues: [], valid: true });
    }

    const primitive = table();
    firstCell(primitive).borders = [];
    expect(validate(primitive).issues).toContainEqual({
      code: 'invalid-scene-document',
      message: 'Expected an object',
      path: `${ELEMENT_PATH}.rows[0].cells[0].borders`,
    });

    const invalid = table();
    firstCell(invalid).borders = {
      diagonal: { color: '#123456', width: 1 },
      top: { extra: true, style: 'double', width: 0 },
    };
    const issues = validate(invalid).issues;
    expect(issues).toEqual(
      expect.arrayContaining([
        {
          code: 'invalid-scene-document',
          message: 'Unknown property',
          path: `${ELEMENT_PATH}.rows[0].cells[0].borders.diagonal`,
        },
        {
          code: 'invalid-scene-document',
          message: 'Unknown property',
          path: `${ELEMENT_PATH}.rows[0].cells[0].borders.top.extra`,
        },
        {
          code: 'invalid-scene-document',
          message: 'Expected a #RRGGBB color',
          path: `${ELEMENT_PATH}.rows[0].cells[0].borders.top.color`,
        },
        {
          code: 'invalid-scene-document',
          message: 'Unknown table border style',
          path: `${ELEMENT_PATH}.rows[0].cells[0].borders.top.style`,
        },
        {
          code: 'invalid-numeric-value',
          message: 'Expected a positive finite number',
          path: `${ELEMENT_PATH}.rows[0].cells[0].borders.top.width`,
        },
      ]),
    );
  });

  it('accepts sparse borders, an omitted style, and a valid cell fill', () => {
    const value = table();
    const target = firstCell(value);
    target.fillColor = '#ABCDEF';
    target.borders = {
      top: { color: '#123456', width: 1 },
    };

    expect(validate(value)).toEqual({ issues: [], valid: true });
  });

  it('reports exact border color, positive EMU, and cell text paths', () => {
    const value = table();
    const target = firstCell(value);
    target.borders = {
      top: { color: 'blue', width: 0.000_01 },
    };
    target.text = null;

    expect(validate(value).issues).toEqual(
      expect.arrayContaining([
        {
          code: 'invalid-scene-document',
          message: 'Expected a #RRGGBB color',
          path: `${ELEMENT_PATH}.rows[0].cells[0].borders.top.color`,
        },
        {
          code: 'invalid-numeric-value',
          message: 'Value must round to a positive OOXML integer',
          path: `${ELEMENT_PATH}.rows[0].cells[0].borders.top.width`,
        },
        {
          code: 'invalid-scene-document',
          message: 'Expected an object',
          path: `${ELEMENT_PATH}.rows[0].cells[0].text`,
        },
      ]),
    );
  });

  it('accepts an exact 2x2 merge rectangle', () => {
    expect(validate(mergedTable())).toEqual({ issues: [], valid: true });
  });

  it('accepts independent horizontal-only and vertical-only merges', () => {
    const horizontal = table();
    firstRow(horizontal).cells = [
      { colSpan: 2, text: text('horizontal-origin') },
      { hMerge: true, text: text('horizontal-tail') },
    ];
    expect(validate(horizontal)).toEqual({ issues: [], valid: true });

    const vertical = table();
    vertical.columns = [300];
    vertical.rows = [
      {
        cells: [{ rowSpan: 2, text: text('vertical-origin') }],
        height: 40,
      },
      {
        cells: [{ text: text('vertical-tail'), vMerge: true }],
        height: 40,
      },
    ];
    expect(validate(vertical)).toEqual({ issues: [], valid: true });
  });

  it.each([
    [0, 1, 'hMerge'],
    [1, 0, 'vMerge'],
    [1, 1, 'hMerge'],
    [1, 1, 'vMerge'],
  ])(
    'requires continuation flag %s at row %i column %i',
    (rowIndex, columnIndex, key) => {
      const value = mergedTable();
      const row = (value.rows as Mutable[])[rowIndex] as Mutable;
      const target = (row.cells as Mutable[])[columnIndex] as Mutable;
      delete target[key];

      expect(validate(value).issues).toContainEqual({
        code: 'invalid-scene-document',
        message: 'Table merge continuation flags do not match its spans',
        path: `${ELEMENT_PATH}.rows[${rowIndex}].cells[${columnIndex}]`,
      });
    },
  );

  it.each([
    [0, 1, 'vMerge'],
    [1, 0, 'hMerge'],
  ])(
    'rejects an extra continuation flag %s at row %i column %i',
    (rowIndex, columnIndex, key) => {
      const value = mergedTable();
      const row = (value.rows as Mutable[])[rowIndex] as Mutable;
      const target = (row.cells as Mutable[])[columnIndex] as Mutable;
      target[key] = true;

      expect(validate(value).issues).toContainEqual({
        code: 'invalid-scene-document',
        message: 'Table merge continuation flags do not match its spans',
        path: `${ELEMENT_PATH}.rows[${rowIndex}].cells[${columnIndex}]`,
      });
    },
  );

  it.each([
    ['colSpan', 3],
    ['rowSpan', 3],
  ])('rejects out-of-bounds %s', (key, replacement) => {
    const value = mergedTable();
    firstCell(value)[key] = replacement;

    expect(validate(value).issues).toContainEqual({
      code: 'invalid-scene-document',
      message: 'Table span exceeds the grid bounds',
      path: `${ELEMENT_PATH}.rows[0].cells[0]`,
    });
  });

  it('rejects overlapping merge rectangles', () => {
    const value = table();
    value.columns = [100, 100, 100];
    value.rows = Array.from({ length: 3 }, (_, rowIndex) => ({
      cells: Array.from({ length: 3 }, (_, columnIndex) =>
        cell(`overlap-${rowIndex}-${columnIndex}`),
      ),
      height: 30,
    }));
    const rows = value.rows as Mutable[];
    const firstOrigin = ((rows[0] as Mutable).cells as Mutable[])[1] as Mutable;
    firstOrigin.colSpan = 2;
    firstOrigin.rowSpan = 3;
    const secondOrigin = (
      (rows[1] as Mutable).cells as Mutable[]
    )[0] as Mutable;
    secondOrigin.colSpan = 3;
    secondOrigin.rowSpan = 2;
    const authored = value.authored as Mutable;
    (authored.transform as Mutable).height = 90;

    expect(validate(value).issues).toContainEqual({
      code: 'invalid-scene-document',
      message: 'Table spans must not overlap',
      path: `${ELEMENT_PATH}.rows[1].cells[0]`,
    });
  });

  it.each([
    ['width', 301, 'column widths'],
    ['height', 81, 'row heights'],
  ])('binds transform %s to exact %s', (key, replacement, source) => {
    const value = table();
    const authored = value.authored as Mutable;
    const transform = authored.transform as Mutable;
    transform[key] = replacement;

    expect(validate(value).issues).toContainEqual({
      code: 'invalid-scene-document',
      message: `Table transform ${key} must equal the sum of its ${source}`,
      path: `${ELEMENT_PATH}.authored.transform.${key}`,
    });
  });

  it('binds exact table dimensions in the preservation scene profile', () => {
    const value = table();
    const authored = value.authored as Mutable;
    (authored.transform as Mutable).width = 301;

    expect(validate(value, 'scene').issues).toContainEqual({
      code: 'invalid-scene-document',
      message: 'Table transform width must equal the sum of its column widths',
      path: `${ELEMENT_PATH}.authored.transform.width`,
    });
  });

  it('never applies table dimension diagnostics to a non-table element', () => {
    const value = table();
    value.type = 'shape';
    value.columns = [100];
    value.rows = [{ cells: [], height: 40 }];
    const messages = validate(value).issues.map(({ message }) => message);

    expect(messages).not.toContain(
      'Table transform width must equal the sum of its column widths',
    );
    expect(messages).not.toContain(
      'Table transform height must equal the sum of its row heights',
    );
  });

  it.each([
    ['column', 'Table transform width must equal the sum of its column widths'],
    ['row', 'Table transform height must equal the sum of its row heights'],
  ])(
    'does not cascade grid-total diagnostics from a malformed %s',
    (kind, message) => {
      const value = table();
      if (kind === 'column') value.columns = ['bad', 200];
      else {
        value.rows = [
          {
            cells: [cell('invalid-row-1'), cell('invalid-row-2')],
            height: 'bad',
          },
          {
            cells: [cell('valid-row-1'), cell('valid-row-2')],
            height: 40,
          },
        ];
      }

      expect(
        validate(value).issues.map((issue) => issue.message),
      ).not.toContain(message);
    },
  );

  it('requires an authored table transform and rejects shape styling', () => {
    const missing = table();
    (missing.authored as Mutable).transform = undefined;
    expect(validate(missing).issues).toContainEqual({
      code: 'unsupported-feature',
      message:
        'Creation profile create-native-v1 requires an authored table transform',
      path: `${ELEMENT_PATH}.authored.transform`,
    });

    for (const key of ['fillColor', 'geometry', 'lineColor', 'lineWidth']) {
      const value = table();
      const authored = value.authored as Mutable;
      authored[key] =
        key === 'geometry' ? 'rect' : key === 'lineWidth' ? 1 : '#123456';
      expect(validate(value).issues).toContainEqual({
        code: 'unsupported-feature',
        message:
          'Creation profile create-native-v1 does not apply shape styling to tables',
        path: `${ELEMENT_PATH}.authored`,
      });
    }
  });
});
