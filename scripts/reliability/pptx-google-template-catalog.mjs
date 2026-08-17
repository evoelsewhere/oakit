const SLIDE_WIDTH = 960;
const SLIDE_HEIGHT = 540;

function wrappedLines(text, width, fontSize) {
  if (text === '') return ['\u200B'];
  const maximumCharacters = Math.max(
    4,
    Math.floor((width - 12) / (fontSize * 0.56)),
  );
  const words = text.split(/\s+/);
  const lines = [];
  let current = '';
  for (const word of words) {
    const candidate = current === '' ? word : `${current} ${word}`;
    if (current !== '' && candidate.length > maximumCharacters) {
      lines.push(current);
      current = word;
    } else {
      current = candidate;
    }
  }
  if (current !== '') lines.push(current);
  return lines;
}

const palettes = {
  aurora: {
    accent: '#67E8F9',
    accent2: '#A78BFA',
    background: '#07111F',
    muted: '#94A3B8',
    surface: '#10233D',
    surface2: '#173252',
    text: '#F8FAFC',
  },
  citrus: {
    accent: '#F97316',
    accent2: '#FACC15',
    background: '#FFF7ED',
    muted: '#9A3412',
    surface: '#FFFFFF',
    surface2: '#FFEDD5',
    text: '#431407',
  },
  cobalt: {
    accent: '#38BDF8',
    accent2: '#FB7185',
    background: '#0B1026',
    muted: '#A5B4FC',
    surface: '#172554',
    surface2: '#1E3A8A',
    text: '#EEF2FF',
  },
  forest: {
    accent: '#A3E635',
    accent2: '#2DD4BF',
    background: '#0C1F17',
    muted: '#A7F3D0',
    surface: '#153C2C',
    surface2: '#1F5740',
    text: '#F0FDF4',
  },
  graphite: {
    accent: '#F43F5E',
    accent2: '#F59E0B',
    background: '#111111',
    muted: '#A3A3A3',
    surface: '#1F1F1F',
    surface2: '#2B2B2B',
    text: '#FAFAFA',
  },
  lavender: {
    accent: '#7C3AED',
    accent2: '#EC4899',
    background: '#FAF5FF',
    muted: '#6B21A8',
    surface: '#FFFFFF',
    surface2: '#F3E8FF',
    text: '#2E1065',
  },
  marine: {
    accent: '#22D3EE',
    accent2: '#FDE047',
    background: '#082F49',
    muted: '#BAE6FD',
    surface: '#0C4A6E',
    surface2: '#075985',
    text: '#F0F9FF',
  },
  mint: {
    accent: '#059669',
    accent2: '#F59E0B',
    background: '#ECFDF5',
    muted: '#047857',
    surface: '#FFFFFF',
    surface2: '#D1FAE5',
    text: '#064E3B',
  },
  paper: {
    accent: '#DC2626',
    accent2: '#2563EB',
    background: '#F7F3EA',
    muted: '#78716C',
    surface: '#FFFCF5',
    surface2: '#E7E0D2',
    text: '#292524',
  },
  sunset: {
    accent: '#FB7185',
    accent2: '#FDBA74',
    background: '#2E1065',
    muted: '#E9D5FF',
    surface: '#4C1D95',
    surface2: '#6D28D9',
    text: '#FFF7ED',
  },
};

function textElement(key, text, x, y, width, height, style = {}) {
  const {
    align = 'left',
    bold = false,
    color = '#111827',
    fillColor,
    fontSize = 14,
    geometry = 'rect',
    italic = false,
    lineColor,
    lineWidth,
    rotation,
  } = style;
  const lines = wrappedLines(text, width, fontSize);
  return {
    authored: {
      ...(fillColor === undefined ? {} : { fillColor }),
      ...(geometry === 'rect' ? {} : { geometry }),
      ...(lineColor === undefined ? {} : { lineColor }),
      ...(lineWidth === undefined ? {} : { lineWidth }),
      transform: {
        height,
        ...(rotation === undefined ? {} : { rotation }),
        width,
        x,
        y,
      },
    },
    key,
    resolved: { hidden: false },
    text: {
      body: { anchor: 'center', autoFit: 'shape', wrap: true },
      paragraphs: lines.map((line, index) => ({
        children: [
          {
            key: `${key}-run-${index + 1}`,
            properties: {
              bold,
              color,
              fontFamily: 'Aptos',
              fontSize,
              italic,
              language: 'en-US',
            },
            text: line,
            type: 'run',
          },
        ],
        key: `${key}-paragraph-${index + 1}`,
        properties: { alignment: align },
      })),
    },
    type: 'text',
  };
}

function add(elements, key, text, box, style) {
  elements.push(textElement(key, text, box[0], box[1], box[2], box[3], style));
}

function heading(elements, spec, palette, eyebrow) {
  add(elements, 'eyebrow', eyebrow.toUpperCase(), [56, 34, 430, 24], {
    bold: true,
    color: palette.accent,
    fontSize: 11,
  });
  add(elements, 'title', spec.title, [56, 64, 760, 64], {
    bold: true,
    color: palette.text,
    fontSize: 30,
  });
  add(elements, 'subtitle', spec.subtitle, [56, 130, 700, 38], {
    color: palette.muted,
    fontSize: 13,
  });
}

function footer(elements, spec, palette) {
  add(elements, 'footer-mark', spec.marker, [56, 504, 310, 20], {
    bold: true,
    color: palette.muted,
    fontSize: 9,
  });
  add(
    elements,
    'footer-page',
    spec.index.toString().padStart(2, '0'),
    [860, 496, 46, 28],
    {
      align: 'center',
      bold: true,
      color: palette.background,
      fillColor: palette.accent,
      fontSize: 11,
      geometry: 'ellipse',
    },
  );
}

function coverSplit(spec, palette) {
  const elements = [];
  add(elements, 'left-panel', '', [0, 0, 590, 540], {
    color: palette.surface,
    fillColor: palette.surface,
  });
  add(elements, 'accent-rail', '', [0, 0, 14, 540], {
    color: palette.accent,
    fillColor: palette.accent,
  });
  add(elements, 'edition', spec.kicker, [58, 64, 240, 28], {
    bold: true,
    color: palette.accent,
    fontSize: 11,
  });
  add(elements, 'cover-title', spec.title, [58, 112, 478, 128], {
    bold: true,
    color: palette.text,
    fontSize: 34,
  });
  add(elements, 'cover-subtitle', spec.subtitle, [58, 256, 430, 62], {
    color: palette.muted,
    fontSize: 14,
  });
  add(elements, 'cover-chip', spec.metric, [58, 352, 176, 52], {
    bold: true,
    color: palette.background,
    fillColor: palette.accent,
    fontSize: 18,
    geometry: 'roundRect',
  });
  add(elements, 'orb-large', '', [660, 76, 196, 196], {
    color: palette.accent2,
    fillColor: palette.accent2,
    geometry: 'ellipse',
  });
  add(elements, 'orb-small', '', [616, 250, 92, 92], {
    color: palette.accent,
    fillColor: palette.accent,
    geometry: 'ellipse',
  });
  add(elements, 'right-card', spec.side, [694, 286, 208, 142], {
    align: 'center',
    bold: true,
    color: palette.text,
    fillColor: palette.surface2,
    fontSize: 16,
    geometry: 'roundRect',
    lineColor: palette.accent,
    lineWidth: 1,
  });
  add(elements, 'slash', '/', [834, 416, 74, 74], {
    align: 'center',
    bold: true,
    color: palette.accent,
    fontSize: 42,
    rotation: -12,
  });
  footer(elements, spec, palette);
  return elements;
}

function kpiDashboard(spec, palette) {
  const elements = [];
  heading(elements, spec, palette, spec.kicker);
  const cards = spec.items.slice(0, 4);
  cards.forEach((item, index) => {
    const x = 56 + index * 218;
    add(elements, `kpi-${index}`, item.value, [x, 190, 194, 86], {
      bold: true,
      color: index === 0 ? palette.background : palette.text,
      fillColor: index === 0 ? palette.accent : palette.surface,
      fontSize: 25,
      geometry: 'roundRect',
      ...(index === 0 ? {} : { lineColor: palette.surface2, lineWidth: 1 }),
    });
    add(elements, `kpi-label-${index}`, item.label, [x + 12, 248, 168, 20], {
      color: index === 0 ? palette.background : palette.muted,
      fontSize: 9,
    });
  });
  add(elements, 'chart-panel', '', [56, 304, 588, 170], {
    color: palette.surface,
    fillColor: palette.surface,
    geometry: 'roundRect',
  });
  spec.bars.forEach((value, index) => {
    const height = Math.round(value * 1.12);
    add(
      elements,
      `bar-${index}`,
      '',
      [84 + index * 72, 450 - height, 40, height],
      {
        color:
          index === spec.bars.length - 1 ? palette.accent : palette.accent2,
        fillColor:
          index === spec.bars.length - 1 ? palette.accent : palette.accent2,
        geometry: 'roundRect',
      },
    );
    add(
      elements,
      `bar-label-${index}`,
      `W${index + 1}`,
      [78 + index * 72, 454, 52, 14],
      {
        align: 'center',
        color: palette.muted,
        fontSize: 8,
      },
    );
  });
  add(elements, 'insight-card', spec.side, [670, 304, 234, 170], {
    bold: true,
    color: palette.text,
    fillColor: palette.surface2,
    fontSize: 18,
    geometry: 'roundRect',
  });
  add(elements, 'insight-dot', '↑', [824, 324, 50, 50], {
    align: 'center',
    bold: true,
    color: palette.background,
    fillColor: palette.accent,
    fontSize: 22,
    geometry: 'ellipse',
  });
  footer(elements, spec, palette);
  return elements;
}

function timeline(spec, palette) {
  const elements = [];
  heading(elements, spec, palette, spec.kicker);
  add(elements, 'timeline-rail', '', [96, 296, 768, 8], {
    color: palette.surface2,
    fillColor: palette.surface2,
    geometry: 'roundRect',
  });
  spec.items.slice(0, 5).forEach((item, index) => {
    const x = 72 + index * 176;
    const above = index % 2 === 0;
    add(
      elements,
      `node-${index}`,
      (index + 1).toString(),
      [x + 34, 270, 52, 52],
      {
        align: 'center',
        bold: true,
        color: index === 2 ? palette.background : palette.text,
        fillColor: index === 2 ? palette.accent : palette.surface2,
        fontSize: 14,
        geometry: 'ellipse',
        lineColor: palette.accent,
        lineWidth: 1.5,
      },
    );
    add(
      elements,
      `milestone-${index}`,
      item.label,
      [x, above ? 208 : 344, 122, 36],
      {
        align: 'center',
        bold: true,
        color: palette.text,
        fontSize: 12,
      },
    );
    add(
      elements,
      `milestone-value-${index}`,
      item.value,
      [x, above ? 240 : 380, 122, 22],
      {
        align: 'center',
        color: palette.muted,
        fontSize: 9,
      },
    );
  });
  add(elements, 'timeline-note', spec.side, [250, 438, 460, 38], {
    align: 'center',
    color: palette.background,
    fillColor: palette.accent,
    fontSize: 12,
    geometry: 'roundRect',
  });
  footer(elements, spec, palette);
  return elements;
}

function cardGrid(spec, palette) {
  const elements = [];
  heading(elements, spec, palette, spec.kicker);
  spec.items.slice(0, 6).forEach((item, index) => {
    const column = index % 3;
    const row = Math.floor(index / 3);
    const x = 56 + column * 292;
    const y = 190 + row * 136;
    add(elements, `card-${index}`, '', [x, y, 264, 112], {
      color: palette.surface,
      fillColor: index === 4 ? palette.accent : palette.surface,
      geometry: 'roundRect',
      lineColor: index === 4 ? palette.accent : palette.surface2,
      lineWidth: 1,
    });
    add(
      elements,
      `card-number-${index}`,
      `0${index + 1}`,
      [x + 16, y + 14, 42, 26],
      {
        bold: true,
        color: index === 4 ? palette.background : palette.accent,
        fontSize: 11,
      },
    );
    add(
      elements,
      `card-title-${index}`,
      item.label,
      [x + 16, y + 46, 220, 28],
      {
        bold: true,
        color: index === 4 ? palette.background : palette.text,
        fontSize: 13,
      },
    );
    add(
      elements,
      `card-value-${index}`,
      item.value,
      [x + 16, y + 78, 220, 18],
      {
        color: index === 4 ? palette.background : palette.muted,
        fontSize: 9,
      },
    );
  });
  footer(elements, spec, palette);
  return elements;
}

function quadrant(spec, palette) {
  const elements = [];
  heading(elements, spec, palette, spec.kicker);
  spec.items.slice(0, 4).forEach((item, index) => {
    const column = index % 2;
    const row = Math.floor(index / 2);
    const x = 90 + column * 410;
    const y = 194 + row * 138;
    add(elements, `quadrant-${index}`, '', [x, y, 370, 110], {
      color: palette.surface,
      fillColor: index === 0 ? palette.surface2 : palette.surface,
      geometry: 'roundRect',
      lineColor: index === 0 ? palette.accent : palette.surface2,
      lineWidth: 1,
    });
    add(
      elements,
      `quadrant-index-${index}`,
      String.fromCharCode(65 + index),
      [x + 16, y + 18, 42, 42],
      {
        align: 'center',
        bold: true,
        color: palette.background,
        fillColor: index % 2 === 0 ? palette.accent : palette.accent2,
        fontSize: 14,
        geometry: 'ellipse',
      },
    );
    add(
      elements,
      `quadrant-title-${index}`,
      item.label,
      [x + 72, y + 16, 270, 30],
      {
        bold: true,
        color: palette.text,
        fontSize: 14,
      },
    );
    add(
      elements,
      `quadrant-copy-${index}`,
      item.value,
      [x + 72, y + 52, 270, 38],
      {
        color: palette.muted,
        fontSize: 10,
      },
    );
  });
  footer(elements, spec, palette);
  return elements;
}

function processFlow(spec, palette) {
  const elements = [];
  heading(elements, spec, palette, spec.kicker);
  spec.items.slice(0, 5).forEach((item, index) => {
    const x = 46 + index * 180;
    const y = index % 2 === 0 ? 214 : 250;
    add(elements, `step-${index}`, '', [x, y, 148, 166], {
      color: palette.surface,
      fillColor: index === 4 ? palette.accent : palette.surface,
      geometry: 'roundRect',
      lineColor: index === 4 ? palette.accent : palette.surface2,
      lineWidth: 1,
    });
    add(
      elements,
      `step-number-${index}`,
      (index + 1).toString(),
      [x + 18, y + 18, 42, 42],
      {
        align: 'center',
        bold: true,
        color: index === 4 ? palette.accent : palette.background,
        fillColor: index === 4 ? palette.background : palette.accent,
        fontSize: 14,
        geometry: 'ellipse',
      },
    );
    add(
      elements,
      `step-title-${index}`,
      item.label,
      [x + 16, y + 76, 116, 30],
      {
        align: 'center',
        bold: true,
        color: index === 4 ? palette.background : palette.text,
        fontSize: 12,
      },
    );
    add(
      elements,
      `step-copy-${index}`,
      item.value,
      [x + 14, y + 112, 120, 34],
      {
        align: 'center',
        color: index === 4 ? palette.background : palette.muted,
        fontSize: 9,
      },
    );
    if (index < 4) {
      add(elements, `step-arrow-${index}`, '→', [x + 148, y + 54, 32, 34], {
        align: 'center',
        bold: true,
        color: palette.accent2,
        fontSize: 20,
      });
    }
  });
  footer(elements, spec, palette);
  return elements;
}

function editorial(spec, palette) {
  const elements = [];
  add(
    elements,
    'editorial-index',
    spec.index.toString().padStart(2, '0'),
    [48, 30, 240, 190],
    {
      bold: true,
      color: palette.surface2,
      fontSize: 92,
    },
  );
  add(elements, 'editorial-rule', '', [54, 226, 852, 4], {
    color: palette.accent,
    fillColor: palette.accent,
  });
  add(elements, 'editorial-title', spec.title, [326, 46, 570, 96], {
    bold: true,
    color: palette.text,
    fontSize: 32,
  });
  add(
    elements,
    'editorial-kicker',
    spec.kicker.toUpperCase(),
    [330, 156, 300, 24],
    {
      bold: true,
      color: palette.accent,
      fontSize: 10,
    },
  );
  add(elements, 'editorial-copy', spec.subtitle, [54, 270, 390, 112], {
    color: palette.text,
    fontSize: 16,
  });
  add(elements, 'editorial-quote', `“${spec.side}”`, [500, 266, 370, 120], {
    bold: true,
    color: palette.background,
    fillColor: palette.accent,
    fontSize: 18,
    geometry: 'roundRect',
  });
  spec.items.slice(0, 3).forEach((item, index) => {
    add(
      elements,
      `editorial-item-${index}`,
      item.label,
      [54 + index * 284, 420, 250, 28],
      {
        bold: true,
        color: palette.text,
        fontSize: 12,
      },
    );
    add(
      elements,
      `editorial-value-${index}`,
      item.value,
      [54 + index * 284, 452, 250, 24],
      {
        color: palette.muted,
        fontSize: 9,
      },
    );
  });
  footer(elements, spec, palette);
  return elements;
}

function dataStory(spec, palette) {
  const elements = [];
  heading(elements, spec, palette, spec.kicker);
  add(elements, 'data-panel', '', [56, 190, 570, 282], {
    color: palette.surface,
    fillColor: palette.surface,
    geometry: 'roundRect',
  });
  spec.bars.forEach((value, index) => {
    const width = value * 4;
    add(
      elements,
      `data-label-${index}`,
      spec.items[index]?.label ?? `Series ${index + 1}`,
      [78, 214 + index * 44, 104, 22],
      {
        color: palette.muted,
        fontSize: 9,
      },
    );
    add(elements, `data-track-${index}`, '', [188, 217 + index * 44, 400, 16], {
      color: palette.surface2,
      fillColor: palette.surface2,
      geometry: 'roundRect',
    });
    add(elements, `data-bar-${index}`, '', [188, 217 + index * 44, width, 16], {
      color: index === 2 ? palette.accent2 : palette.accent,
      fillColor: index === 2 ? palette.accent2 : palette.accent,
      geometry: 'roundRect',
    });
  });
  add(elements, 'data-callout', spec.metric, [662, 198, 240, 126], {
    align: 'center',
    bold: true,
    color: palette.background,
    fillColor: palette.accent,
    fontSize: 34,
    geometry: 'roundRect',
  });
  add(elements, 'data-callout-copy', spec.side, [662, 340, 240, 132], {
    bold: true,
    color: palette.text,
    fillColor: palette.surface2,
    fontSize: 16,
    geometry: 'roundRect',
  });
  footer(elements, spec, palette);
  return elements;
}

function productShowcase(spec, palette) {
  const elements = [];
  add(elements, 'product-band', '', [0, 0, 960, 142], {
    color: palette.surface2,
    fillColor: palette.surface2,
  });
  add(
    elements,
    'product-kicker',
    spec.kicker.toUpperCase(),
    [54, 28, 250, 24],
    {
      bold: true,
      color: palette.accent,
      fontSize: 10,
    },
  );
  add(elements, 'product-title', spec.title, [54, 58, 640, 64], {
    bold: true,
    color: palette.text,
    fontSize: 30,
  });
  add(elements, 'product-device', '', [92, 182, 356, 250], {
    color: palette.surface,
    fillColor: palette.surface,
    geometry: 'roundRect',
    lineColor: palette.accent,
    lineWidth: 2,
    rotation: -2,
  });
  add(elements, 'product-screen', spec.metric, [120, 210, 300, 132], {
    align: 'center',
    bold: true,
    color: palette.background,
    fillColor: palette.accent,
    fontSize: 36,
    geometry: 'roundRect',
  });
  add(elements, 'product-screen-copy', spec.side, [132, 354, 276, 44], {
    align: 'center',
    color: palette.muted,
    fontSize: 10,
  });
  spec.items.slice(0, 3).forEach((item, index) => {
    const y = 192 + index * 92;
    add(elements, `feature-dot-${index}`, '✓', [516, y, 42, 42], {
      align: 'center',
      bold: true,
      color: palette.background,
      fillColor: index === 1 ? palette.accent2 : palette.accent,
      fontSize: 14,
      geometry: 'ellipse',
    });
    add(elements, `feature-title-${index}`, item.label, [576, y - 2, 310, 28], {
      bold: true,
      color: palette.text,
      fontSize: 14,
    });
    add(elements, `feature-copy-${index}`, item.value, [576, y + 28, 310, 32], {
      color: palette.muted,
      fontSize: 10,
    });
  });
  footer(elements, spec, palette);
  return elements;
}

function journey(spec, palette) {
  const elements = [];
  heading(elements, spec, palette, spec.kicker);
  add(elements, 'journey-path', '', [88, 314, 788, 10], {
    color: palette.surface2,
    fillColor: palette.surface2,
    geometry: 'roundRect',
    rotation: -4,
  });
  spec.items.slice(0, 5).forEach((item, index) => {
    const x = 74 + index * 174;
    const y = 254 - index * 12;
    add(elements, `journey-halo-${index}`, '', [x, y, 72, 72], {
      color: index === 4 ? palette.accent : palette.surface,
      fillColor: index === 4 ? palette.accent : palette.surface,
      geometry: 'ellipse',
      lineColor: palette.accent,
      lineWidth: 1.5,
    });
    add(
      elements,
      `journey-index-${index}`,
      (index + 1).toString(),
      [x + 14, y + 14, 44, 44],
      {
        align: 'center',
        bold: true,
        color: index === 4 ? palette.background : palette.text,
        fontSize: 14,
      },
    );
    add(
      elements,
      `journey-title-${index}`,
      item.label,
      [x - 30, y + 88, 132, 28],
      {
        align: 'center',
        bold: true,
        color: palette.text,
        fontSize: 11,
      },
    );
    add(
      elements,
      `journey-copy-${index}`,
      item.value,
      [x - 30, y + 118, 132, 34],
      {
        align: 'center',
        color: palette.muted,
        fontSize: 8,
      },
    );
  });
  add(elements, 'journey-callout', spec.side, [286, 444, 390, 38], {
    align: 'center',
    bold: true,
    color: palette.background,
    fillColor: palette.accent,
    fontSize: 11,
    geometry: 'roundRect',
  });
  footer(elements, spec, palette);
  return elements;
}

const layoutBuilders = [
  coverSplit,
  kpiDashboard,
  timeline,
  cardGrid,
  quadrant,
  processFlow,
  editorial,
  dataStory,
  productShowcase,
  journey,
];

const subjects = [
  [
    'Signal & Story',
    'AI strategy',
    'Turn complexity into a decision system',
    'Q4 / 2026',
    'Design the signal. Remove the noise.',
  ],
  [
    'Northstar Metrics',
    'Executive dashboard',
    'A high-clarity operating review',
    '+34%',
    'Momentum is compounding across the core loop',
  ],
  [
    'Launch Trajectory',
    'Product roadmap',
    'From validated insight to category launch',
    '5 stages',
    'Each milestone earns the next investment',
  ],
  [
    'Creative Systems',
    'Brand platform',
    'Six principles for a recognizable experience',
    '06 rules',
    'Consistency should still leave room for surprise',
  ],
  [
    'Market Tensions',
    'Strategy matrix',
    'Where customer urgency meets durable advantage',
    '4 zones',
    'Invest where differentiation and demand overlap',
  ],
  [
    'From Brief to Breakthrough',
    'Innovation process',
    'A repeatable path for high-conviction ideas',
    '5 moves',
    'Make learning visible at every handoff',
  ],
  [
    'The Quiet Majority',
    'Research editorial',
    'What 2,400 overlooked customers told us',
    '2,400',
    'The strongest signal was hiding in plain sight',
  ],
  [
    'Growth, Reframed',
    'Data narrative',
    'Quality revenue is outpacing raw acquisition',
    '82%',
    'Retention—not reach—is the new growth engine',
  ],
  [
    'Orbit Workspace',
    'Product launch',
    'One calm place for ambitious project teams',
    '3.2×',
    'Less coordination. More meaningful progress.',
  ],
  [
    'Moments That Matter',
    'Customer journey',
    'Designing trust across the full experience',
    '5 moments',
    'The handoff is part of the product',
  ],
  [
    'Climate Positive',
    'Sustainability',
    'A practical operating model for net-positive growth',
    '2030',
    'Build the future into today’s decisions',
  ],
  [
    'Portfolio Pulse',
    'Investment update',
    'Resilient growth across a focused portfolio',
    '18.6%',
    'Disciplined capital is widening the advantage',
  ],
  [
    'Care Without Friction',
    'Healthcare experience',
    'A humane service blueprint for every patient',
    '24/7',
    'Confidence begins before the appointment',
  ],
  [
    'Learning in Motion',
    'Education platform',
    'A modular system for active, social learning',
    '91%',
    'Progress becomes visible and motivating',
  ],
  [
    'City / Sea / Story',
    'Travel editorial',
    'A modern field guide to meaningful escapes',
    '48 hrs',
    'Go closer. Notice more.',
  ],
  [
    'Living, Elevated',
    'Real estate',
    'A new standard for connected urban homes',
    '32 homes',
    'Private calm in the center of everything',
  ],
  [
    'Future of Flavor',
    'Food innovation',
    'Signals reshaping what and how we eat',
    '7 signals',
    'Taste leads; technology follows',
  ],
  [
    'After Dark',
    'Music & culture',
    'An identity system built for live energy',
    '120 BPM',
    'Every beat becomes a visual cue',
  ],
  [
    'Momentum Live',
    'Event keynote',
    'Three ideas designed to move a room',
    '1 day',
    'Make the message impossible to forget',
  ],
  [
    'The New Essential',
    'Fashion campaign',
    'A restrained collection with a bold point of view',
    '12 looks',
    'Less, but unmistakably better',
  ],
  [
    'Pipeline with Purpose',
    'Marketing funnel',
    'How relevance compounds from reach to advocacy',
    '4.8×',
    'Earn attention before asking for action',
  ],
  [
    'Trust by Design',
    'Cybersecurity',
    'Moving protection from policy to product behavior',
    '99.99%',
    'The safest path should also be the easiest',
  ],
  [
    'Intelligence at the Edge',
    'Technology vision',
    'A distributed architecture for instant decisions',
    '12 ms',
    'Put useful intelligence where work happens',
  ],
  [
    'One Team, Many Strengths',
    'People strategy',
    'An operating model for aligned autonomy',
    '86%',
    'Clarity creates room for ownership',
  ],
  [
    'Quarter in Focus',
    'Business review',
    'What changed, what worked, what comes next',
    'Q3',
    'Keep the wins. Correct the drag.',
  ],
  [
    'Circular by Default',
    'Operations',
    'Redesigning value chains around continuous use',
    '−41%',
    'Waste is a design decision',
  ],
  [
    'The Conversion Gap',
    'Commerce insight',
    'Why confident shoppers still abandon the cart',
    '63%',
    'Remove uncertainty, not choice',
  ],
  [
    'Category Creator',
    'Investor narrative',
    'A focused path from wedge to durable platform',
    '$28M',
    'Win a painful workflow, then widen the system',
  ],
  [
    'Signals 2027',
    'Trend report',
    'Ten shifts already changing customer expectation',
    '10 shifts',
    'The future arrives unevenly—watch the edges',
  ],
  [
    'Make the Next Move',
    'Closing story',
    'A decisive final frame for ambitious teams',
    'Now',
    'Small action. Visible proof. Shared momentum.',
  ],
];

const itemSets = [
  [
    ['Discover', 'Find the real tension'],
    ['Frame', 'Choose the valuable question'],
    ['Create', 'Build the smallest proof'],
    ['Test', 'Measure changed behavior'],
    ['Scale', 'Systematize the win'],
    ['Learn', 'Feed evidence forward'],
  ],
  [
    ['Awareness', 'A useful first signal'],
    ['Interest', 'A reason to explore'],
    ['Trial', 'Value before commitment'],
    ['Habit', 'A repeatable reward'],
    ['Advocacy', 'A story worth sharing'],
    ['Renewal', 'Trust compounds'],
  ],
  [
    ['People', 'Clear ownership'],
    ['Process', 'Fast feedback loops'],
    ['Platform', 'Shared building blocks'],
    ['Proof', 'Visible outcomes'],
    ['Practice', 'Skills in the flow'],
    ['Progress', 'A steady cadence'],
  ],
];

export const googleTemplateCatalog = subjects.map((subject, offset) => {
  const index = offset + 1;
  const [title, kicker, subtitle, metric, side] = subject;
  const paletteName =
    Object.keys(palettes)[offset % Object.keys(palettes).length];
  const palette = palettes[paletteName];
  const itemSource = itemSets[offset % itemSets.length];
  const items = itemSource.map(([label, value], itemIndex) => ({
    label: offset % 2 === 0 ? label : `${label} ${itemIndex + 1}`,
    value,
  }));
  const spec = {
    bars: [42, 58, 76, 64, 88, 96].map((value, barIndex) =>
      Math.min(98, value + ((offset * 7 + barIndex * 3) % 13) - 6),
    ),
    index,
    items,
    kicker,
    marker: `OAKIT-GSLIDES-${index.toString().padStart(2, '0')}`,
    metric,
    palette: paletteName,
    side,
    slug: `${index.toString().padStart(2, '0')}-${title
      .toLowerCase()
      .replace(/[^a-z0-9]+/g, '-')
      .replace(/^-|-$/g, '')}`,
    subtitle,
    title,
  };
  const elements = layoutBuilders[offset % layoutBuilders.length](
    spec,
    palette,
  );
  return {
    ...spec,
    scene: {
      layouts: [],
      masters: [],
      media: [],
      schemaVersion: 2,
      size: { height: SLIDE_HEIGHT, width: SLIDE_WIDTH },
      slides: [
        {
          backgroundColor: palette.background,
          elements,
          key: `template-${index}-slide`,
          name: title,
        },
      ],
      themes: [],
    },
  };
});
