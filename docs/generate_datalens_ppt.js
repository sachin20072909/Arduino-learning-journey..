const PptxGenJS = require('pptxgenjs');

const pptx = new PptxGenJS();
pptx.layout = 'LAYOUT_WIDE';
pptx.author = 'Arena.ai Agent Mode';
pptx.company = 'Arena.ai';
pptx.subject = 'DataLens AI professional presentation';
pptx.title = 'DataLens AI — Professional PPT';
pptx.lang = 'en-US';
pptx.theme = {
  headFontFace: 'Aptos Display',
  bodyFontFace: 'Aptos',
  lang: 'en-US',
};
pptx.writeOptions = { compress: true };

const SW = 13.333;
const SH = 7.5;

const C = {
  deep: '07111F',
  navy: '0F172A',
  navy2: '111827',
  card: '0F1B33',
  card2: '111C31',
  panel: '15233E',
  lineDark: '23324B',
  slate: '334155',
  muted: '9FB0C7',
  soft: '8FA3BF',
  border: 'E2E8F0',
  bg: '07111F',
  white: 'FFFFFF',
  textLight: 'EAF2FF',
  textSoft: 'C7D6EA',
  purple: '7C3AED',
  purple2: '8B5CF6',
  cyan: '06B6D4',
  teal: '14B8A6',
  sky: '0EA5E9',
  green: '10B981',
  amber: 'F59E0B',
  rose: 'F43F5E',
  indigo: '4F46E5',
  lightPurple: 'F3E8FF',
  lightCyan: 'ECFEFF',
  lightTeal: 'F0FDFA',
  lightBlue: 'EFF6FF',
  lightAmber: 'FFFBEB',
  lightRose: 'FFF1F2',
};

function addFooter(slide, n) {
  slide.addShape(pptx.ShapeType.line, {
    x: 0.55, y: 7.05, w: 12.2, h: 0,
    line: { color: C.lineDark, pt: 1 }
  });
  slide.addText('DataLens AI', {
    x: 0.6, y: 7.08, w: 1.6, h: 0.18,
    fontFace: 'Aptos', fontSize: 8, color: C.soft,
    margin: 0
  });
  slide.addText(String(n).padStart(2, '0'), {
    x: 12.1, y: 7.06, w: 0.6, h: 0.18,
    fontFace: 'Aptos', fontSize: 8, color: C.soft,
    align: 'right', margin: 0
  });
}

function addThemedBackground(slide) {
  slide.background = { color: C.deep };
  slide.addShape(pptx.ShapeType.rect, {
    x: 0, y: 0, w: SW, h: 0.12,
    line: { color: C.purple, pt: 0 },
    fill: { color: C.purple }
  });
  slide.addShape(pptx.ShapeType.rect, {
    x: 8.2, y: 0.12, w: 5.133, h: 0.04,
    line: { color: C.cyan, pt: 0 },
    fill: { color: C.cyan }
  });
  slide.addShape(pptx.ShapeType.ellipse, {
    x: 10.0, y: -0.75, w: 3.6, h: 3.6,
    line: { color: C.purple2, pt: 0 }, fill: { color: C.purple2, transparency: 78 }
  });
  slide.addShape(pptx.ShapeType.ellipse, {
    x: -0.75, y: 5.1, w: 2.7, h: 2.7,
    line: { color: C.cyan, pt: 0 }, fill: { color: C.cyan, transparency: 82 }
  });
  slide.addShape(pptx.ShapeType.ellipse, {
    x: 10.9, y: 4.65, w: 2.05, h: 2.05,
    line: { color: C.teal, pt: 0 }, fill: { color: C.teal, transparency: 84 }
  });
  slide.addShape(pptx.ShapeType.roundRect, {
    x: 0.42, y: 1.28, w: 12.48, h: 5.5,
    rectRadius: 0.16,
    line: { color: C.lineDark, pt: 1 },
    fill: { color: C.card2, transparency: 12 }
  });
}

function addHeader(slide, title, subtitle, n) {
  addThemedBackground(slide);
  slide.addShape(pptx.ShapeType.roundRect, {
    x: 0.62, y: 0.22, w: 1.62, h: 0.34,
    rectRadius: 0.08,
    line: { color: C.lineDark, pt: 1 },
    fill: { color: C.card }
  });
  slide.addText('DATAlens AI', {
    x: 0.82, y: 0.33, w: 1.2, h: 0.12,
    fontFace: 'Aptos', fontSize: 9.5, bold: true, color: C.textLight,
    margin: 0, align: 'center'
  });
  slide.addText(title, {
    x: 0.62, y: 0.72, w: 8.9, h: 0.4,
    fontFace: 'Aptos Display', fontSize: 24, bold: true, color: C.white,
    margin: 0
  });
  if (subtitle) {
    slide.addText(subtitle, {
      x: 0.62, y: 1.1, w: 10.8, h: 0.26,
      fontFace: 'Aptos', fontSize: 10.5, color: C.textSoft,
      margin: 0
    });
  }
  addFooter(slide, n);
}

function addPill(slide, text, x, y, w, color, textColor = C.white, border = color) {
  slide.addShape(pptx.ShapeType.roundRect, {
    x, y, w, h: 0.34,
    rectRadius: 0.08,
    line: { color: border, pt: 1 },
    fill: { color }
  });
  slide.addText(text, {
    x: x + 0.06, y: y + 0.05, w: w - 0.12, h: 0.22,
    fontFace: 'Aptos', fontSize: 9, bold: true, color: textColor,
    align: 'center', margin: 0
  });
}

function addCard(slide, opts) {
  const {
    x, y, w, h,
    title,
    body,
    fill = C.white,
    line = C.border,
    accent = null,
    titleColor = C.navy,
    bodyColor = C.muted,
    titleSize = 14,
    bodySize = 10,
    icon = null,
    iconFill = C.lightBlue,
    iconColor = C.purple,
    titleY = 0.16,
    bodyY = 0.52,
    radius = 0.08,
  } = opts;

  slide.addShape(pptx.ShapeType.roundRect, {
    x, y, w, h,
    rectRadius: radius,
    line: { color: line, pt: 1 },
    fill: { color: fill }
  });

  if (accent) {
    slide.addShape(pptx.ShapeType.roundRect, {
      x: x + 0.02, y: y + 0.02, w: 0.08, h: h - 0.04,
      rectRadius: 0.04,
      line: { color: accent, pt: 0 },
      fill: { color: accent }
    });
  }

  let titleX = x + 0.16;
  if (icon) {
    slide.addShape(pptx.ShapeType.ellipse, {
      x: x + 0.16, y: y + 0.14, w: 0.34, h: 0.34,
      line: { color: iconFill, pt: 0 },
      fill: { color: iconFill }
    });
    slide.addText(icon, {
      x: x + 0.16, y: y + 0.2, w: 0.34, h: 0.12,
      fontFace: 'Aptos', fontSize: 11, bold: true, color: iconColor,
      align: 'center', margin: 0
    });
    titleX = x + 0.58;
  }

  if (title) {
    slide.addText(title, {
      x: titleX, y: y + titleY, w: w - (titleX - x) - 0.14, h: 0.24,
      fontFace: 'Aptos Display', fontSize: titleSize, bold: true, color: titleColor,
      margin: 0, fit: 'shrink'
    });
  }
  if (body) {
    slide.addText(body, {
      x: x + 0.16, y: y + bodyY, w: w - 0.32, h: h - bodyY - 0.12,
      fontFace: 'Aptos', fontSize: bodySize, color: bodyColor,
      breakLine: false, margin: 0, fit: 'shrink', valign: 'top'
    });
  }
}

function bulletText(items) {
  return items.map(t => `• ${t}`).join('\n');
}

function addStepCard(slide, num, title, body, x, y, w, h, color) {
  slide.addShape(pptx.ShapeType.roundRect, {
    x, y, w, h,
    rectRadius: 0.08,
    line: { color: C.border, pt: 1 },
    fill: { color: C.white }
  });
  slide.addShape(pptx.ShapeType.ellipse, {
    x: x + 0.16, y: y + 0.14, w: 0.42, h: 0.42,
    line: { color, pt: 0 }, fill: { color }
  });
  slide.addText(String(num), {
    x: x + 0.16, y: y + 0.24, w: 0.42, h: 0.12,
    fontFace: 'Aptos', fontSize: 12, bold: true, color: C.white,
    align: 'center', margin: 0
  });
  slide.addText(title, {
    x: x + 0.68, y: y + 0.14, w: w - 0.82, h: 0.24,
    fontFace: 'Aptos Display', fontSize: 12.5, bold: true, color: C.navy,
    margin: 0, fit: 'shrink'
  });
  slide.addText(body, {
    x: x + 0.16, y: y + 0.66, w: w - 0.32, h: h - 0.76,
    fontFace: 'Aptos', fontSize: 9.2, color: C.muted,
    margin: 0, fit: 'shrink', valign: 'top'
  });
}

function addMiniStat(slide, x, y, w, h, label, value, accent, fill) {
  slide.addShape(pptx.ShapeType.roundRect, {
    x, y, w, h,
    rectRadius: 0.08,
    line: { color: accent, pt: 1 },
    fill: { color: fill }
  });
  slide.addText(label, {
    x: x + 0.16, y: y + 0.14, w: w - 0.32, h: 0.18,
    fontFace: 'Aptos', fontSize: 9.5, bold: true, color: accent,
    margin: 0
  });
  slide.addText(value, {
    x: x + 0.16, y: y + 0.38, w: w - 0.32, h: h - 0.5,
    fontFace: 'Aptos Display', fontSize: 18, bold: true, color: C.navy,
    margin: 0, fit: 'shrink'
  });
}

// Slide 1
{
  const slide = pptx.addSlide();
  slide.background = { color: C.navy };
  slide.addShape(pptx.ShapeType.ellipse, {
    x: 8.9, y: -0.8, w: 4.1, h: 4.1,
    line: { color: C.purple2, pt: 0 }, fill: { color: C.purple2, transparency: 68 }
  });
  slide.addShape(pptx.ShapeType.ellipse, {
    x: 10.7, y: 4.6, w: 2.4, h: 2.4,
    line: { color: C.cyan, pt: 0 }, fill: { color: C.cyan, transparency: 70 }
  });
  slide.addShape(pptx.ShapeType.ellipse, {
    x: -0.7, y: 5.1, w: 2.2, h: 2.2,
    line: { color: C.teal, pt: 0 }, fill: { color: C.teal, transparency: 78 }
  });

  addPill(slide, 'AI-powered database intelligence', 0.72, 0.6, 2.9, C.purple, C.white, C.purple);
  slide.addText('DataLens AI', {
    x: 0.72, y: 1.45, w: 5.8, h: 0.72,
    fontFace: 'Aptos Display', fontSize: 29, bold: true, color: C.white,
    margin: 0
  });
  slide.addText('Building Intelligent LLM Agents for\nDatabase Interaction & Visualization', {
    x: 0.72, y: 2.18, w: 6.3, h: 1.1,
    fontFace: 'Aptos', fontSize: 18.5, bold: true, color: 'D8E6FF',
    margin: 0, fit: 'shrink'
  });
  slide.addText('Ask in Natural Language. Query Data. Visualize Insights.', {
    x: 0.72, y: 3.42, w: 5.8, h: 0.28,
    fontFace: 'Aptos', fontSize: 12.5, color: 'D5E7F7',
    margin: 0
  });

  slide.addText('Team Members', {
    x: 0.72, y: 4.12, w: 1.8, h: 0.2,
    fontFace: 'Aptos', fontSize: 10, bold: true, color: 'A5B4FC', margin: 0
  });

  const memberY = 4.48;
  const members = ['Team Member 1', 'Team Member 2', 'Team Member 3', 'Team Member 4'];
  members.forEach((m, i) => {
    const col = i % 2;
    const row = Math.floor(i / 2);
    addCard(slide, {
      x: 0.72 + col * 2.85,
      y: memberY + row * 0.72,
      w: 2.55,
      h: 0.54,
      title: m,
      fill: '172036',
      line: '2A3653',
      titleColor: C.white,
      titleSize: 11.5,
      body: '',
      bodyY: 0.48,
      accent: col === 0 ? C.purple : C.cyan,
      titleY: 0.18,
    });
  });

  const chips = [
    ['AI / LLM', C.purple],
    ['SQL', C.cyan],
    ['React', C.teal],
    ['Data Visualization', C.indigo],
  ];
  chips.forEach((c, i) => addPill(slide, c[0], 7.55 + (i % 2) * 2.35, 5.15 + Math.floor(i / 2) * 0.56, 1.95, c[1]));

  slide.addText('Technology stack', {
    x: 7.55, y: 4.62, w: 2.2, h: 0.2,
    fontFace: 'Aptos', fontSize: 10, bold: true, color: 'A5B4FC', margin: 0
  });

  slide.addShape(pptx.ShapeType.roundRect, {
    x: 7.45, y: 1.2, w: 5.1, h: 2.75,
    rectRadius: 0.14,
    line: { color: '314161', pt: 1 },
    fill: { color: '111C31' }
  });
  slide.addShape(pptx.ShapeType.roundRect, {
    x: 7.85, y: 1.55, w: 4.3, h: 1.95,
    rectRadius: 0.12,
    line: { color: '334155', pt: 1 },
    fill: { color: '0F172A' }
  });
  slide.addText('Natural Language Query', {
    x: 8.1, y: 1.8, w: 1.9, h: 0.18,
    fontFace: 'Aptos', fontSize: 9, bold: true, color: 'CBD5E1', margin: 0
  });
  slide.addText('“Show me the top 5 products by revenue this quarter.”', {
    x: 8.1, y: 2.06, w: 3.55, h: 0.44,
    fontFace: 'Aptos', fontSize: 11, color: C.white, margin: 0, fit: 'shrink'
  });
  slide.addShape(pptx.ShapeType.line, {
    x: 8.15, y: 2.72, w: 3.3, h: 0,
    line: { color: '334155', pt: 1.2, dash: 'dash' }
  });
  slide.addText('Schema  →  SQL  →  Charts  →  Insights', {
    x: 8.1, y: 2.92, w: 3.6, h: 0.18,
    fontFace: 'Aptos', fontSize: 10, bold: true, color: '7DD3FC', margin: 0
  });

  slide.addText('Professional project presentation', {
    x: 0.72, y: 6.92, w: 2.9, h: 0.18,
    fontFace: 'Aptos', fontSize: 8.5, color: '94A3B8', margin: 0
  });
  slide.addText('01', {
    x: 12.08, y: 6.88, w: 0.45, h: 0.18,
    fontFace: 'Aptos', fontSize: 8.5, color: '94A3B8', align: 'right', margin: 0
  });
}

// Slide 2
{
  const slide = pptx.addSlide();
  addHeader(slide, 'Problem Statement', 'Why natural-language database interaction matters', 2);

  addCard(slide, {
    x: 0.7, y: 1.45, w: 4.4, h: 4.6,
    fill: C.white, line: C.border, accent: C.purple,
    title: 'Data access is still too technical',
    body: 'Most business users depend on SQL experts or analysts to answer everyday questions. This slows down decisions and creates friction between curiosity and insight.',
    titleSize: 20, bodySize: 14, bodyY: 1.1,
    icon: '!', iconFill: C.lightPurple, iconColor: C.purple
  });

  const pains = [
    ['SQL dependency', 'Database analysis usually starts with query-writing skills.'],
    ['User friction', 'Non-technical users struggle to retrieve and understand data.'],
    ['Fragmented workflow', 'Querying, charting, and reporting often happen in separate tools.'],
    ['Time-consuming analysis', 'Manual steps make exploration slower and less interactive.'],
  ];
  pains.forEach((p, i) => {
    const col = i % 2;
    const row = Math.floor(i / 2);
    addCard(slide, {
      x: 5.45 + col * 3.45,
      y: 1.5 + row * 1.65,
      w: 3.1,
      h: 1.25,
      title: p[0],
      body: p[1],
      icon: String(i + 1),
      iconFill: [C.lightPurple, C.lightCyan, C.lightTeal, C.lightBlue][i],
      iconColor: [C.purple, C.cyan, C.teal, C.sky][i],
      accent: [C.purple, C.cyan, C.teal, C.indigo][i],
      bodySize: 9.1,
      titleSize: 12.5,
    });
  });

  addCard(slide, {
    x: 5.45, y: 4.9, w: 6.45, h: 1.15,
    fill: 'EEF2FF', line: 'C7D2FE', accent: C.indigo,
    title: 'One-line explanation', titleColor: C.indigo, titleSize: 12,
    body: 'DataLens AI bridges the gap between natural-language users and complex databases.',
    bodyColor: C.navy, bodySize: 14, bodyY: 0.42
  });
}

// Slide 3
{
  const slide = pptx.addSlide();
  addHeader(slide, 'Existing System & Limitations', 'Current workflows depend on technical skill and multiple disconnected steps', 3);

  const workflowY = 1.6;
  const nodes = [
    ['User', C.lightBlue, C.sky],
    ['SQL Knowledge', C.lightPurple, C.purple],
    ['Database', C.lightTeal, C.teal],
    ['Manual Analysis', C.lightAmber, C.amber],
    ['Visualization', C.lightRose, C.rose],
  ];
  nodes.forEach((n, i) => {
    addCard(slide, {
      x: 0.72 + i * 2.45, y: workflowY, w: 1.9, h: 0.95,
      title: n[0], fill: n[1], line: C.border, accent: n[2],
      titleColor: C.navy, titleSize: 13, body: '', titleY: 0.33
    });
    if (i < nodes.length - 1) {
      slide.addShape(pptx.ShapeType.chevron, {
        x: 2.25 + i * 2.45, y: workflowY + 0.27, w: 0.28, h: 0.4,
        line: { color: 'CBD5E1', pt: 0 }, fill: { color: 'CBD5E1' }
      });
    }
  });

  const limits = [
    'Requires SQL expertise',
    'Manual query writing',
    'Separate visualization tools',
    'Difficult for non-technical users',
    'More time spent on analysis',
    'Limited conversational interaction',
  ];
  limits.forEach((l, i) => {
    const col = i % 3;
    const row = Math.floor(i / 3);
    addCard(slide, {
      x: 0.75 + col * 4.2, y: 3.25 + row * 1.2, w: 3.65, h: 0.9,
      title: l, body: '', titleSize: 12.2, titleY: 0.28,
      fill: C.white, line: C.border,
      icon: String(i + 1), iconFill: C.lightPurple, iconColor: C.purple,
      accent: col === 0 ? C.purple : col === 1 ? C.cyan : C.teal,
    });
  });

  addCard(slide, {
    x: 0.72, y: 5.9, w: 11.2, h: 0.65,
    fill: 'F1F5F9', line: 'E2E8F0', accent: C.navy,
    title: 'Key Gap', titleColor: C.navy, titleSize: 12,
    body: 'Existing systems provide data tools, but not a complete conversational AI-driven workflow.',
    bodyColor: C.slate, bodySize: 12.6, bodyY: 0.22, titleY: 0.2
  });
}

// Slide 4
{
  const slide = pptx.addSlide();
  addHeader(slide, 'Proposed Solution', 'DataLens AI turns a plain-language question into a full analytics workflow', 4);

  addCard(slide, {
    x: 0.72, y: 1.45, w: 3.4, h: 2.05,
    fill: 'EEF2FF', line: 'C7D2FE', accent: C.indigo,
    title: 'ChatGPT-like data analyst', titleColor: C.indigo, titleSize: 17,
    body: 'Users simply ask a question such as “Show me the top 5 products by revenue this quarter.”',
    bodyColor: C.navy, bodySize: 12, bodyY: 0.9,
    icon: 'AI', iconFill: 'E0E7FF', iconColor: C.indigo
  });

  const flow = [
    ['Understand', C.purple],
    ['Inspect\nSchema', C.cyan],
    ['Generate\nSQL', C.teal],
    ['Execute\nQuery', C.sky],
    ['Visualize', C.indigo],
    ['Explain', C.green],
  ];
  flow.forEach((f, i) => {
    const x = 4.55 + i * 1.4;
    slide.addShape(pptx.ShapeType.roundRect, {
      x, y: 2.0, w: 1.15, h: 1.15,
      rectRadius: 0.1,
      line: { color: f[1], pt: 1 }, fill: { color: C.white }
    });
    slide.addText(f[0], {
      x: x + 0.08, y: 2.3, w: 0.99, h: 0.28,
      fontFace: 'Aptos Display', fontSize: 11, bold: true, color: C.navy,
      align: 'center', margin: 0, fit: 'shrink'
    });
    slide.addShape(pptx.ShapeType.ellipse, {
      x: x + 0.39, y: 2.12, w: 0.36, h: 0.12,
      line: { color: f[1], pt: 0 }, fill: { color: f[1] }
    });
    if (i < flow.length - 1) {
      slide.addShape(pptx.ShapeType.chevron, {
        x: x + 1.19, y: 2.42, w: 0.18, h: 0.28,
        line: { color: 'CBD5E1', pt: 0 }, fill: { color: 'CBD5E1' }
      });
    }
  });

  addCard(slide, {
    x: 0.72, y: 4.15, w: 11.2, h: 1.05,
    fill: C.white, line: C.border, accent: C.teal,
    title: 'Main Idea', titleColor: C.teal, titleSize: 12,
    body: 'No SQL expertise is required from the user. DataLens AI handles intent understanding, schema discovery, query generation, execution, visualization, and explanation in one continuous flow.',
    bodyColor: C.slate, bodySize: 12, bodyY: 0.22, titleY: 0.2
  });

  addPill(slide, 'Natural language in', 0.95, 5.65, 1.9, C.purple, C.white, C.purple);
  addPill(slide, 'SQL + insights out', 3.0, 5.65, 2.1, C.cyan, C.white, C.cyan);
  addPill(slide, 'Single intelligent workspace', 5.3, 5.65, 2.55, C.teal, C.white, C.teal);
}

// Slide 5
{
  const slide = pptx.addSlide();
  addHeader(slide, 'Objectives', 'What DataLens AI is designed to achieve', 5);

  const objectives = [
    ['Conversational interface', 'Build a chat-based interface for database interaction.', C.purple, C.lightPurple],
    ['Natural language to SQL', 'Convert user questions into accurate database queries.', C.cyan, C.lightCyan],
    ['Schema discovery', 'Automatically understand tables, columns, and relationships.', C.teal, C.lightTeal],
    ['Query execution', 'Run SQL and return structured results safely.', C.sky, C.lightBlue],
    ['Dynamic charts', 'Generate visual analytics from query outputs.', C.indigo, 'EEF2FF'],
    ['Business insights', 'Explain results in language users can understand.', C.green, 'ECFDF5'],
    ['Conversation context', 'Remember prior questions for follow-up analysis.', C.amber, C.lightAmber],
    ['Transparency & errors', 'Show SQL and handle failures clearly.', C.rose, C.lightRose],
  ];
  objectives.forEach((o, i) => {
    const col = i % 4;
    const row = Math.floor(i / 4);
    addCard(slide, {
      x: 0.72 + col * 3.0, y: 1.55 + row * 2.1, w: 2.72, h: 1.65,
      title: o[0], body: o[1], accent: o[2], fill: o[3], line: C.border,
      icon: '✓', iconFill: C.white, iconColor: o[2], titleSize: 12.2, bodySize: 9.2
    });
  });
}

// Slide 6
{
  const slide = pptx.addSlide();
  addHeader(slide, 'System Architecture', 'The LLM agent acts as the decision-maker and selects the right tool for each request', 6);

  addCard(slide, {
    x: 0.82, y: 2.2, w: 1.7, h: 1.15,
    title: 'User', body: 'Asks a question in natural language',
    fill: C.lightBlue, line: C.border, accent: C.sky,
    icon: 'U', iconFill: C.white, iconColor: C.sky, titleSize: 14, bodySize: 9.2
  });
  slide.addShape(pptx.ShapeType.chevron, { x: 2.68, y: 2.55, w: 0.28, h: 0.32, line: { color: 'CBD5E1', pt: 0 }, fill: { color: 'CBD5E1' } });

  addCard(slide, {
    x: 3.05, y: 2.2, w: 2.0, h: 1.15,
    title: 'Chat Interface', body: 'Collects intent, context, and follow-up prompts',
    fill: C.white, line: C.border, accent: C.purple,
    icon: 'C', iconFill: C.lightPurple, iconColor: C.purple, titleSize: 14, bodySize: 9.2
  });
  slide.addShape(pptx.ShapeType.chevron, { x: 5.2, y: 2.55, w: 0.28, h: 0.32, line: { color: 'CBD5E1', pt: 0 }, fill: { color: 'CBD5E1' } });

  addCard(slide, {
    x: 5.58, y: 1.65, w: 2.35, h: 1.9,
    title: 'LLM / AI Agent', body: 'Understands the question, chooses tools, plans actions, and explains outputs to the user.',
    fill: 'EEF2FF', line: 'C7D2FE', accent: C.indigo,
    icon: 'AI', iconFill: C.white, iconColor: C.indigo, titleSize: 17, bodySize: 10.5, bodyY: 0.92
  });

  const toolNames = [
    ['get_schema', C.purple],
    ['execute_query', C.cyan],
    ['generate_chart', C.teal],
    ['generate_flowchart', C.sky],
    ['explain_data', C.green],
  ];
  toolNames.forEach((t, i) => {
    addCard(slide, {
      x: 3.1 + i * 1.92, y: 4.15, w: 1.72, h: 0.92,
      title: t[0], body: '', fill: C.white, line: C.border, accent: t[1],
      titleSize: 10.4, titleY: 0.31
    });
  });

  slide.addShape(pptx.ShapeType.line, {
    x: 6.75, y: 3.6, w: 0, h: 0.52,
    line: { color: '94A3B8', pt: 1.5 }
  });
  slide.addShape(pptx.ShapeType.line, {
    x: 3.98, y: 4.12, w: 5.76, h: 0,
    line: { color: '94A3B8', pt: 1.5 }
  });

  addCard(slide, {
    x: 9.62, y: 1.95, w: 2.1, h: 1.2,
    title: 'Database', body: 'PostgreSQL or another structured data source',
    fill: C.lightTeal, line: C.border, accent: C.teal,
    icon: 'DB', iconFill: C.white, iconColor: C.teal, titleSize: 14, bodySize: 9.2
  });

  addCard(slide, {
    x: 9.62, y: 4.18, w: 2.1, h: 1.15,
    title: 'Results & Visuals', body: 'Tables, charts, ER diagrams, and insights',
    fill: C.lightCyan, line: C.border, accent: C.cyan,
    icon: 'R', iconFill: C.white, iconColor: C.cyan, titleSize: 13.5, bodySize: 9.2
  });
}

// Slide 7
{
  const slide = pptx.addSlide();
  addHeader(slide, 'Intelligent Agent & Tools', 'DataLens AI does not just answer — it performs data operations through tools', 7);

  const tools = [
    ['get_schema', 'Understands tables, columns, primary keys, foreign keys, and relationships.', C.purple, C.lightPurple, 'S'],
    ['execute_query', 'Runs generated SQL against the database and returns structured output.', C.cyan, C.lightCyan, 'Q'],
    ['generate_chart', 'Transforms query results into visual comparisons, trends, and distributions.', C.teal, C.lightTeal, 'C'],
    ['generate_flowchart', 'Creates ER diagrams or process flows for better schema understanding.', C.sky, C.lightBlue, 'F'],
    ['explain_data', 'Converts raw results into business-friendly insights and narratives.', C.green, 'ECFDF5', 'E'],
  ];
  tools.forEach((t, i) => {
    addCard(slide, {
      x: 0.78 + i * 2.48, y: 1.8, w: 2.18, h: 3.15,
      title: t[0], body: t[1], fill: t[3], line: C.border, accent: t[2],
      icon: t[4], iconFill: C.white, iconColor: t[2], titleSize: 14, bodySize: 10.1, bodyY: 0.95
    });
  });

  addCard(slide, {
    x: 1.1, y: 5.38, w: 10.8, h: 0.74,
    title: 'Key Advantage', titleColor: C.indigo, titleSize: 12,
    body: 'The agent does not stop at text generation. It actively inspects the schema, executes SQL, creates visuals, and explains results using dedicated tools.',
    fill: 'EEF2FF', line: 'C7D2FE', accent: C.indigo, bodyColor: C.navy, bodySize: 11.5, bodyY: 0.22, titleY: 0.2
  });
}

// Slide 8
{
  const slide = pptx.addSlide();
  addHeader(slide, 'How DataLens AI Works', 'A step-by-step view of the intelligent analytics workflow', 8);

  const steps = [
    ['User Query', 'The user asks a question in natural language.', C.purple],
    ['Intent Understanding', 'The agent identifies what information is needed.', C.cyan],
    ['Schema Discovery', 'Relevant tables and columns are selected with get_schema.', C.teal],
    ['SQL Generation', 'The agent writes the appropriate SQL query.', C.sky],
    ['Query Execution', 'execute_query retrieves structured data.', C.indigo],
    ['Visualization', 'Charts and diagrams are created from the result.', C.green],
    ['Explanation', 'The AI converts output into business insights.', C.amber],
  ];

  steps.forEach((s, i) => {
    const x = 0.58 + i * 1.82;
    addStepCard(slide, i + 1, s[0], s[1], x, 2.0, 1.55, 2.35, s[2]);
    if (i < steps.length - 1) {
      slide.addShape(pptx.ShapeType.chevron, {
        x: x + 1.6, y: 2.98, w: 0.18, h: 0.28,
        line: { color: 'CBD5E1', pt: 0 }, fill: { color: 'CBD5E1' }
      });
    }
  });
}

// Slide 9
{
  const slide = pptx.addSlide();
  addHeader(slide, 'User Interface', 'A single workspace for querying, visualization, SQL transparency, and business insights', 9);

  slide.addShape(pptx.ShapeType.roundRect, {
    x: 0.82, y: 1.55, w: 8.35, h: 4.95,
    rectRadius: 0.1,
    line: { color: C.border, pt: 1 }, fill: { color: C.white }
  });
  slide.addShape(pptx.ShapeType.roundRect, {
    x: 0.82, y: 1.55, w: 1.72, h: 4.95,
    rectRadius: 0.1,
    line: { color: C.navy, pt: 0 }, fill: { color: C.navy }
  });
  slide.addText('Missions', { x: 1.05, y: 1.88, w: 0.9, h: 0.18, fontFace: 'Aptos', fontSize: 10, bold: true, color: C.white, margin: 0 });
  ['Sales Q4', 'Channel Mix', 'Margin Check', 'Inventory'].forEach((m, i) => {
    slide.addShape(pptx.ShapeType.roundRect, {
      x: 1.0, y: 2.2 + i * 0.68, w: 1.34, h: 0.46,
      rectRadius: 0.06,
      line: { color: i === 0 ? C.cyan : '24324B', pt: 1 },
      fill: { color: i === 0 ? '182941' : '111C31' }
    });
    slide.addText(m, { x: 1.1, y: 2.36 + i * 0.68, w: 1.12, h: 0.12, fontFace: 'Aptos', fontSize: 8.8, color: 'E2E8F0', margin: 0, fit: 'shrink' });
  });

  slide.addShape(pptx.ShapeType.rect, {
    x: 2.54, y: 1.55, w: 6.63, h: 0.52,
    line: { color: 'E2E8F0', pt: 0 }, fill: { color: 'F8FAFC' }
  });
  slide.addText('DataLens AI workspace', { x: 2.82, y: 1.74, w: 2.1, h: 0.14, fontFace: 'Aptos', fontSize: 10, bold: true, color: C.navy, margin: 0 });
  addPill(slide, 'Processing…', 7.82, 1.66, 0.96, C.lightCyan, C.cyan, C.cyan);

  // chat area
  slide.addShape(pptx.ShapeType.roundRect, {
    x: 2.82, y: 2.26, w: 2.95, h: 1.68,
    rectRadius: 0.06,
    line: { color: C.border, pt: 1 }, fill: { color: 'FCFCFF' }
  });
  slide.addText('Conversation', { x: 3.0, y: 2.45, w: 1.0, h: 0.12, fontFace: 'Aptos', fontSize: 9.3, bold: true, color: C.navy, margin: 0 });
  slide.addShape(pptx.ShapeType.roundRect, { x: 3.0, y: 2.72, w: 2.3, h: 0.42, rectRadius: 0.06, line: { color: 'D6E4FF', pt: 1 }, fill: { color: 'EFF6FF' } });
  slide.addText('Show top 5 products by revenue.', { x: 3.12, y: 2.86, w: 2.02, h: 0.11, fontFace: 'Aptos', fontSize: 8.8, color: C.slate, margin: 0, fit: 'shrink' });
  slide.addShape(pptx.ShapeType.roundRect, { x: 3.36, y: 3.2, w: 2.15, h: 0.46, rectRadius: 0.06, line: { color: 'D1FAE5', pt: 1 }, fill: { color: 'ECFDF5' } });
  slide.addText('Generating SQL, chart, and insight.', { x: 3.48, y: 3.35, w: 1.95, h: 0.11, fontFace: 'Aptos', fontSize: 8.7, color: C.slate, margin: 0, fit: 'shrink' });

  // sql terminal
  slide.addShape(pptx.ShapeType.roundRect, {
    x: 5.98, y: 2.26, w: 2.82, h: 1.68,
    rectRadius: 0.06,
    line: { color: C.border, pt: 1 }, fill: { color: '0B1220' }
  });
  slide.addText('SQL Terminal', { x: 6.18, y: 2.45, w: 1.0, h: 0.12, fontFace: 'Aptos', fontSize: 9.3, bold: true, color: 'D8E6FF', margin: 0 });
  slide.addText('SELECT product_name,\nSUM(revenue) AS total\nFROM sales\nGROUP BY product_name\nORDER BY total DESC\nLIMIT 5;', {
    x: 6.16, y: 2.75, w: 2.42, h: 0.92,
    fontFace: 'Consolas', fontSize: 8.4, color: 'A7F3D0', margin: 0, fit: 'shrink'
  });

  // lower panels
  addCard(slide, { x: 2.82, y: 4.18, w: 1.78, h: 1.86, title: 'Schema Browser', body: 'Tables\nColumns\nKeys\nRelationships', fill: C.white, line: C.border, accent: C.purple, titleSize: 12, bodySize: 9.6 });
  addCard(slide, { x: 4.82, y: 4.18, w: 2.1, h: 1.86, title: 'Data Matrix', body: 'Structured query output in tabular format for validation and export.', fill: C.white, line: C.border, accent: C.cyan, titleSize: 12, bodySize: 9.2 });
  addCard(slide, { x: 7.12, y: 4.18, w: 2.06, h: 1.86, title: 'Insights', body: 'Top product\nTrend summary\nChannel mix\nFollow-up suggestions', fill: C.white, line: C.border, accent: C.teal, titleSize: 12, bodySize: 9.4 });

  slide.addText('Integrated modules', { x: 9.6, y: 1.72, w: 2.1, h: 0.16, fontFace: 'Aptos', fontSize: 10, bold: true, color: C.textSoft, margin: 0 });
  addCard(slide, { x: 9.6, y: 2.0, w: 2.0, h: 4.3, title: 'Interface highlights', body: bulletText([
    'ChatGPT-style conversational interface',
    'Mission / conversation history sidebar',
    'Natural-language query composer',
    'Processing indicator',
    'SQL terminal',
    'Schema browser',
    'Data matrix',
    'Embedded charts',
    'ER diagram',
    'Business insights panel',
    'Export options'
  ]), fill: C.white, line: C.border, accent: C.indigo, titleSize: 13, bodySize: 9.4, bodyY: 0.56 });
}

// Slide 10
{
  const slide = pptx.addSlide();
  addHeader(slide, 'Dynamic Visual Analytics', 'DataLens AI supports multiple chart types for faster understanding', 10);

  // Containers
  const boxes = [
    [0.72, 1.55, 3.0, 2.1, 'Bar Chart', 'Product / revenue comparison'],
    [3.95, 1.55, 3.0, 2.1, 'Line Chart', 'Revenue trends over time'],
    [7.18, 1.55, 2.95, 2.1, 'Donut Chart', 'Revenue by channel'],
    [10.36, 1.55, 2.25, 2.1, 'Scatter Plot', 'Discount vs margin'],
  ];
  boxes.forEach((b, i) => {
    slide.addShape(pptx.ShapeType.roundRect, {
      x: b[0], y: b[1], w: b[2], h: b[3],
      rectRadius: 0.08, line: { color: C.border, pt: 1 }, fill: { color: C.white }
    });
    slide.addText(b[4], { x: b[0] + 0.16, y: b[1] + 0.15, w: 1.2, h: 0.14, fontFace: 'Aptos', fontSize: 10, bold: true, color: C.navy, margin: 0 });
    slide.addText(b[5], { x: b[0] + 0.16, y: b[1] + 0.33, w: b[2] - 0.32, h: 0.12, fontFace: 'Aptos', fontSize: 8.4, color: C.muted, margin: 0 });
  });

  slide.addChart(pptx.ChartType.bar, [{
    name: 'Revenue', labels: ['Aurora Pro', 'Nimbus X', 'Pulse Mini', 'Vertex Hub'], values: [482.4, 430.8, 389.7, 352.1]
  }], {
    x: 0.95, y: 1.98, w: 2.55, h: 1.38,
    showLegend: false, showTitle: false, chartColors: [C.purple],
    catAxisLabelFontSize: 8, valAxisLabelFontSize: 8,
    showValue: false, showCatName: false,
    valGridLine: { color: 'E2E8F0', pt: 1 },
    chartArea: { fill: { color: C.white }, line: { color: C.white, pt: 0 } },
    showBorder: false,
  });

  slide.addChart(pptx.ChartType.line, [{
    name: 'Revenue trend', labels: ['W1', 'W2', 'W3', 'W4', 'W5', 'W6', 'W7', 'W8'], values: [78, 92, 88, 104, 116, 123, 119, 134]
  }], {
    x: 4.15, y: 1.98, w: 2.62, h: 1.38,
    showLegend: false, showTitle: false, chartColors: [C.cyan],
    catAxisLabelFontSize: 8, valAxisLabelFontSize: 8,
    valGridLine: { color: 'E2E8F0', pt: 1 },
    chartArea: { fill: { color: C.white }, line: { color: C.white, pt: 0 } },
    lineSize: 2,
    showBorder: false,
  });

  slide.addChart(pptx.ChartType.doughnut, [{
    name: 'Channels', labels: ['Online', 'Retail', 'Partners'], values: [48, 32, 20]
  }], {
    x: 7.36, y: 1.96, w: 2.55, h: 1.42,
    showLegend: true, legendPos: 'r', legendFontSize: 7,
    showTitle: false, chartColors: [C.teal, C.purple, C.sky],
    holeSize: 55,
    chartArea: { fill: { color: C.white }, line: { color: C.white, pt: 0 } },
    showBorder: false,
  });

  // Manual scatter plot
  slide.addShape(pptx.ShapeType.line, { x: 10.68, y: 3.18, w: 1.35, h: 0, line: { color: '94A3B8', pt: 1.2 } });
  slide.addShape(pptx.ShapeType.line, { x: 10.72, y: 3.18, w: 0, h: -1.0, line: { color: '94A3B8', pt: 1.2 } });
  slide.addText('Discount', { x: 11.28, y: 3.22, w: 0.55, h: 0.1, fontFace: 'Aptos', fontSize: 7.5, color: C.soft, margin: 0 });
  slide.addText('Margin', { x: 10.42, y: 2.12, w: 0.48, h: 0.1, fontFace: 'Aptos', fontSize: 7.5, color: C.soft, margin: 0, rotate: 270 });
  [[10.92, 2.9, C.purple], [11.16, 2.55, C.cyan], [11.38, 2.72, C.teal], [11.72, 2.38, C.indigo], [11.9, 2.78, C.amber]].forEach(p => {
    slide.addShape(pptx.ShapeType.ellipse, { x: p[0], y: p[1], w: 0.09, h: 0.09, line: { color: p[2], pt: 0 }, fill: { color: p[2] } });
  });

  addCard(slide, {
    x: 0.9, y: 4.18, w: 11.75, h: 1.15,
    title: 'Why visual analytics?', titleColor: C.indigo, titleSize: 12,
    body: 'The same data can be understood much faster through visual representation. Your frontend already demonstrates bar, line, donut, and scatter visualizations using Recharts.',
    fill: 'EEF2FF', line: 'C7D2FE', accent: C.indigo, bodyColor: C.navy, bodySize: 12, bodyY: 0.22, titleY: 0.2
  });
}

// Slide 11
{
  const slide = pptx.addSlide();
  addHeader(slide, 'Schema Intelligence & ER Diagram', 'The AI understands database structure before generating queries', 11);

  function tableBox(x, y, w, h, title, lines, accent) {
    slide.addShape(pptx.ShapeType.roundRect, { x, y, w, h, rectRadius: 0.06, line: { color: accent, pt: 1 }, fill: { color: C.white } });
    slide.addShape(pptx.ShapeType.rect, { x, y, w, h: 0.36, line: { color: accent, pt: 0 }, fill: { color: accent } });
    slide.addText(title, { x: x + 0.12, y: y + 0.11, w: w - 0.24, h: 0.12, fontFace: 'Aptos', fontSize: 10.5, bold: true, color: C.white, margin: 0 });
    slide.addText(lines.join('\n'), { x: x + 0.14, y: y + 0.48, w: w - 0.28, h: h - 0.58, fontFace: 'Consolas', fontSize: 8.8, color: C.slate, margin: 0, fit: 'shrink' });
  }

  tableBox(0.9, 1.75, 2.25, 1.55, 'CUSTOMERS', ['PK customer_id', 'customer_name', 'email', 'region'], C.purple);
  tableBox(3.5, 1.75, 2.25, 1.72, 'ORDERS', ['PK order_id', 'FK customer_id', 'order_date', 'channel'], C.cyan);
  tableBox(3.5, 4.0, 2.25, 1.8, 'ORDER_ITEMS', ['PK item_id', 'FK order_id', 'FK product_id', 'quantity', 'unit_price'], C.teal);
  tableBox(6.25, 3.0, 2.25, 1.7, 'PRODUCTS', ['PK product_id', 'product_name', 'category', 'list_price'], C.indigo);

  slide.addShape(pptx.ShapeType.line, { x: 3.15, y: 2.5, w: 0.35, h: 0, line: { color: '94A3B8', pt: 1.4 } });
  slide.addShape(pptx.ShapeType.line, { x: 4.62, y: 3.47, w: 0, h: 0.53, line: { color: '94A3B8', pt: 1.4 } });
  slide.addShape(pptx.ShapeType.line, { x: 5.76, y: 4.82, w: 0.49, h: -0.45, line: { color: '94A3B8', pt: 1.4 } });
  slide.addText('1 → many', { x: 3.22, y: 2.3, w: 0.52, h: 0.1, fontFace: 'Aptos', fontSize: 7.8, color: C.soft, margin: 0 });
  slide.addText('1 → many', { x: 4.74, y: 3.64, w: 0.56, h: 0.1, fontFace: 'Aptos', fontSize: 7.8, color: C.soft, margin: 0 });
  slide.addText('many → 1', { x: 5.84, y: 4.32, w: 0.56, h: 0.1, fontFace: 'Aptos', fontSize: 7.8, color: C.soft, margin: 0 });

  addCard(slide, {
    x: 9.15, y: 1.75, w: 3.0, h: 3.55,
    title: 'What DataLens identifies',
    body: bulletText(['Tables', 'Columns', 'Primary keys', 'Foreign keys', 'Relationships', 'Relevant entities for a query']),
    fill: C.white, line: C.border, accent: C.purple, titleSize: 14, bodySize: 10.4, bodyY: 0.62,
    icon: 'ER', iconFill: C.lightPurple, iconColor: C.purple
  });
  addCard(slide, {
    x: 0.95, y: 6.0, w: 11.2, h: 0.58,
    title: 'Example relationship path: CUSTOMERS → ORDERS → ORDER_ITEMS ← PRODUCTS',
    body: '', fill: 'F0FDFA', line: '99F6E4', accent: C.teal, titleColor: C.teal, titleSize: 11.5, titleY: 0.2
  });
}

// Slide 12
{
  const slide = pptx.addSlide();
  addHeader(slide, 'SQL Transparency', 'DataLens AI shows the generated SQL instead of hiding it', 12);

  const flowSteps = ['Natural Language', 'Generated SQL', 'Query Execution', 'Result'];
  flowSteps.forEach((s, i) => {
    addCard(slide, {
      x: 0.88, y: 1.75 + i * 1.02, w: 2.3, h: 0.78,
      title: s, body: '', fill: C.white, line: C.border,
      accent: [C.purple, C.cyan, C.teal, C.indigo][i], titleSize: 12.2, titleY: 0.28
    });
    if (i < flowSteps.length - 1) {
      slide.addShape(pptx.ShapeType.chevron, { x: 1.9, y: 2.55 + i * 1.02, w: 0.22, h: 0.24, rotate: 90, line: { color: 'CBD5E1', pt: 0 }, fill: { color: 'CBD5E1' } });
    }
  });

  slide.addShape(pptx.ShapeType.roundRect, {
    x: 3.65, y: 1.75, w: 5.95, h: 3.95,
    rectRadius: 0.08,
    line: { color: '1E293B', pt: 1 }, fill: { color: '0B1220' }
  });
  slide.addText('Example SQL', { x: 3.9, y: 2.0, w: 1.2, h: 0.16, fontFace: 'Aptos', fontSize: 10, bold: true, color: 'D8E6FF', margin: 0 });
  slide.addText(
`SELECT p.product_name,
       SUM(oi.quantity * oi.unit_price) AS revenue
FROM orders o
JOIN order_items oi ON o.order_id = oi.order_id
JOIN products p ON oi.product_id = p.product_id
WHERE o.order_date >= DATE_TRUNC('quarter', CURRENT_DATE)
GROUP BY p.product_name
ORDER BY revenue DESC
LIMIT 5;`,
    { x: 3.92, y: 2.34, w: 5.42, h: 2.95, fontFace: 'Consolas', fontSize: 10.2, color: 'A7F3D0', margin: 0, fit: 'shrink' }
  );

  addCard(slide, {
    x: 9.92, y: 1.75, w: 2.48, h: 3.95,
    title: 'Benefits',
    body: bulletText(['Builds user trust', 'Makes AI decisions understandable', 'Helps developers verify queries', 'Supports debugging', 'Creates SQL learning opportunities']),
    fill: C.white, line: C.border, accent: C.indigo, titleSize: 14, bodySize: 10, bodyY: 0.6,
    icon: 'SQL', iconFill: 'EEF2FF', iconColor: C.indigo
  });

  addCard(slide, {
    x: 0.88, y: 6.0, w: 11.52, h: 0.58,
    title: 'Important point: DataLens AI combines AI convenience with query transparency.',
    body: '', fill: 'EEF2FF', line: 'C7D2FE', accent: C.indigo, titleColor: C.indigo, titleSize: 11.8, titleY: 0.2
  });
}

// Slide 13
{
  const slide = pptx.addSlide();
  addHeader(slide, 'Sample Use Case — Sales Analysis', 'Example prompt: “Show me the top 5 products by revenue this quarter.”', 13);

  addCard(slide, {
    x: 0.82, y: 1.55, w: 4.0, h: 1.0,
    title: 'Workflow', body: 'Natural Language → Schema → SQL → Database → Chart → Insights',
    fill: C.white, line: C.border, accent: C.purple, titleSize: 13, bodySize: 10.5, bodyY: 0.4
  });

  addMiniStat(slide, 0.82, 2.88, 1.95, 1.08, 'Top Product', 'Aurora Pro', C.purple, C.lightPurple);
  addMiniStat(slide, 2.92, 2.88, 1.9, 1.08, 'Revenue', '$482.4K', C.cyan, C.lightCyan);
  addMiniStat(slide, 0.82, 4.18, 1.95, 1.08, 'Channels', 'Online', C.teal, C.lightTeal);
  addMiniStat(slide, 2.92, 4.18, 1.9, 1.08, 'Follow-up', 'Trend view', C.indigo, 'EEF2FF');

  slide.addChart(pptx.ChartType.bar, [{
    name: 'Revenue', labels: ['Aurora Pro', 'Nimbus X', 'Pulse Mini', 'Vertex Hub', 'Nova Lite'], values: [482.4, 430.8, 389.7, 352.1, 319.6]
  }], {
    x: 5.3, y: 1.85, w: 6.65, h: 3.65,
    showLegend: false, chartColors: [C.purple],
    catAxisLabelFontSize: 10, valAxisLabelFontSize: 9,
    valGridLine: { color: 'E2E8F0', pt: 1 },
    chartArea: { fill: { color: C.white }, line: { color: C.white, pt: 0 } },
    showBorder: false,
  });

  addCard(slide, {
    x: 0.82, y: 5.62, w: 11.13, h: 0.74,
    title: 'Follow-up query', titleColor: C.teal, titleSize: 12,
    body: '“Now show me the trend for these products over the last year.” This demonstrates conversational data exploration with context retention.',
    fill: 'F0FDFA', line: '99F6E4', accent: C.teal, bodyColor: C.navy, bodySize: 11.5, bodyY: 0.22, titleY: 0.2
  });
}

// Slide 14
{
  const slide = pptx.addSlide();
  addHeader(slide, 'Key Features & Innovation', 'Why DataLens AI stands out from conventional dashboards', 14);

  const features = [
    ['LLM-powered database agent', C.purple, C.lightPurple],
    ['Natural-language database querying', C.cyan, C.lightCyan],
    ['Automatic schema understanding', C.teal, C.lightTeal],
    ['Function / tool calling', C.sky, C.lightBlue],
    ['Automatic SQL generation', C.indigo, 'EEF2FF'],
    ['Dynamic visualizations + insights', C.green, 'ECFDF5'],
  ];
  features.forEach((f, i) => {
    const col = i % 2;
    const row = Math.floor(i / 2);
    addCard(slide, {
      x: 0.82 + col * 3.25, y: 1.65 + row * 1.34, w: 2.85, h: 1.0,
      title: f[0], body: '', fill: f[2], line: C.border, accent: f[1], titleSize: 11.5, titleY: 0.31
    });
  });

  slide.addShape(pptx.ShapeType.roundRect, {
    x: 7.65, y: 1.72, w: 4.65, h: 4.55,
    rectRadius: 0.12,
    line: { color: C.navy, pt: 0 }, fill: { color: C.navy }
  });
  slide.addText('Innovation', {
    x: 7.98, y: 2.0, w: 1.2, h: 0.2,
    fontFace: 'Aptos', fontSize: 10.5, bold: true, color: 'A5B4FC', margin: 0
  });
  slide.addText('Instead of building another dashboard, DataLens AI adds an intelligent conversational layer over databases.', {
    x: 7.98, y: 2.42, w: 3.92, h: 1.02,
    fontFace: 'Aptos Display', fontSize: 20, bold: true, color: C.white, margin: 0, fit: 'shrink'
  });
  slide.addText('It unifies question answering, tool usage, SQL generation, visualization, and explanation in a single decision-driven system.', {
    x: 7.98, y: 3.72, w: 3.72, h: 0.76,
    fontFace: 'Aptos', fontSize: 12.5, color: 'D5E7F7', margin: 0, fit: 'shrink'
  });
  addPill(slide, 'Conversational', 8.0, 5.0, 1.5, C.purple);
  addPill(slide, 'Transparent', 9.7, 5.0, 1.4, C.cyan);
  addPill(slide, 'Actionable', 8.0, 5.48, 1.35, C.teal);
  addPill(slide, 'Visual', 9.52, 5.48, 1.05, C.indigo);
}

// Slide 15
{
  const slide = pptx.addSlide();
  addHeader(slide, 'Technology Stack', 'The application combines modern frontend, AI agent, database, and visualization technologies', 15);

  const leftCol = [
    ['Frontend', 'React + TypeScript', C.purple],
    ['UI', 'Tailwind CSS + UI Components', C.cyan],
    ['AI / Agent', 'LLM + AI SDK', C.teal],
    ['Data Visualization', 'Recharts', C.sky],
  ];
  const rightCol = [
    ['Database', 'PostgreSQL', C.indigo],
    ['Query Language', 'SQL', C.green],
    ['Build Tool', 'Vite', C.amber],
    ['Routing', 'TanStack Router', C.rose],
  ];

  leftCol.forEach((r, i) => addCard(slide, {
    x: 0.85, y: 1.65 + i * 1.05, w: 5.3, h: 0.82,
    title: r[0], body: r[1], fill: C.white, line: C.border, accent: r[2], titleSize: 12.2, bodySize: 11, bodyY: 0.3, titleY: 0.17
  }));
  rightCol.forEach((r, i) => addCard(slide, {
    x: 6.48, y: 1.65 + i * 1.05, w: 5.3, h: 0.82,
    title: r[0], body: r[1], fill: C.white, line: C.border, accent: r[2], titleSize: 12.2, bodySize: 11, bodyY: 0.3, titleY: 0.17
  }));

  addCard(slide, {
    x: 1.15, y: 5.95, w: 11.0, h: 0.62,
    title: 'Project flow: Frontend → AI Agent → Tools → Database → Visualization',
    body: '', fill: 'EEF2FF', line: 'C7D2FE', accent: C.indigo, titleColor: C.indigo, titleSize: 11.8, titleY: 0.2
  });
}

// Slide 16
{
  const slide = pptx.addSlide();
  addHeader(slide, 'Advantages', 'How DataLens AI improves speed, usability, and decision-making', 16);

  const adv = [
    ['Easy to use', 'No SQL expertise required.', C.purple, C.lightPurple],
    ['Faster analysis', 'Reduces manual query creation.', C.cyan, C.lightCyan],
    ['Interactive', 'Supports conversational follow-up questions.', C.teal, C.lightTeal],
    ['Visual', 'Converts results into charts and diagrams.', C.sky, C.lightBlue],
    ['Transparent', 'Shows generated SQL to the user.', C.indigo, 'EEF2FF'],
    ['Business-friendly', 'Turns raw data into understandable insights.', C.green, 'ECFDF5'],
  ];
  adv.forEach((a, i) => {
    const col = i % 3;
    const row = Math.floor(i / 3);
    addCard(slide, {
      x: 0.82 + col * 4.0, y: 1.75 + row * 1.8, w: 3.55, h: 1.35,
      title: a[0], body: a[1], fill: a[3], line: C.border, accent: a[2],
      icon: '+', iconFill: C.white, iconColor: a[2], titleSize: 13, bodySize: 9.8
    });
  });
  addCard(slide, {
    x: 1.05, y: 5.75, w: 11.0, h: 0.7,
    title: 'Scalable concept: the same intelligent layer can be extended to other databases and business domains.',
    body: '', fill: 'F8FAFC', line: 'CBD5E1', accent: C.navy, titleColor: C.navy, titleSize: 11.6, titleY: 0.24
  });
}

// Slide 17
{
  const slide = pptx.addSlide();
  addHeader(slide, 'Future Scope', 'A roadmap toward a universal AI interface for enterprise data', 17);

  const cols = [
    {
      title: 'Platform Expansion',
      bullets: ['Support MySQL, MongoDB, and other databases', 'Multi-database connections', 'Dashboard generation from conversation'],
      color: C.purple, fill: C.lightPurple,
    },
    {
      title: 'Smarter Intelligence',
      bullets: ['More advanced AI agents', 'Predictive analytics', 'Advanced anomaly detection', 'Voice-based interaction'],
      color: C.cyan, fill: C.lightCyan,
    },
    {
      title: 'Enterprise Readiness',
      bullets: ['Automated report generation', 'Role-based access and authentication', 'Real-time database monitoring'],
      color: C.teal, fill: C.lightTeal,
    },
  ];
  cols.forEach((col, i) => addCard(slide, {
    x: 0.85 + i * 4.05, y: 1.8, w: 3.45, h: 3.7,
    title: col.title, body: bulletText(col.bullets), fill: col.fill, line: C.border, accent: col.color,
    icon: String(i + 1), iconFill: C.white, iconColor: col.color, titleSize: 14, bodySize: 10, bodyY: 0.66
  }));
  addCard(slide, {
    x: 1.0, y: 5.92, w: 11.0, h: 0.64,
    title: 'Future vision: A universal AI interface for interacting with enterprise data.',
    body: '', fill: 'EEF2FF', line: 'C7D2FE', accent: C.indigo, titleColor: C.indigo, titleSize: 11.8, titleY: 0.22
  });
}

// Slide 18
{
  const slide = pptx.addSlide();
  slide.background = { color: C.navy };
  slide.addShape(pptx.ShapeType.ellipse, {
    x: 9.4, y: -0.6, w: 3.7, h: 3.7,
    line: { color: C.purple2, pt: 0 }, fill: { color: C.purple2, transparency: 72 }
  });
  slide.addShape(pptx.ShapeType.ellipse, {
    x: -0.5, y: 4.7, w: 2.5, h: 2.5,
    line: { color: C.teal, pt: 0 }, fill: { color: C.teal, transparency: 78 }
  });

  addPill(slide, 'Conclusion', 0.8, 0.82, 1.2, C.purple);
  slide.addText('DataLens AI transforms database interaction\nfrom a technical task into a natural conversation.', {
    x: 0.8, y: 1.5, w: 7.0, h: 1.2,
    fontFace: 'Aptos Display', fontSize: 24, bold: true, color: C.white,
    margin: 0, fit: 'shrink'
  });
  slide.addText('It combines LLMs, agent tools, SQL, databases, visualization, and insights to deliver an end-to-end intelligent analytics experience.', {
    x: 0.82, y: 3.02, w: 6.4, h: 0.64,
    fontFace: 'Aptos', fontSize: 13, color: 'D5E7F7', margin: 0, fit: 'shrink'
  });

  const formula = [
    ['LLM', C.purple], ['Agent Tools', C.cyan], ['SQL', C.teal], ['Visualization', C.indigo], ['Insights', C.green]
  ];
  formula.forEach((f, i) => addPill(slide, f[0], 0.9 + i * 1.62, 4.25, 1.38, f[1]));
  slide.addText('+', { x: 2.25, y: 4.32, w: 0.12, h: 0.1, fontFace: 'Aptos', fontSize: 15, bold: true, color: 'CBD5E1', margin: 0 });
  slide.addText('+', { x: 3.86, y: 4.32, w: 0.12, h: 0.1, fontFace: 'Aptos', fontSize: 15, bold: true, color: 'CBD5E1', margin: 0 });
  slide.addText('+', { x: 5.48, y: 4.32, w: 0.12, h: 0.1, fontFace: 'Aptos', fontSize: 15, bold: true, color: 'CBD5E1', margin: 0 });
  slide.addText('+', { x: 7.1, y: 4.32, w: 0.12, h: 0.1, fontFace: 'Aptos', fontSize: 15, bold: true, color: 'CBD5E1', margin: 0 });

  slide.addShape(pptx.ShapeType.roundRect, {
    x: 7.95, y: 1.62, w: 4.4, h: 3.05,
    rectRadius: 0.14,
    line: { color: '314161', pt: 1 }, fill: { color: '111C31' }
  });
  slide.addText('Final line', { x: 8.28, y: 1.96, w: 1.0, h: 0.16, fontFace: 'Aptos', fontSize: 10, bold: true, color: 'A5B4FC', margin: 0 });
  slide.addText('“Instead of learning how to query the data, users can simply ask the data.”', {
    x: 8.28, y: 2.38, w: 3.45, h: 1.0,
    fontFace: 'Aptos Display', fontSize: 20, bold: true, color: C.white,
    margin: 0, fit: 'shrink'
  });
  slide.addText('Thank you', { x: 8.28, y: 4.02, w: 1.4, h: 0.18, fontFace: 'Aptos', fontSize: 11, color: 'D5E7F7', margin: 0 });

  slide.addText('18', {
    x: 12.1, y: 6.88, w: 0.42, h: 0.18,
    fontFace: 'Aptos', fontSize: 8.5, color: '94A3B8', align: 'right', margin: 0
  });
}

pptx.writeFile({ fileName: 'docs/DataLens_AI_Professional_Presentation.pptx' })
  .then(() => console.log('Created docs/DataLens_AI_Professional_Presentation.pptx'))
  .catch((err) => {
    console.error(err);
    process.exit(1);
  });
