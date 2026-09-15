const PptxGenJS = require('pptxgenjs');

const pptx = new PptxGenJS();
pptx.layout = 'LAYOUT_WIDE';
pptx.author = 'Arena.ai Agent Mode';
pptx.company = 'Arena.ai';
pptx.subject = 'DataLens AI premium royal blue presentation';
pptx.title = 'DataLens AI — Premium Royal Blue Presentation';
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
  bg: '061533',
  bg2: '0A1E49',
  royal: '2563EB',
  royal2: '1D4ED8',
  royal3: '3B82F6',
  cyan: '22D3EE',
  teal: '14B8A6',
  indigo: '6366F1',
  purple: '8B5CF6',
  gold: 'F8C146',
  green: '10B981',
  white: 'FFFFFF',
  text: 'EAF2FF',
  text2: 'C9D9F6',
  muted: '94A3C4',
  line: '2A4D89',
  glass: '10254F',
  glass2: '0D2147',
  card: '0F234A',
  dark: '08152F',
  danger: 'F43F5E',
};

const IMG = {
  hero: 'docs/images/hero_ai.png',
  problem: 'docs/images/problem_sql.png',
  agent: 'docs/images/agent_tools.png',
  dashboard: 'docs/images/dashboard_ui.png',
  dataviz: 'docs/images/data_viz.png',
  future: 'docs/images/future_enterprise.png',
  schema: 'docs/images/schema_er.png',
  sales: 'docs/images/sales_usecase.png',
  sql: 'docs/images/sql_transparency.png',
};

function addBg(slide) {
  slide.background = { color: C.bg };
  slide.addShape(pptx.ShapeType.rect, {
    x: 0, y: 0, w: SW, h: SH,
    line: { color: C.bg, pt: 0 }, fill: { color: C.bg }
  });
  slide.addShape(pptx.ShapeType.rect, {
    x: 0, y: 0, w: SW, h: 0.14,
    line: { color: C.royal2, pt: 0 }, fill: { color: C.royal2 }
  });
  slide.addShape(pptx.ShapeType.rect, {
    x: 8.6, y: 0.14, w: 4.733, h: 0.04,
    line: { color: C.cyan, pt: 0 }, fill: { color: C.cyan }
  });
  slide.addShape(pptx.ShapeType.ellipse, {
    x: 10.15, y: -0.65, w: 3.4, h: 3.4,
    line: { color: C.royal3, pt: 0 }, fill: { color: C.royal3, transparency: 82 }
  });
  slide.addShape(pptx.ShapeType.ellipse, {
    x: -0.8, y: 5.05, w: 2.9, h: 2.9,
    line: { color: C.cyan, pt: 0 }, fill: { color: C.cyan, transparency: 88 }
  });
  slide.addShape(pptx.ShapeType.ellipse, {
    x: 10.9, y: 4.9, w: 1.8, h: 1.8,
    line: { color: C.purple, pt: 0 }, fill: { color: C.purple, transparency: 88 }
  });
}

function addFooter(slide, n) {
  slide.addShape(pptx.ShapeType.line, {
    x: 0.58, y: 7.03, w: 12.15, h: 0,
    line: { color: C.line, pt: 1 }
  });
  slide.addText('DataLens AI', {
    x: 0.65, y: 7.07, w: 1.5, h: 0.18,
    fontFace: 'Aptos', fontSize: 8, color: C.muted, margin: 0
  });
  slide.addText(String(n).padStart(2, '0'), {
    x: 12.05, y: 7.07, w: 0.55, h: 0.18,
    fontFace: 'Aptos', fontSize: 8, color: C.muted, margin: 0, align: 'right'
  });
}

function addHeader(slide, title, subtitle, n) {
  addBg(slide);
  slide.addShape(pptx.ShapeType.roundRect, {
    x: 0.62, y: 0.24, w: 1.6, h: 0.34,
    rectRadius: 0.06,
    line: { color: C.line, pt: 1 }, fill: { color: C.card }
  });
  slide.addText('DATAlens AI', {
    x: 0.76, y: 0.34, w: 1.32, h: 0.12,
    fontFace: 'Aptos', fontSize: 9.4, bold: true, color: C.text, margin: 0, align: 'center'
  });
  slide.addText(title, {
    x: 0.64, y: 0.76, w: 8.9, h: 0.42,
    fontFace: 'Aptos Display', fontSize: 24, bold: true, color: C.white, margin: 0
  });
  if (subtitle) {
    slide.addText(subtitle, {
      x: 0.64, y: 1.12, w: 10.8, h: 0.26,
      fontFace: 'Aptos', fontSize: 10.5, color: C.text2, margin: 0
    });
  }
  addFooter(slide, n);
}

function pill(slide, text, x, y, w, fill = C.royal2, line = fill, color = C.white) {
  slide.addShape(pptx.ShapeType.roundRect, {
    x, y, w, h: 0.34, rectRadius: 0.08,
    line: { color: line, pt: 1 }, fill: { color: fill }
  });
  slide.addText(text, {
    x: x + 0.06, y: y + 0.05, w: w - 0.12, h: 0.2,
    fontFace: 'Aptos', fontSize: 9, bold: true, color, margin: 0, align: 'center'
  });
}

function glass(slide, x, y, w, h, accent = C.royal3, fill = C.card) {
  slide.addShape(pptx.ShapeType.roundRect, {
    x, y, w, h, rectRadius: 0.12,
    line: { color: C.line, pt: 1 }, fill: { color: fill, transparency: 6 }
  });
  slide.addShape(pptx.ShapeType.roundRect, {
    x: x + 0.02, y: y + 0.02, w: 0.08, h: h - 0.04, rectRadius: 0.04,
    line: { color: accent, pt: 0 }, fill: { color: accent }
  });
}

function addCard(slide, {x, y, w, h, title, body='', accent=C.royal3, fill=C.card, titleSize=14, bodySize=10, titleColor=C.white, bodyColor=C.text2, icon=null}) {
  glass(slide, x, y, w, h, accent, fill);
  if (icon) {
    slide.addShape(pptx.ShapeType.ellipse, {
      x: x + 0.16, y: y + 0.16, w: 0.34, h: 0.34,
      line: { color: accent, pt: 0 }, fill: { color: accent }
    });
    slide.addText(icon, {
      x: x + 0.16, y: y + 0.22, w: 0.34, h: 0.12,
      fontFace: 'Aptos', fontSize: 10.5, bold: true, color: C.white, margin: 0, align: 'center'
    });
    slide.addText(title, {
      x: x + 0.58, y: y + 0.16, w: w - 0.72, h: 0.22,
      fontFace: 'Aptos Display', fontSize: titleSize, bold: true, color: titleColor, margin: 0, fit: 'shrink'
    });
  } else {
    slide.addText(title, {
      x: x + 0.16, y: y + 0.16, w: w - 0.32, h: 0.22,
      fontFace: 'Aptos Display', fontSize: titleSize, bold: true, color: titleColor, margin: 0, fit: 'shrink'
    });
  }
  if (body) {
    slide.addText(body, {
      x: x + 0.16, y: y + 0.52, w: w - 0.32, h: h - 0.64,
      fontFace: 'Aptos', fontSize: bodySize, color: bodyColor, margin: 0, fit: 'shrink', valign: 'top'
    });
  }
}

function bulletText(items) {
  return items.map(t => `• ${t}`).join('\n');
}

function addImagePanel(slide, imgPath, x=8.1, y=1.72, w=4.6, h=4.9, caption='') {
  slide.addShape(pptx.ShapeType.roundRect, {
    x, y, w, h, rectRadius: 0.12,
    line: { color: C.royal3, pt: 1.2 }, fill: { color: C.glass }
  });
  slide.addImage({ path: imgPath, x: x + 0.08, y: y + 0.08, w: w - 0.16, h: h - 0.16 });
  if (caption) {
    slide.addShape(pptx.ShapeType.roundRect, {
      x: x + 0.2, y: y + h - 0.54, w: w - 0.4, h: 0.28, rectRadius: 0.06,
      line: { color: C.line, pt: 1 }, fill: { color: C.dark, transparency: 20 }
    });
    slide.addText(caption, {
      x: x + 0.28, y: y + h - 0.47, w: w - 0.56, h: 0.12,
      fontFace: 'Aptos', fontSize: 8.7, color: C.text, margin: 0, align: 'center', fit: 'shrink'
    });
  }
}

function addMetric(slide, x, y, w, h, label, value, accent) {
  glass(slide, x, y, w, h, accent, C.card);
  slide.addText(label, {
    x: x + 0.16, y: y + 0.16, w: w - 0.32, h: 0.16,
    fontFace: 'Aptos', fontSize: 9.5, bold: true, color: accent, margin: 0
  });
  slide.addText(value, {
    x: x + 0.16, y: y + 0.42, w: w - 0.32, h: h - 0.52,
    fontFace: 'Aptos Display', fontSize: 18, bold: true, color: C.white, margin: 0, fit: 'shrink'
  });
}

// Slide 1 - Title
{
  const s = pptx.addSlide();
  s.background = { color: C.bg };
  s.addImage({ path: IMG.hero, x: 0, y: 0, w: SW, h: SH });
  s.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: SW, h: SH, line: { color: C.bg, pt: 0 }, fill: { color: C.bg, transparency: 46 } });
  s.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: SW, h: 0.14, line: { color: C.royal2, pt: 0 }, fill: { color: C.royal2 } });
  s.addShape(pptx.ShapeType.roundRect, { x: 0.72, y: 0.72, w: 2.9, h: 0.34, rectRadius: 0.08, line: { color: C.cyan, pt: 1 }, fill: { color: C.royal2, transparency: 8 } });
  s.addText('AI-powered database intelligence', { x: 0.82, y: 0.82, w: 2.7, h: 0.12, fontFace: 'Aptos', fontSize: 9.2, bold: true, color: C.white, margin: 0, align: 'center' });
  s.addText('DataLens AI', { x: 0.78, y: 1.55, w: 5.4, h: 0.7, fontFace: 'Aptos Display', fontSize: 31, bold: true, color: C.white, margin: 0 });
  s.addText('Building Intelligent LLM Agents for\nDatabase Interaction & Visualization', { x: 0.78, y: 2.32, w: 6.15, h: 1.02, fontFace: 'Aptos', fontSize: 18.5, bold: true, color: C.text, margin: 0, fit: 'shrink' });
  s.addText('Ask in Natural Language. Query Data. Visualize Insights.', { x: 0.8, y: 3.52, w: 5.6, h: 0.2, fontFace: 'Aptos', fontSize: 12.8, color: C.text2, margin: 0 });
  s.addText('Team Members', { x: 0.82, y: 4.18, w: 1.6, h: 0.16, fontFace: 'Aptos', fontSize: 10, bold: true, color: 'D7E7FF', margin: 0 });
  ['Team Member 1','Team Member 2','Team Member 3','Team Member 4'].forEach((t,i)=>{
    addCard(s, { x: 0.82 + (i%2)*2.72, y: 4.5 + Math.floor(i/2)*0.76, w: 2.42, h: 0.58, title: t, body: '', accent: i%2===0?C.cyan:C.purple, fill:'0A1C3F', titleSize: 11.4 });
  });
  ['AI / LLM','SQL','React','Data Visualization'].forEach((t,i)=>{
    pill(s, t, 7.9 + (i%2)*2.08, 5.1 + Math.floor(i/2)*0.52, i===3?1.8:1.45, [C.royal2,C.cyan,C.teal,C.indigo][i]);
  });
  s.addText('01', { x: 12.08, y: 6.9, w: 0.4, h: 0.18, fontFace: 'Aptos', fontSize: 8.5, color: C.muted, margin: 0, align: 'right' });
}

// Slide 2 - Problem
{
  const s = pptx.addSlide();
  addHeader(s, 'Problem Statement', 'Why natural-language database interaction matters', 2);
  addCard(s, {
    x: 0.72, y: 1.72, w: 6.8, h: 4.78,
    title: 'Data access is still too technical',
    body: bulletText([
      'Database analysis normally requires SQL knowledge.',
      'Non-technical users struggle to access and understand information.',
      'Traditional BI tools require multiple manual steps.',
      'Generating queries, charts, and reports separately is time-consuming.',
      'Organizations need a conversational AI system that turns questions directly into insights.'
    ]) + '\n\nDataLens AI bridges the gap between natural-language users and complex databases.',
    accent: C.cyan, fill: C.card, titleSize: 20, bodySize: 12.1, icon: '!'
  });
  addImagePanel(s, IMG.problem, 8.0, 1.72, 4.7, 4.78, 'Current experience: too technical, too fragmented, too slow');
}

// Slide 3 - Existing system
{
  const s = pptx.addSlide();
  addHeader(s, 'Existing System & Limitations', 'Current workflows depend on technical skill and disconnected tools', 3);
  const flow = ['User','SQL Knowledge','Database','Manual Analysis','Visualization'];
  flow.forEach((t,i)=>{
    addCard(s, { x: 0.82 + i*1.38, y: 1.92, w: 1.15, h: 0.86, title: t, body: '', accent: [C.cyan,C.purple,C.teal,C.gold,C.indigo][i], fill: '102247', titleSize: 11.1 });
    if (i < flow.length-1) s.addShape(pptx.ShapeType.chevron, { x: 1.92 + i*1.38, y: 2.2, w: 0.18, h: 0.24, line: { color: C.text2, pt: 0 }, fill: { color: C.text2 } });
  });
  addCard(s, {
    x: 0.82, y: 3.15, w: 6.7, h: 3.12,
    title: 'Key limitations',
    body: bulletText([
      'Requires SQL expertise',
      'Manual query writing',
      'Separate visualization tools',
      'Difficult for non-technical users',
      'More time spent on analysis',
      'Limited conversational interaction'
    ]) + '\n\nKey gap: existing systems provide data tools, but not a complete conversational AI-driven workflow.',
    accent: C.royal3, fill: C.card, titleSize: 18, bodySize: 11.7, icon: '1'
  });
  addImagePanel(s, IMG.problem, 8.0, 1.72, 4.7, 4.78, 'Existing approach: skill-heavy and tool-fragmented');
}

// Slide 4 - Proposed solution
{
  const s = pptx.addSlide();
  addHeader(s, 'Proposed Solution', 'DataLens AI turns a plain-language question into a full analytics workflow', 4);
  addCard(s, {
    x: 0.76, y: 1.8, w: 4.0, h: 1.65,
    title: 'ChatGPT-like intelligent data analyst',
    body: 'Users simply ask: “Show me the top 5 products by revenue this quarter.”',
    accent: C.purple, fill: C.card, titleSize: 17, bodySize: 11.6, icon: 'AI'
  });
  const steps = ['Understand','Inspect Schema','Generate SQL','Execute Query','Visualize','Explain'];
  steps.forEach((t,i)=>{
    const x = 0.85 + (i%3)*2.25;
    const y = 3.82 + Math.floor(i/3)*1.15;
    addCard(s, { x, y, w: 1.95, h: 0.88, title: t, body: '', accent: [C.cyan,C.teal,C.royal3,C.indigo,C.purple,C.green][i], fill:'102247', titleSize: 11.3 });
  });
  addCard(s, {
    x: 0.8, y: 6.0, w: 6.75, h: 0.58,
    title: 'Main idea: no SQL expertise is required from the user.',
    body: '', accent: C.cyan, fill: '0B1E42', titleSize: 11.7
  });
  addImagePanel(s, IMG.hero, 7.95, 1.72, 4.75, 4.86, 'Natural language in → SQL, charts, insights out');
}

// Slide 5 - Objectives
{
  const s = pptx.addSlide();
  addHeader(s, 'Objectives', 'What DataLens AI is designed to achieve', 5);
  const items = [
    ['Conversational AI interface', C.cyan],
    ['Natural language to SQL', C.royal3],
    ['Automatic schema discovery', C.teal],
    ['Query execution + structured results', C.indigo],
    ['Dynamic charts and diagrams', C.purple],
    ['Business insight generation', C.green],
    ['Conversation context retention', C.gold],
    ['SQL transparency + error handling', C.cyan],
  ];
  items.forEach((it,i)=>{
    addCard(s, { x: 0.82 + (i%2)*3.38, y: 1.76 + Math.floor(i/2)*1.1, w: 3.0, h: 0.84, title: it[0], body: '', accent: it[1], fill: '102247', titleSize: 11.3 });
  });
  addImagePanel(s, IMG.agent, 7.85, 1.72, 4.85, 4.84, 'Clear goals for a trustworthy AI data workspace');
}

// Slide 6 - Architecture
{
  const s = pptx.addSlide();
  addHeader(s, 'System Architecture', 'The LLM acts as the intelligent decision-maker selecting the right tool for each request', 6);
  const nodes = [
    ['User', 1.15, 2.0, 1.6, 0.86, C.cyan],
    ['Chat Interface', 1.15, 3.15, 1.6, 0.86, C.royal3],
    ['LLM / AI Agent', 3.35, 2.58, 2.1, 1.3, C.purple],
    ['Database', 6.05, 2.58, 1.7, 1.0, C.teal],
    ['Visualization + Insights', 3.35, 4.5, 3.0, 0.96, C.indigo],
  ];
  nodes.forEach(n => addCard(s, { x:n[1], y:n[2], w:n[3], h:n[4], title:n[0], body:'', accent:n[5], fill:'102247', titleSize:n[0].length>16?12:13 }));
  s.addShape(pptx.ShapeType.chevron, { x: 1.78, y: 2.88, w: 0.18, h: 0.22, rotate: 90, line: { color: C.text2, pt: 0 }, fill: { color: C.text2 } });
  s.addShape(pptx.ShapeType.chevron, { x: 2.78, y: 3.02, w: 0.18, h: 0.22, line: { color: C.text2, pt: 0 }, fill: { color: C.text2 } });
  s.addShape(pptx.ShapeType.chevron, { x: 5.55, y: 3.02, w: 0.18, h: 0.22, line: { color: C.text2, pt: 0 }, fill: { color: C.text2 } });
  s.addShape(pptx.ShapeType.chevron, { x: 4.76, y: 4.12, w: 0.18, h: 0.22, rotate: 90, line: { color: C.text2, pt: 0 }, fill: { color: C.text2 } });

  ['get_schema','execute_query','generate_chart','generate_flowchart','explain_data'].forEach((t,i)=>{
    addCard(s, { x: 0.88 + i*1.46, y: 5.86, w: 1.28, h: 0.58, title: t, body:'', accent:[C.cyan,C.teal,C.royal3,C.indigo,C.green][i], fill:'0B1E42', titleSize: 8.8 });
  });
  addImagePanel(s, IMG.agent, 8.0, 1.72, 4.7, 4.84, 'Agent orchestration across schema, SQL, charts, and explanations');
}

// Slide 7 - Tools
{
  const s = pptx.addSlide();
  addHeader(s, 'Intelligent Agent & Tools', 'The agent does not just answer — it performs real operations using tools', 7);
  const tools = [
    ['get_schema', 'Understands tables, columns, and relationships.', C.cyan, 'S'],
    ['execute_query', 'Runs generated SQL queries on the database.', C.teal, 'Q'],
    ['generate_chart', 'Creates business-friendly visual analytics.', C.royal3, 'C'],
    ['generate_flowchart', 'Builds ER diagrams or process flows.', C.indigo, 'F'],
    ['explain_data', 'Converts raw outputs into insights.', C.purple, 'E'],
  ];
  tools.forEach((t,i)=> addCard(s, {
    x: 0.78 + (i%2)*3.45,
    y: 1.76 + Math.floor(i/2)*1.45,
    w: 3.05,
    h: 1.12,
    title: t[0],
    body: t[1],
    accent: t[2],
    fill: '102247',
    titleSize: 13,
    bodySize: 10,
    icon: t[3]
  }));
  addCard(s, {
    x: 0.82, y: 6.0, w: 6.8, h: 0.56,
    title: 'Key advantage: the AI actively inspects schema, executes SQL, creates visuals, and explains results.',
    body: '', accent: C.cyan, fill:'0B1E42', titleSize: 11.3
  });
  addImagePanel(s, IMG.agent, 8.0, 1.72, 4.7, 4.84, 'Tool-enabled intelligence instead of plain text answers');
}

// Slide 8 - Working process
{
  const s = pptx.addSlide();
  addHeader(s, 'How DataLens AI Works', 'A step-by-step intelligent analytics workflow', 8);
  const steps = [
    ['1', 'User Query', 'User asks a question in natural language.', C.cyan],
    ['2', 'Intent', 'AI identifies what information is required.', C.teal],
    ['3', 'Schema', 'Relevant tables and columns are discovered.', C.royal3],
    ['4', 'SQL', 'The agent generates the right SQL query.', C.indigo],
    ['5', 'Execution', 'Data is retrieved from the database.', C.purple],
    ['6', 'Visualization', 'Charts or diagrams are generated.', C.green],
    ['7', 'Explanation', 'Business insights are produced.', C.gold],
  ];
  steps.forEach((st,i)=>{
    const x = 0.72 + (i%2)*3.42;
    const y = 1.78 + Math.floor(i/2)*1.14;
    addCard(s, { x, y, w: 3.05, h: 0.94, title: `${st[0]}  ${st[1]}`, body: st[2], accent: st[3], fill:'102247', titleSize: 12.2, bodySize: 9.5 });
  });
  addImagePanel(s, IMG.dashboard, 7.9, 1.72, 4.8, 4.84, 'Single continuous flow from query to explanation');
}

// Slide 9 - UI
{
  const s = pptx.addSlide();
  addHeader(s, 'User Interface', 'A single workspace for querying, transparency, visualization, and insights', 9);
  addImagePanel(s, IMG.dashboard, 0.78, 1.68, 7.9, 4.95, 'Chat + SQL + charts + insights in one integrated workspace');
  addCard(s, {
    x: 8.95, y: 1.92, w: 3.25, h: 4.35,
    title: 'Interface highlights',
    body: bulletText([
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
    ]),
    accent: C.cyan, fill: C.card, titleSize: 15, bodySize: 10.1, icon: 'UI'
  });
}

// Slide 10 - Visualization
{
  const s = pptx.addSlide();
  addHeader(s, 'Dynamic Visual Analytics', 'The same data becomes easier to understand when presented visually', 10);
  addCard(s, { x: 0.8, y: 1.78, w: 3.0, h: 0.88, title: 'Bar Chart', body: 'Product / revenue comparison', accent: C.cyan, fill:'102247', titleSize: 13 });
  addCard(s, { x: 0.8, y: 2.82, w: 3.0, h: 0.88, title: 'Line Chart', body: 'Revenue trends over time', accent: C.teal, fill:'102247', titleSize: 13 });
  addCard(s, { x: 0.8, y: 3.86, w: 3.0, h: 0.88, title: 'Pie / Donut Chart', body: 'Revenue distribution by channel', accent: C.royal3, fill:'102247', titleSize: 13 });
  addCard(s, { x: 0.8, y: 4.9, w: 3.0, h: 0.88, title: 'Scatter Plot', body: 'Relationship between discount and margin', accent: C.indigo, fill:'102247', titleSize: 13 });
  addImagePanel(s, IMG.dataviz, 4.15, 1.72, 4.2, 4.82, 'Recharts-powered visual outputs for fast decision-making');
  addImagePanel(s, IMG.sales, 8.7, 1.72, 3.95, 4.82, 'Visual analytics accelerate insight discovery');
}

// Slide 11 - Schema and ER
{
  const s = pptx.addSlide();
  addHeader(s, 'Schema Intelligence & ER Diagram', 'The AI understands the database structure before generating queries', 11);
  addCard(s, {
    x: 0.8, y: 1.8, w: 3.65, h: 4.7,
    title: 'What DataLens identifies',
    body: bulletText([
      'Tables',
      'Columns',
      'Primary keys',
      'Foreign keys',
      'Relationships',
      'Relevant entities for the current query'
    ]) + '\n\nExample relationship path:\nCUSTOMERS → ORDERS → ORDER_ITEMS ← PRODUCTS',
    accent: C.teal, fill: C.card, titleSize: 16, bodySize: 11.2, icon: 'ER'
  });
  addImagePanel(s, IMG.schema, 4.78, 1.72, 3.85, 4.84, 'Schema awareness improves SQL accuracy and relevance');
  addImagePanel(s, IMG.sql, 8.88, 1.72, 3.82, 4.84, 'Structure-first reasoning before execution');
}

// Slide 12 - SQL transparency
{
  const s = pptx.addSlide();
  addHeader(s, 'SQL Transparency', 'DataLens AI shows the generated SQL instead of hiding it', 12);
  addCard(s, {
    x: 0.82, y: 1.76, w: 3.25, h: 4.88,
    title: 'Why transparency matters',
    body: bulletText([
      'Builds user trust',
      'Makes AI decisions understandable',
      'Helps developers verify queries',
      'Supports debugging',
      'Creates a SQL learning opportunity'
    ]) + '\n\nFlow:\nNatural Language → Generated SQL → Query Execution → Result',
    accent: C.cyan, fill: C.card, titleSize: 15, bodySize: 11, icon: 'SQL'
  });
  addImagePanel(s, IMG.sql, 4.4, 1.72, 4.15, 4.84, 'Transparent query generation pipeline');
  addImagePanel(s, IMG.dashboard, 8.8, 1.72, 3.88, 4.84, 'Users can see both reasoning and output');
}

// Slide 13 - Use case
{
  const s = pptx.addSlide();
  addHeader(s, 'Sample Use Case — Sales Analysis', 'Example prompt: “Show me the top 5 products by revenue this quarter.”', 13);
  addMetric(s, 0.84, 1.84, 2.25, 1.1, 'Top Product', 'Aurora Pro', C.cyan);
  addMetric(s, 3.28, 1.84, 2.25, 1.1, 'Revenue', '$482.4K', C.royal3);
  addMetric(s, 0.84, 3.18, 2.25, 1.1, 'Channel Leader', 'Online', C.teal);
  addMetric(s, 3.28, 3.18, 2.25, 1.1, 'Follow-up', 'Trend view', C.indigo);
  addCard(s, {
    x: 0.84, y: 4.62, w: 4.7, h: 1.86,
    title: 'DataLens performs',
    body: 'Natural Language → Schema → SQL → Database → Chart → Insights\n\nFollow-up query: “Now show me the trend for these products over the last year.”\n\nThis demonstrates conversational data exploration with context retention.',
    accent: C.purple, fill: C.card, titleSize: 15, bodySize: 10.3, icon: '→'
  });
  addImagePanel(s, IMG.sales, 5.95, 1.72, 6.75, 4.84, 'Top products, trends, channels, and insights in one interaction');
}

// Slide 14 - Features + Tech + Advantages
{
  const s = pptx.addSlide();
  addHeader(s, 'Key Features, Technology Stack & Advantages', 'A modern AI analytics product built for usability, transparency, and speed', 14);
  addCard(s, {
    x: 0.8, y: 1.78, w: 3.7, h: 2.12,
    title: 'Key features',
    body: bulletText([
      'LLM-powered database agent',
      'Natural-language querying',
      'Automatic schema understanding',
      'Function / tool calling',
      'Automatic SQL generation',
      'Dynamic visualizations + insights'
    ]),
    accent: C.cyan, fill: C.card, titleSize: 15, bodySize: 10.2, icon: '★'
  });
  addCard(s, {
    x: 0.8, y: 4.16, w: 3.7, h: 2.12,
    title: 'Technology stack',
    body: bulletText([
      'Frontend: React + TypeScript',
      'UI: Tailwind CSS + components',
      'AI / Agent: LLM + AI SDK',
      'Charts: Recharts',
      'Database: PostgreSQL',
      'Build: Vite + TanStack Router'
    ]),
    accent: C.teal, fill: C.card, titleSize: 15, bodySize: 10.1, icon: 'T'
  });
  addCard(s, {
    x: 4.82, y: 4.16, w: 3.28, h: 2.12,
    title: 'Advantages',
    body: bulletText([
      'Easy to use',
      'Faster analysis',
      'Interactive follow-up queries',
      'Visual outputs',
      'Transparent SQL',
      'Business-friendly insights'
    ]),
    accent: C.royal3, fill: C.card, titleSize: 15, bodySize: 10.1, icon: '+'
  });
  addImagePanel(s, IMG.hero, 4.82, 1.72, 3.28, 2.18, 'AI + SQL + charts + insights');
  addImagePanel(s, IMG.future, 8.42, 1.72, 4.26, 4.56, 'Innovation: a conversational layer over enterprise databases');
}

// Slide 15 - Future + Conclusion
{
  const s = pptx.addSlide();
  addHeader(s, 'Future Scope & Conclusion', 'Toward a universal AI interface for enterprise data', 15);
  addCard(s, {
    x: 0.82, y: 1.76, w: 4.25, h: 4.82,
    title: 'Future enhancements',
    body: bulletText([
      'Support MySQL, MongoDB, and other databases',
      'Multi-database connections',
      'More advanced AI agents',
      'Voice-based database interaction',
      'Predictive analytics',
      'Automated report generation',
      'Dashboard generation from conversation',
      'Role-based access and authentication',
      'Advanced anomaly detection',
      'Real-time database monitoring'
    ]),
    accent: C.cyan, fill: C.card, titleSize: 17, bodySize: 10.1, icon: 'F'
  });
  addImagePanel(s, IMG.future, 5.38, 1.72, 3.18, 4.82, 'Future vision: a universal AI interface for enterprise data');
  addCard(s, {
    x: 8.88, y: 1.76, w: 3.82, h: 4.82,
    title: 'Conclusion',
    body: 'DataLens AI transforms database interaction from a technical task into a natural conversation.\n\nIt combines:\nLLM + Agent Tools + SQL + Database + Visualization + Insights\n\nto provide an end-to-end intelligent data analysis experience.\n\nFinal line:\n“Instead of learning how to query the data, users can simply ask the data.”',
    accent: C.purple, fill: C.card, titleSize: 17, bodySize: 11, icon: '✓'
  });
}

pptx.writeFile({ fileName: 'docs/DataLens_AI_Professional_Presentation.pptx' })
  .then(() => console.log('Created docs/DataLens_AI_Professional_Presentation.pptx'))
  .catch((err) => {
    console.error(err);
    process.exit(1);
  });
