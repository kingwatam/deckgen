const pptxgen = require("pptxgenjs");
const path = require("path");

// ============================================================
// deckgen — generic presentation template
//
// Run:  npm install && node build_deck.js
// Out:  Presentation.pptx (in this directory)
//
// This is a placeholder template. Replace the content in each
// slide with your own research, data, and text. TODO comments
// mark every section you need to customise.
// ============================================================

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13.333" x 7.5"

// ── Palette ──────────────────────────────────────────────────
const NAVY     = "1B3B4F";
const TEAL     = "3D8B8B";
const GOLD     = "C9A95E";
const DARK     = "2D2D2D";
const GREY     = "6B6B6B";
const LIGHT_BG = "F5F5F2";
const WHITE    = "FFFFFF";
const SUBTLE   = "8A8578";
const MID_GREY = "555555";

const FONT_TITLE = "Cambria";
const FONT_BODY  = "Calibri";

// ── Font sizes ───────────────────────────────────────────────
const FS = {
  EYEBROW:    15,   // category labels ("PAPER 1 · CANCER SURVIVORSHIP")
  SLIDE_TITLE: 26,  // slide titles (Cambria)
  SECTION:     16,  // section headers ("Methods", "KEY TAKEAWAY")
  EMPHASIS:    16,  // emphasis body (citations, implication bars)
  BODY:        15,  // primary body (tags, bullets, cards)
  SECONDARY:   14,  // secondary text (authors, stat labels)
  TERTIARY:    14,  // captions/labels only
  STAT_NUM:    28,  // stat card numbers (Cambria)
  TITLE_1:     44,  // name on title slide (Cambria)
  TITLE_2:     28,  // main title on closing slide (Cambria)
  NUM:         40,  // numbered items (Cambria)
  SUBTITLE:    18,  // subtitles (Cambria)
  BLOCK_HEAD:  16,  // block column headers (Cambria)
};

// ── Layout constants ─────────────────────────────────────────
const LAYOUT = {
  // Result slides: table (left) + key box (right)
  resultSlide: {
    tableX: 0.7, tableY: 2.1, tableW: 7.5, colW: [3.3, 2.1, 2.1],
    boxX: 8.5, boxY: 2.1, boxW: 4.2,
    boxH: 3.5, rowH: 0.583,
    noteBarY: 5.9,
  },
  // Two side-by-side panels + bottom bar
  subscaleSlide: {
    leftX: 0.7, leftW: 5.8,
    rightX: 6.8, rightW: 5.9,
    panelY: 1.6, panelH: 3.5, headerH: 0.7,
    contentY: 2.5, contentH: 2.4,
    barY: 5.4, barH: 1.5,
  },
  // Two top panels + two bottom panels
  contextSlide: {
    leftX: 0.7, leftW: 5.8,
    rightX: 6.8, rightW: 5.9,
    topY: 1.6, topH: 2.8,
    botY: 4.7, botH: 2.4,
  },
};

// ── Helpers ──────────────────────────────────────────────────
function slideHeader(s, eyebrow, title) {
  if (eyebrow) {
    s.addText(eyebrow, {
      x: 0.5, y: 0.3, w: 12.3, h: 0.35,
      fontFace: FONT_BODY, fontSize: FS.EYEBROW, color: GOLD,
      bold: true, charSpacing: 3,
    });
  }
  s.addText(title, {
    x: 0.5, y: 0.65, w: 12.3, h: 0.65,
    fontFace: FONT_TITLE, fontSize: FS.SLIDE_TITLE, color: NAVY, bold: true,
  });
}

function addText(s, text, opts) {
  s.addText(text, Object.assign({ wrap: true, valign: "top" }, opts));
}

function notes(s, text) {
  s.addNotes(text);
}

// Key box: grey bg + gold bar + title + bullet items
function addKeyBox(slide, layout, title, items, opts = {}) {
  const { boxX, boxY, boxW, boxH } = Object.assign({}, layout, opts);
  const charsPerLine = Math.floor((boxW - 0.4) / 0.1);
  const totalLines = items.reduce((sum, t) => sum + Math.ceil(t.length / charsPerLine), 0);
  const estimatedH = totalLines * 0.35 + 0.6;
  if (estimatedH > boxH) {
    console.warn(`"${title}": ${items.length} items may overflow (est ${estimatedH.toFixed(1)}" > ${boxH}")`);
  }
  slide.addShape(pres.shapes.RECTANGLE, {
    x: boxX, y: boxY, w: boxW, h: boxH,
    fill: { color: LIGHT_BG }, line: { type: "none" },
  });
  slide.addShape(pres.shapes.RECTANGLE, {
    x: boxX, y: boxY, w: boxW, h: 0.08,
    fill: { color: GOLD }, line: { type: "none" },
  });
  addText(slide, title, {
    x: boxX + 0.2, y: boxY + 0.2, w: boxW - 0.4, h: 0.3,
    fontFace: FONT_BODY, fontSize: FS.SECTION, color: GOLD, bold: true, charSpacing: 2,
  });
  const textItems = items.map((t) => ({
    text: t,
    options: {
      fontFace: FONT_BODY, fontSize: FS.BODY, color: NAVY,
      bullet: { code: "2022", indent: 10 },
      paraSpaceAfter: 8, breakLine: true, wrap: true,
    },
  }));
  slide.addText(textItems, {
    x: boxX + 0.2, y: boxY + 0.6, w: boxW - 0.4, h: boxH - 0.8,
    valign: "top", wrap: true,
  });
}

// Panel box: colored header bar + bg + accent bar + bullet items
function addPanelBox(slide, opts) {
  const { x, y, w, h, headerH, headerColor, headerTextColor, accentColor, title, items } = opts;
  slide.addShape(pres.shapes.RECTANGLE, {
    x, y, w, h, fill: { color: LIGHT_BG }, line: { type: "none" },
  });
  slide.addShape(pres.shapes.RECTANGLE, {
    x, y, w, h: headerH, fill: { color: headerColor }, line: { type: "none" },
  });
  addText(slide, title, {
    x: x + 0.2, y, w: w - 0.4, h: headerH,
    fontFace: FONT_BODY, fontSize: FS.BLOCK_HEAD, color: headerTextColor,
    bold: true, valign: "middle",
  });
  const textItems = items.map((t) => ({
    text: t,
    options: {
      fontFace: FONT_BODY, fontSize: FS.BODY, color: DARK,
      bullet: { code: "2022", indent: 10 },
      paraSpaceAfter: 8, breakLine: true, wrap: true,
    },
  }));
  slide.addText(textItems, {
    x: x + 0.2, y: y + headerH + 0.1, w: w - 0.4, h: h - headerH - 0.2,
    valign: "top", wrap: true,
  });
}

// Left-accent panel box
function addAccentPanel(slide, opts) {
  const { x, y, w, h, accentColor, title, items } = opts;
  slide.addShape(pres.shapes.RECTANGLE, {
    x, y, w, h, fill: { color: LIGHT_BG }, line: { type: "none" },
  });
  slide.addShape(pres.shapes.RECTANGLE, {
    x, y, w: 0.08, h, fill: { color: accentColor }, line: { type: "none" },
  });
  addText(slide, title, {
    x: x + 0.25, y: y + 0.15, w: w - 0.45, h: 0.35,
    fontFace: FONT_BODY, fontSize: FS.SECTION, color: GOLD, bold: true, charSpacing: 2,
  });
  const textItems = items.map((t) => ({
    text: t,
    options: {
      fontFace: FONT_BODY, fontSize: FS.BODY, color: DARK,
      bullet: { code: "2022", indent: 10 },
      paraSpaceAfter: 8, breakLine: true, wrap: true,
    },
  }));
  slide.addText(textItems, {
    x: x + 0.25, y: y + 0.6, w: w - 0.45, h: h - 0.8,
    valign: "top", wrap: true,
  });
}

// Result table
function addResultTable(slide, layout, rows) {
  slide.addTable(rows, {
    x: layout.tableX, y: layout.tableY, w: layout.tableW, colW: layout.colW,
    rowH: layout.rowH, fontSize: FS.BODY, fontFace: FONT_BODY, color: DARK,
    border: { type: "solid", pt: 0.5, color: "DDDDD8" }, valign: "middle",
  });
}

// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// SLIDE 1: TITLE
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
{
  const s = pres.addSlide();
  s.background = { color: NAVY };
  s.addShape(pres.shapes.RECTANGLE, {
    x: 12.7, y: 0, w: 0.633, h: 7.5,
    fill: { color: GOLD }, line: { type: "none" },
  });

  // TODO: Replace with your own eyebrow text
  addText(s, "YOUR EYEBROW TEXT", {
    x: 1.0, y: 1.35, w: 11, h: 0.4,
    fontFace: FONT_BODY, fontSize: FS.EYEBROW, color: GOLD, charSpacing: 4,
  });

  // TODO: Replace with your presentation title
  addText(s, "Your Presentation Title Goes Here", {
    x: 1.0, y: 1.85, w: 11, h: 1.0,
    fontFace: FONT_TITLE, fontSize: FS.TITLE_2, color: WHITE, bold: true,
  });

  // TODO: Replace with your subtitle
  addText(s, "Research Focus: Your Field  \u00b7  Your Method", {
    x: 1.0, y: 3.0, w: 11, h: 0.5,
    fontFace: FONT_BODY, fontSize: FS.SUBTITLE, color: GOLD,
  });

  // TODO: Replace with your name
  addText(s, "First Author", {
    x: 1.0, y: 3.55, w: 11, h: 0.5,
    fontFace: FONT_BODY, fontSize: FS.SUBTITLE, color: WHITE,
  });

  // TODO: Replace with your institution and event
  addText(s, "Your Institution  \u00b7  Your Event Name", {
    x: 1.0, y: 4.1, w: 11, h: 0.5,
    fontFace: FONT_BODY, fontSize: FS.BLOCK_HEAD, color: "AAAAAA",
  });

  // TODO: Replace with your topic tags
  const tags = ["Topic one", "Topic two", "Topic three"];
  tags.forEach((t, i) => {
    const x = 0.79 + i * 3.9;
    s.addShape(pres.shapes.RECTANGLE, {
      x, y: 5.39, w: 3.7, h: 0.6,
      fill: { color: LIGHT_BG }, line: { type: "none" },
    });
    s.addShape(pres.shapes.RECTANGLE, {
      x, y: 5.39, w: 0.06, h: 0.6,
      fill: { color: i === 0 ? MID_GREY : i === 1 ? TEAL : GOLD },
      line: { type: "none" },
    });
    addText(s, t, {
      x: x + 0.2, y: 5.39, w: 3.5, h: 0.6,
      fontFace: FONT_BODY, fontSize: FS.BODY, color: DARK, bold: true, valign: "middle",
    });
  });

  notes(s, "Good morning/afternoon everyone. My name is [Name] and today I will be presenting our study on [topic]. This work was conducted at [Institution] in collaboration with [Collaborators]. I will walk you through our research question, methods, key findings, and what these results mean for practice and future research.");
}

// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// SLIDE 2: PAPER TITLE & CITATION
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
{
  const s = pres.addSlide();
  s.background = { color: WHITE };
  slideHeader(s, "THE STUDY", "Title & citation");

  // Title card
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.7, y: 1.7, w: 12, h: 1.6,
    fill: { color: LIGHT_BG }, line: { type: "none" },
  });

  // TODO: Replace with your paper title
  addText(s, "Your Paper Title Goes Here: A Descriptive Subtitle If Needed", {
    x: 0.9, y: 1.8, w: 11.6, h: 1.4,
    fontFace: FONT_TITLE, fontSize: FS.BLOCK_HEAD, color: NAVY, bold: true, italic: true,
    valign: "middle",
  });

  // TODO: Replace with your author list
  addText(s, "Author A., Author B., Author C., Author D.", {
    x: 0.7, y: 3.5, w: 12, h: 0.4,
    fontFace: FONT_BODY, fontSize: FS.SECONDARY, color: GREY, italic: true,
  });

  // TODO: Replace with your citation details
  addText(s, "Journal Name  \u00b7  Volume(Issue): Pages  \u00b7  Month Year", {
    x: 0.7, y: 3.9, w: 12, h: 0.4,
    fontFace: FONT_BODY, fontSize: FS.EMPHASIS, color: DARK, bold: true,
  });

  // TODO: Replace with your three tag cards
  const tags = [
    { text: "Study design", color: NAVY, textColor: WHITE },
    { text: "Setting", color: TEAL, textColor: WHITE },
    { text: "Key metric", color: GOLD, textColor: NAVY },
  ];
  tags.forEach((t, i) => {
    const x = 0.7 + i * 4.07;
    s.addShape(pres.shapes.RECTANGLE, {
      x, y: 4.7, w: 3.87, h: 0.6,
      fill: { color: t.color }, line: { type: "none" },
    });
    addText(s, t.text, {
      x, y: 4.7, w: 3.87, h: 0.6,
      fontFace: FONT_BODY, fontSize: FS.BODY, color: t.textColor, bold: true,
      align: "center", valign: "middle",
    });
  });

  // TODO: Replace with your research question
  addText(s, "What is the main question your study addresses?", {
    x: 0.7, y: 5.6, w: 12, h: 0.8,
    fontFace: FONT_BODY, fontSize: FS.BODY, color: GREY, italic: true,
  });

  notes(s, "This slide shows the full citation for the paper we are discussing. The study was published in [Journal] and represents [X] years of work. The key contribution is [one-sentence summary of what the paper does that is novel].");
}

// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// SLIDE 3: BACKGROUND & RESEARCH QUESTION
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
{
  const s = pres.addSlide();
  s.background = { color: WHITE };
  slideHeader(s, "BACKGROUND", "Why this study matters");

  // Left: Context (accent panel)
  addAccentPanel(s, {
    x: 0.7, y: 1.6, w: 5.8, h: 3.0, accentColor: NAVY,
    title: "CONTEXT",
    // TODO: Replace with your background points
    items: [
      "Background point one",
      "Background point two",
      "Background point three",
      "Gap in existing evidence",
    ],
  });

  // Right: Research question (navy box)
  s.addShape(pres.shapes.RECTANGLE, {
    x: 6.8, y: 1.6, w: 5.9, h: 3.0,
    fill: { color: NAVY }, line: { type: "none" },
  });
  addText(s, "RESEARCH QUESTION", {
    x: 7.0, y: 1.75, w: 5.5, h: 0.35,
    fontFace: FONT_BODY, fontSize: FS.SECTION, color: GOLD, bold: true, charSpacing: 2,
  });
  // TODO: Replace with your research question
  addText(s, "What is your main research question?", {
    x: 7.0, y: 2.2, w: 5.5, h: 2.0,
    fontFace: FONT_TITLE, fontSize: FS.SUBTITLE, color: WHITE, bold: true, valign: "top",
  });

  // Bottom: Gap & hypothesis (accent panel)
  addAccentPanel(s, {
    x: 0.7, y: 4.9, w: 12, h: 2.2, accentColor: TEAL,
    title: "GAP & HYPOTHESIS",
    // TODO: Replace with your gap and hypothesis
    items: [
      "What gap in the literature does your study address?",
      "What is your hypothesis?",
    ],
  });

  notes(s, "The motivation for this study comes from [brief context]. Despite growing attention to [topic], there remains a gap in [specific gap]. Our study addresses this by asking [research question]. This matters because [why the audience should care].");
}

// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// SLIDE 4: STUDY DESIGN
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
{
  const s = pres.addSlide();
  s.background = { color: WHITE };
  slideHeader(s, "STUDY DESIGN", "Cohort, measures, and analytical approach");

  // Stat cards
  const cardY = 1.6, cardH = 1.8, cardW = 3.87, gap = 0.27;
  // TODO: Replace with your study statistics
  const stats = [
    { num: "1,234", label: "participants in your sample", color: NAVY },
    { num: "3", label: "measurement time points", color: TEAL },
    { num: "OLS", label: "your analytical method", color: GOLD },
  ];
  stats.forEach((st, i) => {
    const x = 0.7 + i * (cardW + gap);
    s.addShape(pres.shapes.RECTANGLE, {
      x, y: cardY, w: cardW, h: cardH,
      fill: { color: LIGHT_BG }, line: { type: "none" },
    });
    s.addShape(pres.shapes.RECTANGLE, {
      x, y: cardY, w: cardW, h: 0.08,
      fill: { color: st.color }, line: { type: "none" },
    });
    addText(s, st.num, {
      x: x + 0.1, y: cardY + 0.25, w: cardW - 0.2, h: 0.65,
      fontFace: FONT_TITLE, fontSize: FS.STAT_NUM, color: NAVY, bold: true,
      align: "center", valign: "middle",
    });
    addText(s, st.label, {
      x: x + 0.1, y: cardY + 0.95, w: cardW - 0.2, h: 0.8,
      fontFace: FONT_BODY, fontSize: FS.SECONDARY, color: GREY,
      align: "center", valign: "top",
    });
  });

  // Left: Measures
  addAccentPanel(s, {
    x: 0.7, y: 3.7, w: 5.8, h: 3.0, accentColor: NAVY,
    title: "MEASURES",
    // TODO: Replace with your measures
    items: [
      "Outcome measure one: instrument name",
      "Outcome measure two: instrument name",
      "Outcome measure three: instrument name",
      "Covariates: list here",
    ],
  });

  // Right: Analytical approach
  addAccentPanel(s, {
    x: 6.8, y: 3.7, w: 5.9, h: 3.0, accentColor: TEAL,
    title: "ANALYTICAL APPROACH",
    // TODO: Replace with your analytical methods
    items: [
      "Your primary analytical method",
      "Model specification details",
      "Sensitivity analyses conducted",
      "Software used",
    ],
  });

  notes(s, "We used a [study design] with [N] participants measured at [X] time points. Our primary outcome was [outcome], assessed using [instrument]. We analysed the data using [method], adjusting for [covariates]. All analyses were conducted in [software].");
}

// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// SLIDE 5: RESULTS — TABLE 1
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
{
  const s = pres.addSlide();
  s.background = { color: WHITE };
  slideHeader(s, "RESULTS", "Your first set of results");

  // TODO: Replace with your results subtitle
  addText(s, "Description of what this table shows", {
    x: 0.7, y: 1.6, w: 12, h: 0.35,
    fontFace: FONT_BODY, fontSize: FS.SECTION, color: GOLD, bold: true, charSpacing: 2,
  });

  // TODO: Replace with your result table data
  const rows = [
    [
      { text: "Outcome", options: { bold: true, color: WHITE, fill: { color: NAVY } } },
      { text: "Group A", options: { bold: true, color: WHITE, fill: { color: NAVY }, align: "center" } },
      { text: "Group B", options: { bold: true, color: WHITE, fill: { color: NAVY }, align: "center" } },
    ],
    ["Outcome one",  { text: "\u22120.15 **", options: { align: "center" } }, { text: "\u22120.08", options: { align: "center" } }],
    ["Outcome two",  { text: "\u22120.13 *",  options: { align: "center" } }, { text: "\u22120.11 *", options: { align: "center" } }],
    ["Outcome three",{ text: "\u22120.18 ***", options: { align: "center" } }, { text: "\u22120.12 **", options: { align: "center" } }],
  ];
  addResultTable(s, LAYOUT.resultSlide, rows);

  // Key takeaway box
  addKeyBox(s, LAYOUT.resultSlide, "KEY TAKEAWAY", [
    // TODO: Replace with your key takeaways
    "Main finding one",
    "Main finding two",
    "Main finding three",
  ]);

  // Note bar
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.7, y: 5.9, w: 12, h: 0.6,
    fill: { color: NAVY }, line: { type: "none" },
  });
  // TODO: Replace with your significance note
  addText(s, "* p<.05   ** p<.01   *** p<.001   NS = not significant   \u03b2 = standardized coefficient", {
    x: 0.9, y: 5.95, w: 11.6, h: 0.5,
    fontFace: FONT_BODY, fontSize: FS.SECONDARY, color: WHITE,
  });

  notes(s, "Turning to our first set of results. This table shows [what the table presents]. The key pattern here is [describe the main trend]. Notice that [specific observation about a row or column]. This suggests that [interpretation].");
}

// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// SLIDE 6: RESULTS — TABLE 2
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
{
  const s = pres.addSlide();
  s.background = { color: WHITE };
  slideHeader(s, "RESULTS", "Your second set of results");

  // TODO: Replace with your results subtitle
  addText(s, "Description of what this table shows", {
    x: 0.7, y: 1.6, w: 12, h: 0.35,
    fontFace: FONT_BODY, fontSize: FS.SECTION, color: GOLD, bold: true, charSpacing: 2,
  });

  // TODO: Replace with your result table data
  const rows = [
    [
      { text: "Predictor", options: { bold: true, color: WHITE, fill: { color: NAVY } } },
      { text: "Group A", options: { bold: true, color: WHITE, fill: { color: NAVY }, align: "center" } },
      { text: "Group B", options: { bold: true, color: WHITE, fill: { color: NAVY }, align: "center" } },
    ],
    [
      { text: "Predictor one", options: { bold: true } },
      { text: "\u22120.21 *", options: { align: "center", bold: true } },
      { text: "\u22120.23 ***", options: { align: "center", bold: true } },
    ],
    ["Predictor two", { text: "NS", options: { align: "center" } }, { text: "\u22120.16 ***", options: { align: "center" } }],
    ["Predictor three",{ text: "\u22120.15 ***", options: { align: "center" } }, { text: "\u22120.10 *", options: { align: "center" } }],
  ];
  addResultTable(s, LAYOUT.resultSlide, rows);

  // Key finding box
  addKeyBox(s, LAYOUT.resultSlide, "KEY FINDING", [
    // TODO: Replace with your key findings
    "Finding one",
    "Finding two",
    "Finding three",
  ]);

  // Note bar
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.7, y: 5.9, w: 12, h: 0.6,
    fill: { color: NAVY }, line: { type: "none" },
  });
  addText(s, "* p<.05   ** p<.01   *** p<.001   NS = not significant   \u03b2 = standardized coefficient", {
    x: 0.9, y: 5.95, w: 11.6, h: 0.5,
    fontFace: FONT_BODY, fontSize: FS.SECONDARY, color: WHITE,
  });

  notes(s, "This second table presents [what Table 2 shows]. The most striking finding is [highlight key result]. Compared to Table 1, we see [difference or consistency]. This is important because [why it matters].");
}

// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// SLIDE 7: KEY NUANCE — TWO-PANEL COMPARISON
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
{
  const s = pres.addSlide();
  s.background = { color: WHITE };
  slideHeader(s, "KEY NUANCE", "Comparing two dimensions");

  // Left panel (teal)
  addPanelBox(s, {
    x: LAYOUT.subscaleSlide.leftX, y: LAYOUT.subscaleSlide.panelY,
    w: LAYOUT.subscaleSlide.leftW, h: LAYOUT.subscaleSlide.panelH,
    headerH: LAYOUT.subscaleSlide.headerH, headerColor: TEAL, headerTextColor: WHITE,
    title: "DIMENSION A",
    // TODO: Replace with your first dimension's findings
    items: [
      "Finding for dimension A, time point 1",
      "Finding for dimension A, time point 2",
      "Pattern description",
      "Interpretation",
    ],
  });

  // Right panel (gold)
  addPanelBox(s, {
    x: LAYOUT.subscaleSlide.rightX, y: LAYOUT.subscaleSlide.panelY,
    w: LAYOUT.subscaleSlide.rightW, h: LAYOUT.subscaleSlide.panelH,
    headerH: LAYOUT.subscaleSlide.headerH, headerColor: GOLD, headerTextColor: NAVY,
    title: "DIMENSION B",
    // TODO: Replace with your second dimension's findings
    items: [
      "Finding for dimension B, time point 1",
      "Finding for dimension B, time point 2",
      "Pattern description",
      "Interpretation",
    ],
  });

  // Bottom: Implication bar
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.7, y: 5.4, w: 12, h: 1.5,
    fill: { color: LIGHT_BG }, line: { type: "none" },
  });
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.7, y: 5.4, w: 0.08, h: 1.5,
    fill: { color: NAVY }, line: { type: "none" },
  });
  addText(s, "IMPLICATION", {
    x: 0.95, y: 5.5, w: 11.5, h: 0.3,
    fontFace: FONT_BODY, fontSize: FS.SECTION, color: GOLD, bold: true, charSpacing: 2,
  });
  // TODO: Replace with your implication
  const implItems = [
    "What this comparison means for theory or practice",
    "Which dimension matters more and why",
  ];
  const implText = implItems.map((t) => ({
    text: t,
    options: {
      fontFace: FONT_BODY, fontSize: FS.BODY, color: DARK,
      bullet: { code: "2022", indent: 10 },
      paraSpaceAfter: 8, breakLine: true, wrap: true,
    },
  }));
  s.addText(implText, {
    x: 0.95, y: 5.85, w: 11.5, h: 0.9, valign: "top", wrap: true,
  });

  notes(s, "A closer look at [Dimension A] versus [Dimension B] reveals an important nuance. On the left, [Dimension A] shows [pattern]. On the right, [Dimension B] shows [pattern]. The contrast between these two dimensions tells us [implication]. This distinction has practical consequences for [field or application].");
}

// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// SLIDE 8: DISCUSSION — FOUR-PANEL CONTEXT
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
{
  const s = pres.addSlide();
  s.background = { color: WHITE };
  slideHeader(s, "DISCUSSION", "Contextualising the findings");

  // Top left: Your finding in context
  addAccentPanel(s, {
    x: 0.7, y: 1.6, w: 5.8, h: 2.8, accentColor: NAVY,
    title: "FINDING IN CONTEXT",
    // TODO: Replace with how your findings compare to prior work
    items: [
      "How your results compare to existing literature",
      "Consistent or contradictory findings",
      "Possible explanations",
    ],
  });

  // Top right: Why effects may differ
  addAccentPanel(s, {
    x: 6.8, y: 1.6, w: 5.9, h: 2.8, accentColor: TEAL,
    title: "INTERPRETATION",
    // TODO: Replace with your interpretation
    items: [
      "Why your findings might differ from prior work",
      "Methodological differences that matter",
      "Theoretical implications",
    ],
  });

  // Bottom left: Strengths
  addAccentPanel(s, {
    x: 0.7, y: 4.7, w: 5.8, h: 2.4, accentColor: MID_GREY,
    title: "STRENGTHS",
    // TODO: Replace with your study strengths
    items: [
      "Strength one: sample size, design, etc.",
      "Strength two: novelty, measures, etc.",
    ],
  });

  // Bottom right: Limitations
  addAccentPanel(s, {
    x: 6.8, y: 4.7, w: 5.9, h: 2.4, accentColor: GOLD,
    title: "LIMITATIONS",
    // TODO: Replace with your study limitations
    items: [
      "Limitation one: design, generalisability, etc.",
      "Limitation two: measures, attrition, etc.",
    ],
  });

  notes(s, "In the broader context, our findings [align with / diverge from] prior work by [Author, Year]. The top-left panel summarises this comparison. Our interpretation is that [explanation for pattern]. Key strengths of this study include [strengths]. We acknowledge limitations such as [limitations], which should be considered when interpreting these results.");
}

// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// SLIDE 9: FIGURE / RESULTS VISUALISATION
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
{
  const s = pres.addSlide();
  s.background = { color: WHITE };
  slideHeader(s, "KEY FIGURE", "Results visualisation");

  // TODO: To use a different image, change the filename below.
  // Place your .png/.jpg in assets/ and reference by filename.
  s.addImage({
    path: path.join(__dirname, "assets", "figure-1.png"),
    x: 0.7, y: 1.6, w: 12, h: 4.8,
  });

  // Figure caption
  addText(s, "Figure 1. Replace this chart with your own using s.addImage(). Place the file in assets/.", {
    x: 0.7, y: 6.5, w: 12, h: 0.5,
    fontFace: FONT_BODY, fontSize: FS.SECONDARY, color: GREY, italic: true,
    align: "center",
  });

  notes(s, "This figure visualises [what the figure shows]. The x-axis represents [variable] and the y-axis represents [variable]. The key pattern to notice is [describe the visual trend]. The error bars indicate [what they represent]. This visual confirms the statistical findings we discussed earlier.");
}

// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// SLIDE 10: IMPLICATIONS — THREE COLUMNS
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
{
  const s = pres.addSlide();
  s.background = { color: WHITE };
  slideHeader(s, "IMPLICATIONS", "Translating findings into practice");

  const colW = 3.87, gap = 0.27;
  // TODO: Replace with your three implication columns
  const cols = [
    {
      head: "For practice",
      items: [
        "Implication one for practitioners",
        "Implication two for practitioners",
        "Implication three for practitioners",
      ],
      color: NAVY,
    },
    {
      head: "For policy",
      items: [
        "Implication one for policy",
        "Implication two for policy",
        "Implication three for policy",
      ],
      color: TEAL,
    },
    {
      head: "For research",
      items: [
        "Implication one for future research",
        "Implication two for future research",
        "Implication three for future research",
      ],
      color: GOLD,
    },
  ];
  cols.forEach((col, i) => {
    const x = 0.7 + i * (colW + gap);
    s.addShape(pres.shapes.RECTANGLE, {
      x, y: 1.6, w: colW, h: 4.5,
      fill: { color: LIGHT_BG }, line: { type: "none" },
    });
    s.addShape(pres.shapes.RECTANGLE, {
      x, y: 1.6, w: colW, h: 0.7,
      fill: { color: col.color }, line: { type: "none" },
    });
    addText(s, col.head, {
      x: x + 0.2, y: 1.6, w: colW - 0.4, h: 0.7,
      fontFace: FONT_BODY, fontSize: FS.BLOCK_HEAD, color: i === 2 ? NAVY : WHITE,
      bold: true, valign: "middle",
    });
    const items = col.items.map((t) => ({
      text: t,
      options: {
        fontFace: FONT_BODY, fontSize: FS.BODY, color: DARK,
        bullet: { code: "2022", indent: 10 },
        paraSpaceAfter: 10, breakLine: true, wrap: true,
      },
    }));
    s.addText(items, {
      x: x + 0.3, y: 2.5, w: colW - 0.5, h: 3.4, valign: "top", wrap: true,
    });
  });

  // Bottom bar
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.7, y: 6.4, w: 12, h: 0.7,
    fill: { color: NAVY }, line: { type: "none" },
  });
  // TODO: Replace with your cross-cutting message
  addText(s, "Your key message that ties the implications together", {
    x: 0.9, y: 6.5, w: 11.6, h: 0.5,
    fontFace: FONT_BODY, fontSize: FS.EMPHASIS, color: WHITE, bold: true,
  });

  notes(s, "These findings carry implications across three domains. For practice, [practical implication]. For policy, [policy implication]. For future research, [research implication]. The overarching message is that [single takeaway that ties all three together].");
}

// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// SLIDE 11: THANK YOU & Q&A
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
{
  const s = pres.addSlide();
  s.background = { color: NAVY };
  s.addShape(pres.shapes.RECTANGLE, {
    x: 12.7, y: 0, w: 0.633, h: 7.5,
    fill: { color: GOLD }, line: { type: "none" },
  });

  addText(s, "THANK YOU", {
    x: 1.0, y: 2.5, w: 11, h: 1.0,
    fontFace: FONT_TITLE, fontSize: FS.TITLE_1, color: WHITE, bold: true,
    align: "center",
  });
  addText(s, "Questions?", {
    x: 1.0, y: 3.6, w: 11, h: 0.7,
    fontFace: FONT_TITLE, fontSize: FS.TITLE_2, color: GOLD, bold: true,
    align: "center",
  });

  notes(s, "Thank you for your attention. I am happy to take questions. If you would like to know more, the full paper is available at [DOI or URL] and I can be reached at [email].");
}

// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// SAVE
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
const out = path.join(__dirname, "Presentation.pptx");
pres.writeFile({ fileName: out }).then((file) => {
  console.log("Saved:", file);
}).catch((err) => {
  console.error("Error:", err);
});
