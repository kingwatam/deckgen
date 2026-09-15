# Skill: PowerPoint Generation with pptxgenjs

Generate professional `.pptx` presentations programmatically using `pptxgenjs`
(JavaScript). This skill produces slides with a consistent visual system: navy/teal/gold
palette, Cambria + Calibri fonts, stat cards, result tables, key takeaway boxes, and
two-panel comparisons.

Use this when a user asks to create a presentation, slide deck, research poster, or
any `.pptx` file from structured content.

## Quick start

```bash
npm install          # one-time, installs pptxgenjs
node build_deck.js   # generates Presentation.pptx
```

The template file `build_deck.js` contains placeholder content. Replace the
`TODO`-marked sections with the user's content, then run the build.

## Working from source documents

When the user provides a paper, thesis chapter, report, or other document, the
agent should:

1. **Extract the structure**: identify research question, methods, key results
   (numbers, p-values, effect sizes), and main contributions.
2. **Map to slide types**: stat cards for sample/key metrics, result tables for
   effect sizes, two-panel comparisons for contrasting findings, accent panels
   for context and limitations.
3. **Generate speaker notes from the source**: write natural-language talking
   points for each slide that explain the content in the user's voice. Notes
   should reference specific findings, contextualise results, and guide the
   audience through the narrative. Do not use generic placeholder notes.
4. **Prompt the user for refinement**: after generating the initial deck, ask
   whether slides should be added, removed, reordered, or restructured.

### Speaker notes

Every slide must have speaker notes via `notes(s, "...")`. When generating from
a source document:

- **Title slide**: introduce yourself and the work in 2-3 sentences.
- **Background**: state why the problem matters and what gap you address.
- **Methods**: explain the design in plain language, highlight what's novel.
- **Results**: walk through the numbers, point to the table/figure, state what
  it means.
- **Discussion**: contextualise against prior work, note limitations honestly.
- **Closing**: summarise the one thing the audience should remember.

## Images and figures

`pptxgenjs` supports images via `slide.addImage()`. To embed a figure from the
user's document:

```javascript
s.addImage({
  path: path.join(__dirname, "assets", "figure-1.png"),
  x: 0.7, y: 1.6, w: 12, h: 5.0,
});
```

Place image files in `assets/` and reference them by path. If no image file is
available, use the placeholder pattern: a light-grey dashed-border rectangle with
centred "[ Your figure here ]" text and a caption below. The template includes a
built-in figure slide (slide 9) demonstrating this layout.

Contributions welcome for: automatic figure extraction from PDFs, chart-to-slide
generation, and responsive image sizing.

## Architecture

### Palette

| Constant | Hex | Usage |
|----------|-----|-------|
| `NAVY` | `1B3B4F` | Slide backgrounds, table headers, accent bars |
| `TEAL` | `3D8B8B` | Secondary accent, panel headers |
| `GOLD` | `C9A95E` | Eyebrow labels, section headers, emphasis |
| `DARK` | `2D2D2D` | Body text |
| `GREY` | `6B6B6B` | Secondary text (authors, labels) |
| `LIGHT_BG` | `F5F5F2` | Card/panel backgrounds |
| `WHITE` | `FFFFFF` | Slide backgrounds, table text on navy |
| `SUBTLE` | `8A8578` | Tertiary accent |
| `MID_GREY` | `555555` | Quaternary accent, tag left-bar |

### Fonts

| Constant | Font | Usage |
|----------|------|-------|
| `FONT_TITLE` | Cambria | Slide titles, stat numbers, section names |
| `FONT_BODY` | Calibri | Body text, bullets, labels, descriptions |

### Font size hierarchy

| Constant | Size | Usage |
|----------|------|-------|
| `EYEBROW` | 15pt | Category labels above slide titles |
| `SLIDE_TITLE` | 26pt | Slide titles (Cambria, bold) |
| `SECTION` | 16pt | Section headers ("KEY TAKEAWAY", "METHODS") |
| `EMPHASIS` | 16pt | Emphasis text (citation bars, implication bars) |
| `BODY` | 15pt | Primary body text, bullets, tags |
| `SECONDARY` | 14pt | Authors, stat labels, detail text |
| `STAT_NUM` | 28pt | Large numbers in stat cards (Cambria) |
| `TITLE_1` | 44pt | Name on title slide |
| `TITLE_2` | 28pt | Main title on closing slide |
| `SUBTITLE` | 18pt | Subtitles |
| `BLOCK_HEAD` | 16pt | Column block headers |

### Layout constants

The `LAYOUT` object defines positions once, reused across slides:

```javascript
const LAYOUT = {
  resultSlide: {
    tableX: 0.7, tableY: 2.1, tableW: 7.5, colW: [3.3, 2.1, 2.1],
    boxX: 8.5, boxY: 2.1, boxW: 4.2, boxH: 3.5, rowH: 0.583,
    noteBarY: 5.9,
  },
  subscaleSlide: {
    leftX: 0.7, leftW: 5.8,
    rightX: 6.8, rightW: 5.9,
    panelY: 1.6, panelH: 3.5, headerH: 0.7,
  },
  contextSlide: {
    leftX: 0.7, leftW: 5.8,
    rightX: 6.8, rightW: 5.9,
    topY: 1.6, topH: 2.8,
    botY: 4.7, botH: 2.4,
  },
};
```

When creating new slides, reference `LAYOUT` for positions rather than
hardcoding x/y values. This keeps slides visually consistent.

## Helper functions

### `slideHeader(s, eyebrow, title)`

Adds the standard eyebrow label (gold, charSpaced) and slide title (Cambria, navy, bold).

```javascript
slideHeader(s, "PAPER 1  \u00b7  METHODS", "Study design");
```

### `addText(s, text, opts)`

Wrapper around `s.addText()` with `wrap: true, valign: "top"` defaults.

### `notes(s, text)`

Adds speaker notes to the current slide.

### `addKeyBox(slide, layout, title, items, opts)`

Grey background + gold top bar + title + bullet list. Warns to console if items
may overflow the box height.

```javascript
addKeyBox(s, LAYOUT.resultSlide, "KEY TAKEAWAY", [
  "Finding one",
  "Finding two",
  "Finding three",
]);
```

### `addPanelBox(slide, opts)`

Colored header bar + light background + bullet items. Used for side-by-side
comparison panels.

```javascript
addPanelBox(s, {
  x: 0.7, y: 1.6, w: 5.8, h: 3.5,
  headerH: 0.7, headerColor: TEAL, headerTextColor: WHITE,
  title: "DIMENSION A",
  items: ["Finding one", "Finding two"],
});
```

### `addAccentPanel(slide, opts)`

Left accent bar + light background + section title + bullets. Used for context
boxes, measures, methods, strengths, limitations.

```javascript
addAccentPanel(s, {
  x: 0.7, y: 1.6, w: 5.8, h: 3.0,
  accentColor: NAVY,
  title: "CONTEXT",
  items: ["Point one", "Point two"],
});
```

### `addResultTable(slide, layout, rows)`

Renders a table using `LAYOUT.resultSlide` positions. Row 0 should be the
header with navy fill and white text.

```javascript
const rows = [
  [
    { text: "Outcome", options: { bold: true, color: WHITE, fill: { color: NAVY } } },
    { text: "\u03b2", options: { bold: true, color: WHITE, fill: { color: NAVY }, align: "center" } },
  ],
  ["Variable one", { text: "\u22120.15 **", options: { align: "center" } }],
];
addResultTable(s, LAYOUT.resultSlide, rows);
```

## Slide types

### 1. Title slide

Navy background, gold vertical accent bar on right, eyebrow text, main title,
subtitle, author name, institution, three topic tags at bottom.

```
Background: NAVY
Accent: GOLD vertical bar at x=12.7
Content: eyebrow → title → subtitle → name → institution → tags
```

### 2. Paper title & citation

White background, `slideHeader`, title card (light bg), authors, citation,
three coloured tag cards, research question in italic grey.

### 3. Background & research question

Two-column layout with bottom panel:
- Left: `addAccentPanel` with NAVY accent (CONTEXT)
- Right: Navy box with gold header (RESEARCH QUESTION)
- Bottom: `addAccentPanel` with TEAL accent (GAP & HYPOTHESIS)

### 4. Study design

Three stat cards across the top (number + label, coloured top bar), then
two-column `addAccentPanel` boxes below (MEASURES left, ANALYTICAL APPROACH right).

### 5. Results table

`addResultTable` on the left, `addKeyBox` on the right, navy note bar at
bottom with significance legend.

### 6. Two-panel comparison

Two `addPanelBox` calls side by side (TEAL header left, GOLD header right),
with a bottom implication bar (light bg + navy left accent).

### 7. Discussion / four-panel context

Four `addAccentPanel` boxes: top-left (NAVY), top-right (TEAL), bottom-left
(MID_GREY for strengths), bottom-right (GOLD for limitations).

### 8. Figure / results visualisation

Full-width figure area with dashed placeholder border, centred label, and italic
caption below. To use a real image, replace the placeholder block with
`s.addImage({ path: ... })`.

### 9. Implications / three columns

Three column cards with coloured header bars (NAVY, TEAL, GOLD), bullet items,
and a navy bottom bar with a unifying message.

### 10. Thank you & Q&A

Navy background, gold vertical bar, centred "THANK YOU" in white (44pt),
"Questions?" in gold (28pt).

## Generating a new presentation

1. Copy `build_deck.js` to a new file (e.g., `build_mydeck.js`).
2. Replace all `TODO`-marked content with the user's actual content.
3. Keep the palette, fonts, helpers, and layout constants unchanged.
4. Add or remove slides as needed, reusing the existing helpers.
5. Run `node build_mydeck.js` to generate the `.pptx`.

### Content guidelines

- **Slide titles**: use `slideHeader(s, "EYEBROW", "Title")`. Keep titles under
  6 words.
- **Stat cards**: three across, each with a large number (28pt Cambria) and a
  short label. Use for sample size, time points, key metrics.
- **Tables**: navy header row, alternating content rows, max 6 data rows per
  slide. Use `align: "center"` for numeric columns.
- **Key takeaway boxes**: max 4 bullet items. The helper warns if text may
  overflow.
- **Speaker notes**: always add them. The agent should generate notes that
  explain the slide content in natural language.

### Adding new slide types

To create a new slide type, compose the existing helpers:

```javascript
{
  const s = pres.addSlide();
  s.background = { color: WHITE };
  slideHeader(s, "YOUR EYEBROW", "Your title");

  // Use addAccentPanel for context boxes
  addAccentPanel(s, { ... });

  // Use addResultTable for data tables
  addResultTable(s, LAYOUT.resultSlide, rows);

  // Use addKeyBox for takeaway boxes
  addKeyBox(s, LAYOUT.resultSlide, "KEY POINT", [...]);

  notes(s, "Speaker notes for this slide.");
}
```

## Anti-patterns

- **Do not** add accent lines under slide titles (AI-generated look).
- **Do not** hardcode x/y positions when `LAYOUT` constants exist.
- **Do not** put more than 6 rows in a result table (overflow).
- **Do not** put more than 4 bullets in a key takeaway box (overflow warning).
- **Do not** use different font families — stick to Cambria + Calibri.
- **Do not** change the palette mid-presentation — use the defined constants.
- **Do not** forget speaker notes on every slide.

## File structure

```
build_deck.js    # Main template — edit this
package.json     # pptxgenjs dependency
SKILL.md         # This file — agent instructions
README.md        # Human-facing documentation
assets/          # Preview screenshots
```
