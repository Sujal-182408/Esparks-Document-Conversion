# Esparks Document Conversion

A browser-based Next.js application for converting Word, PDF, PowerPoint, and
Markdown documents.

## Development

```bash
npm install
npm run dev
```

Run the production checks with:

```bash
npm run lint
npm run build
```

## PowerPoint and Markdown interchange

PowerPoint slides are separated in generated Markdown with an HTML comment:

```markdown
# Presentation title

## Optional subtitle

<!-- slide -->

## Next slide
```

The first heading in each section becomes that slide's title. A level-two
heading immediately after the first level-one heading is treated as the title
slide's subtitle. The previous `---` slide separator is also recognized for
compatibility; horizontal rules in Markdown therefore act as slide boundaries
when converting to PowerPoint.

Supported cached chart data is written as a `pptx-chart` fenced block and
reconstructed as an editable chart:

````markdown
```pptx-chart
{"type":"bar","barDir":"col","title":"Performance","categories":["Conversions","Uploads"],"series":[{"name":"2025","values":[8,12]},{"name":"2026","values":[14,18]}]}
```
````

Chart types currently reconstructed are bar, line, pie, doughnut, and area.
Other chart types are exported as data tables when possible. PowerPoint chart
data must be present in the chart's embedded XML cache; when it is unavailable,
the Markdown contains an explicit note instead of silently omitting the chart.
Hyperlinks with HTTP(S) or `mailto:` URLs are retained in Markdown and restored
when converting back to PowerPoint. Plain-text URLs remain plain text. Bold,
italic, underline, explicit RGB colors, font sizes, paragraph alignment, and
manual line breaks are represented with Markdown emphasis, inline HTML
metadata, and alignment comments. Nested lists and code-formatted text are
preserved.

Embedded PNG, JPEG, and GIF pictures are stored as data URLs in generated
Markdown. Supported PowerPoint shapes are represented by `pptx-shape` fenced
blocks, while table merge relationships are stored in an HTML comment beside
the readable Markdown table. These metadata blocks are used when converting
back to PowerPoint and should be retained for best fidelity. Unsupported
visual features, including SmartArt and pixel-identical slide layouts, are not
guaranteed to round-trip. Markdown files can be up to 28 MB to accommodate the
Base64 size increase from embedded images; other uploads remain limited to
20 MB.
