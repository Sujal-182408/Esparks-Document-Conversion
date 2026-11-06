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
when converting back to PowerPoint. Bold, italic, underline, explicit RGB
colors, font sizes, and paragraph alignment are represented with Markdown
emphasis, inline HTML metadata, and alignment comments. The in-app preview
removes those metadata wrappers while preserving the visible text.

PowerPoint visual styling and non-text objects such as SmartArt are not fully
round-tripped. The converter preserves readable text, tables, supported chart
data, supported links, and selected text styling; it does not guarantee
pixel-identical slide layouts.
