
import JSZip from "jszip";

const MAX_FILE_SIZE = 20 * 1024 * 1024;

function escapeMarkdown(text: string): string {
  return text
    .replace(/\\/g, "\\\\")
    .replace(/([*_`\[\]])/g, "\\$1")
    .replace(/\b(https?|ftp|mailto):\/\//gi, "$1&#58;//");
}

function escapeTableCell(text: string): string {
  return escapeMarkdown(text)
    .replace(/\|/g, "\\|")
    .replace(/\r?\n/g, "<br>");
}

type TableCellMetadata = {
  text: string;
  colspan?: number;
  rowspan?: number;
  merged?: boolean;
};

type TableMetadata = {
  rows: TableCellMetadata[][];
};

function encodeMetadata(value: unknown): string {
  const bytes = new TextEncoder().encode(JSON.stringify(value));
  let binary = "";
  for (const byte of bytes) binary += String.fromCharCode(byte);
  return btoa(binary);
}

function getTableMetadata(table: Element): TableMetadata {
  const rows = getElementsByLocalName(table, "tr").map((row) => {
    const cells = getElementsByLocalName(row, "tc");
    const getCellAttribute = (cell: Element, name: string) => {
      const properties = getElementsByLocalName(cell, "tcPr")[0];
      return cell.getAttribute(name) || properties?.getAttribute(name) || "";
    };
    const isMerged = (cell: Element) => {
      const hMerge = getCellAttribute(cell, "hMerge").toLowerCase();
      const vMerge = getCellAttribute(cell, "vMerge").toLowerCase();
      return (
        hMerge === "1" ||
        hMerge === "true" ||
        vMerge === "1" ||
        vMerge === "true"
      );
    };
    return cells.flatMap((cell, index) => {
      const colspan = Number(getCellAttribute(cell, "gridSpan") || 1);
      const rowspan = Number(getCellAttribute(cell, "rowSpan") || 1);
      const text = getElementsByLocalName(cell, "t")
        .map((node) => node.textContent ?? "")
        .join("");
      if (isMerged(cell)) return [{ text: "", merged: true }];
      const mergedPlaceholders = cells.slice(index + 1, index + colspan);
      const hasMergePlaceholders =
        colspan > 1 &&
        mergedPlaceholders.length === colspan - 1 &&
        mergedPlaceholders.every(isMerged);
      const metadata: TableCellMetadata = {
        text,
        ...(colspan > 1 ? { colspan } : {}),
        ...(rowspan > 1 ? { rowspan } : {}),
      };
      return [
        metadata,
        ...(!hasMergePlaceholders
          ? Array.from({ length: Math.max(0, colspan - 1) }, () => ({
              text: "",
              merged: true,
            }))
          : []),
      ];
    });
  });
  return { rows };
}

function getElementsByLocalName(element: Element | Document, name: string): Element[] {
  return Array.from(element.getElementsByTagName("*")).filter(
    (child) => child.localName === name
  );
}

function getChildElements(element: Element): Element[] {
  return Array.from(element.childNodes).filter(
    (child): child is Element => child.nodeType === 1,
  );
}

function getRelationshipId(element: Element): string | null {
  return [
    element.getAttributeNS(
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
      "id"
    ),
    element.getAttributeNS(
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
      "embed"
    ),
    element.getAttribute("r:id"),
    element.getAttribute("r:embed"),
  ].find((value) => value !== null && value !== "") ?? null;
}

function getColor(
  fill: Element | undefined,
  themeColors: Map<string, string>,
): string | undefined {
  if (!fill) return undefined;
  for (const color of getChildElements(fill)) {
    const value =
      color.localName === "sysClr"
        ? color.getAttribute("lastClr") || color.getAttribute("val")
        : color.getAttribute("val");
    if (!value) continue;
    if (
      color.localName === "srgbClr" ||
      color.localName === "sysClr"
    ) {
      return /^[0-9a-f]{6}$/i.test(value) ? value : undefined;
    }
    if (color.localName === "schemeClr") {
      return themeColors.get(value);
    }
  }
  return undefined;
}

function getParagraphMarkdown(
  paragraph: Element,
  relationships: Map<string, string>,
  themeColors: Map<string, string>,
  shapeLevel = 0,
): {
  text: string;
  bullet: boolean;
  ordered: boolean;
  numberingStart: number;
  level: number;
  alignment: "left" | "center" | "right" | "justify" | null;
} {
  const paragraphProperties = getElementsByLocalName(paragraph, "pPr")[0];
  const paragraphLevel = Number(paragraphProperties?.getAttribute("lvl") ?? 0);
  const level = Math.min(
    5,
    Math.max(Number.isFinite(paragraphLevel) ? paragraphLevel : 0, shapeLevel, 0),
  );
  const autoNumber = getElementsByLocalName(paragraph, "buAutoNum")[0];
  const ordered = Boolean(autoNumber);
  const bullet =
    !!getElementsByLocalName(paragraph, "buChar").length ||
    ordered ||
    level > 0;
  const rawNumberingStart = Number(autoNumber?.getAttribute("startAt") ?? 1);
  const numberingStart = Number.isFinite(rawNumberingStart)
    ? Math.max(1, rawNumberingStart)
    : 1;
  const alignmentValue = paragraphProperties?.getAttribute("algn");
  const alignment =
    alignmentValue === "ctr"
      ? "center"
      : alignmentValue === "r"
        ? "right"
        : alignmentValue === "just" || alignmentValue === "dist"
          ? "justify"
          : null;
  const formatRun = (run: Element) => {
    const value = getElementsByLocalName(run, "t")
      .map((node) => node.textContent ?? "")
      .join("");
    if (!value) return "";

    const properties = getElementsByLocalName(run, "rPr")[0];
    const bold = ["1", "true"].includes(
      properties?.getAttribute("b")?.toLowerCase() ?? "",
    );
    const italic = ["1", "true"].includes(
      properties?.getAttribute("i")?.toLowerCase() ?? "",
    );
    const underlineValue = properties?.getAttribute("u");
    const underline = Boolean(underlineValue && underlineValue !== "none");
    const fontSize = Number(properties?.getAttribute("sz")) / 100;
    const size =
      Number.isFinite(fontSize) && fontSize > 0 ? fontSize : undefined;
    const solidFill = properties
      ? getElementsByLocalName(properties, "solidFill")[0]
      : undefined;
    const color = getColor(solidFill, themeColors);
    const escaped = escapeMarkdown(value);
    const relationshipId = properties
      ? getRelationshipId(
          getElementsByLocalName(properties, "hlinkClick")[0] ?? properties,
        )
      : null;
    const href = relationshipId ? relationships.get(relationshipId) : undefined;
    let formatted = underline ? `<u>${escaped}</u>` : escaped;
    if (color && /^[0-9a-f]{6}$/i.test(color)) {
      formatted = `<span data-pptx-color="#${color.toUpperCase()}"${size ? ` data-pptx-font-size="${size}"` : ""}>${formatted}</span>`;
    } else if (size) {
      formatted = `<span data-pptx-font-size="${size}">${formatted}</span>`;
    }
    const linkedText =
      href && /^(https?:|mailto:)/i.test(href)
        ? `[${formatted}](<${href.replace(/>/g, "%3E")}>)`
        : formatted;

    if (bold && italic) return `***${linkedText}***`;
    if (bold) return `**${linkedText}**`;
    if (italic) return `*${linkedText}*`;
    return linkedText;
  };
  const paragraphParts = getChildElements(paragraph).flatMap((node) => {
    if (node.localName === "br") return ["  \n"];
    if (node.localName === "r" || node.localName === "fld") {
      const formatted = formatRun(node);
      return formatted ? [formatted] : [];
    }
    return [];
  });
  const text = paragraphParts.length
    ? paragraphParts.join("")
    : getElementsByLocalName(paragraph, "t")
        .map((node) => escapeMarkdown(node.textContent ?? ""))
        .join("");

  return {
    text: text.replace(/^[ \t]+|[ \t]+$/g, ""),
    bullet,
    ordered,
    numberingStart,
    level,
    alignment,
  };
}

function getPlainParagraphText(paragraph: Element): string {
  return getChildElements(paragraph)
    .map((node) => {
      if (node.localName === "br") return "\n";
      if (node.localName !== "r" && node.localName !== "fld") return "";
      return getElementsByLocalName(node, "t")
        .map((text) => text.textContent ?? "")
        .join("");
    })
    .join("");
}

function getPlaceholderType(shape: Element): string | null {
  return getElementsByLocalName(shape, "ph")[0]?.getAttribute("type") ?? null;
}

function getTableMarkdown(table: Element): string {
  const metadata = getTableMetadata(table);
  const rows = metadata.rows.map((row) =>
    row.map((cell) =>
      escapeTableCell(cell.text.trim()),
    ),
  );

  if (!rows.length) return "";

  const columnCount = Math.max(...rows.map((row) => row.length));
  const normalizedRows = rows.map((row) =>
    Array.from({ length: columnCount }, (_, index) => row[index] ?? "")
  );
  const separator = normalizedRows[0].map(() => "---");

  const metadataComment = `<!-- pptx-table:${encodeMetadata(metadata)} -->`;
  return [
    metadataComment,
    "",
    `| ${normalizedRows[0].join(" | ")} |`,
    `| ${separator.join(" | ")} |`,
    ...normalizedRows.slice(1).map((row) => `| ${row.join(" | ")} |`),
  ].join("\n");
}

function getShapeMetadata(
  shape: Element,
  slideWidth: number,
  slideHeight: number,
  themeColors: Map<string, string>,
): Record<string, unknown> | null {
  const geometry = getElementsByLocalName(shape, "prstGeom")[0];
  const preset = geometry?.getAttribute("prst");
  const supportedShapes: Record<string, string> = {
    rect: "rect",
    roundRect: "roundRect",
    ellipse: "ellipse",
    diamond: "diamond",
    triangle: "triangle",
    hexagon: "hexagon",
    pentagon: "pentagon",
    parallelogram: "parallelogram",
    trapezoid: "trapezoid",
  };
  if (!preset || !supportedShapes[preset]) return null;

  const shapeProperties = getElementsByLocalName(shape, "spPr")[0];
  if (!shapeProperties) return null;
  const transform = getElementsByLocalName(shapeProperties, "xfrm")[0];
  const offset = transform && getElementsByLocalName(transform, "off")[0];
  const extent = transform && getElementsByLocalName(transform, "ext")[0];
  const x = Number(offset?.getAttribute("x"));
  const y = Number(offset?.getAttribute("y"));
  const width = Number(extent?.getAttribute("cx"));
  const height = Number(extent?.getAttribute("cy"));
  if (![x, y, width, height].every(Number.isFinite) || width <= 0 || height <= 0) {
    return null;
  }

  const fill = getElementsByLocalName(shapeProperties, "solidFill")[0];
  const fillColor = getColor(fill, themeColors);
  const line = getElementsByLocalName(shapeProperties, "ln")[0];
  const lineFill = line && getElementsByLocalName(line, "solidFill")[0];
  const lineColor = getColor(lineFill, themeColors);
  if (!fillColor && !lineColor) return null;
  const label = getElementsByLocalName(shape, "t")
    .map((node) => node.textContent ?? "")
    .join("");

  return {
    type: supportedShapes[preset],
    x: (x / slideWidth) * 13.333,
    y: (y / slideHeight) * 7.5,
    w: (width / slideWidth) * 13.333,
    h: (height / slideHeight) * 7.5,
    ...(fillColor && /^[0-9a-f]{6}$/i.test(fillColor)
      ? { fill: fillColor.toUpperCase() }
      : {}),
    ...(lineColor && /^[0-9a-f]{6}$/i.test(lineColor)
      ? { line: lineColor.toUpperCase() }
      : {}),
    ...(label ? { label } : {}),
  };
}

function getPictureMarkdown(
  picture: Element,
  relationships: Map<string, string>,
  relationshipTypes: Map<string, string>,
  zip: JSZip,
  slidePath: string,
): Promise<string> {
  const imageId = getRelationshipId(getElementsByLocalName(picture, "blip")[0] ?? picture);
  if (!imageId) return Promise.resolve("");
  const target = relationships.get(imageId);
  if (!target || !relationshipTypes.get(imageId)?.endsWith("/image")) {
    return Promise.resolve("");
  }
  const imagePath = resolvePackagePath(slidePath, target);
  const image = zip.file(imagePath);
  if (!image) return Promise.resolve("");
  const extension = imagePath.split(".").pop()?.toLowerCase();
  const mimeTypes: Record<string, string> = {
    png: "image/png",
    jpg: "image/jpeg",
    jpeg: "image/jpeg",
    gif: "image/gif",
  };
  const mimeType = extension ? mimeTypes[extension] : undefined;
  if (!mimeType) {
    return Promise.resolve("_An embedded image uses an unsupported format and could not be exported._");
  }
  const altText =
    getElementsByLocalName(picture, "cNvPr")[0]?.getAttribute("descr") ||
    getElementsByLocalName(picture, "cNvPr")[0]?.getAttribute("name") ||
    "Embedded image";
  return image.async("base64").then(
    (base64) => `![${escapeMarkdown(altText)}](data:${mimeType};base64,${base64})`,
  );
}

function getThemeColors(xml: Document): Map<string, string> {
  const colors = new Map<string, string>();
  const scheme = getElementsByLocalName(xml, "clrScheme")[0];
  if (!scheme) return colors;
  for (const entry of getChildElements(scheme)) {
    const color = getColor(entry, new Map());
    if (color) colors.set(entry.localName, color.toUpperCase());
  }
  return colors;
}

function resolvePackagePath(baseFile: string, target: string): string {
  const parts = target.startsWith("/")
    ? []
    : baseFile.split("/").slice(0, -1);
  for (const part of target.split("/")) {
    if (!part || part === ".") continue;
    if (part === "..") parts.pop();
    else parts.push(decodeURIComponent(part));
  }
  return parts.join("/");
}

function getRelationships(xml: Document): Map<string, string> {
  const relationships = new Map<string, string>();
  for (const relationship of getElementsByLocalName(xml, "Relationship")) {
    const id = relationship.getAttribute("Id");
    const target = relationship.getAttribute("Target");
    if (id && target) relationships.set(id, target);
  }
  return relationships;
}

function getRelationshipTypes(xml: Document): Map<string, string> {
  const relationshipTypes = new Map<string, string>();
  for (const relationship of getElementsByLocalName(xml, "Relationship")) {
    const id = relationship.getAttribute("Id");
    const type = relationship.getAttribute("Type");
    if (id && type) relationshipTypes.set(id, type);
  }
  return relationshipTypes;
}

function getCachedValues(element: Element): string[] {
  const points = getElementsByLocalName(element, "pt").map((point) => ({
    index: Number(point.getAttribute("idx") ?? 0),
    value: getElementsByLocalName(point, "v")[0]?.textContent ?? "",
  }));
  const maxIndex = Math.max(-1, ...points.map((point) => point.index));
  const values = Array.from({ length: maxIndex + 1 }, () => "");
  for (const point of points) values[point.index] = point.value;
  return values;
}

function chartToMarkdown(chartXml: Document): string {
  const chart = getElementsByLocalName(chartXml, "chart")[0];
  if (!chart) return "";

  const title = getElementsByLocalName(chart, "title")[0];
  const chartTitle = title
    ? getElementsByLocalName(title, "t")
        .map((node) => node.textContent ?? "")
        .join("")
    : "";
  const plotArea = getElementsByLocalName(chart, "plotArea")[0];
  if (!plotArea) {
    return [
      chartTitle ? `### Chart data: ${escapeMarkdown(chartTitle)}` : "### Chart data",
      "",
      "_Chart plot data could not be read from the PowerPoint file._",
    ].join("\n");
  }

  const chartKinds = [
    ["barChart", "bar"],
    ["lineChart", "line"],
    ["pieChart", "pie"],
    ["doughnutChart", "doughnut"],
    ["areaChart", "area"],
  ] as const;
  const allChartKinds = [
    ...chartKinds.map(([name]) => name),
    "scatterChart",
    "radarChart",
    "stockChart",
    "surfaceChart",
    "bubbleChart",
    "ofPieChart",
    "boxWhiskerChart",
    "waterfallChart",
    "funnelChart",
    "histogramChart",
    "treemapChart",
    "sunburstChart",
  ];
  const chartNodes = allChartKinds.flatMap((name) =>
    getElementsByLocalName(plotArea, name)
  );
  const chartKind = chartKinds.find(
    ([name]) =>
      chartNodes.length === 1 &&
      getElementsByLocalName(plotArea, name).length > 0
  );
  const chartNode = chartKind
    ? getElementsByLocalName(plotArea, chartKind[0])[0]
    : undefined;
  const barDirection =
    chartNode && chartKind?.[1] === "bar"
      ? getElementsByLocalName(chartNode, "barDir")[0]?.getAttribute("val")
      : undefined;
  const seriesElements = getElementsByLocalName(plotArea, "ser");
  const series = seriesElements.map((item, index) => {
    const nameNode = getElementsByLocalName(item, "tx")[0];
    const nameText = nameNode
      ? getElementsByLocalName(nameNode, "v")[0]?.textContent ??
        getElementsByLocalName(nameNode, "t")
          .map((node) => node.textContent ?? "")
          .join("")
      : "";
    const name = nameText || `Series ${index + 1}`;
    const category =
      getElementsByLocalName(item, "cat")[0] ??
      getElementsByLocalName(item, "xVal")[0];
    const categoryData = category ? getCachedValues(category) : [];
    const valuesNode =
      getElementsByLocalName(item, "val")[0] ??
      getElementsByLocalName(item, "yVal")[0];
    const values = valuesNode ? getCachedValues(valuesNode) : [];
    return {
      name,
      categories: categoryData,
      values,
    };
  });
  const heading = chartTitle
    ? `### Chart data: ${escapeMarkdown(chartTitle)}`
    : "### Chart data";
  if (!series.length || !series.some((item) => item.values.length)) {
    return [
      heading,
      "",
      "_Chart data is not cached in the PowerPoint file, so values could not be extracted._",
    ].join("\n");
  }

  const valueCount = Math.max(...series.map((item) => item.values.length));
  const cachedCategories =
    series.find((item) => item.categories.length)?.categories ?? [];
  const categories = Array.from({ length: valueCount }, (_, index) =>
    cachedCategories[index] ?? String(index + 1)
  );
  const numericSeries = series.every(
    (item) =>
      item.values.length === valueCount &&
      item.values.every(
        (value) => value !== "" && Number.isFinite(Number(value))
      )
  );
  const matchingCategories = series.every(
    (item) =>
      !item.categories.length ||
      item.categories.length === valueCount &&
        item.categories.every(
          (category, index) => category === categories[index]
        )
  );

  if (
    chartKind &&
    numericSeries &&
    matchingCategories &&
    (chartKind[1] !== "bar" || !barDirection || ["bar", "col"].includes(barDirection))
  ) {
    return [
      "```pptx-chart",
      JSON.stringify({
        type: chartKind[1],
        ...(barDirection ? { barDir: barDirection } : {}),
        ...(chartTitle ? { title: chartTitle } : {}),
        categories,
        series: series.map(({ name, values }) => ({
          name,
          values: values.map(Number),
        })),
      }),
      "```",
    ].join("\n");
  }

  const headers = ["Category", ...series.map((item) => item.name)];
  const rows = Array.from({ length: valueCount }, (_, index) => [
    categories[index],
    ...series.map((item) => item.values[index] ?? ""),
  ]);
  const table = [
    `| ${headers.map(escapeTableCell).join(" | ")} |`,
    `| ${headers.map(() => "---").join(" | ")} |`,
    ...rows.map(
      (row) =>
        `| ${row.map(escapeTableCell).join(" | ")} |`
    ),
  ].join("\n");
  const kind =
    chartNodes.length === 1
      ? chartNodes[0].localName.replace(/Chart$/, "")
      : "complex";
  return [
    heading,
    "",
    `_This ${kind} chart is preserved as data because its type or cached values cannot be reconstructed as an editable chart._`,
    "",
    table,
  ].join("\n");
}

export async function convertPowerPointToMarkdown(
  file: File
): Promise<string> {
  if (!file.name.toLowerCase().endsWith(".pptx")) {
    throw new Error("Please select a PowerPoint .pptx file.");
  }

  if (file.size === 0) {
    throw new Error("The selected PowerPoint file is empty.");
  }

  if (file.size > MAX_FILE_SIZE) {
    throw new Error("The maximum file size is 20 MB.");
  }

  const buffer = await file.arrayBuffer();
  const bytes = new Uint8Array(buffer);

  if (bytes[0] !== 0x50 || bytes[1] !== 0x4b) {
    throw new Error("This does not appear to be a valid PPTX file.");
  }

  try {
    const zip = await JSZip.loadAsync(buffer);

    const parser = new DOMParser();
    const presentationFile = zip.file("ppt/presentation.xml");
    const relationshipsFile = zip.file(
      "ppt/_rels/presentation.xml.rels"
    );

    if (!presentationFile || !relationshipsFile) {
      throw new Error("Required PowerPoint presentation files are missing.");
    }

    const presentationXml = parser.parseFromString(
      await presentationFile.async("text"),
      "application/xml"
    );

    const relationshipsXml = parser.parseFromString(
      await relationshipsFile.async("text"),
      "application/xml"
    );

    const relationshipMap = getRelationships(relationshipsXml);
    const themeFile = zip.file("ppt/theme/theme1.xml");
    const themeColors = themeFile
      ? getThemeColors(parser.parseFromString(await themeFile.async("text"), "application/xml"))
      : new Map<string, string>();
    const slideSize = getElementsByLocalName(presentationXml, "sldSz")[0];
    const slideWidth = Number(slideSize?.getAttribute("cx") ?? 9144000);
    const slideHeight = Number(slideSize?.getAttribute("cy") ?? 5143500);

    const slideIds = Array.from(
      presentationXml.getElementsByTagName("p:sldId")
    );

    if (slideIds.length === 0) {
      throw new Error("No slides were found in this presentation.");
    }

    const markdownSlides: string[] = [];

    for (let index = 0; index < slideIds.length; index++) {
      const relationshipId = getRelationshipId(slideIds[index]);

      if (!relationshipId) continue;

      const target = relationshipMap.get(relationshipId);

      if (!target) continue;

      const slidePath = resolvePackagePath("ppt/presentation.xml", target);

      const slideFile = zip.file(slidePath);

      if (!slideFile) continue;

      const slideXml = parser.parseFromString(
        await slideFile.async("text"),
        "application/xml"
      );
      const slideRelationshipsFile = zip.file(
        `${slidePath.slice(0, slidePath.lastIndexOf("/") + 1)}_rels/${slidePath
          .slice(slidePath.lastIndexOf("/") + 1)}.rels`
      );
      const slideRelationshipsXml = slideRelationshipsFile
        ? parser.parseFromString(
            await slideRelationshipsFile.async("text"),
            "application/xml"
          )
        : null;
      const slideRelationships = slideRelationshipsXml
        ? getRelationships(slideRelationshipsXml)
        : new Map<string, string>();
      const slideRelationshipTypes = slideRelationshipsXml
        ? getRelationshipTypes(slideRelationshipsXml)
        : new Map<string, string>();

      const shapes = getElementsByLocalName(slideXml, "sp");
      const hasTitlePlaceholder = shapes.some((shape) => {
        const type = getPlaceholderType(shape);
        return type === "title" || type === "ctrTitle";
      });

      const slideLines: string[] = [];
      let headingAdded = false;
      let lastBulletLine = -1;

      for (const shape of shapes) {
        const placeholderType = getPlaceholderType(shape);
        const shapeName =
          getElementsByLocalName(shape, "cNvPr")[0]?.getAttribute("name");
        const generatedListLevel = shapeName?.match(
          /^Document Conversion List Item level (\d+)$/,
        );
        if (
          placeholderType === "sldNum" ||
          shapeName === "Document Conversion Slide Number" ||
          (shapeName?.startsWith("Document Conversion ") && !generatedListLevel)
        ) {
          continue;
        }
        const isTitlePlaceholder =
          placeholderType === "title" || placeholderType === "ctrTitle";
        const isSubtitlePlaceholder = placeholderType === "subTitle";
        const shapeMetadata = getShapeMetadata(
          shape,
          slideWidth,
          slideHeight,
          themeColors,
        );
        if (shapeMetadata) {
          slideLines.push(
            "```pptx-shape",
            JSON.stringify(shapeMetadata),
            "```",
          );
          lastBulletLine = -1;
          continue;
        }
        const paragraphs = getElementsByLocalName(shape, "p");
        const fonts = getElementsByLocalName(shape, "latin")
          .map((node) => node.getAttribute("typeface") ?? "")
          .filter(Boolean);
        const isCodeShape =
          fonts.length > 0 &&
          fonts.every((font) => /^(consolas|courier new|menlo|monaco)$/i.test(font));
        if (isCodeShape && paragraphs.length) {
          const code = paragraphs
            .map(getPlainParagraphText)
            .join("\n");
          const fence = "`".repeat(Math.max(3, ...Array.from(code.matchAll(/`+/g), (match) => match[0].length + 1)));
          slideLines.push(`${fence}text`, code, fence);
          lastBulletLine = -1;
          continue;
        }

        for (const [paragraphIndex, paragraph] of paragraphs.entries()) {
          const {
            text,
            bullet,
            ordered,
            numberingStart,
            level,
            alignment,
          } = getParagraphMarkdown(
            paragraph,
            slideRelationships,
            themeColors,
            Number(generatedListLevel?.[1] ?? 0),
          );

          if (!text) continue;

          if (isTitlePlaceholder && paragraphIndex === 0) {
            slideLines.push(`# ${text}`);
            headingAdded = true;
            lastBulletLine = -1;
          } else if (isSubtitlePlaceholder && paragraphIndex === 0) {
            slideLines.push(`## ${text}`);
            headingAdded = true;
            lastBulletLine = -1;
          } else if (!hasTitlePlaceholder && !headingAdded && !bullet) {
            slideLines.push(`# ${text}`);
            headingAdded = true;
            lastBulletLine = -1;
          } else {
            if (alignment) {
              slideLines.push(`<!-- pptx-align:${alignment} -->`);
            }
            if (bullet) {
              const marker = ordered ? `${numberingStart}.` : "-";
              const bulletLine = `${"  ".repeat(level)}${marker} ${text}`;
              if (lastBulletLine >= 0) {
                slideLines[lastBulletLine] += `\n${bulletLine}`;
              } else {
                slideLines.push(bulletLine);
                lastBulletLine = slideLines.length - 1;
              }
            } else {
              slideLines.push(text);
              lastBulletLine = -1;
            }
          }
        }
      }

      const tables = getElementsByLocalName(slideXml, "tbl")
        .map(getTableMarkdown)
        .filter(Boolean);
      if (tables.length) slideLines.push(...tables);
      if (tables.length) lastBulletLine = -1;

      const pictures = getElementsByLocalName(slideXml, "pic");
      for (const picture of pictures) {
        const markdown = await getPictureMarkdown(
          picture,
          slideRelationships,
          slideRelationshipTypes,
          zip,
          slidePath,
        );
        if (markdown) slideLines.push(markdown);
      }

      const charts: string[] = [];
      for (const [id, chartTarget] of slideRelationships) {
        if (!slideRelationshipTypes.get(id)?.endsWith("/chart")) continue;
        const chartPath = resolvePackagePath(slidePath, chartTarget);
        const chartFile = zip.file(chartPath);
        if (!chartFile) {
          charts.push(
            "### Chart data\n\n_The referenced chart file is missing from the PowerPoint package._"
          );
          continue;
        }
        const chartXml = parser.parseFromString(
          await chartFile.async("text"),
          "application/xml"
        );
        const chartMarkdown = chartToMarkdown(chartXml);
        charts.push(
          chartMarkdown ||
            "### Chart data\n\n_Chart data could not be read from the PowerPoint file._"
        );
      }
      if (charts.length) slideLines.push(...charts);

      if (slideLines.length === 0) {
        slideLines.push("_No extractable text on this slide._");
      }

      markdownSlides.push(
        headingAdded
          ? slideLines.join("\n\n")
          : `${index === 0 ? "#" : "##"} Slide ${index + 1}\n\n${slideLines.join("\n\n")}`
      );
    }

    if (markdownSlides.length === 0) {
      throw new Error("No readable slide content was found.");
    }

    return markdownSlides.join("\n\n<!-- slide -->\n\n") + "\n";
  } catch (error) {
    if (
      error instanceof Error &&
      (
        error.message.startsWith("Required PowerPoint") ||
        error.message.startsWith("No slides") ||
        error.message.startsWith("No readable")
      )
    ) {
      throw error;
    }

    console.error("PowerPoint conversion failed:", error);

    throw new Error(
      "Unable to read this PowerPoint file. It may be corrupt, encrypted, or unsupported."
    );
  }
}
