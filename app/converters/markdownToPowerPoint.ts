
import { marked, type Token, type Tokens } from "marked";
import type PptxGenJS from "pptxgenjs";

export type MdToPptxOptions = {
  fontSize?: number; // body, 10-24, default 18
  title?: string; // doc title fallback
  author?: string;
};

/* ---------- Slide geometry — PowerPoint Windows default 16:9 ---------- */
const SW = 13.333,
  SH = 7.5;
const MARGIN_X = 0.75; // left/right
const CONTENT_W = SW - MARGIN_X * 2; // 11.833
const HEADER_TOP = 0.45; // title baseline area (Title & Content)
const HEADER_H = 0.95;
const DIVIDER_Y = 1.48;
const CONTENT_TOP = 1.75; // first content y
const CONTENT_BOTTOM = 6.85; // last content y
const CONTENT_H = CONTENT_BOTTOM - CONTENT_TOP; // 5.10
const FOOTER_Y = 7.08;
const GAP = 0.18; // gap between blocks — never 0

const C = {
  bg: "FFFFFF",
  primary: "101828",
  body: "475467",
  muted: "667085",
  accent: "0891B2",
  accentDark: "0369A1",
  border: "E4E7EC",
  codeBg: "F2F4F7",
  quoteBg: "F0F9FF",
  tableHead: "E0F2FE",
  tableAlt: "F8FAFC",
  link: "0284C7",
  codeText: "0369A1",
};
const FONT = "Calibri"; // PowerPoint Windows default font (not Arial)
const FONT_MONO = "Consolas"; // Windows default mono

type Run = {
  text: string;
  bold?: boolean;
  italic?: boolean;
  strike?: boolean;
  color?: string;
  fontFace?: string;
  link?: string;
};
type Line = Run[];

/* ---------- Measurement: canvas is the only honest way ---------- */
let _canvas: HTMLCanvasElement | null = null;
// SSR-SAFE: never touch `document` at import time, and never during render.
// This function only runs inside the Convert click handler (client-only).
// If called on server, fall back to char-width estimate instead of crashing.
function measureWidth(
  text: string,
  fontSize: number,
  bold: boolean,
  fontFace: string,
): number {
  if (typeof document === "undefined") {
    // ~0.55 * fontSize per char in inches — same fallback as fallbackChars()
    return (text.length * fontSize * 0.55) / 72;
  }
  if (!_canvas) _canvas = document.createElement("canvas");
  const ctx = _canvas.getContext("2d")!;
  ctx.font = `${bold ? "700" : "400"} ${fontSize}px ${fontFace}, Arial, sans-serif`;
  return ctx.measureText(text).width / 96;
}
/** chars that fit — fallback if canvas unavailable (SSR) */
function fallbackChars(widthIn: number, fontSize: number) {
  return Math.max(12, Math.floor((widthIn * 72) / (fontSize * 0.55)));
}

/** Wrap runs into lines that FIT width. Never splits a word unless word > width. */
function wrapRuns(
  runs: Run[],
  widthIn: number,
  fontSize: number,
  fontFace: string,
): Line[] {
  // Expand runs into word tokens, preserving run style
  type Tok = { text: string; run: Run; isSpace: boolean };
  const toks: Tok[] = [];
  for (const r of runs) {
    const parts = r.text.split(/(\s+)/);
    for (const p of parts)
      if (p) toks.push({ text: p, run: r, isSpace: /^\s+$/.test(p) });
  }
  const lines: Line[] = [];
  let cur: Run[] = [];
  let curW = 0;
  const pushLine = () => {
    if (cur.length) lines.push(cur);
    cur = [];
    curW = 0;
  };
  const pushRun = (text: string, run: Run) => {
    const last = cur[cur.length - 1];
    if (
      last &&
      last.bold === run.bold &&
      last.italic === run.italic &&
      last.color === run.color &&
      last.fontFace === run.fontFace &&
      last.link === run.link
    )
      last.text += text;
    else cur.push({ ...run, text });
  };
  for (const t of toks) {
    if (t.text === "\n") {
      pushLine();
      continue;
    }
    const ff = t.run.fontFace ?? fontFace;
    let remaining = t.text;
    // Hard-split overlong words char-by-char
    while (remaining) {
      const w = measureWidth(remaining, fontSize, !!t.run.bold, ff);
      if (curW + w <= widthIn) {
        pushRun(remaining, t.run);
        curW += w;
        remaining = "";
      } else if (curW > 0 && !t.isSpace) {
        pushLine();
      } else {
        // take as many chars as fit
        let lo = 0,
          hi = remaining.length;
        // binary search max fitting prefix
        while (lo < hi) {
          const mid = Math.ceil((lo + hi) / 2);
          if (
            measureWidth(remaining.slice(0, mid), fontSize, !!t.run.bold, ff) <=
            widthIn - curW
          )
            lo = mid;
          else hi = mid - 1;
        }
        const take = Math.max(1, lo);
        pushRun(remaining.slice(0, take), t.run);
        curW += measureWidth(
          remaining.slice(0, take),
          fontSize,
          !!t.run.bold,
          ff,
        );
        remaining = remaining.slice(take);
        if (remaining) pushLine();
      }
    }
    // trailing space of a line is dropped visually — keep width calc honest
    if (t.isSpace && curW >= widthIn) pushLine();
  }
  pushLine();
  return lines.length ? lines : [[]];
}

function plainText(runs: Run[]) {
  return runs.map((r) => r.text).join("");
}
function lineText(l: Line) {
  return l.map((r) => r.text).join("");
}

/* ---------- Markdown -> Runs ---------- */
function inlineRuns(tokens: Token[], inherited: Partial<Run> = {}): Run[] {
  const out: Run[] = [];
  const add = (text: string, o: Partial<Run> = {}) => {
    if (text) out.push({ text, ...inherited, ...o });
  };
  for (const t of tokens) {
    switch (t.type) {
      case "text":
      case "escape":
        if ("tokens" in t && Array.isArray(t.tokens))
          out.push(...inlineRuns(t.tokens, inherited));
        else if ("text" in t) add(t.text);
        break;
      case "strong":
        out.push(...inlineRuns(t.tokens ?? [], { ...inherited, bold: true }));
        break;
      case "em":
        out.push(...inlineRuns(t.tokens ?? [], { ...inherited, italic: true }));
        break;
      case "del":
        out.push(...inlineRuns(t.tokens ?? [], { ...inherited, strike: true }));
        break;
      case "codespan":
        add(t.text, { fontFace: FONT_MONO, color: C.codeText });
        break;
      case "link": {
        const href = (t as Tokens.Link).href ?? "";
        const safe = /^(https?:|mailto:|#|\/)/i.test(href) ? href : undefined;
        out.push(
          ...inlineRuns(t.tokens ?? [], {
            ...inherited,
            color: C.link,
            link: safe,
          }),
        );
        break;
      }
      case "image":
        add(`[Image: ${t.text || (t as Tokens.Image).href}]`, {
          italic: true,
          color: C.muted,
        });
        break;
      case "br":
        out.push({ text: "\n", ...inherited });
        break;
      default:
        if ("tokens" in t && Array.isArray(t.tokens))
          out.push(...inlineRuns(t.tokens, inherited));
        else if ("text" in t && typeof t.text === "string") add(t.text);
    }
  }
  return out;
}
const isList = (t: Token): t is Tokens.List =>
  t.type === "list" && "items" in t && Array.isArray((t as any).items);
const isImage = (t: Token): t is Tokens.Image =>
  t.type === "image" && "href" in t && typeof (t as any).href === "string";
const isTable = (t: Token): t is Tokens.Table =>
  t.type === "table" && "header" in t && Array.isArray((t as any).header);

/* ---------- Main ---------- */
export async function convertMarkdownToPowerPoint(
  markdown: string,
  opts: MdToPptxOptions = {},
): Promise<Blob> {
  if (!markdown?.trim()) throw new Error("The Markdown file is empty.");
  if (new Blob([markdown]).size > 20 * 1024 * 1024)
    throw new Error("The maximum file size is 20 MB.");
  // Strip front-matter + HTML comments (outside code — marked handles code, so do it token-aware: simple front-matter only)
  let clean = markdown
    .replace(/^\uFEFF/, "")
    .replace(/\r\n?/g, "\n")
    .trim();
  clean = clean.replace(/^---\n[\s\S]*?\n---\n/, "");
  if (!clean)
    throw new Error("The Markdown file contains no readable content.");

  const bodySize = Math.min(24, Math.max(10, opts.fontSize ?? 18));
  const lineH = (fs: number) => (fs / 72) * 1.32; // Calibri line height incl. spacing
  const blockH = (lines: number, fs: number) => lines * lineH(fs) + 0.14; // + padding

  const { default: PptxGenJS } = await import("pptxgenjs");
  const pptx = new PptxGenJS();
  pptx.layout = "LAYOUT_WIDE";
  pptx.author = opts.author ?? "Document Conversion";
  pptx.title = opts.title ?? "Markdown Presentation";
  pptx.theme = { headFontFace: FONT, bodyFontFace: FONT };

  let slide!: PptxGenJS.Slide;
  let slideNo = 0;
  let slideTitle = opts.title ?? "Markdown Presentation";
  let hasContent = false;
  let y = CONTENT_TOP;
  let isTitleSlide = false; // tracks Title Slide vs Title & Content — used by addTitleSlide/addContentSlide

  /* ----- Slide builders: proper PPTX layouts ----- */
  function footer() {
    slide.addText(String(slideNo), {
      x: SW - 1.1,
      y: FOOTER_Y,
      w: 0.42,
      h: 0.2,
      fontFace: FONT,
      fontSize: 9,
      color: C.muted,
      align: "right",
      margin: 0,
    });
  }
  /** Layout 1: Title Slide — ONLY for first H1 / doc title */
  function addTitleSlide(title: string) {
    slideNo++;
    slideTitle = title || "Untitled";
    isTitleSlide = true;
    hasContent = false;
    slide = pptx.addSlide();
    slide.background = { color: C.bg };
    // centred accent bar + big centred title (never overflows: fit shrink ONLY here, box is huge)
    slide.addShape(pptx.ShapeType.rect, {
      x: 0,
      y: 0,
      w: SW,
      h: 0.14,
      fill: { color: C.accent },
      line: { color: C.accent, transparency: 100 },
    });
    slide.addShape(pptx.ShapeType.rect, {
      x: (SW - 1.2) / 2,
      y: 2.35,
      w: 1.2,
      h: 0.06,
      fill: { color: C.accent },
      line: { color: C.accent, transparency: 100 },
    });
    slide.addText(slideTitle, {
      x: 1.2,
      y: 2.6,
      w: SW - 2.4,
      h: 1.6,
      fontFace: FONT,
      fontSize: 40,
      bold: true,
      color: C.primary,
      align: "center",
      valign: "middle",
      margin: 0,
      breakLine: true,
      fit: "shrink",
    });
    slide.addText("Converted from Markdown", {
      x: 1.2,
      y: 4.3,
      w: SW - 2.4,
      h: 0.4,
      fontFace: FONT,
      fontSize: 16,
      color: C.muted,
      align: "center",
      margin: 0,
    });
    footer();
    y = CONTENT_TOP;
  }
  /** Layout 2: Title & Content — every other slide. This is Windows default. */
  function addContentSlide(title: string) {
    slideNo++;
    slideTitle =
      title.replace(/(?: \(continued\))+$/g, "").trim() || "Untitled section";
    isTitleSlide = false;
    hasContent = false;
    slide = pptx.addSlide();
    slide.background = { color: C.bg };
    slide.addShape(pptx.ShapeType.rect, {
      x: 0,
      y: 0,
      w: SW,
      h: 0.12,
      fill: { color: C.accent },
      line: { color: C.accent, transparency: 100 },
    });
    // Title box is TALL (0.95") + fit shrink, so 2-line titles never hit divider
    slide.addText(slideTitle, {
      x: MARGIN_X,
      y: HEADER_TOP,
      w: CONTENT_W,
      h: HEADER_H,
      fontFace: FONT,
      fontSize: 28,
      bold: true,
      color: C.primary,
      margin: 0,
      valign: "middle",
      breakLine: true,
      fit: "shrink",
    });
    slide.addShape(pptx.ShapeType.line, {
      x: MARGIN_X,
      y: DIVIDER_Y,
      w: CONTENT_W,
      h: 0,
      line: { color: C.border, width: 1 },
    });
    footer();
    y = CONTENT_TOP;
  }
  function ensureSlide() {
    if (slideNo === 0) addContentSlide(slideTitle);
  }
  function continueSlide() {
    addContentSlide(`${slideTitle} (continued)`);
  }
  /** Reserve height. If it doesn't fit, move to fresh slide. Never throws. */
  function ensureSpace(h: number) {
    ensureSlide();
    if (y + h <= CONTENT_BOTTOM) return;
    if (hasContent) continueSlide();
    // On fresh slide: if STILL too tall, caller must have split it. We never throw.
  }
  function remainingH() {
    return CONTENT_BOTTOM - y;
  }

  /* ----- Core: draw wrapped lines, paginated, no overlap ----- */
  function drawTextBlock(
    runs: Run[],
    o: {
      fontSize?: number;
      bold?: boolean;
      italic?: boolean;
      color?: string;
      indent?: number;
      quote?: boolean;
    } = {},
  ) {
    const text = plainText(runs);
    if (!text.trim()) return;
    ensureSlide();
    const fs = o.fontSize ?? bodySize;
    const indent = o.indent ?? 0;
    const quotePad = o.quote ? 0.22 : 0;
    const x = MARGIN_X + indent + quotePad;
    const w = CONTENT_W - indent - quotePad;
    const allLines = wrapRuns(
      runs.map((r) => ({
        ...r,
        bold: r.bold || o.bold,
        italic: r.italic || o.italic,
        color: r.color ?? o.color ?? C.body,
      })),
      w,
      fs,
      FONT,
    );
    let idx = 0;
    while (idx < allLines.length) {
      // How many lines fit in remaining space? At least 1.
      const maxLines = Math.max(
        1,
        Math.floor((remainingH() - 0.14) / lineH(fs)),
      );
      // If slide is full, go to next slide FIRST (don't draw 1 line + overflow)
      if (maxLines < 1) {
        continueSlide();
        continue;
      }
      // If even 1 line doesn't fit because y is at bottom, continue
      if (y + blockH(1, fs) > CONTENT_BOTTOM && hasContent) {
        continueSlide();
        continue;
      }
      const pageLines = allLines.slice(idx, idx + maxLines);
      const h = blockH(pageLines.length, fs);
      ensureSpace(h);
      // Recompute after possible continueSlide (y reset)
      const h2 = blockH(pageLines.length, fs);
      if (o.quote) {
        slide.addShape(pptx.ShapeType.roundRect, {
          x: MARGIN_X,
          y,
          w: CONTENT_W,
          h: h2 + 0.12,
          rectRadius: 0.08,
          fill: { color: C.quoteBg },
          line: { color: C.quoteBg, transparency: 100 },
        });
        slide.addShape(pptx.ShapeType.rect, {
          x: MARGIN_X,
          y: y + 0.02,
          w: 0.055,
          h: Math.max(0.2, h2 - 0.04),
          fill: { color: C.accent },
          line: { color: C.accent, transparency: 100 },
        });
      }
      // Convert lines back to pptxgenjs runs with explicit line breaks.
      // NO fit:"shrink" — height is exact, so nothing overflows.
      const pptxRuns: PptxGenJS.TextProps[] = [];
      pageLines.forEach((line, li) => {
        line.forEach((r) =>
          pptxRuns.push({
            text: r.text,
            options: {
              bold: r.bold,
              italic: r.italic,
              strike: r.strike,
              color: r.color,
              fontFace: r.fontFace,
              ...(r.link ? { hyperlink: { url: r.link } } : {}),
            },
          }),
        );
        if (li < pageLines.length - 1)
          pptxRuns.push({ text: "", options: { breakLine: true } });
      });
      slide.addText(pptxRuns, {
        x,
        y: y + (o.quote ? 0.06 : 0),
        w,
        h: h2,
        fontFace: FONT,
        fontSize: fs,
        color: o.color ?? C.body,
        margin: 0.02,
        breakLine: true,
        valign: "top",
        lineSpacingMultiple: 1.15,
        paraSpaceAfter: 2,
      });
      y += h2 + GAP;
      hasContent = true;
      idx += pageLines.length;
      if (idx < allLines.length) continueSlide();
    }
  }

  function addHeading(text: string, level: number) {
    const runs = inlineRuns(marked.lexer(text) as Token[]);
    const title = plainText(runs) || "Untitled section";
    if (level === 1) {
      if (slideNo === 0)
        addTitleSlide(title); // first H1 = Title Slide
      else addContentSlide(title); // later H1 = new Title & Content
      return;
    }
    if (level === 2) {
      addContentSlide(title);
      return;
    }
    const sizes: Record<number, number> = { 3: 22, 4: 20, 5: 18, 6: 16 };
    drawTextBlock(runs, {
      fontSize: sizes[level] ?? 16,
      bold: true,
      color: C.primary,
    });
  }

  function addCodeBlock(code: string, lang = "") {
    ensureSlide();
    const fs = 10;
    // Wrap code to fit width using measured Consolas width (no horizontal overflow)
    const maxW = CONTENT_W - 0.4;
    const rawLines = code.replace(/\t/g, "  ").split("\n");
    const wrapped: string[] = [];
    for (const l of rawLines) {
      if (!l) {
        wrapped.push("");
        continue;
      }
      let rest = l;
      while (measureWidth(rest, fs, false, FONT_MONO) > maxW) {
        let lo = 1,
          hi = rest.length;
        while (lo < hi) {
          const mid = Math.ceil((lo + hi) / 2);
          if (measureWidth(rest.slice(0, mid), fs, false, FONT_MONO) <= maxW)
            lo = mid;
          else hi = mid - 1;
        }
        wrapped.push(rest.slice(0, Math.max(1, lo)));
        rest = rest.slice(Math.max(1, lo));
      }
      wrapped.push(rest);
    }
    const lh = 0.22;
    let idx = 0;
    if (!wrapped.length) wrapped.push("");
    while (idx < wrapped.length) {
      const maxLines = Math.max(1, Math.floor((remainingH() - 0.42) / lh));
      if (y + 0.55 > CONTENT_BOTTOM && hasContent) {
        continueSlide();
        continue;
      }
      const page = wrapped.slice(idx, idx + maxLines);
      const h = Math.min(CONTENT_H, page.length * lh + 0.42);
      ensureSpace(h);
      slide.addShape(pptx.ShapeType.roundRect, {
        x: MARGIN_X,
        y,
        w: CONTENT_W,
        h,
        rectRadius: 0.06,
        fill: { color: C.codeBg },
        line: { color: C.border, width: 0.5 },
      });
      if (lang)
        slide.addText(lang.toUpperCase(), {
          x: MARGIN_X + 0.2,
          y: y + 0.08,
          w: CONTENT_W - 0.4,
          h: 0.18,
          fontFace: FONT,
          fontSize: 7,
          bold: true,
          color: C.muted,
          margin: 0,
        });
      // One textbox, exact height, breakLine:true — no overflow
      slide.addText(page.join("\n"), {
        x: MARGIN_X + 0.2,
        y: y + 0.28,
        w: CONTENT_W - 0.4,
        h: Math.max(0.22, h - 0.34),
        fontFace: FONT_MONO,
        fontSize: fs,
        color: C.primary,
        margin: 0,
        breakLine: true,
        valign: "top",
      });
      y += h + GAP;
      hasContent = true;
      idx += page.length;
      if (idx < wrapped.length) continueSlide();
    }
  }

  async function resolveImage(src: string): Promise<string> {
    if (/^data:image\/(png|jpe?g|gif);base64,/.test(src)) return src;
    if (/^data:image\/webp/i.test(src))
      throw new Error("WebP not supported — use PNG/JPEG.");
    if (!/^https?:\/\//i.test(src))
      throw new Error(`Unsupported image: ${src}`);
    const res = await fetch(src);
    if (!res.ok) throw new Error(`Image load failed: ${src}`);
    const blob = await res.blob();
    if (!blob.type.startsWith("image/")) throw new Error("URL is not an image");
    if (blob.size > 8 * 1024 * 1024) throw new Error("Image too large (>8MB)");
    return new Promise((resolve, reject) => {
      const fr = new FileReader();
      fr.onload = () =>
        typeof fr.result === "string"
          ? resolve(fr.result)
          : reject(new Error("read failed"));
      fr.onerror = () => reject(new Error("read failed"));
      fr.readAsDataURL(blob);
    });
  }
  function imgSize(data: string) {
    return new Promise<{ w: number; h: number }>((resolve, reject) => {
      const im = new Image();
      im.onload = () => resolve({ w: im.naturalWidth, h: im.naturalHeight });
      im.onerror = () => reject(new Error("decode failed"));
      im.src = data;
    });
  }
  async function addImage(tok: Tokens.Image) {
    ensureSlide();
    try {
      const data = await resolveImage(tok.href);
      const d = await imgSize(data);
      const scale = Math.min(CONTENT_W / d.w, (CONTENT_H - 0.2) / d.h, 1);
      const w = d.w * scale,
        h = d.h * scale;
      ensureSpace(h + GAP);
      slide.addImage({
        data,
        x: MARGIN_X + (CONTENT_W - w) / 2,
        y,
        w,
        h,
        altText: tok.text || "Markdown image",
      });
      y += h + GAP;
      hasContent = true;
    } catch {
      drawTextBlock([
        {
          text: `[Image could not be loaded: ${tok.text || tok.href}]`,
          italic: true,
          color: C.muted,
        },
      ]);
    }
  }

  async function addList(tok: Tokens.List, level = 0) {
    for (const [i, item] of tok.items.entries()) {
      const prefix = (item as any).task
        ? (item as any).checked
          ? "☑ "
          : "☐ "
        : tok.ordered
          ? `${Number((tok as any).start || 1) + i}. `
          : "• ";
      const children = (item.tokens ?? []) as Token[];
      const imgs = children.filter(isImage);
      const flushText = (toks: Token[], first: boolean) => {
        const runs = inlineRuns(toks);
        if (!runs.length) return;
        runs.unshift({ text: prefix, bold: true, color: C.accent });
        drawTextBlock(runs, { indent: level * 0.32 });
      };
      if (!imgs.length)
        flushText(
          children.filter((c) => !isList(c)),
          true,
        );
      else {
        let buf: Token[] = [];
        let first = true;
        for (const c of children) {
          if (isImage(c)) {
            if (buf.length) {
              flushText(buf, first);
              first = false;
              buf = [];
            }
            await addImage(c);
          } else if (!isList(c)) buf.push(c);
        }
        if (buf.length) flushText(buf, first);
      }
      for (const c of children) if (isList(c)) await addList(c, level + 1);
    }
  }

  function addTable(tok: Tokens.Table) {
    ensureSlide();
    // Column widths proportional to longest content — no squeezed / overflowed column
    const maxLens = tok.header.map((h, ci) =>
      Math.max(
        h.text.length,
        ...tok.rows.map((r) => (r[ci]?.text ?? "").length),
        4,
      ),
    );
    const total = maxLens.reduce((a, b) => a + b, 0);
    const colW = maxLens.map((l) => Math.max(0.9, (CONTENT_W * l) / total));
    // Normalise to exactly CONTENT_W
    const sum = colW.reduce((a, b) => a + b, 0);
    const widths = colW.map((w) => (w / sum) * CONTENT_W);
    const headers = tok.header.map((c) => c.text);
    const rows = tok.rows.map((r) => headers.map((_, ci) => r[ci]?.text ?? ""));
    const fs = Math.min(bodySize, 16); // tables use slightly smaller font to avoid overflow
    const rowH = (vals: string[]) => {
      const lines = vals.map((v, ci) => {
        const linesForCell = wrapRuns(
          [{ text: v }],
          widths[ci] - 0.16,
          fs,
          FONT,
        );
        return Math.max(1, linesForCell.length);
      });
      return Math.max(0.5, Math.max(...lines) * lineH(fs) + 0.24);
    };
    const headH = rowH(headers);
    const rowHeights = rows.map(rowH);
    const makeRow = (
      vals: string[],
      isHead: boolean,
      ri: number,
    ): PptxGenJS.TableRow =>
      vals.map((v, ci) => ({
        text: v,
        options: {
          bold: isHead,
          color: C.primary,
          fill: {
            color: isHead ? C.tableHead : ri % 2 === 0 ? C.bg : C.tableAlt,
          },
          align: (tok.align?.[ci] as any) ?? "left",
          valign: "middle",
          margin: 0.06,
        },
      }));
    let idx = 0;
    // Header-only table still draws once
    while (idx < rows.length || idx === 0) {
      const pageRows: PptxGenJS.TableRow[] = [makeRow(headers, true, 0)];
      const pageH: number[] = [headH];
      let totalH = headH;
      const start = idx;
      while (
        idx < rows.length &&
        totalH + rowHeights[idx] + GAP <= remainingH()
      ) {
        pageRows.push(makeRow(rows[idx], false, idx));
        pageH.push(rowHeights[idx]);
        totalH += rowHeights[idx];
        idx++;
      }
      if (idx === start && idx < rows.length) {
        if (hasContent) {
          continueSlide();
          continue;
        }
        // Single huge row on fresh slide: clamp (never overflow bottom)
        pageRows.push(makeRow(rows[idx], false, idx));
        pageH.push(Math.min(rowHeights[idx], remainingH() - headH));
        totalH += pageH[pageH.length - 1];
        idx++;
      }
      slide.addTable(pageRows, {
        x: MARGIN_X,
        y,
        w: CONTENT_W,
        h: totalH,
        colW: widths,
        rowH: pageH,
        fontFace: FONT,
        fontSize: fs,
        color: C.primary,
        margin: 0.06,
        border: { type: "solid", color: C.border, pt: 0.7 },
      });
      y += totalH + GAP;
      hasContent = true;
      if (idx < rows.length) continueSlide();
      if (!rows.length) break;
    }
  }

  async function addInline(toks: Token[], o: { quote?: boolean } = {}) {
    let buf: Token[] = [];
    const flush = () => {
      if (buf.length) {
        drawTextBlock(inlineRuns(buf), {
          quote: o.quote,
          italic: o.quote,
          color: o.quote ? C.body : undefined,
        });
        buf = [];
      }
    };
    for (const t of toks) {
      if (isImage(t)) {
        flush();
        await addImage(t);
      } else buf.push(t);
    }
    flush();
  }

  async function addBlock(tok: Token): Promise<void> {
    switch (tok.type) {
      case "heading":
        addHeading((tok as Tokens.Heading).text, (tok as Tokens.Heading).depth);
        break;
      case "paragraph":
        await addInline((tok as Tokens.Paragraph).tokens ?? []);
        break;
      case "text": {
        const t = tok as Tokens.Text;
        if ("tokens" in t && Array.isArray(t.tokens)) await addInline(t.tokens);
        else if ("text" in t) drawTextBlock([{ text: t.text }]);
        break;
      }
      case "list":
        if (isList(tok)) await addList(tok);
        break;
      case "blockquote":
        await addInline((tok as Tokens.Blockquote).tokens ?? [], {
          quote: true,
        });
        break;
      case "code":
        addCodeBlock(
          (tok as Tokens.Code).text,
          (tok as Tokens.Code).lang ?? "",
        );
        break;
      case "table":
        if (isTable(tok)) addTable(tok);
        break;
      case "hr":
        ensureSlide();
        ensureSpace(0.32);
        slide.addShape(pptx.ShapeType.line, {
          x: MARGIN_X,
          y: y + 0.08,
          w: CONTENT_W,
          h: 0,
          line: { color: C.border, width: 1 },
        });
        y += 0.32;
        hasContent = true;
        break;
      case "space":
      case "def":
      case "html":
        break;
      default:
        if ("tokens" in tok && Array.isArray((tok as any).tokens))
          await addInline((tok as any).tokens);
        else if ("text" in tok && typeof (tok as any).text === "string")
          drawTextBlock([{ text: (tok as any).text }]);
    }
  }

  const tokens = marked.lexer(clean);
  for (const t of tokens) await addBlock(t);
  if (slideNo === 0) {
    addTitleSlide(opts.title ?? "Markdown Presentation");
    drawTextBlock([
      { text: "No readable content was found in the Markdown file." },
    ]);
  }

  const out = await pptx.write({ outputType: "blob" });
  if (!(out instanceof Blob) || out.size === 0)
    throw new Error("Could not generate the PowerPoint presentation.");
  return out;
}
