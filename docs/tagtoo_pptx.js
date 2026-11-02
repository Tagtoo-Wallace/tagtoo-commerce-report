// SPDX-License-Identifier: Apache-2.0
/**
 * tagtoo_pptx.js — Canonical Tagtoo presentation theme for PptxGenJS
 * ------------------------------------------------------------------
 * Reproduces the gold-standard "sales analysis" deck EXACTLY.
 * Every project (KOL, Ad Report, Brand Voice, Competitive material) must
 * build its deck through these helpers — never hand-draw backgrounds or
 * fake the logo with text/shapes.
 *
 * Slide archetypes:
 *   addCover, addAgenda, addSection, addContent, addClosing
 *
 * Design system reverse-engineered from the reference deck:
 *   Layout 10 x 5.63 in (25.4 x 14.29 cm, 16:9)
 *   Backgrounds = solid fill + translucent ellipse decorations (NOT images)
 *   Logo = embedded transparent PNG (white on red slides, dark on white slides)
 *
 * Usage:
 *   const pptxgen = require("pptxgenjs");
 *   const T = require("./tagtoo_pptx.js");
 *   const pptx = new pptxgen();
 *   T.init(pptx);
 *   T.addCover(pptx, {...});
 *   ...
 *   await pptx.writeFile({ fileName: "out.pptx" });
 */

const fs = require("fs");
const { LOGO_WHITE, LOGO_DARK } = require("./tagtoo_assets.js");

// ── Layout (inches) ───────────────────────────────────────────────────────────
const LAYOUT_NAME = "TAGTOO_16x9";
const PAGE_W = 10.0;
const PAGE_H = 5.63;

// ── Colors ──────────────────────────────────────────────────────────────────
const C = {
  coverBg:   "66272A",   // cover / section background
  agendaBg:  "8E291B",   // agenda background
  contentBg: "FDFDFD",   // content background
  closeBg:   "873222",   // closing solid fill (brand red)
  red:       "873222",   // brand red
  rose:      "E0B9B3",   // pale rose accent (subtitles/captions on red slides)
  white:     "FFFFFF",
  dark:      "333333",   // primary text
  gray:      "888888",   // secondary text
  caption:   "B2B2B2",   // caption / data source
  deco:      "EBEBEB",   // content decorative arc
  footLine:  "E2E8F0",   // content footer divider line
};

// ── Fonts (content-aware) ─────────────────────────────────────────────────────
// PptxGenJS sets a run's Latin + East-Asian + Complex-Script faces ALL to one
// `fontFace`, so a single run cannot mix a Latin font and a CJK font. We therefore
// pick the face PER RUN by content:
//   • Latin-only text  → LATIN_FONT (crisp for English-only decks)
//   • text with any CJK → CJK_FONT  (Microsoft JhengHei 微軟正黑體: the most universally
//     installed Traditional-Chinese face — built into every Windows since Vista and into
//     Office on macOS — with a real Bold weight → no glyph substitution, no faux-bold
//     doubling. Chosen over PingFang TC because PingFang is macOS-ONLY, so Windows
//     recipients (the majority) saw substituted / garbled CJK. On a bare macOS without
//     Office, PowerPoint falls back to PingFang TC, which is also fine.)
// "Arial" alone has NO Chinese glyphs, which caused the overlapping / garbled CJK.
const LATIN_FONT = "Arial";
const CJK_FONT   = "Microsoft JhengHei";   // 微軟正黑體
const _CJK_RE = /[⺀-鿿豈-﫿＀-￯]/;   // CJK ideographs, kana, fullwidth, compat
function font(text) {
  return _CJK_RE.test(String(text == null ? "" : text)) ? CJK_FONT : LATIN_FONT;
}
const FONT = CJK_FONT;   // back-compat alias (default = CJK font)

// ── Logo placement (inches) ───────────────────────────────────────────────────
const LOGO_RED  = { x: 7.3, y: 0.15, w: 2.1, h: 0.65 };   // white logo, red slides
const LOGO_WHT  = { x: 8.2, y: 0.12, w: 1.2, h: 0.37 };   // dark logo, content slides

// alpha (a:alpha %) → PptxGenJS transparency (% transparent) = 100 - alpha
const tr = (alphaPct) => 100 - alphaPct;

function init(pptx) {
  pptx.defineLayout({ name: LAYOUT_NAME, width: PAGE_W, height: PAGE_H });
  pptx.layout = LAYOUT_NAME;
}

function ell(slide, x, y, w, h, color, alphaPct) {
  slide.addShape("ellipse", {
    x, y, w, h,
    fill: { color, transparency: tr(alphaPct) },
    line: { type: "none" },
  });
}

function pageNum(slide, n, total, color) {
  if (n == null || total == null) return;
  slide.addText(`${n} / ${total}`, {
    x: 8.8, y: 5.38, w: 1.0, h: 0.18,
    fontFace: LATIN_FONT, fontSize: 8, color, align: "right",
  });
}

// ── Cover ──────────────────────────────────────────────────────────────────
function addCover(pptx, o = {}) {
  const s = pptx.addSlide();
  s.background = { color: C.coverBg };
  ell(s, 6.0, -2.0, 8.0, 10.0, C.red,  72);
  ell(s, 7.5,  2.2, 4.5,  4.5, C.rose, 42);
  s.addImage({ data: LOGO_WHITE, ...LOGO_RED });

  if (o.mainTitle)  s.addText(o.mainTitle,  { x: 0.5, y: 1.3,  w: 7.0, h: 0.9,  fontFace: font(o.mainTitle), fontSize: 36, bold: true, color: C.white });
  if (o.subtitleEN) s.addText(o.subtitleEN, { x: 0.5, y: 2.25, w: 7.0, h: 0.45, fontFace: font(o.subtitleEN), fontSize: 16, color: C.rose });
  if (o.reportTitle)s.addText(o.reportTitle,{ x: 0.5, y: 2.8,  w: 7.0, h: 0.6,  fontFace: font(o.reportTitle), fontSize: 22, color: C.white });
  s.addShape("line", { x: 0.5, y: 3.5, w: 5.0, h: 0, line: { color: C.white, width: 1 } });
  if (o.tagline)    s.addText(o.tagline,    { x: 0.5, y: 3.65, w: 6.5, h: 0.3,  fontFace: font(o.tagline), fontSize: 11, color: C.rose });
  if (o.period)     s.addText(o.period,     { x: 0.5, y: 4.0,  w: 6.5, h: 0.3,  fontFace: font(o.period), fontSize: 11, color: C.white });
  if (o.confidential !== false)
    s.addText(o.confidential || "CONFIDENTIAL", { x: 0.5, y: 5.1, w: 3.0, h: 0.25, fontFace: font(o.confidential || "CONFIDENTIAL"), fontSize: 9, color: C.rose, charSpacing: 2 });
  return s;
}

// ── Agenda ───────────────────────────────────────────────────────────────────
function addAgenda(pptx, o = {}) {
  const s = pptx.addSlide();
  s.background = { color: C.agendaBg };
  ell(s, 5.5, 0.5, 6.5, 6.5, C.red, 52);
  s.addImage({ data: LOGO_WHITE, ...LOGO_RED });
  s.addText(o.title || "目錄 Agenda", { x: 0.5, y: 0.35, w: 5.0, h: 0.55, fontFace: font(o.title || "目錄 Agenda"), fontSize: 22, bold: true, color: C.white });

  const items = o.items || [];
  let y = 1.1;
  const step = 0.58;
  items.forEach((it, i) => {
    const no = it.no || String(i + 1).padStart(2, "0");
    s.addText(no,         { x: 0.5,  y, w: 0.5, h: 0.4, fontFace: font(no), fontSize: 11,  bold: true, color: C.rose });
    s.addText(it.title || "", { x: 1.05, y, w: 3.0, h: 0.4, fontFace: font(it.title), fontSize: 11,  bold: true, color: C.white });
    s.addText(it.desc || "",  { x: 4.15, y, w: 4.5, h: 0.4, fontFace: font(it.desc), fontSize: 9.5, color: C.white });
    s.addShape("line", { x: 0.5, y: y + 0.46, w: 8.0, h: 0, line: { color: C.rose, width: 0.75, transparency: 60 } });
    y += step;
  });
  pageNum(s, o.pageNum, o.totalPages, C.white);
  return s;
}

// ── Section divider ────────────────────────────────────────────────────────────
function addSection(pptx, o = {}) {
  const s = pptx.addSlide();
  s.background = { color: C.coverBg };
  ell(s, 5.5, -1.5, 8.0, 9.0, C.red, 65);
  s.addImage({ data: LOGO_WHITE, ...LOGO_RED });
  if (o.number) s.addText(o.number, { x: 0.5, y: 1.6,  w: 1.5, h: 0.7, fontFace: font(o.number), fontSize: 40, bold: true, color: C.rose });
  if (o.title)  s.addText(o.title,  { x: 0.5, y: 2.35, w: 6.8, h: 0.9, fontFace: font(o.title), fontSize: 30, bold: true, color: C.white });
  s.addShape("line", { x: 0.5, y: 3.35, w: 4.0, h: 0, line: { color: C.white, width: 1 } });
  if (o.subtitle) s.addText(o.subtitle, { x: 0.5, y: 3.5, w: 6.8, h: 0.4, fontFace: font(o.subtitle), fontSize: 13, color: C.white });
  pageNum(s, o.pageNum, o.totalPages, C.white);
  return s;
}

// ── Content (white) ─────────────────────────────────────────────────────────
// Returns the slide so the caller can add body content in the area
// roughly x:0.35–9.65, y:1.0–5.2.
function addContent(pptx, o = {}) {
  const s = pptx.addSlide();
  s.background = { color: C.contentBg };
  ell(s,  7.0,  3.0, 5.5, 5.0, C.deco, 25);
  ell(s, -1.2, -0.8, 3.5, 3.5, C.deco, 20);
  s.addImage({ data: LOGO_DARK, ...LOGO_WHT });

  if (o.title)    s.addText(o.title,    { x: 0.35, y: 0.15, w: 9.3, h: 0.42, fontFace: font(o.title), fontSize: 17, bold: true, color: C.dark });
  if (o.subtitle) s.addText(o.subtitle, { x: 0.35, y: 0.6,  w: 9.3, h: 0.26, fontFace: font(o.subtitle), fontSize: 9.5, color: C.gray });

  // Footer
  s.addShape("line", { x: 0.35, y: 5.28, w: 9.3, h: 0, line: { color: C.footLine, width: 0.75 } });
  s.addText("塔圖科技 Tagtoo", { x: 0.35, y: 5.33, w: 2.0, h: 0.18, fontFace: font("塔圖科技 Tagtoo"), fontSize: 7.5, color: C.dark });
  if (o.footerCenter) s.addText(o.footerCenter, { x: 3.4, y: 5.33, w: 3.5, h: 0.18, fontFace: font(o.footerCenter), fontSize: 7.5, color: C.gray, align: "center" });
  pageNum(s, o.pageNum, o.totalPages, C.gray);
  s.addText(o.dataSource || "Data source: Tagtoo", { x: 0.35, y: 5.48, w: 9.3, h: 0.14, fontFace: font(o.dataSource || "Data source: Tagtoo"), fontSize: 6.5, color: C.caption });
  return s;
}

// ── Closing ───────────────────────────────────────────────────────────────────
function addClosing(pptx, o = {}) {
  const s = pptx.addSlide();
  s.background = { color: C.closeBg };          // solid #873222
  ell(s, 5.5, -1.5, 8.0, 9.0, C.coverBg, 62);
  s.addImage({ data: LOGO_WHITE, ...LOGO_RED });
  s.addText(o.title || "Thank You", { x: 0.5, y: 1.7, w: 7.5, h: 1.3, fontFace: font(o.title || "Thank You"), fontSize: 54, bold: true, color: C.white });
  if (o.subtitle) s.addText(o.subtitle, { x: 0.5, y: 3.1, w: 7.0, h: 0.5, fontFace: font(o.subtitle), fontSize: 14, color: C.white });
  s.addShape("line", { x: 0.5, y: 3.75, w: 4.0, h: 0, line: { color: C.rose, width: 1 } });
  if (o.poweredBy) s.addText(o.poweredBy, { x: 0.5, y: 3.9, w: 6.0, h: 0.35, fontFace: font(o.poweredBy), fontSize: 11, color: C.rose });
  pageNum(s, o.pageNum, o.totalPages, C.white);
  return s;
}

// ── Image cover-fit (preserve aspect, no distortion) ─────────────────────────
// Read pixel dimensions from a JPEG/PNG buffer (no external deps).
function _imageSize(buf) {
  if (buf.length > 24 && buf[0] === 0x89 && buf[1] === 0x50) {       // PNG
    return { w: buf.readUInt32BE(16), h: buf.readUInt32BE(20) };
  }
  if (buf[0] === 0xFF && buf[1] === 0xD8) {                          // JPEG
    let o = 2;
    while (o < buf.length - 8) {
      if (buf[o] !== 0xFF) { o++; continue; }
      const m = buf[o + 1];
      if (m >= 0xC0 && m <= 0xCF && m !== 0xC4 && m !== 0xC8 && m !== 0xCC) {
        return { h: buf.readUInt16BE(o + 5), w: buf.readUInt16BE(o + 7) };
      }
      o += 2 + buf.readUInt16BE(o + 2);
    }
  }
  return null;
}

/**
 * Place an image so it COVERS the box {x,y,w,h} with NO distortion (aspect preserved):
 *   - landscape source → fills the box height, cropped evenly left/right (centered)
 *   - portrait source  → fills the box width, cropped evenly top/bottom (centered)
 *   - square           → exact fill
 * src is { path } (local file) or { data } (base64 data URI). Returns true if placed.
 */
function coverImage(slide, src, box) {
  try {
    let buf = null;
    const imgRef = src.data ? { data: src.data } : { path: src.path };
    if (src.data && /^data:/.test(String(src.data))) {
      buf = Buffer.from(String(src.data).split(",")[1] || "", "base64");
    } else if (src.path && fs.existsSync(src.path)) {
      buf = fs.readFileSync(src.path);
    } else {
      return false;
    }
    const dim = _imageSize(buf);
    if (!dim || !dim.w || !dim.h) {
      // Dimensions unreadable → fall back to PptxGenJS centered cover (still no distortion).
      slide.addImage({ ...imgRef, x: box.x, y: box.y, w: box.w, h: box.h,
        sizing: { type: "cover", w: box.w, h: box.h } });
      return true;
    }
    const s  = Math.max(box.w / dim.w, box.h / dim.h);
    const fw = dim.w * s, fh = dim.h * s;
    slide.addImage({
      ...imgRef,
      x: box.x, y: box.y, w: fw, h: fh,
      sizing: { type: "crop", w: box.w, h: box.h, x: (fw - box.w) / 2, y: (fh - box.h) / 2 },
    });
    return true;
  } catch (e) {
    return false;
  }
}

module.exports = {
  LAYOUT_NAME, PAGE_W, PAGE_H, COLORS: C,
  FONT, LATIN_FONT, CJK_FONT, font,
  init, addCover, addAgenda, addSection, addContent, addClosing,
  coverImage,
};
