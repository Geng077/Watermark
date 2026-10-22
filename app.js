(function () {
  "use strict";

  const PDFLib = window.PDFLib;
  const fontkit = window.fontkit;
  const pdfjsLib = window.pdfjsLib;
  const WM_FONTS = window.WM_FONTS || {};
  const PDF_WORKER_SRC = window.__WM_PDF_WORKER_URL__ || "vendor/pdf.worker.min.js";

  /* ============================ 字体注册表 ============================ */
  const FONT_MAP = {
    "noto-sans-sc": { name: "思源黑体", family: "Noto Sans CJK SC", weight: 400, kind: "ttf" },
    "noto-sans-sc-bold": { name: "思源黑体 · 粗", family: "Noto Sans CJK SC", weight: 700, kind: "ttf" },
    "ma-shan-zheng": { name: "马善政 · 楷体", family: "Ma Shan Zheng", weight: 400, kind: "ttf" },
    "helvetica": { name: "Helvetica（拉丁）", family: "Helvetica, Arial, sans-serif", weight: 400, builtin: "Helvetica" },
    "helvetica-bold": { name: "Helvetica Bold（拉丁）", family: "Helvetica, Arial, sans-serif", weight: 700, builtin: "HelveticaBold" },
    "times": { name: "Times Roman（拉丁）", family: "Times New Roman, Georgia, serif", weight: 400, builtin: "TimesRoman" },
    "times-bold": { name: "Times Roman Bold（拉丁）", family: "Times New Roman, Georgia, serif", weight: 700, builtin: "TimesRomanBold", hidden: true },
    "courier": { name: "Courier（拉丁）", family: "Courier New, Courier, monospace", weight: 400, builtin: "Courier" },
    "courier-bold": { name: "Courier Bold（拉丁）", family: "Courier New, Courier, monospace", weight: 700, builtin: "CourierBold", hidden: true },
  };

  Object.keys(FONT_MAP).forEach((id) => {
    if (!FONT_MAP[id].builtin) FONT_MAP[id].custom = true;
  });

  const SWATCH_COLORS = [
    "#e11d48", "#dc2626", "#f97316", "#f59e0b", "#16a34a",
    "#0d9488", "#2563eb", "#6366f1", "#9333ea", "#64748b",
    "#334155", "#000000", "#ffffff",
  ];

  const BOLD_DIRS = [[1, 0], [-1, 0], [0, 1], [0, -1], [0.7, 0.7], [-0.7, 0.7], [0.7, -0.7], [-0.7, -0.7]];

  const PRESETS = {
    default: { type: "text", text: "水印", fontId: "noto-sans-sc", fontSize: 64, color: "#94a3b8", opacity: 30, angle: 45, letterSpacing: 0, bold: false, outline: false, shadow: false, layout: "tile", tileGapX: 80, tileGapY: 80 },
    confidential: { type: "text", text: "机密", fontId: "noto-sans-sc-bold", fontSize: 88, color: "#e11d48", opacity: 35, angle: 0, letterSpacing: 0, bold: true, outline: true, shadow: false, layout: "single", position: "c", margin: 40 },
    stamp: { type: "text", text: "机密", fontId: "ma-shan-zheng", fontSize: 120, color: "#dc2626", opacity: 92, angle: 0, letterSpacing: 10, bold: false, outline: false, shadow: false, layout: "single", position: "c", margin: 40 },
    diagonal: { type: "text", text: "水印", fontId: "noto-sans-sc", fontSize: 56, color: "#3b82f6", opacity: 28, angle: -35, letterSpacing: 0, bold: false, outline: false, shadow: false, layout: "tile", tileGapX: 130, tileGapY: 90 },
    light: { type: "text", text: "内部资料", fontId: "noto-sans-sc", fontSize: 48, color: "#cbd5e1", opacity: 12, angle: 0, letterSpacing: 0, bold: false, outline: false, shadow: false, layout: "tile", tileGapX: 60, tileGapY: 70 },
  };

  /* ============================ 状态 ============================ */
  const state = {
    file: null,
    fileType: null, // 'pdf' | 'image'
    fileBuffer: null,
    imageEl: null,
    imageType: null, // 'png' | 'jpeg' | 'webp'
    pdfJsDoc: null,
    pageCount: 1,
    currentPage: 1,
    pageW: 0,
    pageH: 0,
    scale: 1,
    baseCanvas: null,
    overlayDoc: null,
    overlayFontCache: new Map(),
    overlayImageCache: {},

    type: "text",
    text: "水印",
    fontId: "noto-sans-sc",
    fontSize: 64,
    color: "#e11d48",
    opacity: 40,
    angle: -30,
    letterSpacing: 0,
    bold: false,
    outline: false,
    shadow: false,
    lineSpacing: 1.1,

    wmImageEl: null,
    wmImageBytes: null,
    wmImageType: null,
    imageScale: 30,

    layout: "tile",
    position: "c",
    margin: 40,
    offsetX: 0,
    offsetY: 0,
    tileGapX: 60,
    tileGapY: 60,
    layer: "under",

    pagesMode: "all",
    pagesRange: "1",
  };

  /* ============================ DOM ============================ */
  const $ = (id) => document.getElementById(id);
  const els = {
    dropzone: $("dropzone"),
    fileInput: $("fileInput"),
    fileInfo: $("fileInfo"),
    fileName: $("fileName"),
    fileMeta: $("fileMeta"),
    changeFileBtn: $("changeFileBtn"),
    settingsCard: $("settingsCard"),
    typeSeg: $("typeSeg"),
    textSection: $("textSection"),
    imageSection: $("imageSection"),
    presetChips: $("presetChips"),
    textInput: $("textInput"),
    fontSelect: $("fontSelect"),
    fontSizeRange: $("fontSizeRange"),
    fontSizeInput: $("fontSizeInput"),
    colorInput: $("colorInput"),
    colorHex: $("colorHex"),
    colorSwatches: $("colorSwatches"),
    opacityRange: $("opacityRange"),
    opacityInput: $("opacityInput"),
    angleRange: $("angleRange"),
    angleInput: $("angleInput"),
    letterSpacingRange: $("letterSpacingRange"),
    letterSpacingInput: $("letterSpacingInput"),
    boldToggle: $("boldToggle"),
    outlineToggle: $("outlineToggle"),
    shadowToggle: $("shadowToggle"),
    imageDropzone: $("imageDropzone"),
    watermarkImageInput: $("watermarkImageInput"),
    watermarkImagePreview: $("watermarkImagePreview"),
    imageDropzoneEmpty: $("imageDropzoneEmpty"),
    imageRemoveBtn: $("imageRemoveBtn"),
    imageScaleRange: $("imageScaleRange"),
    imageScaleInput: $("imageScaleInput"),
    layoutSeg: $("layoutSeg"),
    positionSection: $("positionSection"),
    positionGrid: $("positionGrid"),
    marginRange: $("marginRange"),
    marginInput: $("marginInput"),
    offsetXRange: $("offsetXRange"),
    offsetXInput: $("offsetXInput"),
    offsetYRange: $("offsetYRange"),
    offsetYInput: $("offsetYInput"),
    tileSection: $("tileSection"),
    tileGapXRange: $("tileGapXRange"),
    tileGapXInput: $("tileGapXInput"),
    tileGapYRange: $("tileGapYRange"),
    tileGapYInput: $("tileGapYInput"),
    layerSeg: $("layerSeg"),
    layerField: $("layerField"),
    pagesSeg: $("pagesSeg"),
    pagesRangeWrap: $("pagesRangeWrap"),
    pagesRange: $("pagesRange"),
    previewCanvas: $("previewCanvas"),
    emptyState: $("emptyState"),
    canvasScroll: $("canvasScroll"),
    pageInfo: $("pageInfo"),
    pager: $("pager"),
    prevPageBtn: $("prevPageBtn"),
    nextPageBtn: $("nextPageBtn"),
    pageIndicator: $("pageIndicator"),
    downloadBtn: $("downloadBtn"),
    statusText: $("statusText"),
    toast: $("toast"),
  };

  /* ============================ 字体测量（独立文档） ============================ */
  let scratchDoc = null;
  const measureFontCache = new Map();
  const fontBytesCache = new Map();
  const fontFaceCache = new Map();

  function base64ToBytes(b64) {
    const bin = atob(b64);
    const bytes = new Uint8Array(bin.length);
    for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
    return bytes;
  }

  function getFontBytes(fontId) {
    const f = FONT_MAP[fontId];
    if (!f || f.builtin) return null;
    if (fontBytesCache.has(fontId)) return fontBytesCache.get(fontId);
    const b64 = WM_FONTS[fontId];
    if (!b64) throw new Error("缺少字体数据：" + fontId);
    const bytes = base64ToBytes(b64);
    fontBytesCache.set(fontId, bytes);
    return bytes;
  }

  async function getScratchDoc() {
    if (!scratchDoc) {
      scratchDoc = await PDFLib.PDFDocument.create();
      scratchDoc.registerFontkit(fontkit);
    }
    return scratchDoc;
  }

  async function getMeasureFont(fontId) {
    if (measureFontCache.has(fontId)) return measureFontCache.get(fontId);
    const doc = await getScratchDoc();
    const f = FONT_MAP[fontId] || FONT_MAP["noto-sans-sc"];
    let font;
    if (f.builtin) font = await doc.embedFont(PDFLib.StandardFonts[f.builtin]);
    else font = await doc.embedFont(getFontBytes(fontId));
    measureFontCache.set(fontId, font);
    return font;
  }

  async function getOutputFont(doc, cache, fontId) {
    if (cache.has(fontId)) return cache.get(fontId);
    const f = FONT_MAP[fontId] || FONT_MAP["noto-sans-sc"];
    let font;
    if (f.builtin) font = await doc.embedFont(PDFLib.StandardFonts[f.builtin]);
    else font = await doc.embedFont(getFontBytes(fontId));
    cache.set(fontId, font);
    return font;
  }

  async function ensureCanvasFont(fontId) {
    const f = FONT_MAP[fontId];
    if (!f || !f.custom) return;
    if (fontFaceCache.has(fontId)) return fontFaceCache.get(fontId);
    const p = (async () => {
      const b64 = WM_FONTS[fontId];
      if (!b64) return;
      const mime = f.kind === "ttf" ? "font/ttf" : "font/otf";
      const ff = new FontFace(f.family, `url(data:${mime};base64,${b64})`, {
        weight: String(f.weight),
        style: "normal",
      });
      await ff.load();
      document.fonts.add(ff);
    })();
    fontFaceCache.set(fontId, p);
    return p;
  }

  function preloadFonts() {
    Object.keys(FONT_MAP).forEach((id) => {
      if (FONT_MAP[id].custom) ensureCanvasFont(id).catch(() => {});
    });
  }

  /* ============================ 工具函数 ============================ */
  function clamp(v, min, max) {
    return Math.min(max, Math.max(min, v));
  }

  function rgb01(r, g, b) {
    return PDFLib.rgb(r / 255, g / 255, b / 255);
  }

  function hexToRgb01(hex) {
    const h = hex.replace("#", "");
    const full = h.length === 3 ? h.split("").map((c) => c + c).join("") : h;
    const n = parseInt(full, 16);
    return rgb01((n >> 16) & 255, (n >> 8) & 255, n & 255);
  }

  function contrastColor(hex) {
    const h = hex.replace("#", "");
    const full = h.length === 3 ? h.split("").map((c) => c + c).join("") : h;
    const n = parseInt(full, 16);
    const r = (n >> 16) & 255, g = (n >> 8) & 255, b = n & 255;
    return 0.299 * r + 0.587 * g + 0.114 * b > 160 ? "#000000" : "#ffffff";
  }

  function formatBytes(n) {
    if (n < 1024) return n + " B";
    if (n < 1024 * 1024) return (n / 1024).toFixed(1) + " KB";
    return (n / 1024 / 1024).toFixed(2) + " MB";
  }

  function replaceExt(name, suffix) {
    const i = name.lastIndexOf(".");
    return (i > 0 ? name.slice(0, i) : name) + suffix;
  }

  function loadImageElement(url) {
    return new Promise((res, rej) => {
      const img = new Image();
      img.onload = () => res(img);
      img.onerror = () => rej(new Error("图片解码失败"));
      img.src = url;
    });
  }

  function setStatus(msg) {
    els.statusText.textContent = msg;
  }

  let toastTimer = null;
  function toast(msg, isError) {
    const t = els.toast;
    t.textContent = msg;
    t.classList.toggle("error", !!isError);
    t.classList.add("show");
    clearTimeout(toastTimer);
    toastTimer = setTimeout(() => t.classList.remove("show"), 2800);
  }

  function saveBlob(blob, name) {
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = name;
    document.body.appendChild(a);
    a.click();
    a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 5000);
  }

  function updateRangeFill(input) {
    const min = +input.min, max = +input.max, val = +input.value;
    const p = max > min ? ((val - min) / (max - min)) * 100 : 0;
    input.style.setProperty("--fill", p + "%");
  }

  function textWidthPts(font, text, size, letterSpacing) {
    if (letterSpacing <= 0) return font.widthOfTextAtSize(text, size);
    const chars = Array.from(text);
    let w = 0;
    for (const ch of chars) w += font.widthOfTextAtSize(ch, size);
    return w + letterSpacing * (chars.length - 1);
  }

  function resolveFont(fontId, bold) {
    const id = FONT_MAP[fontId] ? fontId : "noto-sans-sc";
    if (!bold) return { id, syntheticBold: false };
    const boldMap = {
      "noto-sans-sc": "noto-sans-sc-bold",
      "noto-sans-sc-bold": "noto-sans-sc-bold",
      "helvetica": "helvetica-bold",
      "helvetica-bold": "helvetica-bold",
      "times": "times-bold",
      "times-bold": "times-bold",
      "courier": "courier-bold",
      "courier-bold": "courier-bold",
    };
    if (boldMap[id]) return { id: boldMap[id], syntheticBold: false };
    return { id, syntheticBold: id === "ma-shan-zheng" };
  }

  /* ============================ 水印布局 ============================ */
  function centerToOrigin(cx, cy, w, h, angleRad) {
    const c = Math.cos(angleRad), s = Math.sin(angleRad);
    return { x: cx - (w / 2) * c + (h / 2) * s, y: cy - (w / 2) * s - (h / 2) * c };
  }

  function getCenters(pageW, pageH, w, h) {
    const s = state;
    if (s.layout === "tile") {
      const a = (s.angle * Math.PI) / 180;
      const rw = Math.abs(w * Math.cos(a)) + Math.abs(h * Math.sin(a));
      const rh = Math.abs(w * Math.sin(a)) + Math.abs(h * Math.cos(a));
      const stepX = Math.max(rw + s.tileGapX, 1);
      const stepY = Math.max(rh + s.tileGapY, 1);
      const out = [];
      for (let cx = rw / 2; cx <= pageW + rw / 2 + 0.001; cx += stepX) {
        for (let cy = rh / 2; cy <= pageH + rh / 2 + 0.001; cy += stepY) {
          out.push({ x: cx, y: cy });
        }
      }
      return out;
    }

    const m = s.margin;
    const xs = { nw: m, n: pageW / 2, ne: pageW - m, w: m, c: pageW / 2, e: pageW - m, sw: m, s: pageW / 2, se: pageW - m };
    const ys = { nw: pageH - m, n: pageH - m, ne: pageH - m, w: pageH / 2, c: pageH / 2, e: pageH / 2, sw: m, s: m, se: m };
    const cx = xs[s.position] ?? pageW / 2;
    const cy = ys[s.position] ?? pageH / 2;
    return [{ x: cx + s.offsetX, y: cy + s.offsetY }];
  }

  async function computeWatermarkItems(pageW, pageH) {
    const s = state;
    const items = [];

    if (s.type === "text") {
      if (!s.text || !s.text.trim()) return items;
      const lines = s.text.split(/\r?\n/);
      const resolved = resolveFont(s.fontId, s.bold);
      const font = await getMeasureFont(resolved.id);
      const lineH = s.fontSize * s.lineSpacing;
      const widths = lines.map((l) => textWidthPts(font, l, s.fontSize, s.letterSpacing));
      const blockW = Math.max.apply(null, widths);
      const blockH = lineH * (lines.length - 1) + s.fontSize;
      const centers = getCenters(pageW, pageH, blockW, blockH);
      const f = FONT_MAP[resolved.id];

      for (const c of centers) {
        lines.forEach((line, i) => {
          if (!line) return;
          const lineY = c.y + blockH / 2 - s.fontSize / 2 - i * lineH;
          items.push({
            kind: "text",
            text: line,
            x: c.x,
            y: lineY,
            size: s.fontSize,
            angle: s.angle,
            color: s.color,
            opacity: s.opacity,
            fontId: resolved.id,
            letterSpacing: s.letterSpacing,
            outline: s.outline,
            shadow: s.shadow,
            syntheticBold: resolved.syntheticBold,
            weight: f.weight,
            cssFamily: f.custom ? '"' + f.family + '"' : f.family,
            outlineWidth: Math.max(1.2, s.fontSize * 0.035),
            outlineColor: contrastColor(s.color),
            shadowOffset: Math.max(1.5, s.fontSize * 0.05),
          });
        });
      }
      return items;
    }

    // 图片水印
    if (!s.wmImageEl) return items;
    const w = pageW * (s.imageScale / 100);
    const h = w * (s.wmImageEl.naturalHeight / s.wmImageEl.naturalWidth);
    const centers = getCenters(pageW, pageH, w, h);
    for (const c of centers) {
      items.push({
        kind: "image",
        img: s.wmImageEl,
        x: c.x,
        y: c.y,
        w,
        h,
        angle: s.angle,
        opacity: s.opacity,
      });
    }
    return items;
  }

  /* ============================ 绘制（预览） ============================ */
  function drawTextOnCanvas(ctx, it, scale) {
    const px = it.size * scale;
    ctx.font = it.weight + " " + px + "px " + it.cssFamily;
    ctx.fillStyle = it.color;
    if ("letterSpacing" in ctx) ctx.letterSpacing = (it.letterSpacing * scale) + "px";

    if (it.shadow) {
      const so = it.shadowOffset * scale;
      ctx.save();
      ctx.fillStyle = "rgba(0,0,0,0.55)";
      ctx.globalAlpha = (it.opacity / 100) * 0.55;
      ctx.fillText(it.text, so, so);
      ctx.restore();
    }

    if (it.syntheticBold) {
      const b = Math.max(0.5, px * 0.014);
      for (const d of BOLD_DIRS) ctx.fillText(it.text, d[0] * b, d[1] * b);
    }

    if (it.outline) {
      const o = it.outlineWidth * scale;
      ctx.save();
      ctx.fillStyle = it.outlineColor;
      for (const d of BOLD_DIRS) ctx.fillText(it.text, d[0] * o, d[1] * o);
      ctx.restore();
    }

    ctx.fillText(it.text, 0, 0);
    if ("letterSpacing" in ctx) ctx.letterSpacing = "0px";
  }

  function drawItemsOnCanvas(ctx, items, scale, pageH) {
    ctx.save();
    ctx.textAlign = "center";
    ctx.textBaseline = "middle";
    for (const it of items) {
      const px = it.x * scale;
      const py = (pageH - it.y) * scale;
      ctx.save();
      ctx.translate(px, py);
      if (it.angle) ctx.rotate((it.angle * Math.PI) / 180);
      ctx.globalAlpha = it.opacity / 100;
      if (it.kind === "text") {
        drawTextOnCanvas(ctx, it, scale);
      } else {
        const w = it.w * scale, h = it.h * scale;
        ctx.drawImage(it.img, -w / 2, -h / 2, w, h);
      }
      ctx.restore();
    }
    ctx.restore();
  }

  function watermarkReady() {
    return state.type === "text" ? !!state.text && !!state.text.trim() : !!state.wmImageEl;
  }

  function pageInRange(pageNumber) {
    if (state.pagesMode === "all") return true;
    const indices = getPageIndices(state.pageCount);
    return indices ? indices.includes(pageNumber - 1) : true;
  }

  let overlaySeq = 0;
  let overlayChain = Promise.resolve();

  function requestOverlayPreview() {
    const seq = ++overlaySeq;
    overlayChain = overlayChain
      .then(async () => {
        if (seq !== overlaySeq) return;
        await renderOverlayPreview(seq);
      })
      .catch((e) => console.error(e));
  }

  async function renderOverlayPreview(seq) {
    if (!state.baseCanvas || state.fileType !== "pdf") return;
    const base = state.baseCanvas;
    const canvas = els.previewCanvas;
    canvas.width = base.width;
    canvas.height = base.height;
    canvas.classList.remove("hidden");
    els.emptyState.classList.add("hidden");
    const ctx = canvas.getContext("2d");

    if (!watermarkReady() || !pageInRange(state.currentPage)) {
      ctx.clearRect(0, 0, canvas.width, canvas.height);
      ctx.drawImage(base, 0, 0);
      return;
    }

    try {
      if (!state.overlayDoc) {
        state.overlayDoc = await PDFLib.PDFDocument.create();
        state.overlayDoc.registerFontkit(fontkit);
      }
      if (seq !== overlaySeq) return;

      const doc = state.overlayDoc;
      while (doc.getPageCount() > 0) doc.removePage(0);
      const page = doc.addPage([state.pageW, state.pageH]);
      const items = await computeWatermarkItems(state.pageW, state.pageH);
      if (seq !== overlaySeq) return;
      for (const it of items) {
        await drawItemOnPdf(doc, page, it, state.overlayFontCache, state.overlayImageCache);
      }
      if (seq !== overlaySeq) return;

      const bytes = await doc.save();
      if (seq !== overlaySeq) return;

      const pdoc = await pdfjsLib.getDocument({ data: bytes.slice(0) }).promise;
      try {
        const p = await pdoc.getPage(1);
        const vp = p.getViewport({ scale: state.scale });
        const overlay = document.createElement("canvas");
        overlay.width = Math.round(vp.width);
        overlay.height = Math.round(vp.height);
        const octx = overlay.getContext("2d");
        await p.render({ canvasContext: octx, viewport: vp, background: "rgba(255,255,255,0)" }).promise;
        if (seq !== overlaySeq) return;

        ctx.clearRect(0, 0, canvas.width, canvas.height);
        if (state.layer === "under") {
          ctx.drawImage(overlay, 0, 0, canvas.width, canvas.height);
          ctx.drawImage(base, 0, 0);
        } else {
          ctx.drawImage(base, 0, 0);
          ctx.drawImage(overlay, 0, 0, canvas.width, canvas.height);
        }
      } finally {
        await pdoc.destroy();
      }
    } catch (e) {
      console.error(e);
      drawPreview();
    }
  }

  async function drawPreview() {
    if (!state.baseCanvas) return;
    const canvas = els.previewCanvas;
    canvas.width = state.baseCanvas.width;
    canvas.height = state.baseCanvas.height;
    canvas.classList.remove("hidden");
    els.emptyState.classList.add("hidden");

    const ctx = canvas.getContext("2d");
    ctx.clearRect(0, 0, canvas.width, canvas.height);
    const under = state.layer === "under" && state.fileType === "pdf";
    if (watermarkReady()) {
      try {
        if (state.type === "text") await ensureCanvasFont(resolveFont(state.fontId, state.bold).id);
        const items = await computeWatermarkItems(state.pageW, state.pageH);
        if (under) {
          drawItemsOnCanvas(ctx, items, state.scale, state.pageH);
          ctx.drawImage(state.baseCanvas, 0, 0);
        } else {
          ctx.drawImage(state.baseCanvas, 0, 0);
          drawItemsOnCanvas(ctx, items, state.scale, state.pageH);
        }
        return;
      } catch (e) {
        console.error(e);
      }
    }
    ctx.drawImage(state.baseCanvas, 0, 0);
  }

  let previewTimer = null;
  function schedulePreview() {
    clearTimeout(previewTimer);
    previewTimer = setTimeout(() => {
      if (state.fileType === "pdf") requestOverlayPreview();
      else drawPreview();
    }, 260);
  }

  /* ============================ 文件加载 ============================ */
  async function loadPDF(file) {
    const buf = await file.arrayBuffer();
    pdfjsLib.GlobalWorkerOptions.workerSrc = PDF_WORKER_SRC;
    const doc = await pdfjsLib.getDocument({ data: buf.slice(0) }).promise;
    state.pdfJsDoc = doc;
    state.fileBuffer = buf;
    state.pageCount = doc.numPages;
    state.currentPage = 1;
    await renderCurrentPage();
  }

  async function renderCurrentPage() {
    const page = await state.pdfJsDoc.getPage(state.currentPage);
    const vp1 = page.getViewport({ scale: 1 });
    state.pageW = vp1.width;
    state.pageH = vp1.height;
    const maxDim = Math.max(vp1.width, vp1.height);
    state.scale = clamp(2200 / maxDim, 0.75, 2.5);
    const vp = page.getViewport({ scale: state.scale });
    const canvas = document.createElement("canvas");
    canvas.width = Math.floor(vp.width);
    canvas.height = Math.floor(vp.height);
    const ctx = canvas.getContext("2d");
    await page.render({ canvasContext: ctx, viewport: vp, background: "rgba(255,255,255,0)" }).promise;
    state.baseCanvas = canvas;
    requestOverlayPreview();
  }

  async function loadImage(file) {
    const buf = await file.arrayBuffer();
    const url = URL.createObjectURL(new Blob([buf], { type: file.type }));
    const img = await loadImageElement(url);
    state.imageEl = img;
    state.fileBuffer = buf;
    state.imageType = file.type.includes("png") ? "png" : file.type.includes("webp") ? "webp" : "jpeg";
    state.pageW = img.naturalWidth;
    state.pageH = img.naturalHeight;
    state.scale = 1;
    state.pageCount = 1;
    state.currentPage = 1;
    const canvas = document.createElement("canvas");
    canvas.width = img.naturalWidth;
    canvas.height = img.naturalHeight;
    const ctx = canvas.getContext("2d");
    ctx.drawImage(img, 0, 0);
    state.baseCanvas = canvas;
    drawPreview();
  }

  async function handleFiles(files) {
    const f = files && files[0];
    if (!f) return;
    const ext = (f.name.split(".").pop() || "").toLowerCase();
    const isPdf = ext === "pdf" || f.type === "application/pdf";
    const isImg = ["png", "jpg", "jpeg", "webp"].includes(ext) || /^image\//.test(f.type);
    if (!isPdf && !isImg) {
      toast("不支持该文件类型，请选择 PDF 或图片", true);
      return;
    }

    state.file = f;
    state.fileType = isPdf ? "pdf" : "image";
    state.baseCanvas = null;
    state.overlayDoc = null;
    state.overlayFontCache = new Map();
    state.overlayImageCache = {};
    overlaySeq++;
    els.settingsCard.classList.remove("hidden");
    els.emptyState.classList.add("hidden");
    setStatus("正在加载文件…");
    try {
      if (isPdf) await loadPDF(f);
      else await loadImage(f);
      updateFileInfo();
      updatePager();
      setStatus("文件已就绪");
    } catch (e) {
      console.error(e);
      toast("文件加载失败：" + (e.message || e), true);
      setStatus("就绪");
    }
  }

  function updateFileInfo() {
    els.fileName.textContent = state.file.name;
    const size = formatBytes(state.file.size);
    const desc = state.fileType === "pdf" ? state.pageCount + " 页" : state.pageW + " × " + state.pageH + " px";
    els.fileMeta.textContent = desc + " · " + size;
    els.dropzone.classList.add("hidden");
    els.fileInfo.classList.remove("hidden");
    els.downloadBtn.disabled = false;
    syncLayerUI();
  }

  function updatePager() {
    const isPdf = state.fileType === "pdf";
    els.pager.classList.toggle("hidden", !isPdf || state.pageCount <= 1);
    els.pageIndicator.textContent = state.currentPage + " / " + state.pageCount;
    els.pageInfo.textContent = isPdf ? "第 " + state.currentPage + " / " + state.pageCount + " 页" : "";
    els.pageInfo.classList.toggle("hidden", !isPdf);
  }

  async function goPage(n) {
    if (state.fileType !== "pdf") return;
    n = clamp(n, 1, state.pageCount);
    if (n === state.currentPage) return;
    overlaySeq++;
    state.currentPage = n;
    updatePager();
    await renderCurrentPage();
  }

  /* ============================ 水印图片上传 ============================ */
  async function handleWmImage(file) {
    if (!file) return;
    const isSvg = file.type === "image/svg+xml" || /\.svg$/i.test(file.name);
    const buf = await file.arrayBuffer();
    const url = URL.createObjectURL(new Blob([buf], { type: file.type }));
    const img = await loadImageElement(url);

    if (isSvg) {
      const c = document.createElement("canvas");
      c.width = Math.max(1, img.naturalWidth || 512);
      c.height = Math.max(1, img.naturalHeight || 512);
      const cctx = c.getContext("2d");
      cctx.drawImage(img, 0, 0, c.width, c.height);
      const pngBlob = await new Promise((r) => c.toBlob(r, "image/png"));
      state.wmImageBytes = new Uint8Array(await pngBlob.arrayBuffer());
      state.wmImageType = "png";
    } else {
      state.wmImageBytes = new Uint8Array(buf);
      state.wmImageType = file.type.includes("png") ? "png" : "jpeg";
    }

    state.wmImageEl = img;
    els.watermarkImagePreview.src = url;
    els.watermarkImagePreview.classList.remove("hidden");
    els.imageDropzoneEmpty.classList.add("hidden");
    els.imageRemoveBtn.classList.remove("hidden");
    schedulePreview();
  }

  function clearWmImage() {
    state.wmImageEl = null;
    state.wmImageBytes = null;
    state.wmImageType = null;
    els.watermarkImagePreview.removeAttribute("src");
    els.watermarkImagePreview.classList.add("hidden");
    els.imageDropzoneEmpty.classList.remove("hidden");
    els.imageRemoveBtn.classList.add("hidden");
    els.watermarkImageInput.value = "";
    schedulePreview();
  }

  /* ============================ 生成与下载 ============================ */
  function getPageIndices(total) {
    if (state.pagesMode === "all") {
      const arr = [];
      for (let i = 0; i < total; i++) arr.push(i);
      return arr;
    }
    const set = new Set();
    const parts = state.pagesRange.split(/[,，]/);
    for (let p of parts) {
      p = p.trim();
      if (!p) continue;
      if (p.includes("-")) {
        let seg = p.split("-");
        let a = Number(seg[0]), b = Number(seg[1]);
        if (isNaN(a) || isNaN(b)) continue;
        if (a > b) [a, b] = [b, a];
        for (let i = a; i <= b; i++) if (i >= 1 && i <= total) set.add(i - 1);
      } else {
        const n = Number(p);
        if (!isNaN(n) && n >= 1 && n <= total) set.add(n - 1);
      }
    }
    const arr = Array.from(set).sort((a, b) => a - b);
    return arr.length ? arr : null;
  }

  function drawTextPdf(page, font, text, O, opts) {
    const { size, angleDeg, angleRad, color, opacity, letterSpacing } = opts;
    const rotate = PDFLib.degrees(angleDeg);
    if (letterSpacing <= 0) {
      page.drawText(text, { x: O.x, y: O.y, size, font, color, opacity, rotate });
      return;
    }
    let offset = 0;
    const c = Math.cos(angleRad), s = Math.sin(angleRad);
    for (const ch of Array.from(text)) {
      page.drawText(ch, { x: O.x + c * offset, y: O.y + s * offset, size, font, color, opacity, rotate });
      offset += font.widthOfTextAtSize(ch, size) + letterSpacing;
    }
  }

  async function embedWatermarkImage(doc, cache) {
    if (cache.img) return cache.img;
    const bytes = state.wmImageBytes;
    if (!bytes) throw new Error("未选择水印图片");
    cache.img = state.wmImageType === "png" ? await doc.embedPng(bytes) : await doc.embedJpg(bytes);
    return cache.img;
  }

  async function drawItemOnPdf(doc, page, it, fontCache, imageCache) {
    const angleDeg = it.angle;
    const angleRad = (angleDeg * Math.PI) / 180;
    if (it.kind === "text") {
      const font = await getOutputFont(doc, fontCache, it.fontId);
      const color = hexToRgb01(it.color);
      const opacity = it.opacity / 100;
      const w = textWidthPts(font, it.text, it.size, it.letterSpacing);
      const h = it.size;
      const O = centerToOrigin(it.x, it.y, w, h, angleRad);
      const base = { size: it.size, angleDeg, angleRad, color, opacity, letterSpacing: it.letterSpacing };

      if (it.shadow) {
        const so = it.shadowOffset;
        drawTextPdf(page, font, it.text, { x: O.x + so, y: O.y - so }, { ...base, color: rgb01(0, 0, 0), opacity: opacity * 0.55 });
      }
      if (it.syntheticBold) {
        const b = it.size * 0.014;
        const c = Math.cos(angleRad), s = Math.sin(angleRad);
        for (const d of BOLD_DIRS) {
          const ox = d[0] * b * c - d[1] * b * s;
          const oy = d[0] * b * s + d[1] * b * c;
          drawTextPdf(page, font, it.text, { x: O.x + ox, y: O.y + oy }, base);
        }
      }
      if (it.outline) {
        const o = it.outlineWidth;
        const oc = hexToRgb01(it.outlineColor);
        for (const d of BOLD_DIRS) drawTextPdf(page, font, it.text, { x: O.x + d[0] * o, y: O.y + d[1] * o }, { ...base, color: oc });
      }
      drawTextPdf(page, font, it.text, O, base);
    } else {
      const img = await embedWatermarkImage(doc, imageCache);
      const O = centerToOrigin(it.x, it.y, it.w, it.h, angleRad);
      page.drawImage(img, {
        x: O.x,
        y: O.y,
        width: it.w,
        height: it.h,
        rotate: PDFLib.degrees(angleDeg),
        opacity: it.opacity / 100,
      });
    }
  }

  async function generatePDF() {
    if (state.type === "text") {
      const f = FONT_MAP[state.fontId];
      if (f.builtin && /[^\u0000-\u00FF]/.test(state.text)) {
        throw new Error("当前字体为拉丁字体，无法显示中文，请换用「思源黑体」等中文字体");
      }
    }

    const doc = await PDFLib.PDFDocument.load(state.fileBuffer, { ignoreEncryption: true });
    doc.registerFontkit(fontkit);
    const pages = doc.getPages();
    let indices = getPageIndices(pages.length);
    if (!indices) throw new Error("页码范围无效，请检查后重试");

    const fontCache = new Map();
    const imageCache = {};
    for (const idx of indices) {
      const page = pages[idx];
      const { width, height } = page.getSize();
      const items = await computeWatermarkItems(width, height);
      if (!items.length) continue;
      if (state.layer === "under") {
        await drawItemsUnderPage(doc, page, items, fontCache, imageCache);
      } else {
        for (const it of items) await drawItemOnPdf(doc, page, it, fontCache, imageCache);
      }
    }
    return await doc.save();
  }

  async function drawItemsUnderPage(doc, page, items, fontCache, imageCache) {
    const { width, height } = page.getSize();
    const tempPage = doc.addPage([width, height]);
    const tempIndex = doc.getPageCount() - 1;
    try {
      for (const it of items) await drawItemOnPdf(doc, tempPage, it, fontCache, imageCache);

      const tempEntries = tempPage.node.normalizedEntries();
      const contentStream = tempPage.getContentStream();
      contentStream.dict.set(PDFLib.PDFName.of("Type"), PDFLib.PDFName.of("XObject"));
      contentStream.dict.set(PDFLib.PDFName.of("Subtype"), PDFLib.PDFName.of("Form"));
      contentStream.dict.set(PDFLib.PDFName.of("BBox"), doc.context.obj([0, 0, width, height]));
      contentStream.dict.set(PDFLib.PDFName.of("Resources"), tempEntries.Resources);
      const formRef = doc.context.register(contentStream);

      const targetEntries = page.node.normalizedEntries();
      let key = "WM0";
      let n = 0;
      while (targetEntries.XObject.get(PDFLib.PDFName.of(key)) !== undefined) {
        n++;
        key = "WM" + n;
      }
      targetEntries.XObject.set(PDFLib.PDFName.of(key), formRef);

      const pre = doc.context.stream("q\n/" + key + " Do\nQ\n");
      const preRef = doc.context.register(pre);
      targetEntries.Contents.insert(0, preRef);
    } finally {
      doc.removePage(tempIndex);
    }
  }

  async function renderWatermarkedImage() {
    const canvas = document.createElement("canvas");
    canvas.width = state.imageEl.naturalWidth;
    canvas.height = state.imageEl.naturalHeight;
    const ctx = canvas.getContext("2d");
    ctx.drawImage(state.imageEl, 0, 0);
    const items = await computeWatermarkItems(state.pageW, state.pageH);
    drawItemsOnCanvas(ctx, items, 1, state.pageH);
    return await new Promise((res) => {
      if (state.imageType === "png") canvas.toBlob(res, "image/png");
      else if (state.imageType === "webp") canvas.toBlob(res, "image/webp", 0.9);
      else canvas.toBlob(res, "image/jpeg", 0.92);
    });
  }

  async function handleDownload() {
    if (!state.file) return;
    els.downloadBtn.disabled = true;
    setStatus("正在处理…");
    try {
      if (state.fileType === "pdf") {
        const bytes = await generatePDF();
        saveBlob(new Blob([bytes], { type: "application/pdf" }), replaceExt(state.file.name, "-水印.pdf"));
        toast("已生成水印 PDF，开始下载");
      } else {
        const blob = await renderWatermarkedImage();
        const ext = state.imageType === "png" ? ".png" : state.imageType === "webp" ? ".webp" : ".jpg";
        saveBlob(blob, replaceExt(state.file.name, "-水印" + ext));
        toast("已生成水印图片，开始下载");
      }
    } catch (e) {
      console.error(e);
      toast("处理失败：" + (e.message || e), true);
    } finally {
      els.downloadBtn.disabled = false;
      setStatus("就绪");
    }
  }

  /* ============================ 界面同步 ============================ */
  function syncTypeUI() {
    els.typeSeg.querySelectorAll(".seg-btn").forEach((b) => b.classList.toggle("active", b.dataset.type === state.type));
    els.textSection.classList.toggle("hidden", state.type !== "text");
    els.imageSection.classList.toggle("hidden", state.type !== "image");
  }

  function syncLayoutUI() {
    els.layoutSeg.querySelectorAll(".seg-btn").forEach((b) => b.classList.toggle("active", b.dataset.layout === state.layout));
    els.positionSection.classList.toggle("hidden", state.layout !== "single");
    els.tileSection.classList.toggle("hidden", state.layout !== "tile");
  }

  function syncPositionUI() {
    els.positionGrid.querySelectorAll("button").forEach((b) => b.classList.toggle("active", b.dataset.pos === state.position));
  }

  function syncPagesUI() {
    els.pagesSeg.querySelectorAll(".seg-btn").forEach((b) => b.classList.toggle("active", b.dataset.pages === state.pagesMode));
    els.pagesRangeWrap.classList.toggle("hidden", state.pagesMode !== "range");
  }

  function syncLayerUI() {
    els.layerSeg.querySelectorAll(".seg-btn").forEach((b) => b.classList.toggle("active", b.dataset.layer === state.layer));
    els.layerField.classList.toggle("hidden", state.fileType === "image");
  }

  function syncSwatches() {
    els.colorSwatches.querySelectorAll(".swatch").forEach((s) => s.classList.toggle("active", s.dataset.color.toLowerCase() === state.color.toLowerCase()));
  }

  function setRange(id, valueId, val, fmt) {
    const input = $(id);
    input.value = val;
    $(valueId).value = fmt ? fmt(val) : val;
    updateRangeFill(input);
  }

  function syncAllUI() {
    els.textInput.value = state.text;
    els.fontSelect.value = state.fontId;
    setRange("fontSizeRange", "fontSizeInput", state.fontSize);
    els.colorInput.value = state.color;
    els.colorHex.textContent = state.color;
    syncSwatches();
    setRange("opacityRange", "opacityInput", state.opacity);
    setRange("angleRange", "angleInput", state.angle);
    setRange("letterSpacingRange", "letterSpacingInput", state.letterSpacing);
    els.boldToggle.checked = state.bold;
    els.outlineToggle.checked = state.outline;
    els.shadowToggle.checked = state.shadow;
    setRange("imageScaleRange", "imageScaleInput", state.imageScale);
    setRange("marginRange", "marginInput", state.margin);
    setRange("offsetXRange", "offsetXInput", state.offsetX);
    setRange("offsetYRange", "offsetYInput", state.offsetY);
    setRange("tileGapXRange", "tileGapXInput", state.tileGapX);
    setRange("tileGapYRange", "tileGapYInput", state.tileGapY);
    els.pagesRange.value = state.pagesRange;
    syncTypeUI();
    syncLayoutUI();
    syncPositionUI();
    syncPagesUI();
    syncLayerUI();
  }

  function applyPreset(id) {
    const p = PRESETS[id];
    if (!p) return;
    Object.assign(state, p);
    els.presetChips.querySelectorAll(".chip").forEach((c) => c.classList.toggle("active", c.dataset.preset === id));
    syncAllUI();
    schedulePreview();
  }

  /* ============================ 事件绑定 ============================ */
  function bindRange(id, valueId, setter, fmt) {
    const input = $(id);
    const update = () => {
      $(valueId).value = fmt ? fmt(+input.value) : input.value;
      updateRangeFill(input);
      setter(+input.value);
      schedulePreview();
    };
    input.addEventListener("input", update);
  }

  function bindNumberInput(inputId, rangeId, setter) {
    const input = $(inputId);
    const range = $(rangeId);
    const apply = () => {
      let v = parseFloat(input.value);
      if (!isFinite(v)) v = parseFloat(range.value);
      const min = parseFloat(range.min);
      const max = parseFloat(range.max);
      const step = parseFloat(range.step) || 1;
      v = clamp(v, min, max);
      v = Math.round(v / step) * step;
      input.value = v;
      range.value = v;
      updateRangeFill(range);
      setter(v);
      schedulePreview();
    };
    input.addEventListener("change", apply);
    input.addEventListener("keydown", (e) => {
      if (e.key === "Enter") {
        e.preventDefault();
        apply();
        input.blur();
      }
    });
    input.addEventListener("focus", () => input.select());
  }

  function bindEvents() {
    // 文件
    els.dropzone.addEventListener("click", () => els.fileInput.click());
    els.fileInput.addEventListener("change", (e) => handleFiles(e.target.files));
    els.changeFileBtn.addEventListener("click", () => els.fileInput.click());
    ["dragenter", "dragover"].forEach((ev) => els.dropzone.addEventListener(ev, (e) => { e.preventDefault(); els.dropzone.classList.add("drag"); }));
    ["dragleave", "drop"].forEach((ev) => els.dropzone.addEventListener(ev, (e) => { e.preventDefault(); els.dropzone.classList.remove("drag"); }));
    els.dropzone.addEventListener("drop", (e) => handleFiles(e.dataTransfer.files));
    els.dropzone.addEventListener("keydown", (e) => { if (e.key === "Enter" || e.key === " ") { e.preventDefault(); els.fileInput.click(); } });

    // 类型
    els.typeSeg.querySelectorAll(".seg-btn").forEach((b) => b.addEventListener("click", () => {
      state.type = b.dataset.type;
      syncTypeUI();
      schedulePreview();
    }));

    // 预设
    els.presetChips.querySelectorAll(".chip").forEach((c) => c.addEventListener("click", () => applyPreset(c.dataset.preset)));

    // 文字
    els.textInput.addEventListener("input", () => { state.text = els.textInput.value; schedulePreview(); });
    els.fontSelect.addEventListener("change", () => { state.fontId = els.fontSelect.value; schedulePreview(); });
    bindRange("fontSizeRange", "fontSizeInput", (v) => { state.fontSize = v; });
    bindNumberInput("fontSizeInput", "fontSizeRange", (v) => { state.fontSize = v; });
    els.colorInput.addEventListener("input", () => { state.color = els.colorInput.value; els.colorHex.textContent = state.color; syncSwatches(); schedulePreview(); });
    bindRange("opacityRange", "opacityInput", (v) => { state.opacity = v; });
    bindNumberInput("opacityInput", "opacityRange", (v) => { state.opacity = v; });
    bindRange("angleRange", "angleInput", (v) => { state.angle = v; });
    bindNumberInput("angleInput", "angleRange", (v) => { state.angle = v; });
    bindRange("letterSpacingRange", "letterSpacingInput", (v) => { state.letterSpacing = v; });
    bindNumberInput("letterSpacingInput", "letterSpacingRange", (v) => { state.letterSpacing = v; });
    els.boldToggle.addEventListener("change", () => { state.bold = els.boldToggle.checked; schedulePreview(); });
    els.outlineToggle.addEventListener("change", () => { state.outline = els.outlineToggle.checked; schedulePreview(); });
    els.shadowToggle.addEventListener("change", () => { state.shadow = els.shadowToggle.checked; schedulePreview(); });

    // 图片水印
    els.imageDropzone.addEventListener("click", () => els.watermarkImageInput.click());
    els.watermarkImageInput.addEventListener("change", (e) => handleWmImage(e.target.files[0]));
    els.imageRemoveBtn.addEventListener("click", clearWmImage);
    bindRange("imageScaleRange", "imageScaleInput", (v) => { state.imageScale = v; });
    bindNumberInput("imageScaleInput", "imageScaleRange", (v) => { state.imageScale = v; });

    // 布局
    els.layoutSeg.querySelectorAll(".seg-btn").forEach((b) => b.addEventListener("click", () => {
      state.layout = b.dataset.layout;
      syncLayoutUI();
      schedulePreview();
    }));
    els.positionGrid.querySelectorAll("button").forEach((b) => b.addEventListener("click", () => {
      state.position = b.dataset.pos;
      syncPositionUI();
      schedulePreview();
    }));
    bindRange("marginRange", "marginInput", (v) => { state.margin = v; });
    bindNumberInput("marginInput", "marginRange", (v) => { state.margin = v; });
    bindRange("offsetXRange", "offsetXInput", (v) => { state.offsetX = v; });
    bindNumberInput("offsetXInput", "offsetXRange", (v) => { state.offsetX = v; });
    bindRange("offsetYRange", "offsetYInput", (v) => { state.offsetY = v; });
    bindNumberInput("offsetYInput", "offsetYRange", (v) => { state.offsetY = v; });
    bindRange("tileGapXRange", "tileGapXInput", (v) => { state.tileGapX = v; });
    bindNumberInput("tileGapXInput", "tileGapXRange", (v) => { state.tileGapX = v; });
    bindRange("tileGapYRange", "tileGapYInput", (v) => { state.tileGapY = v; });
    bindNumberInput("tileGapYInput", "tileGapYRange", (v) => { state.tileGapY = v; });

    // 层级
    els.layerSeg.querySelectorAll(".seg-btn").forEach((b) => b.addEventListener("click", () => {
      state.layer = b.dataset.layer;
      syncLayerUI();
      schedulePreview();
    }));

    // 页面
    els.pagesSeg.querySelectorAll(".seg-btn").forEach((b) => b.addEventListener("click", () => {
      state.pagesMode = b.dataset.pages;
      syncPagesUI();
    }));
    els.pagesRange.addEventListener("input", () => { state.pagesRange = els.pagesRange.value; });

    // 翻页
    els.prevPageBtn.addEventListener("click", () => goPage(state.currentPage - 1));
    els.nextPageBtn.addEventListener("click", () => goPage(state.currentPage + 1));

    // 下载
    els.downloadBtn.addEventListener("click", handleDownload);
  }

  function populateFontSelect() {
    els.fontSelect.innerHTML = "";
    Object.keys(FONT_MAP).forEach((id) => {
      if (FONT_MAP[id].hidden) return;
      const opt = document.createElement("option");
      opt.value = id;
      opt.textContent = FONT_MAP[id].name;
      els.fontSelect.appendChild(opt);
    });
    els.fontSelect.value = state.fontId;
  }

  function populateSwatches() {
    els.colorSwatches.innerHTML = "";
    SWATCH_COLORS.forEach((c) => {
      const b = document.createElement("button");
      b.type = "button";
      b.className = "swatch";
      b.dataset.color = c;
      b.style.background = c;
      b.title = c;
      b.addEventListener("click", () => {
        state.color = c;
        els.colorInput.value = c;
        els.colorHex.textContent = c;
        syncSwatches();
        schedulePreview();
      });
      els.colorSwatches.appendChild(b);
    });
  }

  function init() {
    pdfjsLib.GlobalWorkerOptions.workerSrc = PDF_WORKER_SRC;
    populateFontSelect();
    populateSwatches();
    syncAllUI();
    bindEvents();
    preloadFonts();
  }

  init();
})();
