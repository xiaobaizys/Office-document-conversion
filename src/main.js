import "./styles.css";
import mammoth from "mammoth/mammoth.browser";
import html2canvas from "html2canvas";
import { jsPDF } from "jspdf";
import DOMPurify from "dompurify";
import * as XLSX from "xlsx";
import JSZip from "jszip";
import * as pdfjsLib from "pdfjs-dist";
import pdfWorkerUrl from "pdfjs-dist/build/pdf.worker.min.mjs?url";
import { Document, Packer, Paragraph, TextRun } from "docx";
import PptxGenJS from "pptxgenjs";
import {
  AlertCircle,
  Check,
  Download,
  FileSpreadsheet,
  FileText,
  Gauge,
  Image,
  Info,
  Loader2,
  Lock,
  Presentation,
  Settings,
  UploadCloud,
  X,
  createIcons
} from "lucide";

pdfjsLib.GlobalWorkerOptions.workerSrc = pdfWorkerUrl;

const MAX_FILE_SIZE = 35 * 1024 * 1024;

const state = {
  status: "idle",
  selectedFile: null,
  sourceKind: null,
  target: "pdf",
  messages: [],
  quality: "standard",
  pageTone: "paper",
  outputName: "converted.pdf",
  pdfPages: [],
  extractedLines: []
};

const qualityPresets = {
  light: { scale: 1, imageQuality: 0.82 },
  standard: { scale: 1.5, imageQuality: 0.9 },
  high: { scale: 2, imageQuality: 0.95 }
};

const app = document.querySelector("#app");

app.innerHTML = `
  <main class="app-shell">
    <aside class="control-panel" aria-label="转换控制面板">
      <header class="brand-block">
        <div class="brand-mark" aria-hidden="true">
          <i data-lucide="file-text"></i>
        </div>
        <div>
          <p class="eyebrow">Local document converter</p>
          <h1>办公文档转换</h1>
        </div>
      </header>

      <section class="privacy-strip">
        <i data-lucide="lock"></i>
        <span>文件在浏览器本地处理。高保真 PDF 反转 Office 需要专业转换引擎，当前提供可用的前端重建方案。</span>
      </section>

      <section class="upload-card">
        <input
          id="file-input"
          class="sr-only"
          type="file"
          accept=".docx,.xlsx,.xls,.csv,.pptx,.pdf,application/pdf,application/vnd.openxmlformats-officedocument.wordprocessingml.document,application/vnd.openxmlformats-officedocument.presentationml.presentation,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        />
        <label id="drop-zone" class="drop-zone" for="file-input">
          <span class="drop-icon"><i data-lucide="upload-cloud"></i></span>
          <strong>选择或拖入文件</strong>
          <small>支持 DOCX、XLSX、XLS、CSV、PPTX、PDF</small>
        </label>
        <div id="file-meta" class="file-meta is-empty">
          <i data-lucide="info"></i>
          <span>尚未选择文件</span>
        </div>
      </section>

      <section id="target-panel" class="target-panel is-hidden">
        <div class="section-heading">
          <i data-lucide="settings"></i>
          <h2>转换目标</h2>
        </div>
        <div class="target-grid" role="radiogroup" aria-label="PDF 转换目标">
          <button class="target-option is-active" type="button" data-target="docx">
            <i data-lucide="file-text"></i>
            <span>Word</span>
          </button>
          <button class="target-option" type="button" data-target="pptx">
            <i data-lucide="presentation"></i>
            <span>PPT</span>
          </button>
          <button class="target-option" type="button" data-target="xlsx">
            <i data-lucide="file-spreadsheet"></i>
            <span>Excel</span>
          </button>
        </div>
        <p id="target-note" class="target-note"></p>
      </section>

      <section class="settings-panel">
        <div class="section-heading">
          <i data-lucide="settings"></i>
          <h2>PDF 导出设置</h2>
        </div>

        <div class="field-group">
          <span class="field-label">质量</span>
          <div class="segmented" role="radiogroup" aria-label="导出质量">
            <button class="quality-option" type="button" data-quality="light">轻量</button>
            <button class="quality-option is-active" type="button" data-quality="standard">标准</button>
            <button class="quality-option" type="button" data-quality="high">高清</button>
          </div>
        </div>

        <div class="field-group">
          <span class="field-label">纸张底色</span>
          <div class="tone-grid" role="radiogroup" aria-label="纸张底色">
            <button class="tone-option is-active" type="button" data-tone="paper" aria-label="白色纸张"><span class="swatch swatch-paper"></span></button>
            <button class="tone-option" type="button" data-tone="warm" aria-label="暖色纸张"><span class="swatch swatch-warm"></span></button>
            <button class="tone-option" type="button" data-tone="cool" aria-label="冷色纸张"><span class="swatch swatch-cool"></span></button>
          </div>
        </div>

        <button id="export-button" class="primary-button" type="button" disabled>
          <i data-lucide="download"></i>
          <span>导出</span>
        </button>
      </section>

      <section class="capability-panel">
        <div class="section-heading">
          <i data-lucide="info"></i>
          <h2>能力说明</h2>
        </div>
        <div class="capability-list">
          <span>Office -> PDF：本地预览后导出 A4 PDF。</span>
          <span>PDF -> Word：提取文本重建 DOCX。</span>
          <span>PDF -> PPT：按页生成图片幻灯片，视觉保真优先。</span>
          <span>PDF -> Excel：按页和文本行整理为表格。</span>
        </div>
      </section>

      <section class="status-panel">
        <div class="section-heading">
          <i data-lucide="gauge"></i>
          <h2>状态</h2>
        </div>
        <div class="progress-rail" aria-hidden="true">
          <span id="progress-bar" class="progress-bar"></span>
        </div>
        <div id="status-line" class="status-line">等待上传文档</div>
        <ol class="step-list" aria-label="转换步骤">
          <li class="step-item" data-step="reading"><span></span>读取</li>
          <li class="step-item" data-step="converting"><span></span>解析</li>
          <li class="step-item" data-step="preview"><span></span>预览</li>
          <li class="step-item" data-step="exporting"><span></span>导出</li>
        </ol>
        <div id="message-list" class="message-list" aria-live="polite"></div>
      </section>
    </aside>

    <section class="workspace" aria-label="文档预览区">
      <div class="workspace-topbar">
        <div>
          <p id="preview-kicker" class="eyebrow">A4 preview</p>
          <h2 id="preview-title">预览画布</h2>
        </div>
        <div id="status-pill" class="status-pill">
          <i data-lucide="info"></i>
          <span>Idle</span>
        </div>
      </div>

      <div class="preview-scroll">
        <div class="preview-stage">
          <article id="pdf-preview" class="pdf-page tone-paper">
            <div class="empty-preview">
              <div class="empty-icon">
                <i data-lucide="image"></i>
              </div>
              <h2>上传文件后在这里检查版面</h2>
              <p>Office 文件会渲染为可导出的 PDF 预览；PDF 文件会生成页面预览，并可导出为 Word、PPT 或 Excel。</p>
            </div>
          </article>
        </div>
        <div id="work-overlay" class="work-overlay" aria-hidden="true">
          <div class="work-card">
            <span class="work-spinner"></span>
            <strong id="work-title">正在处理</strong>
            <small id="work-detail">请稍候</small>
          </div>
        </div>
      </div>
    </section>
  </main>
`;

createIcons({
  icons: {
    AlertCircle,
    Check,
    Download,
    FileSpreadsheet,
    FileText,
    Gauge,
    Image,
    Info,
    Loader2,
    Lock,
    Presentation,
    Settings,
    UploadCloud,
    X
  }
});

const fileInput = document.querySelector("#file-input");
const dropZone = document.querySelector("#drop-zone");
const fileMeta = document.querySelector("#file-meta");
const preview = document.querySelector("#pdf-preview");
const exportButton = document.querySelector("#export-button");
const statusLine = document.querySelector("#status-line");
const statusPill = document.querySelector("#status-pill");
const messageList = document.querySelector("#message-list");
const previewTitle = document.querySelector("#preview-title");
const previewKicker = document.querySelector("#preview-kicker");
const progressBar = document.querySelector("#progress-bar");
const workOverlay = document.querySelector("#work-overlay");
const workTitle = document.querySelector("#work-title");
const workDetail = document.querySelector("#work-detail");
const targetPanel = document.querySelector("#target-panel");
const targetNote = document.querySelector("#target-note");

fileInput.addEventListener("change", (event) => {
  const [file] = event.target.files;
  if (file) handleFile(file);
});

dropZone.addEventListener("dragover", (event) => {
  event.preventDefault();
  dropZone.classList.add("is-dragging");
});

dropZone.addEventListener("dragleave", () => {
  dropZone.classList.remove("is-dragging");
});

dropZone.addEventListener("drop", (event) => {
  event.preventDefault();
  dropZone.classList.remove("is-dragging");
  const [file] = event.dataTransfer.files;
  if (file) handleFile(file);
});

fileMeta.addEventListener("click", (event) => {
  const clearButton = event.target.closest("#clear-file-button");
  if (!clearButton) return;
  clearSelectedFile();
});

document.querySelectorAll(".quality-option").forEach((button) => {
  button.addEventListener("click", () => {
    state.quality = button.dataset.quality;
    document.querySelectorAll(".quality-option").forEach((item) => item.classList.remove("is-active"));
    button.classList.add("is-active");
  });
});

document.querySelectorAll(".tone-option").forEach((button) => {
  button.addEventListener("click", () => {
    state.pageTone = button.dataset.tone;
    document.querySelectorAll(".tone-option").forEach((item) => item.classList.remove("is-active"));
    button.classList.add("is-active");
    preview.classList.remove("tone-paper", "tone-warm", "tone-cool");
    preview.classList.add(`tone-${state.pageTone}`);
  });
});

document.querySelectorAll(".target-option").forEach((button) => {
  button.addEventListener("click", () => {
    state.target = button.dataset.target;
    document.querySelectorAll(".target-option").forEach((item) => item.classList.remove("is-active"));
    button.classList.add("is-active");
    updateTargetNote();
    updateExportButton();
  });
});

exportButton.addEventListener("click", async () => {
  try {
    setStatus("exporting", "正在导出文件，请稍候");
    exportButton.disabled = true;

    if (state.sourceKind === "pdf") {
      await exportPdfToOffice();
    } else {
      await waitForPreviewReady(preview);
      const canvas = await renderElementToCanvas(preview);
      await canvasToA4Pdf(canvas, state.outputName);
    }

    setStatus("success", "文件已生成并开始下载");
  } catch (error) {
    showError(error);
  } finally {
    exportButton.disabled = state.status !== "preview" && state.status !== "success";
  }
});

async function handleFile(file) {
  try {
    clearMessages();
    resetDocumentState();
    validateFile(file);
    renderFileMeta(file);

    setStatus("reading", "正在读取文件");
    const arrayBuffer = await file.arrayBuffer();
    const kind = getFileKind(file);

    state.selectedFile = file;
    state.sourceKind = kind;
    state.outputName = getOutputName(file, kind === "pdf" ? state.target : "pdf");
    previewTitle.textContent = file.name;

    setStatus("converting", "正在解析文件内容");

    if (kind === "docx") {
      await renderDocx(arrayBuffer);
    } else if (kind === "sheet") {
      await renderSpreadsheet(arrayBuffer, file.name);
    } else if (kind === "pptx") {
      await renderPptx(arrayBuffer);
    } else if (kind === "pdf") {
      await renderPdf(arrayBuffer);
    }

    targetPanel.classList.toggle("is-hidden", kind !== "pdf");
    previewKicker.textContent = kind === "pdf" ? "PDF reverse export" : "A4 PDF preview";
    updateTargetNote();
    updateExportButton();
    setStatus("preview", "预览已就绪，可以导出");

    preview.animate(
      [{ transform: "translateY(10px)", opacity: 0.82 }, { transform: "translateY(0)", opacity: 1 }],
      { duration: 220, easing: "cubic-bezier(0.23, 1, 0.32, 1)" }
    );
  } catch (error) {
    showError(error);
  }
}

function validateFile(file) {
  const kind = getFileKind(file);
  if (!kind) {
    throw new Error("暂不支持该文件类型。请选择 DOCX、XLSX、XLS、CSV、PPTX 或 PDF。");
  }

  if (file.size > MAX_FILE_SIZE) {
    throw new Error("文件超过 35MB，建议压缩图片或拆分文档后再转换。");
  }
}

function getFileKind(file) {
  const name = file.name.toLowerCase();
  if (name.endsWith(".docx")) return "docx";
  if (name.endsWith(".xlsx") || name.endsWith(".xls") || name.endsWith(".csv")) return "sheet";
  if (name.endsWith(".pptx")) return "pptx";
  if (name.endsWith(".pdf")) return "pdf";
  return null;
}

async function renderDocx(arrayBuffer) {
  const result = await mammoth.convertToHtml(
    { arrayBuffer },
    {
      convertImage: mammoth.images.imgElement(async (image) => {
        const base64 = await image.read("base64");
        return { src: `data:${image.contentType};base64,${base64}` };
      })
    }
  );

  const cleanHtml = DOMPurify.sanitize(result.value || "<p>文档没有可显示的正文内容。</p>", {
    USE_PROFILES: { html: true },
    ADD_ATTR: ["target"]
  });

  setPreviewHtml(cleanHtml);
  renderMessages(result.messages || []);
}

async function renderSpreadsheet(arrayBuffer, fileName) {
  const workbook = XLSX.read(arrayBuffer, { type: "array" });
  const sheetNames = workbook.SheetNames.slice(0, 8);

  const html = sheetNames
    .map((sheetName) => {
      const worksheet = workbook.Sheets[sheetName];
      const table = XLSX.utils.sheet_to_html(worksheet, { editable: false });
      return `<section class="sheet-block"><h2>${escapeHtml(sheetName)}</h2>${table}</section>`;
    })
    .join("");

  setPreviewHtml(html || `<p>${escapeHtml(fileName)} 没有可读取的表格内容。</p>`, "sheet-preview");
  renderMessages([{ message: "Excel/CSV 会按工作表渲染为 PDF 预览，复杂公式和宏不会执行。" }]);
}

async function renderPptx(arrayBuffer) {
  const zip = await JSZip.loadAsync(arrayBuffer);
  const slideFiles = Object.keys(zip.files)
    .filter((name) => /^ppt\/slides\/slide\d+\.xml$/.test(name))
    .sort((a, b) => Number(a.match(/slide(\d+)/)?.[1] || 0) - Number(b.match(/slide(\d+)/)?.[1] || 0));

  if (!slideFiles.length) {
    throw new Error("未能在 PPTX 中读取到幻灯片内容。");
  }

  const slides = await Promise.all(
    slideFiles.map(async (path, index) => {
      const xml = await zip.files[path].async("text");
      const doc = new DOMParser().parseFromString(xml, "application/xml");
      const texts = Array.from(doc.querySelectorAll("a\\:t, t"))
        .map((node) => node.textContent.trim())
        .filter(Boolean);

      return `
        <section class="slide-page">
          <div class="slide-number">Slide ${index + 1}</div>
          ${
            texts.length
              ? texts.map((text, textIndex) => `<p class="${textIndex === 0 ? "slide-title" : ""}">${escapeHtml(text)}</p>`).join("")
              : "<p>该页没有可提取的文本内容。</p>"
          }
        </section>
      `;
    })
  );

  setPreviewHtml(slides.join(""), "deck-preview");
  renderMessages([{ message: "PPTX 前端解析以文本结构重建为主，不等同于 PowerPoint 原始版式渲染。" }]);
}

async function renderPdf(arrayBuffer) {
  const pdf = await pdfjsLib.getDocument({ data: arrayBuffer.slice(0) }).promise;
  state.pdfPages = [];
  state.extractedLines = [];

  for (let pageNumber = 1; pageNumber <= pdf.numPages; pageNumber += 1) {
    const page = await pdf.getPage(pageNumber);
    const viewport = page.getViewport({ scale: 1.45 });
    const canvas = document.createElement("canvas");
    canvas.width = Math.floor(viewport.width);
    canvas.height = Math.floor(viewport.height);

    const context = canvas.getContext("2d");
    await page.render({ canvasContext: context, viewport }).promise;

    const textContent = await page.getTextContent();
    const lines = groupPdfTextLines(textContent.items || []);
    state.extractedLines.push(...lines.map((line) => ({ page: pageNumber, text: line })));
    state.pdfPages.push({
      page: pageNumber,
      dataUrl: canvas.toDataURL("image/jpeg", 0.92),
      width: canvas.width,
      height: canvas.height,
      lines
    });
  }

  const html = state.pdfPages
    .map(
      (page) => `
        <section class="pdf-source-page">
          <img src="${page.dataUrl}" alt="PDF 第 ${page.page} 页" />
        </section>
      `
    )
    .join("");

  setPreviewHtml(html, "pdf-source-preview");
  renderMessages([
    {
      message:
        "PDF 转 Word/Excel 会基于文本提取重建结构；PDF 转 PPT 会把每页作为图片放入幻灯片，视觉更稳定但不可逐字编辑。"
    }
  ]);
}

function groupPdfTextLines(items) {
  const rows = new Map();

  items.forEach((item) => {
    const text = String(item.str || "").trim();
    if (!text) return;

    const y = Math.round(item.transform?.[5] || 0);
    const bucket = Math.round(y / 5) * 5;
    const row = rows.get(bucket) || [];
    row.push({ x: item.transform?.[4] || 0, text });
    rows.set(bucket, row);
  });

  return Array.from(rows.entries())
    .sort((a, b) => b[0] - a[0])
    .map(([, row]) =>
      row
        .sort((a, b) => a.x - b.x)
        .map((item) => item.text)
        .join(" ")
        .replace(/\s+/g, " ")
        .trim()
    )
    .filter(Boolean);
}

async function exportPdfToOffice() {
  if (!state.pdfPages.length) {
    throw new Error("请先上传并解析 PDF 文件。");
  }

  if (state.target === "docx") {
    await exportPdfToDocx();
  } else if (state.target === "pptx") {
    await exportPdfToPptx();
  } else if (state.target === "xlsx") {
    exportPdfToXlsx();
  }
}

async function exportPdfToDocx() {
  const children = [];

  state.pdfPages.forEach((page) => {
    children.push(new Paragraph({ children: [new TextRun({ text: `第 ${page.page} 页`, bold: true })] }));
    page.lines.forEach((line) => {
      children.push(new Paragraph({ children: [new TextRun(line)] }));
    });
    children.push(new Paragraph(""));
  });

  const doc = new Document({
    sections: [{ properties: {}, children }]
  });

  const blob = await Packer.toBlob(doc);
  downloadBlob(blob, getOutputName(state.selectedFile, "docx"));
}

async function exportPdfToPptx() {
  const pptx = new PptxGenJS();
  pptx.layout = "LAYOUT_WIDE";
  pptx.author = "Local document converter";
  pptx.subject = "PDF pages exported as slide images";
  pptx.title = state.selectedFile?.name || "PDF export";

  state.pdfPages.forEach((page) => {
    const slide = pptx.addSlide();
    slide.background = { color: "FFFFFF" };
    slide.addImage({ data: page.dataUrl, x: 0, y: 0, w: 13.333, h: 7.5, sizing: { type: "contain", x: 0, y: 0, w: 13.333, h: 7.5 } });
  });

  await pptx.writeFile({ fileName: getOutputName(state.selectedFile, "pptx") });
}

function exportPdfToXlsx() {
  const rows = [["页码", "行号", "文本"]];

  state.pdfPages.forEach((page) => {
    page.lines.forEach((line, index) => {
      rows.push([page.page, index + 1, line]);
    });
  });

  const workbook = XLSX.utils.book_new();
  const worksheet = XLSX.utils.aoa_to_sheet(rows);
  worksheet["!cols"] = [{ wch: 8 }, { wch: 8 }, { wch: 90 }];
  XLSX.utils.book_append_sheet(workbook, worksheet, "PDF文本");
  XLSX.writeFile(workbook, getOutputName(state.selectedFile, "xlsx"));
}

function setPreviewHtml(html, modeClass = "") {
  preview.className = `pdf-page tone-${state.pageTone} ${modeClass}`.trim();
  preview.innerHTML = html;
}

async function waitForPreviewReady(element) {
  const images = Array.from(element.querySelectorAll("img"));

  await Promise.all(
    images.map((img) => {
      if (img.complete) return Promise.resolve();
      return new Promise((resolve) => {
        img.onload = resolve;
        img.onerror = resolve;
      });
    })
  );

  if (document.fonts?.ready) {
    await document.fonts.ready;
  }

  await new Promise((resolve) => requestAnimationFrame(resolve));
}

async function renderElementToCanvas(element) {
  const preset = qualityPresets[state.quality];

  return await html2canvas(element, {
    scale: preset.scale,
    useCORS: true,
    allowTaint: false,
    backgroundColor: getComputedStyle(element).backgroundColor || "#ffffff",
    logging: false
  });
}

async function canvasToA4Pdf(canvas, fileName) {
  const preset = qualityPresets[state.quality];
  const pdf = new jsPDF("p", "mm", "a4");
  const pdfWidth = 210;
  const pdfHeight = 297;
  const pageCanvasHeight = Math.floor((canvas.width * pdfHeight) / pdfWidth);

  let renderedHeight = 0;
  let pageIndex = 0;

  while (renderedHeight < canvas.height) {
    const sliceHeight = Math.min(pageCanvasHeight, canvas.height - renderedHeight);
    const pageCanvas = document.createElement("canvas");
    pageCanvas.width = canvas.width;
    pageCanvas.height = sliceHeight;

    const context = pageCanvas.getContext("2d");
    context.fillStyle = "#ffffff";
    context.fillRect(0, 0, pageCanvas.width, pageCanvas.height);
    context.drawImage(canvas, 0, renderedHeight, canvas.width, sliceHeight, 0, 0, canvas.width, sliceHeight);

    const imageData = pageCanvas.toDataURL("image/jpeg", preset.imageQuality);
    const imageHeight = (sliceHeight * pdfWidth) / canvas.width;

    if (pageIndex > 0) pdf.addPage();
    pdf.addImage(imageData, "JPEG", 0, 0, pdfWidth, imageHeight);

    renderedHeight += sliceHeight;
    pageIndex += 1;
  }

  pdf.save(fileName);
}

function renderFileMeta(file) {
  fileMeta.classList.remove("is-empty");
  const kind = getFileKind(file);
  const kindLabel = {
    docx: "Word",
    sheet: "Excel",
    pptx: "PPT",
    pdf: "PDF"
  }[kind] || "文件";

  fileMeta.innerHTML = `
    <i data-lucide="file-text"></i>
    <span class="file-meta-text">
      <strong>${escapeHtml(file.name)}</strong>
      <small>${kindLabel} · ${formatFileSize(file.size)} · 已识别</small>
    </span>
    <button id="clear-file-button" class="clear-file-button" type="button" aria-label="删除当前文件">
      <i data-lucide="x"></i>
    </button>
  `;
  createIcons({ icons: { FileText, X } });
}

function renderMessages(messages) {
  const normalized = messages.filter(Boolean);

  if (!normalized.length) {
    messageList.innerHTML = `
      <div class="message-item is-ok">
        <i data-lucide="check"></i>
        <span>解析完成，未发现明显警告。</span>
      </div>
    `;
    createIcons({ icons: { Check } });
    return;
  }

  messageList.innerHTML = normalized
    .slice(0, 6)
    .map(
      (message) => `
        <div class="message-item">
          <i data-lucide="alert-circle"></i>
          <span>${escapeHtml(message.message || String(message))}</span>
        </div>
      `
    )
    .join("");
  createIcons({ icons: { AlertCircle } });
}

function clearMessages() {
  state.messages = [];
  messageList.innerHTML = "";
}

function resetDocumentState() {
  state.selectedFile = null;
  state.sourceKind = null;
  state.pdfPages = [];
  state.extractedLines = [];
  targetPanel.classList.add("is-hidden");
}

function clearSelectedFile() {
  fileInput.value = "";
  resetDocumentState();
  clearMessages();
  state.outputName = "converted.pdf";
  previewTitle.textContent = "预览画布";
  previewKicker.textContent = "A4 preview";
  preview.className = `pdf-page tone-${state.pageTone}`;
  preview.innerHTML = `
    <div class="empty-preview">
      <div class="empty-icon">
        <i data-lucide="image"></i>
      </div>
      <h2>上传文件后在这里检查版面</h2>
      <p>Office 文件会渲染为可导出的 PDF 预览；PDF 文件会生成页面预览，并可导出为 Word、PPT 或 Excel。</p>
    </div>
  `;
  fileMeta.className = "file-meta is-empty";
  fileMeta.innerHTML = `
    <i data-lucide="info"></i>
    <span>尚未选择文件</span>
  `;
  updateExportButton();
  setStatus("idle", "等待上传文档");
  createIcons({ icons: { Image, Info } });
}

function setStatus(status, text) {
  state.status = status;
  statusLine.textContent = text;
  statusPill.className = `status-pill status-${status}`;

  const icon = status === "success" || status === "preview" ? "check" : status === "error" ? "x" : status === "idle" ? "info" : "loader-2";
  const label = {
    idle: "Idle",
    reading: "Reading",
    converting: "Converting",
    preview: "Ready",
    exporting: "Exporting",
    success: "Done",
    error: "Error"
  }[status];

  statusPill.innerHTML = `<i data-lucide="${icon}"></i><span>${label}</span>`;
  createIcons({ icons: { Check, Info, Loader2, X } });
  exportButton.disabled = status !== "preview" && status !== "success";
  updateProgress(status);
  updateWorkOverlay(status, text);
}

function updateProgress(status) {
  const progress = {
    idle: 0,
    reading: 18,
    converting: 48,
    preview: 76,
    exporting: 92,
    success: 100,
    error: 100
  }[status];

  progressBar.style.width = `${progress}%`;
  progressBar.classList.toggle("is-error", status === "error");
  progressBar.classList.toggle("is-complete", status === "success");

  const stepOrder = ["reading", "converting", "preview", "exporting"];
  const activeIndex = stepOrder.indexOf(status);
  document.querySelectorAll(".step-item").forEach((item) => {
    const itemIndex = stepOrder.indexOf(item.dataset.step);
    item.classList.toggle("is-active", item.dataset.step === status);
    item.classList.toggle("is-done", status === "success" || (activeIndex > -1 && itemIndex < activeIndex) || (status === "preview" && itemIndex <= 2));
    item.classList.toggle("is-error", status === "error");
  });
}

function updateWorkOverlay(status, text) {
  const visible = status === "reading" || status === "converting" || status === "exporting";
  workOverlay.classList.toggle("is-visible", visible);
  preview.classList.toggle("is-processing", status === "reading" || status === "converting");

  if (!visible) return;
  workTitle.textContent = status === "exporting" ? "正在导出文件" : "正在转换文档";
  workDetail.textContent = text || "请稍候";
}

function updateTargetNote() {
  if (state.sourceKind !== "pdf") {
    targetNote.textContent = "";
    return;
  }

  const notes = {
    docx: "导出 Word 时会提取 PDF 文本并按页重建段落，适合二次编辑基础内容。",
    pptx: "导出 PPT 时会把 PDF 每页作为幻灯片图片，视觉还原优先，但页面内容不可逐字编辑。",
    xlsx: "导出 Excel 时会按页和文本行整理为表格，适合提取清单类文本。"
  };

  targetNote.textContent = notes[state.target];
}

function updateExportButton() {
  if (!state.selectedFile) {
    exportButton.disabled = true;
    exportButton.querySelector("span").textContent = "导出";
    return;
  }

  const target = state.sourceKind === "pdf" ? state.target.toUpperCase() : "PDF";
  exportButton.querySelector("span").textContent = `导出 ${target}`;
  state.outputName = getOutputName(state.selectedFile, state.sourceKind === "pdf" ? state.target : "pdf");
  exportButton.disabled = state.status !== "preview" && state.status !== "success";
}

function showError(error) {
  const message = error instanceof Error ? error.message : String(error);
  setStatus("error", message);
  messageList.innerHTML = `
    <div class="message-item is-error">
      <i data-lucide="alert-circle"></i>
      <span>${escapeHtml(message)}</span>
    </div>
  `;
  createIcons({ icons: { AlertCircle } });
}

function getOutputName(file, extension) {
  const base = file?.name ? file.name.replace(/\.[^.]+$/i, "") : "converted";
  return `${base}.${extension}`;
}

function downloadBlob(blob, fileName) {
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = fileName;
  document.body.appendChild(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(url);
}

function formatFileSize(size) {
  if (size < 1024) return `${size} B`;
  if (size < 1024 * 1024) return `${(size / 1024).toFixed(1)} KB`;
  return `${(size / 1024 / 1024).toFixed(1)} MB`;
}

function escapeHtml(value) {
  return String(value)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}
