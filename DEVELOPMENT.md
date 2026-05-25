# 办公文档转换工具 - 开发文档

## 1. 项目目标

本项目是一个**纯前端运行的办公文档转换工具**，所有文件处理在浏览器本地完成，无需上传到服务器，保护用户隐私。

### 支持的转换路径

**Office → PDF：**
- `.docx` → PDF
- `.xlsx`/`.xls`/`.csv` → PDF  
- `.pptx` → PDF

**PDF → Office：**
- PDF → `.docx`
- PDF → `.pptx`
- PDF → `.xlsx`

### 核心特点

- **隐私优先**：文件不上传，本地处理
- **轻量部署**：纯静态资源，直接托管即可
- **实时预览**：转换前可检查版面效果
- **质量可选**：提供轻量/标准/高清三档导出质量

---

## 2. 技术选型

| 依赖 | 版本 | 用途 |
|------|------|------|
| `mammoth` | ^9.7.0 | DOCX 解析为 HTML |
| `html2canvas` | ^1.4.1 | DOM 节点转 Canvas |
| `jsPDF` | ^3.0.1 | Canvas 生成 PDF |
| `DOMPurify` | ^3.2.5 | HTML 安全清理 |
| `xlsx` | ^0.18.5 | Excel 文件读写 |
| `JSZip` | ^3.10.1 | ZIP/PDF 解析 |
| `pdfjs-dist` | ^5.7.284 | PDF 渲染与文本提取 |
| `docx` | ^9.7.0 | DOCX 文档生成 |
| `pptxgenjs` | ^4.0.1 | PPTX 幻灯片生成 |
| `lucide` | ^0.468.0 | 图标组件 |

---

## 3. 项目结构

```
格式转换/
├── src/
│   ├── main.js          # 主应用逻辑（状态管理、文件处理、导出）
│   └── styles.css       # 完整样式（响应式布局、动画效果）
├── dist/                # Vite 构建产物
├── index.html           # 入口页面
├── package.json         # 依赖配置
├── DEVELOPMENT.md       # 开发文档
└── SKILL.md             # UI 设计指南
```

---

## 4. 核心转换流程

### 4.1 Office → PDF 通用流程

```
用户上传文件
    ↓
File API 读取 ArrayBuffer
    ↓
对应解析器转换为 HTML
    ↓
DOMPurify 安全清理
    ↓
渲染到 A4 预览容器
    ↓
等待图片/字体加载完成
    ↓
html2canvas 截取 Canvas
    ↓
jsPDF 按 A4 尺寸切割多页
    ↓
浏览器下载 PDF
```

### 4.2 DOCX → PDF

```js
// 使用 mammoth 解析 DOCX
const result = await mammoth.convertToHtml(
  { arrayBuffer },
  {
    convertImage: mammoth.images.imgElement(async (image) => {
      const base64 = await image.read("base64");
      return { src: `data:${image.contentType};base64,${base64}` };
    })
  }
);
```

### 4.3 Excel → PDF

```js
// 使用 xlsx 解析工作表
const workbook = XLSX.read(arrayBuffer, { type: "array" });
const table = XLSX.utils.sheet_to_html(worksheet, { editable: false });
```

### 4.4 PPTX → PDF

```js
// 使用 JSZip 解析 PPTX 内部结构
const zip = await JSZip.loadAsync(arrayBuffer);
const slideFiles = Object.keys(zip.files)
  .filter((name) => /^ppt\/slides\/slide\d+\.xml$/.test(name));
```

### 4.5 PDF → Office

```js
// 使用 pdfjs-dist 解析 PDF
const pdf = await pdfjsLib.getDocument({ data: arrayBuffer }).promise;
const page = await pdf.getPage(pageNumber);
const textContent = await page.getTextContent(); // 提取文本
```

---

## 5. 状态管理

应用使用简单的状态对象管理全局状态：

```js
const state = {
  status: "idle",           // idle/reading/converting/preview/exporting/success/error
  selectedFile: null,       // 当前选择的文件
  sourceKind: null,         // docx/sheet/pptx/pdf
  target: "pdf",            // 转换目标格式
  quality: "standard",      // 导出质量
  pageTone: "paper",        // 纸张底色
  outputName: "converted.pdf",
  pdfPages: [],             // PDF页面数据（图片+文本行）
  extractedLines: []        // 提取的文本内容
};
```

---

## 6. 导出质量预设

| 档位 | scale | imageQuality | 适用场景 |
|------|-------|--------------|----------|
| light | 1 | 0.82 | 快速分享，小体积 |
| standard | 1.5 | 0.9 | 日常办公，平衡质量 |
| high | 2 | 0.95 | 打印输出，高清需求 |

---

## 7. 页面布局

### 7.1 双栏布局

```
┌─────────────────────────────────────────────────────────────┐
│  左侧控制面板 (360px)           │  右侧预览区 (剩余空间)      │
│  ┌─────────────────────────┐   │                           │
│  │ 品牌标识                │   │  ┌─────────────────────┐  │
│  │ 隐私提示                │   │  │   工具栏            │  │
│  │ 文件上传区              │   │  └─────────────────────┘  │
│  │ 转换目标选择 (PDF时显示) │   │                           │
│  │ 导出设置                │   │  ┌─────────────────────┐  │
│  │ 能力说明                │   │  │   A4 预览画布       │  │
│  │ 状态面板                │   │  │   (794px × 1123px) │  │
│  └─────────────────────────┘   │  └─────────────────────┘  │
└─────────────────────────────────┴───────────────────────────┘
```

### 7.2 响应式适配

- **桌面端**：双栏并排显示
- **平板端**：控制面板在上，预览区在下
- **移动端**：紧凑布局，A4 预览自适应宽度

---

## 8. 核心组件说明

### 8.1 预览容器

```html
<article id="pdf-preview" class="pdf-page tone-paper">
  <!-- 解析后的 HTML 内容 -->
</article>
```

**样式特点**：
- 固定宽度：794px（96 DPI 下的 A4 宽度）
- 背景色可选：paper(白)/warm(暖白)/cool(冷白)
- 边框阴影模拟纸张效果

### 8.2 状态面板

包含进度条和步骤指示器：

```
读取 ──► 解析 ──► 预览 ──► 导出
  ●         ●         ●         ●
```

---

## 9. 安全与隐私

### 9.1 隐私保护

- 所有文件在浏览器本地处理，不上传服务器
- 不记录文件内容或转换历史
- 不调用第三方转换 API

### 9.2 HTML 安全

```js
// 使用 DOMPurify 清理用户上传文档中的 HTML
const cleanHtml = DOMPurify.sanitize(result.value, {
  USE_PROFILES: { html: true },
  ADD_ATTR: ["target"]
});
```

### 9.3 文件大小限制

```js
const MAX_FILE_SIZE = 35 * 1024 * 1024; // 35MB
```

---

## 10. 运行命令

```bash
# 开发模式（本地服务器）
npm run dev

# 构建生产版本
npm run build

# 预览生产版本
npm run preview
```

---

## 11. 测试清单

### 11.1 功能测试

- [ ] DOCX 上传后能生成预览
- [ ] Excel 表格显示正常
- [ ] PPTX 幻灯片内容提取
- [ ] PDF 页面渲染和文本提取
- [ ] 导出按钮功能正常
- [ ] 多页文档生成多页 PDF
- [ ] 错误文件能显示提示

### 11.2 样式测试

- [ ] A4 预览宽度稳定
- [ ] 中文字体显示正常
- [ ] 图片不溢出页面
- [ ] 表格自适应宽度
- [ ] 长链接自动换行

### 11.3 兼容性测试

- [ ] Chrome 转换成功
- [ ] Firefox 转换成功
- [ ] Safari 转换成功
- [ ] Edge 转换成功
- [ ] 移动端布局正常

---

## 12. 已知限制

### 12.1 Office → PDF 限制

- 复杂 Word 排版（页眉页脚、脚注、文本框）可能无法完全还原
- Excel 复杂公式和宏不会执行
- PPTX 仅提取文本结构，不等同于原始版式渲染

### 12.2 PDF → Office 限制

> **重要**：浏览器纯前端方案无法实现高保真可编辑还原

| 转换路径 | 实现方式 | 限制 |
|----------|----------|------|
| PDF → DOCX | 文本提取重建 | 仅保留纯文本和分页结构 |
| PDF → PPTX | 每页作为图片 | 视觉稳定但不可逐字编辑 |
| PDF → XLSX | 文本行整理为表格 | 无复杂表格结构 |

### 12.3 技术限制

- 大文件（>35MB）受浏览器内存限制
- 特殊字体依赖系统字体，可能被替换
- Canvas 截图式 PDF 文本不可复制搜索

---

## 13. 代码组织

### 13.1 main.js 结构

```
├── 依赖导入 (第1-28行)
├── 常量定义 (第30-51行)
├── 状态对象 (第34-45行)
├── DOM 模板 (第55-209行)
├── 图标注册 (第211-228行)
├── DOM 引用 (第230-245行)
├── 事件监听绑定 (第247-321行)
├── 文件处理主流程 (第323-364行)
│   └── handleFile()
├── 格式解析器 (第386-501行)
│   ├── renderDocx()
│   ├── renderSpreadsheet()
│   ├── renderPptx()
│   └── renderPdf()
├── PDF转Office导出 (第503-593行)
│   ├── exportPdfToDocx()
│   ├── exportPdfToPptx()
│   └── exportPdfToXlsx()
├── PDF导出工具 (第620-664行)
│   ├── renderElementToCanvas()
│   └── canvasToA4Pdf()
└── 辅助函数 (第666-884行)
    ├── 文件验证、状态更新、UI操作等
```

---

## 14. 后续优化建议

### 14.1 第一阶段（已完成）

- DOCX/XLSX/PPTX/PDF 基础解析
- HTML 预览和 PDF 导出
- 响应式布局
- 基础错误处理

### 14.2 第二阶段（建议）

- 按 DOM 块级元素智能分页
- 支持更多图片格式优化
- 添加导出进度详细反馈
- 支持自定义页面边距

### 14.3 服务端扩展

如需高保真 PDF→Office 转换，建议增加服务端转换层，使用：
- LibreOffice/OpenOffice 转换引擎
- 商业转换 API（如 Adobe ExportPDF）
- 专业文档处理 SDK