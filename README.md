# PDF / DOCX reader

DeepSeek Harness 的本地文档阅读插件。为 agent 提供两个工具：

- `pdf_docx_read`：把 PDF / DOCX 读成**带行号、页码/标题定位的 Markdown 窗口**，支持续读、目录、表格聚焦。
- `pdf_docx_search`：在已解析文本中做字面搜索，返回**行号、页码/标题路径和上下文**，可直接用返回的行号继续读。

不联网、不调用 MinerU、不依赖 Python/LibreOffice、不做 OCR。

## 安装

```sh
dsh plugin --profile desktop add github:ToCyrene/dsh-pdf-docx-reader
```

本地开发时，也可以用 `link:<REPO_PATH>` 安装当前工作区，把 `<REPO_PATH>` 替换为本仓库的绝对路径，或在 Web 的 **Plugins → Add plugin** 中选择本目录。安装后在 Plugins 页面可见；模型工具在下一步或新会话中出现。

## 使用

```js
pdf_docx_read({ file_path: 'report.docx' })
pdf_docx_read({ file_path: 'report.pdf', pages: '3-5' })
pdf_docx_read({ file_path: 'report.docx', focus: 'outline' })
pdf_docx_read({ file_path: 'report.pdf', focus: 'tables' })
pdf_docx_read({ file_path: 'report.docx', offset: 301, limit: 300 })
pdf_docx_search({ file_path: 'report.pdf', query: '指标3.3', context: 200 })
```

### `pdf_docx_read`

| 参数 | 默认 | 说明 |
|---|---|---|
| `file_path` | 必填 | `.pdf` / `.docx` / `.docm`；相对路径按会话工作目录解析 |
| `offset` | 1 | 1-based 内容行号；`focus=outline` 时忽略 |
| `limit` | 300 | 最多 1000 行；另有 48 KB 字符上限 |
| `pages` | 全部 | 仅 PDF；支持 `3`、`3-5`、`1,3-5`、`第3-5页`、`3~5` |
| `focus` | `content` | `content` / `outline` / `tables` |

输出带页/行锚点、续读脚注和明确截断提示。PDF 每页有 `[pN]` 标记；DOCX 明确标注“无固定页码”。

### `pdf_docx_search`

| 参数 | 默认 | 说明 |
|---|---|---|
| `query` | 必填 | 1–200 字符字面文本（v1 不支持正则） |
| `context` | 160 | 匹配前后各取多少字符，最大 1000 |
| `max_hits` | 10 | 最多显示多少条，最大 50 |
| `case_sensitive` | false | 仅影响 ASCII 大小写 |
| `pages` | 全部 | 仅 PDF |

命中格式：`L301 p3 [标题路径] …上下文⟦匹配⟧上下文…`。

## 设计要点

- **虚拟 Markdown 文档 + 行号**：搜索返回的 `L` 行号就是 `offset`，不需要重新猜位置。
- **DOCX 不是按样式硬转 Markdown**：真实文档经常全用 `Normal` 样式，靠自动编号和加粗表达标题。插件解析 `numbering.xml`、`styles.xml`、`outlineLvl`、手动编号前缀和加粗短行，还原 `一、` / `1.1` / `2.2.1`；同时保留表格（含合并单元格）、图片占位、目录识别、脚注/尾注、文本框。
- **PDF 文本优先**：用 pdf.js 的坐标做行/栏/段落重排，字号和编号前缀判标题，重复页眉页脚/页码/水印移除。表格密集页不伪造 GFM 表格，而是输出保留列分隔的 `[layout]` 文本块，避免串列乱码。
- **省 token**：默认 300 行/48 KB，目录默认省略（保留在 outline），页眉页脚去重，图片只给占位，截断有续读提示。
- **缓存**：内存 LRU（默认 8 个文档），同一文件解析一次；搜索和续读直接复用。

## 配置

在 profile 的 `cordis.patch.yml` 中覆盖插件行配置：

```yaml
- id: pdf-docx-reader
  config:
    readLimit: 300
    readMaxChars: 49152
    searchContext: 160
    searchMaxHits: 10
    cacheDocuments: 8
    inferHeadings: true
    pageMarkers: true
    maxFileBytes: 104857600
```

无设置页；修改后由 Loader 重新加载。

## 限制

- **PDF 表格不保证单元格边界**：复杂表格页输出 `[layout]` 列布局文本；需要精确表格时请用原页/OCR 工具。
- **扫描件不 OCR**：没有文本层的 PDF 会明确报错并建议用 OCR 能力。
- **DOCX 没有固定页码**：页码由 Word 分页决定；需要页码时先转 PDF。
- **图片不内联**：只输出 `[image N: alt, W×H]` 占位。
- **旧 `.doc` 不支持**：请先转 `.docx`。
- **DOCX 标题是启发式**：outline 标注 `inferred`；`inferHeadings: false` 可关闭加粗短行推断。

## 开发

```sh
npm install --ignore-scripts
node test/pages.test.js
node test/numbering.test.js
node test/docx.test.js
node test/pdf.test.js
node test/tools.test.js
node test/cache.test.js
```

测试使用本仓库的 `test.pdf` / `test.docx`，以及 `test/fixtures/structured.docx`。真实文件（第5章.docx、课题3实施方案.docx、两份任务书 PDF）的验收测试是可选的：把它们放到任意本地目录，并设置环境变量 `DSH_PDF_DOCX_TEST_DATA` 指向该目录即可；未设置或缺少文件时自动跳过。

依赖：`pdfjs-dist`（Apache-2.0）、`fflate`（MIT）、`fast-xml-parser`（MIT）。
