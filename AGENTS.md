# AGENTS.md

本文件面向在本仓库工作的编码 agent。目标是修改这个 DSH 插件时，不破坏它的核心契约：**本地解析、按行号可寻址、输出有界、搜索与续读闭环**。

## 1. 项目边界

- 项目根：本文件所在目录（仓库根目录）。
- 仓库根目录本身就是一个可安装的 DSH Host-only 插件包，包名 `dsh-pdf-docx-reader`，版本 `0.1.0`。
- 根目录的 `test.pdf`、`test.docx`：测试样本；测试通过 `sample()` / `fixture()` 解析路径，不依赖机器绝对路径。
- 插件只做三件事：读取本地 PDF/DOCX → 构建虚拟 Markdown 文档 → 暴露 `pdf_docx_read` / `pdf_docx_search`。
- 明确不做：OCR、联网、调用 MinerU/Python/LibreOffice、把页面渲染成图片、内联图片、修改源文件。
- 支持 `.pdf`、`.docx`、`.docm`；`.doc` / `.dot` 明确拒绝并提示先转换。
- 这是 Host 插件，没有客户端 bundle，也没有设置页。

## 2. 常用命令

所有开发命令都在项目根目录执行：

```powershell
npm ci --ignore-scripts
```

日常迭代用 `npm install --ignore-scripts` 也可以。**使用 npm，不要改用 pnpm**：本仓库曾遇到 pnpm isolated linker 下 `pdfjs-dist` / `fast-xml-parser` 传递依赖解析失败的问题。

运行全部测试：

```powershell
Get-ChildItem test -Filter *.test.js | Sort-Object Name | ForEach-Object {
  node $_.FullName
  if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }
}
```

或逐个运行：

```powershell
node test/pages.test.js
node test/numbering.test.js
node test/docx.test.js
node test/pdf.test.js
node test/tools.test.js
node test/cache.test.js
```

注意：

- 仓库**没有 build、lint、typecheck 步骤**，也没有 `npm test` 脚本；不要声称这些命令存在。
- 在 Windows 沙箱里直接跑 `node --test test` 可能因 `spawn EPERM` 失败；用上面的“每个文件单独传 `node`”方式。
- `docx.test.js` / `pdf.test.js` 的真实文档验收测试从环境变量 `DSH_PDF_DOCX_TEST_DATA` 指向的目录读取；未设置或文件不存在时自动 skip，这是预期行为。

## 3. 代码地图

| 路径 | 职责 |
|---|---|
| `index.js` | Cordis 入口：注册两个工具、创建缓存、注册 system-prompt 引导段。 |
| `lib/load.js` | 解析主流程：resolve → readBytes → magic → parse → buildDocument → cache。 |
| `lib/resolve.js` | 扩展名检查、工作目录解析、`ctx.fs.stat`、PDF/ZIP/OLE magic 校验。 |
| `lib/document.js` | **唯一的行号来源**：把 parser 的 blocks 转成 `doc.lines`、outline、meta、blockRanges。 |
| `lib/window.js` | `pdf_docx_read` 的 content/outline/tables 渲染、截断提示、续读脚注。 |
| `lib/read-tool.js` | `pdf_docx_read` 的工具定义、参数校验、错误规范化。 |
| `lib/search-tool.js` | `pdf_docx_search` 的工具定义、行内搜索、上下文片段。 |
| `lib/pdf/extract.js` | pdf.js 文本提取、坐标重排、标题/页眉页脚/分栏/表格页处理。 |
| `lib/docx/unzip.js` | fflate 解包 OOXML，执行 zip-bomb / part / entry 上限。 |
| `lib/docx/styles.js` | `styles.xml` 继承解析、标题样式识别。 |
| `lib/docx/numbering.js` | `numbering.xml` 计数器、编号格式、中文数字、项目符号。 |
| `lib/docx/blocks.js` | DOCX 文档主体遍历、标题推断、表格/图片/脚注/TOC 处理。 |
| `lib/cache.js` | 内存 LRU（默认 8 个文档，总字符上限 20M）+ 并发加载去重。 |
| `lib/config.js` | 配置默认值和严格校验。 |
| `lib/pages.js` | PDF 页码选择解析：`3`、`3-5`、`1,3-5`、`第3-5页`、`3~5`。 |
| `lib/text.js` | CJK 判断、空白折叠、Markdown 转义、搜索用 plain text 工具。 |
| `lib/xml.js` | `fast-xml-parser` preserveOrder 树与 undefined-safe helpers。 |
| `lib/errors.js` | `DocReaderError` 及错误包装。 |
| `test/helpers.js` | 最小假 Host（`ctx.fs` / `ctx.tools` / `ctx.workingDirectory`）。 |

## 4. 数据流与不可破坏的契约

```text
tool args
  → parseReadArguments / parseSearchArguments
  → loadDocument
      → resolveTarget（workingDirectory.ensure → ctx.fs.resolve/stat）
      → ctx.fs.readBytes
      → verifyMagic
      → parsePdf / parseDocx  → blocks + meta
      → buildDocument         → doc.lines
      → DocumentCache         → 复用已解析文档
  → renderRead / searchDocument
  → { text }
```

### 行号契约

- `doc.lines[i].number === i + 1`，且 `offset` 直接对应这个 number。
- 搜索返回的 `L<number>` 必须能直接用于下一次 `pdf_docx_read({ offset })`。
- 不要为了加新 block 而重新编号；新增内容只能在 `buildDocument` 的 push 流程里自然追加。
- `line.text` 是面向模型的渲染文本（可含 Markdown），`line.plain` 是搜索和标题路径用的纯文本。新增 block 类型时两者都要维护。
- `document.js` 当前处理 `page`、`heading`、`paragraph`、`list`、`table`、`image`、`layout`、`note` 八种 block。加新类型时同步加 case，并补 `buildDocument` 测试。
- 搜索片段是**行内**的，匹配用字面量 `indexOf`，不要引入正则或跨行 snippet。
- 任何会改变最终 `doc.lines` 文本的改动，都要提升 `load.js` 的 `PARSER_VERSION`，否则旧缓存会返回旧行号。
- 缓存 key 已经覆盖 kind、绝对路径、mtime、size、`inferHeadings`、`pageMarkers` 和 parser 版本。新增会影响输出的配置时，必须把该配置加入 key。

### 输出预算

- 默认 `readLimit: 300` 行，`readMaxChars: 48 KB`；硬上限 1000 行。
- 不要一次性返回全文；窗口截断时必须给出 `Continue with offset=...`。
- 图片只输出占位，例如 `[image 1, 640×480]`。
- DOCX 不伪造页码；`pages` 参数只允许 PDF。

## 5. 工具接口约定

- 工具名保持 `pdf_docx_read` / `pdf_docx_search`，不要重命名。
- 使用普通 ESM + JSDoc，不引入 TypeScript、构建步骤或 `@deepseek-ai/*` 运行时框架导入；Host 边界上的 `ctx` 用 `any` 即可。
- 参数 schema 使用原始 JSON Schema，`additionalProperties: false`，必填项只放真正必需的字段。
- `execute` 返回 `{ text }`；`output.render` 把 `text` 渲染成单个 `{ type: 'text', text }`。
- 面向模型的错误用 `DocReaderError`，消息要可操作（例如告诉用户转换 `.doc`、解密 PDF、改用 offset）。保留 `FS_*` 和 `AbortError`，其余未知错误走 `normalizeToolError`。
- `isConcurrencySafe: () => true` 依赖 `DocumentCache` 的 in-flight 去重；不要让缓存绕过这条路径。
- 新增工具或参数时，更新 `README.md` 的参数表，并在 `test/tools.test.js` 加回归测试。

### system prompt 引导不是硬钩子

`index.js` 注册的 `tool:pdf-docx-reader` section 只是一句软引导：

> For .pdf and .docx files, use pdf_docx_read or pdf_docx_search instead of read, shell, or Python extraction.

它只在插件启用且工具可见时生效，**不会拦截** `read`、`pwsh` 或 Python。当前没有 `tools/pre-execute` / `tools.guard` 之类的强制拦截。若用户明确要求加强制 guard，需要单独设计、说明副作用并补测试，不要默认添加全局拦截。

## 6. PDF 解析要点

- 懒加载 `pdfjs-dist/legacy/build/pdf.mjs`，保持插件激活便宜。
- **pdf.js 只接受 `Uint8Array`，不接受 Node `Buffer`**；`load.js` 中的 `Uint8Array.from(raw)` 是必要兼容层，不要删。
- 提取基于 text item 的坐标：先按 top-down viewport 归一化，再做行分组、分栏拆分、标题判断。
- 书签标题优先用于标题识别；正文列表项不能被误判成标题。
- 重复页眉/页脚、页码、水印会被移除；PDF 表格密集页输出 `[layout pN]` 的保列文本，不伪造 GFM 表格。
- 没有文本层的页面输出明确占位；整份扫描件应报“需要 OCR”，不要静默返回空文档。
- 改 `extract.js` 的启发式规则时，至少用一个真实 18 页任务书 PDF 和 `test.pdf` 做回归。

## 7. DOCX 解析要点

- 不按样式硬转 Markdown；真实文档经常整篇 `Normal`，标题靠自动编号、`outlineLvl` 或加粗短行表达。
- 解析顺序：`word/document.xml` 为主体，`styles.xml` 做继承，`numbering.xml` 维护计数器；表格处理 `gridSpan` / `vMerge`。
- TOC 和“图/表目录”按设计省略，避免重复内容占 token；正文目录信息保留在 outline。
- 标题识别来源：样式级别、`outlineLvl`、自动编号、手动前缀（`一、` / `1.1` / `（一）`）、加粗短行。`inferHeadings: false` 只关闭最后的加粗推断。
- 解析器必须对 undefined / 缺失节点安全；`xml.js` 的 `attr` / `child` / `textContent` 已做守卫，不要假设节点一定存在。
- 解包上限：总解压 200 MB、单 part 64 MB、5000 个 entry；不要放宽而不评估 zip bomb 风险。
- DOCX 没有固定页码，这是产品限制，不要尝试从 XML 推算“页码”。

## 8. 配置与文档

配置默认值在 `lib/config.js` 的 `DEFAULT_CONFIG` 和 `cordis.patch.yml` 中各有一份，校验逻辑统一在 `resolveConfig` 中。当前键：

```yaml
readLimit: 300
readMaxChars: 49152
searchContext: 160
searchMaxHits: 10
cacheDocuments: 8
inferHeadings: true
pageMarkers: true
maxFileBytes: 104857600
```

新增/修改配置时要同步：

1. `lib/config.js`：默认值、类型/范围校验。
2. `cordis.patch.yml`：composition 默认值。
3. `README.md`：配置示例和参数行为。
4. 如果影响工具 schema 文案，更新 `read-tool.js` / `search-tool.js` 的 description。
5. 如果改变输出行，提升 `PARSER_VERSION`。

`locale/*.json` 只有 `meta.title` / `meta.description`，没有设置页；改 UI 文案时保持中英文对应。新增包内顶层文件时，更新 `package.json` 的 `files`。

## 9. 测试与验收

- 测试框架是 Node 内置 `node:test` + `node:assert/strict`，没有 Jest/Vitest。
- `test/helpers.js` 是假 Host；如果插件新增 `ctx` 服务依赖，必须同步扩展 harness。
- `test/fixtures/structured.docx` 是生成好的固定样本；新 fixture 放到 `test/fixtures/`，用 `fixture('name')` 引用。
- 解析启发式的改动要有“正向 + 反向”用例：既断言识别出的标题，也断言正文/表格不会被误判。
- 修改 `read-tool.js` / `search-tool.js` 后，用 `createHarness()` 做端到端调用，不只看 parser 单元测试。
- 最低验收：
  - 六个测试文件全部通过；
  - `pdf_docx_read` 能读根目录 `test.pdf` 和 `test.docx`；
  - `pdf_docx_search` 用一个短字面量能返回行号，且该行号能续读；
  - `focus: outline` / `focus: tables` 不报错；
  - 无效文件、缺失文件、错误页码给出可操作错误。
- 修改公共输出后，同步更新 `README.md` 和对应测试的期望值；不要只改测试来“通过”。

## 10. 安装与重载

如果本机 profile 已用 link 安装本插件，入口应指向仓库根目录。修改 Host 代码后，通常需要让 Loader 重新加载（Plugins 页面禁用/启用，或重启 DSH），再到新会话或下一步确认 `pdf_docx_read` 可见；**只刷新 Web 页面不够**。

安装/移除命令（把 `<REPO_PATH>` 换成仓库的绝对路径；如果 `dsh` 不在 PATH，用 DSH 安装目录的 CLI 或 Web 的 Plugins 页面）：

```powershell
dsh plugin --profile desktop add link:<REPO_PATH>
dsh plugin --profile desktop remove dsh-pdf-docx-reader
```

不要为了排查本插件而改动其他已安装插件的启用状态。

## 11. 已知限制与常见坑

- **测试路径必须可移植**：仓库内样本用 `sample()`，固定 fixture 用 `fixture()`，真实文档用 `DSH_PDF_DOCX_TEST_DATA`；发布前不要重新引入 Windows 盘符、POSIX 家目录或任何本机绝对路径。
- **Windows 路径/编码**：项目路径和真实样本包含中文；保持 UTF-8，处理 `\` 与 `/` 时不要假设只有一种。
- **无 OCR**：扫描 PDF 没有文本层时只能报错并建议 OCR，不能假装成功。
- **PDF 表格不保证单元格边界**：`[layout]` 文本保留列顺序，不是精确表格；需要精确表格请用原页/OCR 工具。
- **DOCX 表格**：支持 `gridSpan` / `vMerge`，但嵌套表格、复杂浮动对象仍可能退化。
- **`.doc` 不支持**：只支持 OOXML `.docx` / `.docm`；OLE/加密文件会被 magic 检查拒绝。
- **图片不内联**：只输出占位和可选尺寸；没有图像理解能力。
- **缓存是进程内内存**：不是持久缓存；重启或插件重载后失效。缓存满时会按 LRU 淘汰，这是预期行为。
- **行号不是永久稳定 ID**：同一文件在 `inferHeadings` / `pageMarkers` / parser 版本变化后，行号可能变化；跨调用续读应使用当前返回的 `offset`。
- **不要编辑 `node_modules/`**：依赖问题通过 `package.json` / lockfile 解决。
- **当前源码目录没有 `.git`**（截至生成时）：不要假设可以 `git diff` / `git checkout` 恢复文件；若之后初始化了仓库，以实际状态为准。

## 12. 完成标准

修改后至少做到：

1. 相关测试通过，全量六个测试文件通过。
2. 用 `test.pdf` 和 `test.docx` 做一次真实工具调用冒烟。
3. 若改了输出格式、工具 schema 或配置，更新 `README.md`、locale 或 `cordis.patch.yml`。
4. 若改了行号/渲染逻辑，提升 `PARSER_VERSION` 并补回归测试。
5. 不引入联网、OCR、Python/LibreOffice 依赖；保持 Host-only。
