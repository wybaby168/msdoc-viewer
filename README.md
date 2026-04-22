# msdoc-viewer

一个面向 **MS-DOC（`.doc` 二进制 Word）** 的 **TypeScript 严格类型** 解析与预览库。它把旧版 Word 文档解析为结构化 AST，再渲染为可直接嵌入任意项目的 `HTML + CSS`，同时提供 DOM Viewer 与 Web Worker 通道。

> 本项目处理的是 **`.doc`**，不是 `.docx`。

## 特性

- **整仓 TypeScript 重构**：源码、公开 API、Worker 协议、解析 AST 全部使用 TypeScript，并开启 `strict` 模式
- **完整声明文件产物**：构建后输出 `dist/*.d.ts`，便于社区阅读、联想、二次开发
- **纯前端方案**：不依赖 LibreOffice、COM、ActiveX 或服务器端转换
- **按 MS-DOC 主链路实现**：CFB/OLE → FIB → CLX/Piece Table → FKP → PAPX/CHPX/TAPX → STSH → HTML
- **可扩展架构**：把二进制读取、属性解码、样式继承、表格布局、渲染层完全拆分
- **支持 Web Worker**：避免大文档解析阻塞 UI 线程
- **生产向安全加固**：渲染层会过滤 `javascript:` / `file://` 等不安全链接，避免生成无效或危险的 HTML
- **CI 就绪**：内置 GitHub Actions，自动执行 typecheck + smoke test

## 当前能力

已经覆盖的主要能力：

- CFB/OLE 容器读取（支持 512 / 4096 扇区 compound file，包含对尾部非整扇区流的容错读取，避免因最后一个未用满的扇区误报越界）
- FIB（`FibBase` / `FibRgLw` / `FibRgFcLcb`）解析
- CLX / Piece Table 文本恢复
- FKP（PAPX / CHPX）属性读取
- STSH 样式表解析与样式继承
- 字体表读取
- 常见字符样式：粗体、斜体、下划线、删除线、字号、颜色、高亮、大小写、上下标、字体切换
- 常见段落样式：对齐、缩进、段前后距、行距、分页控制、边框
- 表格：`sprmTDefTable`、单元格宽度、横向/纵向合并、边框、垂直对齐、nowrap、fitText
- 图片：按 `PICFAndOfficeArtData` / `PICF` / `OfficeArtInlineSpContainer` / `OfficeArtBStoreContainerFileBlock` / `OfficeArtFBSE` / `OfficeArtBlip*` 解析，优先提取浏览器可直接显示的 PNG/JPEG/BMP 等位图
- 多 story 内容：主文档、脚注、尾注、批注、页眉页脚、textbox、header textbox
- floating shape 锚点：解析 `PlcfSpaMom` / `PlcfSpaHdr` 的 `Spa` 结构，把包围框、锚定基准、环绕方式、前后文关系暴露到 AST，并自动关联可匹配的 textbox
- 行内引用：脚注/尾注引用、批注引用、复杂域嵌套、修订插入/删除标记
- 矢量图：对常见 `EMF` / `WMF` 形状记录做纯前端 SVG 转换，浏览器可直接显示
- 链接图片：识别 `stPicName` / 外链目标，无法内嵌的本地 `file://` 图片会以非点击型回退占位展示，而不是输出损坏的 `<img>` 或不可用链接
- OLE / ObjectPool 附件提取
- 域代码的基础处理（例如超链接、`INCLUDEPICTURE` 外链图片）
- 浏览器 Viewer 与 Worker Client

## 工程结构

```text
src/
  core/
    binary.ts        # 二进制读取器
    cfb.ts           # CFB/OLE 容器解析
    utils.ts         # 字节、编码、范围等底层工具
  msdoc/
    fib.ts           # FIB 读取
    clx.ts           # CLX / Piece Table
    fkp.ts           # FKP 页读取
    sprm.ts          # SPRM 操作数解码
    properties.ts    # PAPX / CHPX / TAPX 到状态对象的归并
    styles.ts        # STSH / 样式继承
    fonts.ts         # 字体表
    objects.ts       # 图片 / OfficeArt / OLE / ObjectPool
    parser.ts        # 主解析流程，输出 AST
  render/
    html.ts          # AST -> HTML/CSS
  types.ts           # 完整类型声明与公共接口
  viewer.ts          # DOM Viewer
  worker-client.ts   # Worker 客户端
  worker.ts          # Worker 入口
  index.ts           # 公共导出

demo/
  index.html

test/
  smoke.mjs
```

## 安装与构建

```bash
npm install
npm run typecheck
npm run build
```

构建产物输出到 `dist/`：

- `dist/index.js`
- `dist/index.d.ts`
- `dist/worker.js`
- 其余模块的 JS / d.ts

## 快速开始

### 1）直接解析并渲染

```ts
import { parseMsDoc, renderMsDoc } from 'msdoc-viewer';

const buffer = await file.arrayBuffer();
const parsed = parseMsDoc(buffer);
const rendered = renderMsDoc(parsed);

document.getElementById('app')!.innerHTML = `
  <style>${rendered.css}</style>
  <div class="msdoc-root">${rendered.html}</div>
`;
```

### 2）使用内置 Viewer

```ts
import { createMsDocViewer } from 'msdoc-viewer';

const container = document.getElementById('app')!;
const viewer = createMsDocViewer(container);

await viewer.load(file);
```

### 3）使用 Worker 提升体验

```ts
import { createMsDocViewer, MsDocWorkerClient } from 'msdoc-viewer';

const workerClient = MsDocWorkerClient.create(
  new URL('./node_modules/msdoc-viewer/dist/worker.js', import.meta.url)
);

const viewer = createMsDocViewer(document.getElementById('app')!, {
  workerClient,
});

await viewer.load(file);
```

## API 概览

### `parseMsDoc(input, options?)`

把 `.doc` 解析成结构化结果：

```ts
import type { MsDocParseResult } from 'msdoc-viewer';
```

核心字段：

- `warnings`: 解析期告警
- `meta`: 文档元数据与统计
- `fonts`: 字体表摘要
- `styles`: 样式表摘要
- `blocks`: 结构化块级 AST（段落、表格、脚注/尾注、批注、页眉页脚、textbox、floating shape、附件）
- `assets`: 图片 / 附件资产

### `renderMsDoc(parsed, options?)`

把 AST 渲染为：

- `html`
- `css`
- `warnings`
- `meta`
- `assets`
- `parsed`

### `parseMsDocToHtml(input, options?)`

一步完成解析与渲染。

### `convertMetafileToSvg(mime, bytes)`

低层辅助函数，用于把常见 `EMF` / `WMF` 字节流转换成 SVG。

适合在你自己实现自定义图片策略、调试 OfficeArt 资产，或者单独处理历史矢量图时使用。

### `createMsDocViewer(container, config?)`

返回一个面向浏览器的 viewer：

- `load(input, options?)`
- `mount(rendered)`
- `clear()`
- `destroy()`
- `value`

### `MsDocWorkerClient`

为 Worker 提供强类型 RPC 包装：

- `parse(input, options?)`
- `render(parsed, options?)`
- `parseToHtml(input, options?)`
- `destroy()`

## 关键类型

所有公共类型都从根入口导出，例如：

```ts
import type {
  MsDocParseResult,
  MsDocRenderResult,
  MsDocBlock,
  ParagraphBlock,
  TableBlock,
  MsDocAsset,
  MsDocParseOptions,
  MsDocRenderOptions,
  MsDocViewer,
} from 'msdoc-viewer';
```

这意味着：

- 社区开发者可以直接消费 `AST` 与 `asset` 数据自行渲染
- 也可以在现有渲染器之外实现 Markdown、Canvas、PDF 等输出层
- Worker 主线程通信协议也有完整类型约束，避免“隐式 any 消息格式”问题
- 多 story 结果、修订标记、批注/脚注引用节点也都有独立类型，可按需自定义渲染


## 图片解析说明

当前图片链路不再把“从 Data 流里扫 PNG/JPEG 魔数”当成主路径，而是按规范做结构化解析；只有在极少数结构异常文档里才退回签名扫描：

- 通过字符属性里的 `sprmCPicLocation` 找到 Data 流偏移
- 按 `PICFAndOfficeArtData` 读取 `PICF` 头和可选的 `stPicName`
- 递归遍历 `OfficeArtInlineSpContainer` / `OfficeArt` 记录头
- 支持 `OfficeArtInlineSpContainer` 后续的 `rgfb` 文件块，识别 `OfficeArtBStoreContainerFileBlock` / `OfficeArtFBSE` 中嵌入的 `OfficeArtBlip*`
- 命中 `OfficeArtBlipPNG`、`OfficeArtBlipJPEG`、`OfficeArtBlipDIB`、`OfficeArtBlipTIFF` 等 BLIP 记录后，按各自记录布局提取真正的图片负载
- 对 DIB 自动转成浏览器可显示的 BMP
- 对只包含本地外链路径而不包含实际位图数据的图片，保留链接元数据并渲染为回退占位，避免错误地输出损坏图片

这一步专门修复了旧实现里“把 OfficeArt 容器里的非图片字节误判成 EMF，导致图片无法显示”的问题。

对于真正的矢量图（尤其是 `OfficeArtBlipEMF` / `OfficeArtBlipWMF`），库会优先把可直接读取的 EMF/WMF 记录转换成 SVG，再交给浏览器渲染；如果遇到压缩 metafile 或暂未覆盖的记录类型，则会保留原始元数据并以非浏览器直显资源处理。

## 严格类型与可维护性说明

这次重构的重点不只是把文件后缀改成 `.ts`，而是把整条主链路都显式类型化：

- 二进制结构：CFB/FIB/CLX/FKP
- 样式模型：字符、段落、表格状态对象
- 中间 AST：段落、行内节点、表格、附件
- 渲染结果：HTML/CSS/资产
- Viewer / Worker 协议

同时给关键模块补充了注释，尤其是：

- `BinaryReader` 的边界读取行为
- Worker 请求/响应通道
- parser 主流程
- 样式与属性归并逻辑


## 生产可用性补充

这版额外做了几项生产向收口：

- **外链图片安全降级**：`file://`、盘符路径、UNC 路径不会再被渲染成可点击链接，只保留占位信息
- **链接协议白名单**：超链接默认只放行 `http(s)`、`mailto`、`tel`、`ftp` 与文档内 `#anchor`
- **避免无效 HTML**：附件与图片回退在带文档超链接时，不再生成嵌套 `<a>`
- **结构化告警**：对本地外链图片、非浏览器可显示图片、无法识别的图片负载，都会在 `warnings` 里给出结构化告警
- **附件 MIME 推断**：对常见 `doc/docx/xls/xlsx/ppt/pptx/pdf/png/jpg/zip/txt` 附件补充 MIME 推断

## 本地验证

```bash
npm run smoke
```

这个命令会：

1. 先运行 TypeScript 构建
2. 读取 `test/test.doc`、`test/fixtures/image-embedded.doc`、`test/fixtures/image-linked.doc`
3. 验证嵌入图片 fixture 是否被解析成可直接渲染的 PNG/JPEG
4. 验证 `WMF` / `EMF` 会被转换为 SVG
5. 验证 story helper（story window / 批注引用元数据）
6. 验证本地外链图片不会生成 `file://` 可点击链接
7. 验证 `PlcfSpaMom` / `PlcfSpaHdr` 浮动 shape 锚点会被解析并渲染为结构化元数据
8. 验证渲染层会清理 `javascript:` 等不安全链接，并避免嵌套 `<a>`
9. 输出 `test/rendered-sample.html` 与 `test/rendered-image-sample.html`

如果你要手工打开 demo：

```bash
npm run build
# 然后用任意静态服务器打开 demo/index.html
```

## 已知边界

MS-DOC 是历史包袱很重的二进制格式，虽然主链路已经完整重构，但以下场景仍建议用更多真实样本持续压测：

- 极端复杂的嵌套表格
- 压缩的 EMF / WMF metafile BLIP（当前已支持直接读取的常见 EMF/WMF 记录到 SVG，但未内置通用 deflate 解压器）
- 复杂 OfficeArt shape / textbox 的高级几何、旋转、裁剪和更接近 Word 的文本环绕
- 修订/批注/页眉页脚在极少数兼容模式文档中的边角差异
- 非常规域代码组合
- 少见的 OLE 嵌入形式
- 某些兼容模式下的边角 sprm 变体

这些场景并不影响当前的整体架构，后续继续扩展时可直接在现有 TS 类型体系上增量补强。对于常见业务型 `.doc`、包含表格/批注/脚注/页眉页脚/内嵌图片/常见 EMF-WMF 矢量图的历史文档，这个版本已经更适合在生产环境灰度接入并持续用真实样本回归。

## 许可证

MIT

## CI

仓库已包含 `.github/workflows/ci.yml`，在 GitHub 上会自动执行：

- `npm ci`
- `npm run typecheck`
- `npm run smoke`
