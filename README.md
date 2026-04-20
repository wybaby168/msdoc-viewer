# msdoc-html-viewer

纯 JavaScript 的旧版 `.doc` 读取与 HTML/CSS 渲染插件，目标是把 **MS-DOC（二进制 Word）** 文档解析为可嵌入任意前端项目的结构化数据与可直接挂载的 HTML。

## 设计目标

- 不依赖 `mammoth`、`LibreOffice`、`ActiveX`、`COM`
- 纯 JS 解析 CFB/OLE、FIB、CLX、FKP、STSH、TAPX/PAPX/CHPX
- 输出结构化 AST，再渲染为 `html + css`
- 支持 Web Worker，把 `.doc` 解析放到后台线程
- 允许任何项目直接引用：原生 JS、Vue、React、Vite、Webpack 均可接入

## 已实现能力

- CFB/OLE 容器读取
- FIB 正确读取（含 `FibBase` / `FibRgLw` / `FibRgFcLcb`）
- CLX / Piece Table 文本恢复
- PAPX / CHPX / FKP 读取
- STSH 样式表解析与样式继承
- 常见段落样式：对齐、缩进、段前后距、行距、分页控制、边框
- 常见字符样式：粗体、斜体、下划线、删除线、字号、颜色、高亮、大小写、上下标、字体
- 表格：`sprmTDefTable`、单元格宽度、横向/纵向合并、垂直对齐、边框、nowrap、fitText
- PICF + Data 流图片提取
- ObjectPool/OLE 附件提取
- 字段基础处理：超链接
- DOM Viewer 与 Worker Client

## 安装 / 引用

把整个项目目录拷贝到你的仓库，或发布为 npm 包后使用：

```js
import {
  parseMsDoc,
  renderMsDoc,
  parseMsDocToHtml,
  createMsDocViewer,
  MsDocWorkerClient,
} from 'msdoc-html-viewer';
```

## 最简单用法

```js
import { createMsDocViewer } from 'msdoc-html-viewer';

const viewer = createMsDocViewer(document.getElementById('app'));
const file = document.querySelector('input[type=file]').files[0];
await viewer.load(file);
```

## 直接拿到解析结果

```js
import { parseMsDoc, renderMsDoc } from 'msdoc-html-viewer';

const buffer = await file.arrayBuffer();
const parsed = parseMsDoc(buffer);
const rendered = renderMsDoc(parsed);

console.log(parsed.blocks);
console.log(rendered.html);
console.log(rendered.css);
```

## Worker 用法

```js
import { MsDocWorkerClient, createMsDocViewer } from 'msdoc-html-viewer';

const workerClient = MsDocWorkerClient.create(
  new URL('./node_modules/msdoc-html-viewer/src/worker.js', import.meta.url)
);

const viewer = createMsDocViewer(document.getElementById('app'), {
  workerClient,
});

await viewer.load(file);
```

## API

### `parseMsDoc(input, options?)`

返回结构化对象：

```ts
{
  kind: 'msdoc',
  version: 1,
  warnings: Array<{ message: string }>,
  meta: {
    fib: {...},
    counts: {...}
  },
  fonts: Array<...>,
  styles: Array<...>,
  blocks: Array<...>,
  assets: Array<...>
}
```

### `renderMsDoc(parsed, options?)`

返回：

```ts
{
  html: string,
  css: string,
  warnings: [...],
  meta: {...},
  assets: [...],
  parsed
}
```

### `parseMsDocToHtml(input, options?)`

一步完成解析与渲染。

### `createMsDocViewer(container, config?)`

创建一个面向 DOM 的 viewer：

- `load(input)`
- `mount(rendered)`
- `clear()`
- `destroy()`
- `value`

## 当前工程结构

```text
src/
  core/
    binary.js
    cfb.js
    utils.js
  msdoc/
    clx.js
    fib.js
    fkp.js
    fonts.js
    objects.js
    parser.js
    properties.js
    sprm.js
    styles.js
  render/
    html.js
  index.js
  viewer.js
  worker-client.js
  worker.js
```

## 注意事项

- 这个项目面向 **`.doc` 二进制格式**，不是 `.docx`
- 已对大量常见样式、表格、图片/OLE 场景做了解析，但 Word 的历史兼容分支很多，极端文档仍可能触发 `warnings`
- `OfficeArt` / 旧版 OLE / 域代码 / 嵌套表格 的边角行为已预留扩展点，后续可以继续把更多 spec 细节补齐

## 本地验证

```bash
npm run smoke
```

它会读取 `test/test.doc` 并输出 `test/rendered-sample.html`。
