# dsh-word-editor

[![npm version](https://img.shields.io/npm/v/@xd-hhz/dsh-word-editor.svg)](https://www.npmjs.com/package/@xd-hhz/dsh-word-editor)
[![npm downloads](https://img.shields.io/npm/dm/@xd-hhz/dsh-word-editor.svg)](https://www.npmjs.com/package/@xd-hhz/dsh-word-editor)
[![CI](https://github.com/XD-HHZ/dsh-word-editor/actions/workflows/ci.yml/badge.svg)](https://github.com/XD-HHZ/dsh-word-editor/actions/workflows/ci.yml)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)
[![DSH web plugin](https://img.shields.io/badge/DSH-web%20plugin-4C6EF5.svg)](https://github.com/deepseek-ai/deepseek-harness)

在 [DeepSeek Harness](https://github.com/deepseek-ai/deepseek-harness)（DSH）的右侧栏「文档」标签页里**直接打开、编辑并保存 `.docx`** 的 Web 插件。

不需要安装 Word，也不需要额外的文档服务：浏览器负责解析与重建 docx，主机半负责字节精确落盘。

## 它做什么

- 在右侧栏文件树里点击任意 `.docx`，直接进入可编辑视图（标题、段落、加粗/斜体/下划线、有序/无序列表）；
- 点「保存」写回**同一个文件**；
- **只有你改动过的段落会被重新生成**，其余段落与所有其他部件（样式表、主题、字体表、settings、页眉页脚、图片、批注…）原样回写，逐字节不变；
- 同一文档只允许一个编辑器实例可编辑，另开的窗口显示接管提示，避免两边各写一份。

## 架构

```
浏览器半 (lib/client.js)                       主机半 (lib/index.js)
  __ModuleLoader__ 模块表加载                    cordis 插件（真实 Node 环境）
  ├ documentPreviews.register(extensions: docx)  ├ node:fs 直接读写二进制
  ├ sidebar.right.tab.document 正文与工具栏       ├ POST /word-editor/read
  └ fetch ─────── /word-editor/* ─────────────►  ├ POST /word-editor/save（写完回读校验）
                                                 └ 同源校验：拒绝跨站写入
```

- **没有构建步骤**：`lib/client.js` 是按 DSH 模块表格式手写的 bundle（`window.__ModuleLoader__.load({ id, factory })`），改完直接生效；
- 浏览器半用 `DecompressionStream` / `CompressionStream` 处理 `deflate-raw`，自实现 zip 读取与 CRC32；
- 主机半用 `node:fs` 写盘，因此不受 DSH 文本文件服务的限制，字节精确。

## 安装

要求：DSH 已按 Web 方式运行（`dsh web`，界面在 `http://127.0.0.1:3080`）。

安装只有两步：**把包放到能被 profile 解析的位置** + **在 profile 的 patch 层加一行**。

### 方式 A：克隆进 profile 的 `node_modules`（已实测，推荐）

```bash
mkdir -p ~/.dsh/profiles/node_modules/@XD-HHZ
git clone https://github.com/XD-HHZ/dsh-word-editor.git \
  ~/.dsh/profiles/node_modules/@xd-hhz/dsh-word-editor
```

```yaml
# 编辑 ~/.dsh/profiles/web/cordis.patch.yml
# 注意：包名以 @ 开头，在 YAML 里必须加引号（@ 是保留指示符）
- insert:
    - id: ui-word-editor
      name: '@xd-hhz/dsh-word-editor'
```

保存后 DSH 会**热加载**这一行（profile 的 `patchReload: live`），然后**刷新浏览器页面**。

> ⚠️ 这个位置在 profile 里执行 `pnpm install` 时可能被裁剪；若你经常动 profile 依赖，用方式 B。

### 方式 B：npm / pnpm 依赖（最抗裁剪）

已发布到 npm 时：

```bash
cd ~/.dsh/profiles/web
pnpm add @xd-hhz/dsh-word-editor
# 或直接从 GitHub 装：
pnpm add github:XD-HHZ/dsh-word-editor
```

patch 行同上（`name: '@xd-hhz/dsh-word-editor'`）。

### ❌ 不要用 `file:` URL 命名这一行

```yaml
# 这样写：主机半能起来，但浏览器半不会入启动图 → 界面里不会出现 docx 编辑器
- insert:
    - id: ui-word-editor
      name: file:///C:/Users/you/plugins/dsh-word-editor/lib/index.js
```

实测结论：DSH 的 client-modules 用 Node 模块解析定位包清单，**路径形式的行名解析不到 `package.json`**，于是这个包不会进入浏览器启动名单（还会把扫描缓存指向旧路径）。**必须让 row 写裸包名**，包本身放在 profile 能解析到的 `node_modules` 下。

### 验证安装

```bash
curl http://127.0.0.1:3080/word-editor/ping
# {"ok":true,"version":"1.1.0"}
```

返回该 JSON 说明主机半已挂载；再刷新 `http://127.0.0.1:3080/`，在右侧栏文件树里点开一个 `.docx` 应出现编辑器。

### 卸载

删除 `cordis.patch.yml` 里的那段 insert，删掉克隆目录，重启 DSH。

### 更新

方式 A：`cd ~/.dsh/profiles/node_modules/@xd-hhz/dsh-word-editor && git pull`，方式 B：`npm update @xd-hhz/dsh-word-editor`。之后分两种情况：

- 只改了**浏览器半**（`lib/client.js`）：刷新页面即可。DSH 对插件的 client bundle 装了文件监听，改动会重新读盘并刷新 bundle 的 `rev`（内容哈希，bundle 不是每次请求读盘，所以必须等监听器察觉到改动）；
- 改了**主机半**（`lib/index.js`）：要让 DSH 重新加载这个插件（重启最稳）。实测过：改完 `lib/index.js` 不重启，下面的 `ping` 仍然报旧版本号——所以主机半不会被自动重新加载。

判断主机半是不是新版，看它自报的版本号：

```bash
curl http://127.0.0.1:3080/word-editor/ping
# {"ok":true,"version":"1.1.0"}
```

版本号是运行时从 `package.json` 读的，所以这里能直接看出 `git pull` 有没有真的生效。

## 使用

右侧栏文件树 → 点击 `.docx`，出现工具栏与正文：

| 控件 | 作用 |
| --- | --- |
| B / I / U | 加粗 / 斜体 / 下划线 |
| 字号 | 给**选中的文字**设字号（9 / 10.5 / 12 / 14 / 16 / 18 / 24 / 36 pt），选「字号」回到默认 |
| 颜色 | 给选中的文字设颜色 |
| A | 清除所选文字的字号与颜色 |
| H1 / H2 / H3 / P | 一级 / 二级 / 三级标题 / 正文 |
| • 列表 / 1. 列表 | 无序 / 有序列表 |
| ◉ 检查 | 按当前编辑区重新读取（诊断用） |
| 保存 | 写回原文件，状态栏显示段数、XML 长度与字节数 |

字号与颜色是**按选区**生效的：先选中文字再点。选区**必须在同一段内**，跨段或选中只读内容会被拒绝并在状态栏说明——浏览器无法让一个内联元素跨过段落边界，硬做会把段落结构搅乱甚至合并。

正文里显示的是**文档自己的字号与颜色**（`w:sz` 折算成 pt、`w:color` 直接上色），所以界面不会把 14pt 的红字显示成默认大小的黑字。

保存时只重写**你真正动过**的段落：没碰的段落按原字节回写，动过的段落继承原有的缩进、对齐、行距与样式，段内没改的文字也保持原格式。

编辑器管不了的内容**会显示出来但不可编辑**：表格渲染成只读表格，含图片、超链接、域、换行、内容控制或修订的段落右侧带一条黄色标记、整段只读（鼠标悬停有说明）。这样你不会在「看起来能编辑」的地方改了却白改，也不会因为它们而丢内容。

## 已知边界

**数据安全**

- 打开再保存、什么都没改，`document.xml` **逐字节不变**（真实 Word 文件上验证过，且是走「界面 → 采集 → 回写」整条链路）；
- **非段落内容不会被删除**：表格、块级内容控制、书签/批注范围标记这类 body 子元素在界面上不进 DOM，但保存时以文件里解析出的结构为准，原样保留——这是刻意的设计，早期版本在这里会把表格整个删掉；
- **只读段落永远不会被重写**，即使浏览器把它的 DOM 弄乱了也一样（判定不依赖 DOM）；
- 未改动的书签、批注范围、权限标记、校对标记会**原位放回**，所以含书签的标题仍然可以正常编辑；
- 改动过的段落继承原有 `pPr`（缩进、对齐、行距、自定义样式），`pStyle`/`numPr` 只做定点替换；段内未改动的 run 逐字节复用；
- 改字号时 `w:szCs`（复杂文种字号）会跟着一起改，避免同一段里两种文种字号不一致。

**做不到的**

- 表格、图片、页眉页脚、批注、修订**只读**，不能编辑，也不能新增；
- 字符格式只到**字号、颜色、加粗/斜体/下划线**：字体名、突出显示色、上下标、字距还不能设（原有值会被保留，只是改不了）；
- 主机路由按页面给出的路径用 `node:fs` 写入，只校验同源（`Origin`/`Host` 一致），不受 DSH 文件沙箱约束——它是由你安装的本地插件，请自行判断是否可接受；
- 仅支持 `.docx`（Word 2007+），不支持旧版 `.doc`。

## 开发

```bash
npm test              # 下面五个全跑
npm run check         # node --check lib/index.js && node --check lib/client.js
npm run smoke         # 用假模块表加载 bundle，校验 apply/inject 与两处注册
npm run host          # 用假 ctx 驱动 apply，校验路由集合与 ping 自报的版本号
npm run roundtrip     # docx 编解码的无头测试（47 项断言，不需要浏览器）
npm run structure     # 真实文档结构的防数据丢失测试（40 项断言）
```

`roundtrip` 覆盖最容易悄悄坏掉的部分：无改动时逐字节不变、改文案时 `pPr` 与未改动 run 原样保留、**只改格式（加粗、改字号、改颜色，都不动文字）也必须被检测到**、样式/列表切换时 `pPr` 是定点修补而不是重建、字号与颜色在 `rPr` 里按 schema 顺序落位（`w:color` 先于 `w:sz`）、改字号时 `w:szCs` 一起动、新段落与空段落的输出形状。

`structure` 是防数据丢失的那一道：夹具（`test/fixtures.mjs`，XML 字符串，仓库里不放二进制）覆盖表格、超链接、图片、域、换行、书签与块级内容控制，断言**打开+保存逐字节不变**、非段落子元素不丢、只读段落即使被伪造的 DOM 数据篡改也保持原字节。它特意调用保存按钮真正调用的那个 `saveDocumentXml`，而不是底层的 emitter——早期测试就是这里出了偏差：绕过保存路径去测 emitter，于是「保存时表格被整个丢掉」这个 bug 一直没被发现。它还用一个**按 surface 结构搭出来的假 DOM** 驱动真实的采集器，所以「界面显示的格式 ⇄ 文档里的格式」对不上时，会表现为字节差异而不是悄悄通过。

一条容易踩的坑，改格式相关代码时留意：**「采集器没报告」不等于「文档里没有」**。字号/颜色只有工具栏的显式重置才会被清除，未声明的值一律视为「不要动」——否则一个浏览器没保留住的 span 就会把用户文档里的字号悄悄抹掉。

另外还有一个针对**真实 Word 文件**的不变式检查（仓库里没有测试用 docx，所以不进 CI，自己指一个文件跑）：

```bash
node test/noop-invariant.mjs "D:\\path\\to\\your.docx"
# document.xml    : 4041 chars -> 4041 chars
# paragraphs      : 8 (0 read-only)
# sized/coloured  : 0 runs
# other children  : 0 (emitted verbatim)
# surface round   : byte-exact
# NO-OP INVARIANT OK (byte-for-byte, including through the surface)
```

它做的事就是：解析再原样输出，要求每个段落都与原文逐字节相同。**改编解码层前后都值得跑一次**——任何“顺手规范化一下 XML”的改动都会在这里立刻暴露。

两条硬性约定，改代码时别踩：

1. **bundle 里 `__ModuleLoader__.load({ id })` 必须等于 `package.json` 的 `name`**（启动图按包名请求 `<name>/client.js`，模块表按这个 id 注册）。`npm run smoke` 会断言这一点；
2. **`package.json` 的 `name` 决定了 patch 行里的写法**，改名后要同步改 `cordis.patch.yml`（并给 `@` 开头的名字加引号）。

### 让运行中的插件就是你的仓库

**不要直接改 `node_modules` 里的文件**——下一次 `pnpm install` 会静默覆盖。把安装位置换成指向仓库的链接（Windows 用 junction，macOS/Linux 用 `ln -s`，或 pnpm 的 `link:` 协议）：

```powershell
# Windows（junction 不需要管理员权限）
Remove-Item "$env:USERPROFILE\.dsh\profiles\node_modules\@xd-hhz\dsh-word-editor" -Recurse -Force
New-Item -ItemType Junction `
  -Path   "$env:USERPROFILE\.dsh\profiles\node_modules\@xd-hhz\dsh-word-editor" `
  -Target "<你的仓库绝对路径>"
```

```bash
# 或者用 pnpm 的 link: 协议（在 profile 里执行，保证是符号链接，装依赖也不会被换掉）
cd ~/.dsh/profiles/web && pnpm add link:<你的仓库绝对路径>
```

之后的开发环路：

| 动作 | 生效方式 |
| --- | --- |
| 改 `lib/client.js`（浏览器半） | 刷新页面即可（无构建步骤） |
| 改 `lib/index.js`（主机半） | profile 的 `patchReload: live` 会重启该行 |
| `git pull` | 等于更新运行中的插件 |
| `git checkout .` | 回滚 |
| 发布 | `npm version patch && git push --follow-tags && npm publish` |

注意：DSH 的客户端 HMR（`pnpm run dev:web`）只服务 DSH 自身源码树的重建，不覆盖手写 bundle，因此浏览器半的改动需要**刷新页面**；另外别移动或删除被链接的仓库目录，否则链接悬空、插件会加载失败（移动后重新指向即可）。

## 许可

MIT，见 [LICENSE](LICENSE)。
