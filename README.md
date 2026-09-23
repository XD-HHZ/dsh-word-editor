# dsh-word-editor

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
# {"ok":true,"version":"1.0.0"}
```

返回该 JSON 说明主机半已挂载；再刷新 `http://127.0.0.1:3080/`，在右侧栏文件树里点开一个 `.docx` 应出现编辑器。

### 卸载

删除 `cordis.patch.yml` 里的那段 insert，删掉克隆目录，重启 DSH。

### 更新

方式 A：`cd ~/.dsh/profiles/node_modules/@xd-hhz/dsh-word-editor && git pull`，然后刷新页面；只改主机半时 patch 热加载也会生效。

## 使用

右侧栏文件树 → 点击 `.docx`，出现工具栏与正文：

| 控件 | 作用 |
| --- | --- |
| B / I / U | 加粗 / 斜体 / 下划线 |
| H1 / H2 / H3 / P | 一级 / 二级 / 三级标题 / 正文 |
| • 列表 / 1. 列表 | 无序 / 有序列表 |
| ◉ 检查 | 按当前编辑区重新读取（诊断用） |
| 保存 | 写回原文件，状态栏显示段数、XML 长度与字节数 |

## 已知边界

- **改动过的段落**按 `pStyle` / `numPr` 重新生成，会丢失该段原有的 `pPr`（缩进、对齐、行距、自定义样式名）；未改动的段落不受影响；
- 表格、图片、页眉页脚、批注、修订、分页符**保留但不参与编辑**；
- 主机路由按页面给出的路径用 `node:fs` 写入，只校验同源（`Origin`/`Host` 一致），不受 DSH 文件沙箱约束——它是由你安装的本地插件，请自行判断是否可接受；
- 仅支持 `.docx`（Word 2007+），不支持旧版 `.doc`。

## 开发

```bash
npm run check    # node --check lib/index.js && node --check lib/client.js
npm run smoke    # 用假模块表加载 bundle，校验 apply/inject 与两处注册
```

两条硬性约定，改代码时别踩：

1. **bundle 里 `__ModuleLoader__.load({ id })` 必须等于 `package.json` 的 `name`**（启动图按包名请求 `<name>/client.js`，模块表按这个 id 注册）。`npm run smoke` 会断言这一点；
2. **`package.json` 的 `name` 决定了 patch 行里的写法**，改名后要同步改 `cordis.patch.yml`（并给 `@` 开头的名字加引号）。

## 许可

MIT，见 [LICENSE](LICENSE)。
