# walioffice-dsh-plugin

`walioffice-dsh-plugin` 是可直接安装到 DeepSeek Harness（DSH）Web profile 的办公工具插件。插件以单个 npm 包发布，安装后一次注册 10 个办公工具，并自动注入 WaLiOffice Web 界面，无需复制源码到 DSH 仓库。

## 兼容要求

- DSH：支持 Web profile，并建议使用与插件 peer dependencies 匹配的最新版
- Node.js：`22.19+` 或 `24+`
- npm：用于安装 DSH 与插件；pnpm 仅用于源码开发
- LLM：基础办公工具使用 DSH 已配置的 LLM provider

插件运行在 DSH 的 `web` profile 中；只安装到其他 profile 不会自动启用 Web 界面。

## 工具清单

| 工具 | 功能 | 输出 |
| --- | --- | --- |
| `ppt_plan` | 生成 PPT 大纲与页面规划 | JSON |
| `ppt_generate` | 生成完整演示文稿 | `.pptx` |
| `doc_generate` | 生成 Word 文档 | `.docx` |
| `md_generate` | 生成 Markdown 文档 | `.md` |
| `sheet_generate` | 生成结构化表格 | `.xlsx` |
| `chart_generate` | 生成 ECharts 图表配置 | JSON |
| `drawio_generate` | 生成 draw.io 图表 | XML |
| `image_prompt` | 调用 Agnes Image API 生成图片 | 图片 URL |
| `video_generate` | 调用 Agnes Video API 生成视频 | 视频 URL |
| `video_storyboard` | 生成视频分镜方案 | JSON |

生成的 `.pptx`、`.docx`、`.xlsx` 和 `.md` 文件默认保存在启动 DSH 时工作目录下的 `output/` 目录。图片与视频工具返回远程资源 URL，不会自动把远程资源下载到该目录。

## Web 界面

安装后，DSH Web 会增加 WaLiOffice 客户端界面：

- 输入框上方的办公类型栏，可选择 Word、Excel、PPT、图表、Draw.io、图片和视频
- 页面右侧的 WaLiOffice 面板，集中展示工具入口与最近产物
- 10 个办公工具的专属执行结果卡片
- 工具执行完成后，产物自动进入右侧预览与汇总列表

## 快速安装

### 1. 安装 DSH

```bash
npm install -g @deepseek-ai/dsh
dsh --version
```

### 2. 安装 WaLiOffice 插件

```bash
dsh plugin --profile web add walioffice-dsh-plugin
```

安装指定版本（将版本替换为目标版本）：

```bash
dsh plugin --profile web add walioffice-dsh-plugin@0.1.13
```

### 3. 启动 DSH Web

```bash
dsh web
```

如果安装插件时 `dsh web` 已经运行，请重启 DSH 并刷新浏览器页面。

## 更新与卸载

更新到最新版（建议先停止正在运行的 `dsh web`）：

```bash
dsh plugin --profile web remove walioffice-dsh-plugin
dsh plugin --profile web add walioffice-dsh-plugin@latest
```

安装后重启 DSH，并刷新浏览器页面；如果仍显示旧界面，先确认安装版本：

```bash
dsh plugin --profile web list
```

卸载：

```bash
dsh plugin --profile web remove walioffice-dsh-plugin
```

## 环境变量

基础文档、表格、PPT、图表和 draw.io 工具使用 DSH 已配置的 LLM 服务。运行时配置优先级为 `WALIOFFICE_*`，其次是 DSH 的同名兼容变量；可以用以下变量覆盖 provider 和模型：

```bash
export WALIOFFICE_LLM_PROVIDER=deepseek
export WALIOFFICE_LLM_MODEL=deepseek-chat
```

图片与视频生成是可选能力，需要额外配置对应 API。插件会优先读取 DSH 的 `~/.dsh/settings.yaml` 与 `~/.dsh/.credentials.yaml`，也可以通过环境变量覆盖。

图片配置示例：

```bash
export AGNES_IMAGE_BASE_URL=https://your-image-api.example.com
export AGNES_IMAGE_API_KEYS=key-1,key-2
export AGNES_IMAGE_MODEL=agnes-image-2.1-flash

export AGNES_VIDEO_BASE_URL=https://your-video-api.example.com
export AGNES_VIDEO_API_KEYS=key-1,key-2
export AGNES_VIDEO_MODEL=agnes-video-v2.5
```

`AGNES_*` 未设置时，图片和视频工具支持使用 `LLM_IMAGE_*`、`LLM_VIDEO_*` 兼容变量；密钥也支持 DSH credentials 中的 `apiKeyEnv`，以及 `AGNES_AI_API_KEY`。仓库内的 [`.env.example`](.env.example) 提供了可复制的变量清单，但 DSH 不会自动加载 `.env`，请在启动 DSH 的同一 shell 中 `export`，或使用 DSH 自身的 credentials 配置。

不要把真实 API key 提交到仓库、README 或 issue 中。

## 常见问题

### 安装后页面没有 WaLiOffice

确认插件安装在 `web` profile，并完整重启 `dsh web`：

```bash
dsh plugin --profile web remove walioffice-dsh-plugin
dsh plugin --profile web add walioffice-dsh-plugin@latest
dsh web
```

### 安装成功但工具不可用

先确认 DSH 的 LLM provider 已配置，再查看 DSH 日志。图片和视频工具还需要对应的 API 地址与密钥；基础文档、表格和演示工具不需要 Agnes API。

### 更新后仍然是旧版本

使用 npm registry 直接确认线上版本，并安装明确版本号，避免本地 profile 或镜像缓存影响：

```bash
npm view walioffice-dsh-plugin version dist-tags.latest --registry=https://registry.npmjs.org/
dsh plugin --profile web add walioffice-dsh-plugin@<version>
```

### 生成文件在哪里

检查启动 DSH 时的当前工作目录下的 `output/`。插件不会把产物写入 npm 全局安装目录。

## 本地开发

```bash
pnpm install
pnpm run typecheck
pnpm run bundle
pnpm run pack:check
```

构建产物：

```text
lib/index.js
lib/index.d.ts
lib/client.js
cordis.patch.yml
```

本地打包并安装测试：

```bash
pnpm pack
dsh plugin --profile web add /absolute/path/to/walioffice-dsh-plugin-0.1.13.tgz
dsh web
```

发布前检查：

```bash
pnpm run typecheck
pnpm run bundle
npm pack --dry-run
```

`npm pack --dry-run` 应至少包含 `README.md`、`lib/index.js`、`lib/index.d.ts`、`lib/client.js`、`cordis.patch.yml` 和 `.env.example`。

## 项目结构

```text
src/index.ts                         对外发布入口
cordis.patch.yml                     DSH 自动挂载配置
packages/extensions/
  walioffice-dsh-office              Office 服务与 LLM helper
  walioffice-dsh-office-bundle       10 个工具的聚合入口
  walioffice-dsh-office-render-*     PPTX/DOCX/XLSX 渲染器
  walioffice-dsh-tool-*              各办公工具实现
```

根包构建时会把所有 `@walioffice/*` workspace 模块合并到 `lib/index.js`，最终安装者只需要安装 `walioffice-dsh-plugin` 一个包。

## 发布信息

- npm 包：`walioffice-dsh-plugin`
- 发布方式：公开 npm 包（`publishConfig.access=public`）
- GitHub：<https://github.com/fuzhengwei/walioffice-dsh-plugin>
- 市场摘要：为 DeepSeek Harness Web 提供文档、表格、演示文稿、图表、Draw.io、图片和视频生成能力
- 推荐关键词：`dsh`、`dsh-plugin`、`deepseek-harness`、`office`、`pptx`、`docx`、`xlsx`

## License

MIT
