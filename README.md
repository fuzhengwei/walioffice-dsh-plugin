# walioffice-dsh-plugin

`walioffice-dsh-plugin` 是可直接安装到 DeepSeek Harness（DSH）的办公工具插件。插件以单个 npm 包发布，安装后一次注册 10 个办公工具，无需复制源码到 DSH 仓库。

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

生成的 `.pptx`、`.docx`、`.xlsx` 和 `.md` 文件默认保存在启动 DSH 时工作目录下的 `output/` 目录。

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

安装指定版本：

```bash
dsh plugin --profile web add walioffice-dsh-plugin@0.1.1
```

### 3. 启动 DSH Web

```bash
dsh web
```

如果安装插件时 `dsh web` 已经运行，请重启 DSH 并刷新浏览器页面。

## 更新与卸载

更新到最新版：

```bash
dsh plugin --profile web remove walioffice-dsh-plugin
dsh plugin --profile web add walioffice-dsh-plugin@latest
```

卸载：

```bash
dsh plugin --profile web remove walioffice-dsh-plugin
```

## 环境变量

基础文档、表格、PPT、图表和 draw.io 工具使用 DSH 已配置的 LLM 服务。可以用以下变量覆盖模型：

```bash
export WALIOFFICE_LLM_PROVIDER=deepseek
export WALIOFFICE_LLM_MODEL=deepseek-chat
```

图片与视频生成是可选能力，需要额外配置 Agnes API：

图片生成会优先从 DSH 的 `~/.dsh/settings.yaml`（`llm-pi-ai.providers`）中查找包含 `image` 的模型，读取对应的 `baseURL`，并从 `~/.dsh/.credentials.yaml` 读取 provider 的 `apiKeyEnv` 对应密钥。例如配置了 `agnes-ai` provider 和 `agnes-image-2.1-flash` 后，无需再重复设置图片地址；如果需要覆盖 DSH 配置，仍可使用以下环境变量：

```bash
export AGNES_IMAGE_BASE_URL=https://your-image-api.example.com
export AGNES_IMAGE_API_KEYS=key-1,key-2
export AGNES_IMAGE_MODEL=agnes-image-2.1-flash

export AGNES_VIDEO_BASE_URL=https://your-video-api.example.com
export AGNES_VIDEO_API_KEYS=key-1,key-2
export AGNES_VIDEO_MODEL=agnes-video-v2.5
```

也可以复制 `.env.example` 查看完整变量说明。DSH 进程必须能够读取这些环境变量。

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
cordis.patch.yml
```

本地打包并安装测试：

```bash
pnpm pack
dsh plugin --profile web add /absolute/path/to/walioffice-dsh-plugin-0.1.1.tgz
dsh web
```

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

## License

MIT
