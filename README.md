# WaLiOffice DSH Plugin

WaLiOffice 办公工具套件，作为 [DeepSeek Harness (DSH)](https://github.com/deepseek-ai/deepseek-harness) 的 TypeScript Cordis 插件运行。

## 工具清单（10 个）

| 工具 | 功能 | 渲染库 |
|------|------|--------|
| `ppt_plan` | PPT 大纲规划（7 场景推断 + LLM 生成 JSON） | — |
| `ppt_generate` | PPT 渲染为 .pptx 文件 | pptxgenjs |
| `doc_generate` | Word 文档生成 | docx (npm) |
| `md_generate` | Markdown 文档生成 | — |
| `sheet_generate` | Excel 表格生成 | exceljs |
| `chart_generate` | ECharts 图表数据（6 种类型） | — |
| `drawio_generate` | draw.io 图表 XML（6 种类型） | — |
| `image_prompt` | AI 文生图/图生图（Agnes Image 2.1 Flash） | 外部 API |
| `video_generate` | AI 视频生成（Agnes Video V2.5，3 种模式） | 外部 API |
| `video_storyboard` | 视频分镜规划 | — |

## 包结构（12 包）

```
packages/extensions/
  walioffice-dsh-office              核心服务定义（LLM helper、场景推断、主题色板、Context 扩展）
  walioffice-dsh-office-bundle       聚合包（一键加载全部工具 + OfficeService）
  walioffice-dsh-office-render-docx  Word 渲染层（docx npm）
  walioffice-dsh-office-render-pptx  PPT 渲染层（pptxgenjs）
  walioffice-dsh-office-render-xlsx  Excel 渲染层（exceljs）
  walioffice-dsh-tool-chart          图表工具
  walioffice-dsh-tool-doc            Word + Markdown 工具
  walioffice-dsh-tool-drawio         DrawIO 工具
  walioffice-dsh-tool-image          图片工具
  walioffice-dsh-tool-ppt            PPT 工具（plan + generate）
  walioffice-dsh-tool-sheet          Excel 工具
  walioffice-dsh-tool-video          视频工具（generate + storyboard）
```

## 快速开始

### 方式一：集成到 DSH 仓库（推荐开发）

将 `packages/extensions/walioffice-*` 复制到 DSH 仓库的 `packages/extensions/` 下，然后在 DSH 根目录执行：

```bash
pnpm install
pnpm run build:lib:host
npx @deepseek-ai/dsh web --preset walioffice
```

### 方式二：独立安装

```bash
cd walioffice-dsh-plugin
pnpm install
pnpm build
```

然后在 DSH 的 `cordis.yml` 中加载：

```yaml
plugins:
  - name: '@walioffice/dsh-office-bundle'
```

或单独加载需要的工具：

```yaml
plugins:
  - name: '@walioffice/dsh-tool-ppt'
  - name: '@walioffice/dsh-tool-doc'
  - name: '@walioffice/dsh-tool-sheet'
```

## 环境变量

```bash
cp .env.example .env
```

| 变量 | 说明 | 默认值 |
|------|------|--------|
| `WALIOFFICE_LLM_PROVIDER` | LLM provider | `deepseek` |
| `WALIOFFICE_LLM_MODEL` | LLM 模型 | `deepseek-chat` |
| `AGNES_IMAGE_BASE_URL` | 图片 API 地址 | — |
| `AGNES_IMAGE_API_KEYS` | 图片 API 密钥（逗号分隔轮询） | — |
| `AGNES_IMAGE_MODEL` | 图片模型 | `agnes-image-2.1-flash` |
| `AGNES_VIDEO_BASE_URL` | 视频 API 地址 | — |
| `AGNES_VIDEO_API_KEYS` | 视频 API 密钥（逗号分隔轮询） | — |
| `AGNES_VIDEO_MODEL` | 视频模型 | `agnes-video-v2.5` |

## DSH 集成配置

### tsconfig.base.json paths 映射

```json
"@walioffice/*": ["./packages/extensions/walioffice-*/src"]
```

### tsconfig.host.json references

```json
{ "path": "./packages/extensions/walioffice-dsh-office" },
{ "path": "./packages/extensions/walioffice-dsh-office-bundle" },
...
```

### Agent Preset

预设位于 `apps/cli/config/agent-presets/walioffice/`，使用：

```bash
npx @deepseek-ai/dsh web --preset walioffice
```

## 设计决策

- **LLM 调用**：通过 `ctx.llm.stream()` 走 DSH 统一 LLM 服务
- **PPTX 渲染**：pptxgenjs 替代 Rust 手写 OOXML
- **DOCX 渲染**：docx npm 替代 docx-rs
- **XLSX 渲染**：exceljs 替代 rust_xlsxwriter
- **进度推送**：`session/event` 事件系统替代 SSE
- **场景推断**：从 Rust 直接翻译为 TS，保持 7 种场景一致

## License

MIT
