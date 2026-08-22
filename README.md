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

- 输入框上方的办公类型栏，可选择 Word、Markdown、Excel、PPT、图表、Draw.io、图片和视频
- 综合模式会按交付物类型主动路由工具；明确要求文件类型时优先遵循用户指定，不会把“只要图表”误判为 Excel
- 页面右侧的 WaLiOffice 面板，集中展示工具入口与最近产物
- 10 个办公工具的专属执行结果卡片
- 工具执行完成后，产物自动进入右侧预览与汇总列表

## 安装前准备

安装前请确认已经具备以下条件：

1. 已安装 Node.js `22.19+` 或 `24+`。
2. 已安装 npm。npm 会随 Node.js 一起安装。
3. DSH 支持 `web` profile。
4. 已在 DSH 中配置基础 LLM provider。文档、表格、PPT、图表和 Draw.io 工具都使用该 provider。

先逐条检查版本：

```bash
node --version
```

```bash
npm --version
```

如果电脑中已经安装 DSH，再检查 DSH 版本：

```bash
dsh --version
```

## 安装

以下命令适用于第一次安装。每条命令执行成功后，再执行下一条。

### 第 1 步：安装 DSH

如果还没有安装 DSH，执行：

```bash
npm install --global @deepseek-ai/dsh
```

确认 DSH 命令可用：

```bash
dsh --version
```

### 第 2 步：安装 WaLiOffice 插件

插件必须安装到 `web` profile，执行：

```bash
dsh plugin --profile web add walioffice-dsh-plugin
```

如果需要安装指定版本，把 `<版本号>` 替换成实际版本，例如 `0.1.15`：

```bash
dsh plugin --profile web add walioffice-dsh-plugin@<版本号>
```

查看插件是否已经安装：

```bash
dsh plugin --profile web list
```

输出中应能看到 `walioffice-dsh-plugin`。不要省略 `--profile web`，否则插件可能被安装到其他 profile，Web 界面不会加载。

### 第 3 步：启动 DSH Web

在安装插件的同一个终端中执行：

```bash
dsh web
```

然后打开 DSH Web 页面。若安装插件前已经启动了 `dsh web`，请先停止旧进程，再重新执行上面的启动命令，并刷新浏览器页面。

## 配置

### 基础办公工具

基础办公工具默认使用 DSH 已配置的 LLM。如果需要明确指定 provider 和模型，请在启动 DSH 之前，在同一个终端逐条执行：

```bash
export WALIOFFICE_LLM_PROVIDER=deepseek
```

```bash
export WALIOFFICE_LLM_MODEL=deepseek-chat
```

然后启动 DSH：

```bash
dsh web
```

插件的配置优先级如下：

1. `WALIOFFICE_LLM_PROVIDER`、`WALIOFFICE_LLM_MODEL`
2. `DSH_LLM_PROVIDER`、`DSH_LLM_MODEL`
3. DSH 自身已配置的 LLM provider

### 图片和视频工具（可选）

只有使用 `image_prompt`、`video_generate` 或 `video_storyboard` 时，才需要额外配置图片或视频 API。配置图片 API 时，逐条执行：

```bash
export AGNES_IMAGE_BASE_URL=https://your-image-api.example.com
```

```bash
export AGNES_IMAGE_API_KEYS=key-1,key-2
```

图片工具使用的模型是 `agnes-image-2.1-flash`。如果 DSH 通过 `~/.dsh/settings.yaml` 管理模型，请确保该模型已经配置。

配置视频 API 时，逐条执行：

```bash
export AGNES_VIDEO_BASE_URL=https://your-video-api.example.com
```

```bash
export AGNES_VIDEO_API_KEYS=key-1,key-2
```

视频工具使用的模型是 `agnes-video-v2.0`。如果 DSH 通过 `~/.dsh/settings.yaml` 管理模型，请确保该模型已经配置。

配置完成后，必须在同一个终端启动 DSH：

```bash
dsh web
```

如果不使用 `AGNES_*` 变量，也可以使用兼容变量 `LLM_IMAGE_*`、`LLM_VIDEO_*`。插件还支持从 DSH 的 `~/.dsh/settings.yaml`、`~/.dsh/.credentials.yaml` 读取配置。仓库中的 [`.env.example`](.env.example) 只是变量清单，DSH 不会自动加载 `.env` 文件。

不要把真实 API key 写入 README、提交到 Git，或发布到 issue。

## 使用

### 通过 Web 界面使用

1. 启动 DSH Web：`dsh web`。
2. 在输入框上方选择办公类型，或选择综合模式。
3. 输入任务要求，例如“根据下面的内容生成一份 PPT”或“把这组数据整理成 Excel”。
4. 等待工具执行完成。
5. 在右侧 WaLiOffice 面板查看预览、最近产物和汇总结果。

可直接尝试以下任务：

```text
生成一份关于人工智能入门的 PPT，包含 8 页，最后输出 pptx 文件。
```

```text
把下面的会议纪要整理成 Word 文档，并保存为 docx 文件。
```

```text
根据以下销售数据生成 Excel 表格，并增加一个按月份统计的图表。
```

### 查看生成文件

`.pptx`、`.docx`、`.xlsx` 和 `.md` 文件默认写入启动 DSH 时所在目录的 `output/` 文件夹。先查看当前目录：

```bash
pwd
```

再查看生成的文件：

```bash
ls -la output
```

图片和视频工具返回远程资源 URL，不会自动下载到 `output/`。

## 升级

### 升级 WaLiOffice 插件到最新版

为避免 profile 缓存或旧进程继续使用旧代码，推荐使用“停止 DSH → 删除旧插件 → 安装新插件 → 重启 DSH”的方式升级。

第 1 步，停止正在运行的 `dsh web`。在运行该命令的终端按 `Ctrl+C`。

第 2 步，删除旧版本插件：

```bash
dsh plugin --profile web remove walioffice-dsh-plugin
```

第 3 步，安装最新版插件：

```bash
dsh plugin --profile web add walioffice-dsh-plugin@latest
```

第 4 步，确认插件版本：

```bash
dsh plugin --profile web list
```

也可以先查看 npm registry 当前的最新版：

```bash
npm view walioffice-dsh-plugin version --registry=https://registry.npmjs.org/
```

然后安装返回的明确版本号：

```bash
dsh plugin --profile web add walioffice-dsh-plugin@<版本号>
```

第 5 步，重新启动 DSH Web：

```bash
dsh web
```

最后刷新浏览器页面。如果页面仍显示旧界面，请完全关闭旧的 DSH 进程后再启动一次。

### 升级 DSH

先查看当前 DSH 版本：

```bash
dsh --version
```

升级到 npm 上的最新版：

```bash
npm install --global @deepseek-ai/dsh@latest
```

再次确认升级结果：

```bash
dsh --version
```

升级 DSH 后重新启动 Web：

```bash
dsh web
```

如果升级 DSH 后出现插件兼容或页面加载问题，按“升级 WaLiOffice 插件到最新版”中的步骤重新安装插件。

### 回退到指定插件版本

第 1 步，删除当前版本：

```bash
dsh plugin --profile web remove walioffice-dsh-plugin
```

第 2 步，安装需要回退的版本：

```bash
dsh plugin --profile web add walioffice-dsh-plugin@<版本号>
```

第 3 步，重启 DSH Web：

```bash
dsh web
```

## 卸载

只卸载 WaLiOffice 插件，不会卸载 DSH：

```bash
dsh plugin --profile web remove walioffice-dsh-plugin
```

卸载后重新启动 DSH Web：

```bash
dsh web
```

如果还要卸载 DSH 本身，执行：

```bash
npm uninstall --global @deepseek-ai/dsh
```

## 常见问题

### 安装后页面没有 WaLiOffice

确认插件安装在 `web` profile，并完整重启 `dsh web`。先检查插件：

```bash
dsh plugin --profile web list
```

如果列表中没有插件，重新安装：

```bash
dsh plugin --profile web add walioffice-dsh-plugin@latest
```

然后重启：

```bash
dsh web
```

### 安装成功但工具不可用

先确认 DSH 的 LLM provider 已配置，再查看 DSH 日志。图片和视频工具还需要对应的 API 地址与密钥；基础文档、表格和演示工具不需要 Agnes API。

### 更新后仍然是旧版本

先查看 npm registry 上的最新版：

```bash
npm view walioffice-dsh-plugin version dist-tags.latest --registry=https://registry.npmjs.org/
```

然后按顺序删除旧版、安装明确版本并重启：

```bash
dsh plugin --profile web remove walioffice-dsh-plugin
```

```bash
dsh plugin --profile web add walioffice-dsh-plugin@<版本号>
```

```bash
dsh web
```

### 生成文件在哪里

先检查启动 DSH 时的当前工作目录：

```bash
pwd
```

再检查 `output/`：

```bash
ls -la output
```

插件不会把产物写入 npm 全局安装目录。

## 本地开发

安装源码依赖：

```bash
pnpm install
```

执行类型检查：

```bash
pnpm run typecheck
```

构建插件 bundle：

```bash
pnpm run bundle
```

检查 npm 打包内容：

```bash
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

先生成本地 tarball：

```bash
pnpm pack
```

再将下面的路径替换为实际生成的 tarball 路径：

```bash
dsh plugin --profile web add /absolute/path/to/walioffice-dsh-plugin-<版本号>.tgz
```

最后启动 DSH Web：

```bash
dsh web
```

发布前检查：

```bash
pnpm run typecheck
```

```bash
pnpm run bundle
```

```bash
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
