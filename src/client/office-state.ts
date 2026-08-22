import type { ToolCallBlock } from '@deepseek-ai/dsh-client-runtime/client'

export type OfficeMode = 'all' | 'doc' | 'markdown' | 'sheet' | 'ppt' | 'chart' | 'drawio' | 'image' | 'video'

export interface OfficeModeDefinition {
  id: OfficeMode
  label: string
  shortLabel: string
  color: string
  prompt: string
}

export interface OfficeArtifact {
  id: string
  toolName: string
  mode: OfficeMode
  title: string
  summary: string
  filePath?: string
  output: string
  createdAt: number
  isError: boolean
  meta?: OfficeArtifactMeta
}

export interface OfficeDownload {
  fileName: string
  mimeType: string
  base64: string
}

export type OfficeArtifactMeta =
  | { kind: 'chart'; title: string; chartType: string; labels: string[]; values: number[]; seriesName: string; summary?: string }
  | {
    kind: 'doc'
    title: string
    markdown: string
    filePath?: string
    download?: OfficeDownload
    sectionCount?: number
    format?: string
    sections?: {
      heading: string
      headingLevel: number
      paragraphs: string[]
      bullets: string[]
      table?: { headers: string[]; rows: string[][] }
    }[]
  }
  | { kind: 'markdown'; title: string; markdown: string; filePath?: string; download?: OfficeDownload }
  | {
    kind: 'sheet'
    title?: string
    filePath?: string
    download?: OfficeDownload
    tableCount?: number
    totalRows?: number
    tables?: {
      title: string
      headers: string[]
      rows?: (string | number)[][]
      rowCount?: number
      summary?: string
    }[]
  }
  | { kind: 'ppt'; title: string; filePath?: string; download?: OfficeDownload; slideCount: number; slides?: { index: number; layout?: string; background?: string; title?: string; goal?: string; visual?: string; points?: string[]; elements?: { shape?: { x: number; y: number; w: number; h: number; fill?: string; shapeType?: string; line?: { color?: string; width?: number } }; text?: { content: string; x: number; y: number; w: number; h: number; fontSize?: number; color?: string; bold?: boolean; align?: string; valign?: string }; table?: { x: number; y: number; w: number; headers: string[]; rows: string[][] } }[] }[] }
  | { kind: 'drawio'; title: string; diagramType: string; xml: string; download?: OfficeDownload }
  | { kind: 'image'; title: string; images: { url: string; style: string }[]; generationMode?: string; provider?: string; model?: string }
  | { kind: 'video'; title: string; videoUrl: string; duration?: number; aspectRatio?: string; mode?: string; provider?: string; model?: string }
  | { kind: 'storyboard'; title: string; totalShots: number; totalSeconds: number; aspectRatio?: string; shots: { index?: number; title?: string; description?: string; prompt?: string; seconds?: number; mode?: string }[] }
  | { kind: 'generic'; [key: string]: unknown }

export const OFFICE_MODES: readonly OfficeModeDefinition[] = [
  {
    id: 'all',
    label: '综合办公',
    shortLabel: '综合',
    color: '#111827',
    prompt: '请根据需求主动选择并调用最合适的办公工具：Word/报告/PRD 用 doc_generate；Markdown/README/知识库/操作手册用 md_generate；Excel/表格/预算/排期用 sheet_generate；PPT/演示/汇报先 ppt_plan 再 ppt_generate；趋势/对比/占比/排名/漏斗/ECharts 用 chart_generate；流程/架构/泳道/ER/思维导图用 drawio_generate；生成图片/海报/插画用 image_prompt；复杂视频先 video_storyboard，再按需 video_generate。除非用户明确要求文件，不要为了凑产物调用多个工具。\n用户需求：',
  },
  { id: 'doc', label: 'Word 文档', shortLabel: 'Word', color: '#16a34a', prompt: '请只使用 doc_generate（不要使用 md_generate）生成一份专业 Word 文档（.docx）。根据需求自动选择报告、计划、总结、文章或 PRD 结构：\n用户需求：' },
  { id: 'markdown', label: 'Markdown 文档', shortLabel: 'MD', color: '#475569', prompt: '请只使用 md_generate 生成一份可直接保存的 Markdown 文档（.md），适合 README、知识库、说明文档、操作手册、会议纪要或调研整理：\n用户需求：' },
  { id: 'sheet', label: 'Excel 表格', shortLabel: 'Excel', color: '#059669', prompt: '请使用 sheet_generate 生成一份结构化 Excel 表格（.xlsx），自动设计合理字段、表头和示例数据：\n用户需求：' },
  { id: 'ppt', label: 'PPT 演示', shortLabel: 'PPT', color: '#2563eb', prompt: '请先调用 ppt_plan 规划大纲，再调用 ppt_generate 生成可下载的完整 PPT（.pptx）。如果没有现成大纲，ppt_generate 可以自动补齐规划：\n用户需求：' },
  { id: 'chart', label: '数据图表', shortLabel: '图表', color: '#7c3aed', prompt: '请使用 chart_generate 生成一份可直接渲染的 ECharts 图表；趋势用 line，对比/排名用 bar，占比/分布用 pie，漏斗用 funnel，达成率用 gauge：\n用户需求：' },
  { id: 'drawio', label: 'Draw.io 图', shortLabel: 'Draw.io', color: '#ea580c', prompt: '请使用 drawio_generate 生成一份可编辑的 draw.io 图表（流程图、架构图、泳道图、拓扑图、ER 图或思维导图）：\n用户需求：' },
  { id: 'image', label: 'AI 图片', shortLabel: '图像', color: '#db2777', prompt: '请使用 image_prompt 生成符合需求的图片、海报、封面、插画或视觉素材：\n用户需求：' },
  { id: 'video', label: 'AI 视频', shortLabel: '视频', color: '#e11d48', prompt: '请使用 video_generate 生成视频；如果需求包含多场景、多镜头或故事线，先使用 video_storyboard 规划分镜：\n用户需求：' },
]

export const OFFICE_TOOLS = [
  'ppt_plan',
  'ppt_generate',
  'doc_generate',
  'md_generate',
  'sheet_generate',
  'chart_generate',
  'drawio_generate',
  'image_prompt',
  'video_generate',
  'video_storyboard',
] as const

export const OFFICE_PANEL_EVENT = 'walioffice:panel'
export const OFFICE_ARTIFACT_EVENT = 'walioffice:artifact'

export function modeForTool(toolName: string): OfficeMode {
  if (toolName.startsWith('ppt_')) return 'ppt'
  if (toolName === 'doc_generate') return 'doc'
  if (toolName === 'md_generate') return 'markdown'
  if (toolName === 'sheet_generate') return 'sheet'
  if (toolName === 'chart_generate') return 'chart'
  if (toolName === 'drawio_generate') return 'drawio'
  if (toolName === 'image_prompt') return 'image'
  if (toolName.startsWith('video_')) return 'video'
  return 'all'
}

export function modeDefinition(mode: OfficeMode): OfficeModeDefinition {
  return OFFICE_MODES.find(item => item.id === mode) ?? OFFICE_MODES[0]!
}

export function openOfficePanel(artifactId?: string): void {
  window.dispatchEvent(new CustomEvent(OFFICE_PANEL_EVENT, { detail: { open: true, artifactId } }))
}

export function toggleOfficePanel(): void {
  window.dispatchEvent(new CustomEvent(OFFICE_PANEL_EVENT, { detail: { toggle: true } }))
}

export function publishArtifact(artifact: OfficeArtifact): void {
  window.dispatchEvent(new CustomEvent<OfficeArtifact>(OFFICE_ARTIFACT_EVENT, { detail: artifact }))
  if (!artifact.isError) openOfficePanel(artifact.id)
}

export function artifactFromTool(toolName: string, callId: string, block: ToolCallBlock): OfficeArtifact | null {
  if (!('kind' in block) || block.kind !== 'tool-result') return null
  const argsRaw = block.call?.argsRaw ?? ''
  const args = parseJsonRecord(argsRaw)
  const output = block.content
    .filter(item => item.type === 'text')
    .map(item => item.text)
    .join('\n')
    .trim()
  const filePath = findFilePath(output)
  const topic = typeof args?.topic === 'string' ? args.topic.trim() : ''
  const mode = modeForTool(toolName)
  const definition = modeDefinition(mode)
  const summary = firstUsefulLine(output) || (block.isError ? '生成失败' : '办公产物已生成')
  return {
    id: callId,
    toolName,
    mode,
    title: metaTitle(block.meta) || topic || fileName(filePath) || definition.label,
    summary: metaSummary(block.meta) || summary,
    filePath,
    output,
    createdAt: block.time,
    isError: block.isError,
    meta: readPresentationMeta(block.meta),
  }
}

function readPresentationMeta(value: unknown): OfficeArtifactMeta | undefined {
  if (typeof value !== 'object' || value === null || Array.isArray(value)) return undefined
  const record = value as Record<string, unknown>
  const kind = typeof record.kind === 'string' ? record.kind : 'generic'
  return { kind, ...record } as OfficeArtifactMeta
}

function metaTitle(value: unknown): string | undefined {
  return typeof (value as Record<string, unknown> | null)?.title === 'string'
    ? (value as Record<string, unknown>).title as string
    : undefined
}

function metaSummary(value: unknown): string | undefined {
  return typeof (value as Record<string, unknown> | null)?.summary === 'string'
    ? (value as Record<string, unknown>).summary as string
    : undefined
}

function parseJsonRecord(value: string): Record<string, unknown> | null {
  if (value.trim() === '') return null
  try {
    const parsed = JSON.parse(value)
    return typeof parsed === 'object' && parsed !== null && !Array.isArray(parsed)
      ? parsed as Record<string, unknown>
      : null
  } catch {
    return null
  }
}

function firstUsefulLine(value: string): string {
  return value.split('\n').map(line => line.trim()).find(Boolean)?.slice(0, 180) ?? ''
}

function findFilePath(value: string): string | undefined {
  const extension = '(?:pptx|docx|xlsx|md|drawio|xml)'
  const unix = value.match(new RegExp(`(/[^\\n\\r]+?\\.${extension})(?=\\s|$|[，。；,;])`, 'i'))?.[1]
  if (unix) return unix.trim()
  const windows = value.match(new RegExp(`([A-Za-z]:\\\\[^\\n\\r]+?\\.${extension})(?=\\s|$|[，。；,;])`, 'i'))?.[1]
  return windows?.trim()
}

function fileName(path: string | undefined): string | undefined {
  if (!path) return undefined
  return path.split(/[\\/]/).filter(Boolean).pop()
}
