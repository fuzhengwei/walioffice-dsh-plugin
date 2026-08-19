/**
 * WaLiOffice service definitions for DSH plugins.
 * 
 * This package defines the capability seams that WaLiOffice tool plugins
 * consume. It declares:
 * - OfficeToolContext: shared context for tool execution (LLM, artifacts, progress)
 * - OfficeArtifact: structured output from office tools
 * - SceneInference: shared scene-detection logic
 * - ThemePalettes: PPT color themes
 * 
 * @module @walioffice/dsh-office
 */

import type { Context, Service } from '@deepseek-ai/cordis'
import { createUserMessage } from '@deepseek-ai/dsh-llm'
import type { ContentBlock, GenerateOptions, StreamChunk } from '@deepseek-ai/dsh-llm'
import type { JsonValue } from '@deepseek-ai/dsh-session'

// ── Artifact types ──────────────────────────────────────────────────────────

export interface OfficeArtifact {
  kind: string
  title: string
  content: JsonValue
}

// ── Tool context (passed through ToolExecution.agent) ───────────────────────

export interface OfficeToolContext {
  /** User ID for per-user config resolution */
  userId: string
  /** Preferred model override */
  preferredModel?: string
  /** Image attachments from the current message (data URLs) */
  attachments: ChatAttachment[]
  /** Prior artifacts from the current session (for cross-tool references) */
  priorArtifacts: OfficeArtifact[]
  /** Scratchpad for cross-tool data passing */
  scratchpad: Map<string, JsonValue>
}

export interface ChatAttachment {
  kind: string
  dataUrl?: string
  filename?: string
}

// ── Scene inference ─────────────────────────────────────────────────────────

export type SceneType =
  | 'product'
  | 'growth'
  | 'sales'
  | 'tech'
  | 'training'
  | 'project'
  | 'default'

const SCENE_KEYWORDS: Record<Exclude<SceneType, 'default'>, string[]> = {
  product: ['产品', '需求', 'prd', 'roadmap', '版本', '迭代', '用户故事', 'feature'],
  growth: ['运营', '增长', '拉新', '留存', '转化', '活动', 'campaign', 'gmv'],
  sales: ['销售', '客户', '商机', '渠道', '业绩', '回款', '签约', '线索'],
  tech: ['技术', '架构', '系统', '平台', '接口', '部署', '微服务', '数据库', 'agent', 'ai'],
  training: ['培训', '课程', '学习', '上手', '入门', '手册', '宣导', 'workshop'],
  project: ['项目', '排期', '里程碑', '实施', '交付', '风险', '计划'],
}

export function inferScene(topic: string): SceneType {
  const lower = topic.toLowerCase()
  for (const [scene, keywords] of Object.entries(SCENE_KEYWORDS)) {
    if (keywords.some(kw => lower.includes(kw))) {
      return scene as SceneType
    }
  }
  return 'default'
}

export function sceneGuide(scene: SceneType, domain: string): string {
  const guides: Record<SceneType, Record<string, string>> = {
    product: {
      ppt: '当前更像产品方案场景，优先组织为：背景与机会、用户与问题、方案设计、核心流程、版本规划、收益与风险。',
      doc: '当前更像产品方案/PRD 场景，请重点补足用户场景、功能设计、流程说明、优先级、边界和验收标准。',
      sheet: '当前更像产品管理场景，优先设计需求池、版本规划、优先级评估、验收清单等表格。',
      drawio: '当前更像产品方案场景，优先体现用户流程、功能模块、角色关系或版本路径。',
      md: '当前更像产品知识沉淀场景，请重点补足背景、用户场景、功能说明、流程、边界和 FAQ。',
    },
    growth: {
      ppt: '当前更像运营复盘/增长场景，优先组织为：目标与结果、核心指标、问题拆解、关键动作、复盘结论、下阶段计划。',
      doc: '当前更像运营分析/复盘场景，请重点补足指标口径、动作拆解、问题诊断、结论和下阶段策略。',
      sheet: '当前更像运营分析场景，优先设计指标明细、渠道效果、活动复盘、周报汇总等表格。',
      drawio: '当前更像运营流程场景，优先体现渠道、动作链路、指标漏斗和复盘闭环。',
      md: '当前更像运营复盘/方法论沉淀场景，请重点补足指标口径、动作拆解、案例和经验总结。',
    },
    sales: {
      ppt: '当前更像销售汇报场景，优先组织为：业绩概览、区域/客户分析、机会与风险、重点动作、预测与资源诉求。',
      doc: '当前更像销售方案/经营汇报场景，请重点补足业绩结构、客户分层、商机推进、风险点和资源诉求。',
      sheet: '当前更像销售管理场景，优先设计线索跟进、客户分层、商机漏斗、区域业绩等表格。',
      drawio: '当前更像销售流程/经营场景，优先体现线索到签约的阶段流转、角色协同和客户分层。',
      md: '当前更像销售资料整理场景，请重点补足客户画像、销售流程、关键话术、阶段策略和常见问题。',
    },
    tech: {
      ppt: '当前更像技术设计/架构汇报场景，优先组织为：建设背景、总体架构、模块分层、关键流程、稳定性与安全、落地计划。',
      doc: '当前更像技术设计文档场景，请重点补足架构分层、关键流程、接口边界、依赖项、稳定性和安全要求。',
      sheet: '当前更像技术项目场景，优先设计接口清单、服务台账、发布计划、风险清单或测试追踪表。',
      drawio: '当前更像技术架构场景，优先体现系统边界、模块层次、调用链路、数据流和基础设施。',
      md: '当前更像技术文档/README 场景，请重点补足架构说明、目录结构、安装步骤、配置示例、调用示例和排错说明。',
    },
    training: {
      ppt: '当前更像培训课件场景，优先组织为：学习目标、核心概念、方法步骤、案例演示、常见误区、行动建议。',
      doc: '当前更像培训手册/课程文档场景，请重点补足学习目标、章节安排、案例说明、常见问题和实践建议。',
      sheet: '当前更像培训管理场景，优先设计课程安排、签到成绩、练习任务、反馈汇总等表格。',
      drawio: '当前更像培训流程场景，优先体现学习路径、步骤、角色分工和知识结构。',
      md: '当前更像培训讲义/操作手册场景，请重点补足学习路径、步骤、示例、练习建议和常见误区。',
    },
    project: {
      ppt: '当前更像项目实施/计划汇报场景，优先组织为：目标范围、阶段计划、里程碑、角色分工、风险依赖、验收标准。',
      doc: '当前更像项目实施/交付文档场景，请重点补足阶段任务、责任分工、里程碑、风险依赖和验收方式。',
      sheet: '当前更像项目管理场景，优先设计排期、里程碑、责任分工、风险跟踪和验收清单等表格。',
      drawio: '当前更像项目实施场景，优先体现阶段、责任人、交付物和依赖关系。',
      md: '当前更像项目执行手册场景，请重点补足实施阶段、责任分工、里程碑、依赖项和验收标准。',
    },
    default: {
      ppt: '未识别到强场景时，默认按商务化正式汇报组织，兼顾背景、分析、方案、价值与下一步。',
      doc: '默认按正式商务文档处理，兼顾背景、问题、方案、价值、风险与下一步。',
      sheet: '默认按真实业务表格处理，兼顾明细、汇总、分析字段和执行字段。',
      drawio: '默认按方案汇报图处理，兼顾层次、关系和对外讲解的清晰度。',
      md: '默认按正式知识文档处理，兼顾背景、核心说明、示例、FAQ 和后续建议。',
    },
  }
  return guides[scene][domain] ?? guides.default[domain] ?? ''
}

// ── PPT theme palettes ──────────────────────────────────────────────────────

export interface Palette {
  name: string
  bg: string
  card: string
  primary: string
  accent: string
  dark: string
}

export function getThemePalettes(theme: string): Palette[] {
  const themes: Record<string, Palette[]> = {
    business: [
      { name: 'navy', bg: 'F8FAFC', card: 'FFFFFF', primary: '1E3A5F', accent: '3B82F6', dark: '1E293B' },
      { name: 'steel', bg: 'F1F5F9', card: 'FFFFFF', primary: '334155', accent: '64748B', dark: '0F172A' },
    ],
    tech: [
      { name: 'cyber', bg: '0F172A', card: '1E293B', primary: '06B6D4', accent: '8B5CF6', dark: 'F1F5F9' },
      { name: 'matrix', bg: '0C4A1E', card: '14532D', primary: '22C55E', accent: 'FCD34D', dark: 'F0FDF4' },
    ],
    warm: [
      { name: 'sunset', bg: 'FEF3C7', card: 'FFFFFF', primary: 'EA580C', accent: 'F59E0B', dark: '7C2D12' },
      { name: 'rose', bg: 'FDF2F8', card: 'FFFFFF', primary: 'BE185D', accent: 'EC4899', dark: '831843' },
    ],
    minimal: [
      { name: 'mono', bg: 'FFFFFF', card: 'F8FAFC', primary: '1E293B', accent: '64748B', dark: '0F172A' },
      { name: 'paper', bg: 'FAFAF9', card: 'FFFFFF', primary: '44403C', accent: 'A8A29E', dark: '292524' },
    ],
    default: [
      { name: 'classic', bg: 'F8FAFC', card: 'FFFFFF', primary: '2563EB', accent: '3B82F6', dark: '1E293B' },
      { name: 'ocean', bg: 'EFF6FF', card: 'FFFFFF', primary: '1D4ED8', accent: '60A5FA', dark: '1E3A5F' },
    ],
  }
  return themes[theme] ?? themes['default']!
}

// ── LLM helper ──────────────────────────────────────────────────────────────

/**
 * Call the DSH LLM service and extract JSON from the response.
 * Uses ctx.llm.stream() — the standard DSH LLM calling convention.
 *
 * Provider/model resolution order:
 * 1. WALIOFFICE_LLM_PROVIDER / WALIOFFICE_LLM_MODEL env vars
 * 2. DSH_LLM_PROVIDER / DSH_LLM_MODEL env vars
 * 3. DSH agent-loop config (if available via ctx.get('agent'))
 * 4. Fallback: 'deepseek' / 'deepseek-chat'
 */
export async function llmGenerateJson(
  ctx: Context,
  systemPrompt: string,
  userPrompt: string,
): Promise<JsonValue> {
  const llm = ctx.get('llm')
  if (!llm) {
    throw new Error(
      'LLM 服务未加载。请在 cordis.yml 中配置 @deepseek-ai/dsh-llm 插件，\n' +
      '并设置 WALIOFFICE_LLM_PROVIDER 和 WALIOFFICE_LLM_MODEL 环境变量。'
    )
  }

  const provider = process.env.WALIOFFICE_LLM_PROVIDER
    ?? process.env.DSH_LLM_PROVIDER
    ?? 'deepseek'
  const model = process.env.WALIOFFICE_LLM_MODEL
    ?? process.env.DSH_LLM_MODEL
    ?? 'deepseek-chat'

  const options: GenerateOptions = {
    provider,
    model,
    system: systemPrompt,
    messages: [
      createUserMessage({
        content: [{ type: 'text', text: userPrompt }],
        source: { kind: 'user' },
      }),
    ],
  }

  // Collect text from the stream
  const text = await collectText(llm.stream(options))
  return extractJson(text)
}

/**
 * Collect all text deltas from an LLM stream into a single string.
 */
async function collectText(stream: AsyncIterable<StreamChunk>): Promise<string> {
  let text = ''
  for await (const chunk of stream) {
    if (chunk.type === 'text-delta') {
      text += chunk.text
    } else if (chunk.type === 'finish') {
      break
    }
  }
  return text
}

/**
 * Call the DSH LLM service and return raw text (non-JSON responses).
 */
export async function llmGenerateText(
  ctx: Context,
  systemPrompt: string,
  userPrompt: string,
): Promise<string> {
  const llm = ctx.get('llm')
  if (!llm) {
    throw new Error(
      'LLM 服务未加载。请在 cordis.yml 中配置 @deepseek-ai/dsh-llm 插件，\n' +
      '并设置 WALIOFFICE_LLM_PROVIDER 和 WALIOFFICE_LLM_MODEL 环境变量。'
    )
  }

  const provider = process.env.WALIOFFICE_LLM_PROVIDER
    ?? process.env.DSH_LLM_PROVIDER
    ?? 'deepseek'
  const model = process.env.WALIOFFICE_LLM_MODEL
    ?? process.env.DSH_LLM_MODEL
    ?? 'deepseek-chat'

  const options: GenerateOptions = {
    provider,
    model,
    system: systemPrompt,
    messages: [
      createUserMessage({
        content: [{ type: 'text', text: userPrompt }],
        source: { kind: 'user' },
      }),
    ],
  }

  return collectText(llm.stream(options))
}

/**
 * Extract JSON from LLM output (handles markdown fences, surrounding text, etc.)
 */
export function extractJson(text: string): JsonValue {
  // Try direct parse first
  try {
    return JSON.parse(text)
  } catch { /* continue */ }

  // Try to find JSON in markdown code block
  const fenceMatch = text.match(/```(?:json)?\s*([\s\S]*?)```/)
  if (fenceMatch?.[1]) {
    try {
      return JSON.parse(fenceMatch[1])
    } catch { /* continue */ }
  }

  // Try to find first { or [ and last } or ]
  const objStart = text.indexOf('{')
  const arrStart = text.indexOf('[')
  const start = objStart === -1 ? arrStart : arrStart === -1 ? objStart : Math.min(objStart, arrStart)
  if (start === -1) {
    throw new Error('No JSON found in LLM response')
  }
  const isObj = text[start] === '{'
  const end = isObj ? text.lastIndexOf('}') : text.lastIndexOf(']')
  if (end === -1 || end < start) {
    throw new Error('Incomplete JSON in LLM response')
  }
  const jsonStr = text.slice(start, end + 1)
  return JSON.parse(jsonStr)
}

// ── Office service declaration ──────────────────────────────────────────────

declare module '@deepseek-ai/cordis' {
  interface Context {
    office: OfficeService
  }
}

export interface OfficeService extends Service {
  /** Emit a progress event to the session */
  emitProgress(phase: string, step: string, detail: string): void
  /** Get the current user ID */
  getUserId(): string
  /** Read from scratchpad */
  getScratchpad(key: string): JsonValue | undefined
  /** Write to scratchpad */
  setScratchpad(key: string, value: JsonValue): void
  /** Get prior artifacts */
  getPriorArtifacts(): OfficeArtifact[]
}

// ── Cordis plugin: provides the 'office' service ─────────────────────────────

export const name = 'walioffice-office'
export const inject = ['llm']

import { Service as CordisService } from '@deepseek-ai/cordis'

class OfficeServiceImpl extends CordisService {
  private scratchpad = new Map<string, unknown>()
  private artifacts: OfficeArtifact[] = []

  constructor(ctx: any) {
    super(ctx, 'office')
  }

  emitProgress(phase: string, step: string, _detail: string): void {
    try { (this.ctx as any).emit?.('session/event', { type: 'office/progress', data: { phase, step, timestamp: Date.now() } }) } catch { /* best effort */ }
  }

  getUserId(): string {
    return process.env.USER || 'default'
  }

  getScratchpad(key: string): JsonValue | undefined {
    return this.scratchpad.get(key) as JsonValue | undefined
  }

  setScratchpad(key: string, value: JsonValue): void {
    this.scratchpad.set(key, value)
  }

  getPriorArtifacts(): OfficeArtifact[] {
    return this.artifacts
  }
}

export function apply(ctx: any): void {
  ctx.plugin(OfficeServiceImpl)
}
