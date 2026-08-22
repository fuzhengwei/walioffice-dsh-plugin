/**
 * PPT Plan tool — first step of PPT generation.
 * Calls LLM to produce a structured outline, stores it in scratchpad.
 * 
 * @module @walioffice/dsh-tool-ppt
 */

import type { Context } from '@deepseek-ai/cordis'
import { defineTool } from '@deepseek-ai/dsh-tools'
import type { ToolResult } from '@deepseek-ai/dsh-tools'
import { inferScene, llmGenerateJson, resolveOfficeService, sceneGuide } from '@walioffice/dsh-office'
import type { Palette } from '@walioffice/dsh-office'

// ── Types ───────────────────────────────────────────────────────────────────

export interface SlidePlan {
  title: string
  layout: 'title' | 'content' | 'section' | 'two-column'
  goal: string
  visual: string
  points: string[]
  [key: string]: string | string[]
}

export interface PresentationPlan {
  title: string
  slides: SlidePlan[]
}

function compactPptTitle(topic: string): string {
  const normalized = topic.replace(/\s+/g, ' ').trim()
  const head = normalized.split(/[：:。！？]/)[0]?.trim() || normalized
  return head.slice(0, 36) || 'WaLiOffice 演示文稿'
}

function cleanSlideText(value: unknown, fallback: string): string {
  return typeof value === 'string' && value.trim() ? value.replace(/\s+/g, ' ').trim() : fallback
}

function isPlaceholderTitle(value: string): boolean {
  return ['title', 'section', 'content', 'slide', 'slide title', '页标题', '章节标题', '页面标题'].includes(value.trim().toLowerCase())
}

export function fallbackPlan(topic: string): PresentationPlan {
  const title = compactPptTitle(topic)
  return {
    title,
    slides: [
      { title, layout: 'title', goal: '明确汇报主题与核心目标', visual: '封面页', points: [`围绕“${title}”展开结构化汇报`] },
      { title: '背景与目标', layout: 'section', goal: '说明为什么做、要解决什么问题', visual: '背景与目标页', points: ['明确现状与业务背景', '统一本次汇报的目标和范围'] },
      { title: '核心问题与判断', layout: 'content', goal: '提炼影响结果的关键问题', visual: '问题卡片与重点数据', points: ['梳理当前最重要的矛盾与挑战', '区分事实、影响和需要决策的事项', '建立后续方案的判断依据'] },
      { title: '方案框架与重点动作', layout: 'content', goal: '展示整体方案和执行重点', visual: '方案框架图与任务列表', points: ['按优先级拆分核心工作模块', '明确角色分工、资源和交付结果', '形成可落地的协作机制'] },
      { title: '实施路径与保障', layout: 'content', goal: '说明推进节奏和风险控制', visual: '时间轴与风险矩阵', points: ['采用准备、执行、检查、优化的推进节奏', '设置关键节点、验收标准和反馈机制', '提前准备风险应对和资源备选方案'] },
      { title: '结论与下一步', layout: 'content', goal: '收束结论并明确行动安排', visual: '结论页与行动清单', points: ['确认核心结论和优先级', '落实负责人、时间节点和交付标准', '安排首轮检查并持续复盘优化'] },
    ],
  }
}

export function normalizePresentationPlan(value: unknown, topic: string): PresentationPlan {
  if (typeof value !== 'object' || value === null) throw new Error('LLM 返回的 PPT 大纲格式不正确')
  const source = value as { title?: unknown; slides?: unknown }
  if (!Array.isArray(source.slides)) throw new Error('LLM 返回的 PPT 大纲格式不正确')
  if (source.slides.some(item => {
    const slide = item && typeof item === 'object' ? item as Record<string, unknown> : {}
    return isPlaceholderTitle(cleanSlideText(slide.title, ''))
  })) throw new Error('LLM 返回了占位 PPT 大纲')

  const slides = source.slides.map((item, index): SlidePlan => {
    const slide = item && typeof item === 'object' ? item as Record<string, unknown> : {}
    const rawLayout = cleanSlideText(slide.layout, index === 0 ? 'title' : 'content')
    const layout = rawLayout === 'title' || rawLayout === 'section' || rawLayout === 'two-column' ? rawLayout : 'content'
    const rawPoints = Array.isArray(slide.points) ? slide.points.map(item => cleanSlideText(item, '')).filter(Boolean) : []
    return {
      title: cleanSlideText(slide.title, `第 ${index + 1} 页`),
      layout,
      goal: cleanSlideText(slide.goal, '说明本页的核心信息'),
      visual: cleanSlideText(slide.visual, layout === 'title' ? '封面页' : '要点列表'),
      points: rawPoints.length > 0 ? rawPoints.slice(0, 4) : ['补充本页核心信息', '说明关键依据或行动建议'],
    }
  }).filter(slide => slide.title.trim())

  if (slides.length < 3) throw new Error('LLM 返回的 PPT 大纲页数不足')
  const rawTitle = cleanSlideText(source.title, '')
  if (isPlaceholderTitle(rawTitle)) throw new Error('LLM 返回了占位 PPT 标题')
  return {
    title: rawTitle && rawTitle.length <= 48 && !rawTitle.includes(topic.trim()) ? rawTitle : compactPptTitle(topic),
    slides: slides.slice(0, 10),
  }
}

// ── Tool definition ─────────────────────────────────────────────────────────

const SYSTEM_PROMPT = `你是资深演示文稿策划。请规划一份可以直接拿去汇报的 PPT 大纲。
只返回 JSON：{"title":"...","slides":[{"title":"...","layout":"title|content|section","goal":"...","visual":"...","points":["..."]}]}
要求：
- 5-10页，结构完整，有封面和结尾页
- 页面顺序要有叙事感，不要东拼西凑
- 标题必须具体，避免空话
- 每页 points 2-4 条，适合上屏展示
- visual 要写明页面版式：如封面、数据卡片、流程、对比、时间轴、结论页
- 如用户没说清楚，就按专业商务汇报默认处理`

export function apply(ctx: Context): void {
  ctx.tools.register(defineTool({
    name: 'ppt_plan',
    description: '规划 PPT 大纲：根据主题和受众规划幻灯片结构（标题、布局、要点）。这是 PPT 生成的第一步，只产出规划，不生成最终幻灯片。生成 PPT 时必须先调用此工具，再调用 ppt_generate。',
    parameters: {
      topic: { type: 'string', required: true, description: 'PPT 主题/用户需求' },
      audience: { type: 'string', description: '目标受众（可选）' },
    },
    output: {
      schema: {
        type: 'object',
        additionalProperties: true,
        properties: {
          title: { type: 'string', required: true },
          slideCount: { type: 'integer', required: true },
          slides: {
            type: 'array',
            required: true,
            items: {
              type: 'object',
              additionalProperties: true,
              properties: {
                title: { type: 'string', required: true },
                layout: { type: 'string', required: true },
                goal: { type: 'string', required: true },
                visual: { type: 'string' },
                points: { type: 'array', items: { type: 'string' } },
              },
            },
          },
        },
      },
      render: (_args, value) => [{
        type: 'text',
        text: `已规划 PPT 大纲《${value.title}》，共 ${value.slideCount} 页`,
      }],
      presentationMeta: (_args, value) => ({ kind: 'ppt', ...value }),
    },
    isConcurrencySafe: () => false,
    async execute(args, exec) {
      const topic = args.topic
      if (!topic?.trim()) {
        throw new Error('topic 不能为空')
      }
      const audience = args.audience ?? ''
      const scene = inferScene(topic)
      const guide = sceneGuide(scene, 'ppt')

      const userPrompt = [
        `标题：${topic}`,
        audience ? `受众：${audience}` : '',
        `场景偏好：${guide}`,
        '',
        '请规划一份完整的 PPT 大纲，只返回 JSON。',
      ].filter(Boolean).join('\n')

      const office = resolveOfficeService(ctx)
      office.emitProgress('running', '规划 PPT 大纲', `正在为《${topic}》规划大纲...`)

      let plan: PresentationPlan
      try {
        const json = await llmGenerateJson(ctx, SYSTEM_PROMPT, userPrompt)
        plan = normalizePresentationPlan(json, topic)
      } catch (err) {
        office.emitProgress('running', '使用 PPT 模板兜底', 'LLM 未返回可解析大纲，已使用内置 6 页结构继续生成。')
        plan = fallbackPlan(topic)
        void err
      }

      // Store in scratchpad for ppt_generate
      office.setScratchpad('ppt_plan', plan as any)

      return {
        title: plan.title,
        slideCount: plan.slides.length,
        slides: plan.slides as any,
      }
    },
  }))
}

export const name = 'tool-ppt-plan'
export const inject = ['tools', 'office']

// Re-export for generate.ts convenience
export type { Palette } from '@walioffice/dsh-office'
