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

      const json = await llmGenerateJson(ctx, SYSTEM_PROMPT, userPrompt)
      const plan = json as unknown as PresentationPlan

      // Validate
      if (!plan.title || !Array.isArray(plan.slides) || plan.slides.length === 0) {
        throw new Error('LLM 返回的 PPT 大纲格式不正确')
      }

      // Store in scratchpad for ppt_generate
      office.setScratchpad('ppt_plan', json)

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
