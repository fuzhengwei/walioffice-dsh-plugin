/**
 * PPT Generate tool — second step: produce a .pptx file from a plan.
 * Reads the plan from scratchpad (set by ppt_plan) or generates one on the fly.
 * Uses pptxgenjs for rendering.
 * 
 * @module @walioffice/dsh-tool-ppt
 */

import type { Context } from '@deepseek-ai/cordis'
import { defineTool } from '@deepseek-ai/dsh-tools'
import { llmGenerateJson, sceneGuide, inferScene, getThemePalettes } from '@walioffice/dsh-office'
import type { PresentationPlan, SlidePlan } from './plan.ts'
import type { Palette } from '@walioffice/dsh-office'
import { renderPptx } from '@walioffice/dsh-office-render-pptx'

// ── Slide element model (mirrors Rust SlideElement) ─────────────────────────

export interface SlideElement {
  shape?: {
    x: number; y: number; w: number; h: number
    fill: string
    shapeType: 'rect' | 'roundRect' | 'ellipse' | 'line'
    line?: { color: string; width: number }
  }
  text?: {
    content: string
    x: number; y: number; w: number; h: number
    fontSize: number
    color: string
    bold?: boolean
    align?: 'left' | 'center' | 'right'
  }
  table?: {
    x: number; y: number; w: number
    headers: string[]
    rows: string[][]
  }
}

export interface Slide {
  layout: string
  elements: SlideElement[]
}

// ── Plan → Slide conversion ─────────────────────────────────────────────────

const SLIDE_W = 13.33 // inches (16:9)
const SLIDE_H = 7.5

function planToSlide(plan: SlidePlan, palette: Palette, index: number, total: number): Slide {
  const elements: SlideElement[] = []

  // Background
  elements.push({
    shape: {
      x: 0, y: 0, w: SLIDE_W, h: SLIDE_H,
      fill: palette.bg,
      shapeType: 'rect',
    },
  })

  if (index === 0 || plan.layout === 'title') {
    // Title slide
    elements.push({
      shape: {
        x: 0.8, y: 2.0, w: 0.15, h: 3.5,
        fill: palette.accent,
        shapeType: 'rect',
      },
    })
    elements.push({
      shape: {
        x: 1.2, y: 2.2, w: 10, h: 3,
        fill: palette.card,
        shapeType: 'roundRect',
      },
    })
    elements.push({
      text: {
        content: plan.title,
        x: 1.8, y: 2.8, w: 9, h: 1.5,
        fontSize: 36, color: palette.primary, bold: true, align: 'left',
      },
    })
    if (plan.points.length > 0 && plan.points[0]) {
      elements.push({
        text: {
          content: plan.points[0],
          x: 1.8, y: 4.3, w: 8, h: 0.8,
          fontSize: 18, color: palette.dark, align: 'left',
        },
      })
    }
  } else if (plan.layout === 'section') {
    // Section divider
    elements.push({
      shape: {
        x: 0, y: 3.0, w: SLIDE_W, h: 2,
        fill: palette.primary,
        shapeType: 'rect',
      },
    })
    elements.push({
      text: {
        content: plan.title,
        x: 1, y: 3.3, w: 11, h: 1.4,
        fontSize: 32, color: palette.bg, bold: true, align: 'center',
      },
    })
    if (plan.goal) {
      elements.push({
        text: {
          content: plan.goal,
          x: 1, y: 4.7, w: 11, h: 0.6,
          fontSize: 16, color: palette.accent, align: 'center',
        },
      })
    }
  } else {
    // Content slide
    elements.push({
      shape: {
        x: 0.8, y: 0.4, w: 0.12, h: 0.8,
        fill: palette.accent,
        shapeType: 'rect',
      },
    })
    elements.push({
      text: {
        content: plan.title,
        x: 1.1, y: 0.4, w: 11, h: 0.8,
        fontSize: 28, color: palette.primary, bold: true, align: 'left',
      },
    })
    elements.push({
      shape: {
        x: 0.8, y: 1.4, w: 11.7, h: 5.5,
        fill: palette.card,
        shapeType: 'roundRect',
      },
    })
    const points = plan.points.slice(0, 4)
    const startY = 1.8
    const lineH = 1.1
    points.forEach((pt, i) => {
      elements.push({
        text: {
          content: `• ${pt}`,
          x: 1.3, y: startY + i * lineH, w: 10.5, h: lineH,
          fontSize: 18, color: palette.dark, align: 'left',
        },
      })
    })
    if (plan.goal) {
      elements.push({
        text: {
          content: plan.goal,
          x: 1.3, y: 6.3, w: 10.5, h: 0.5,
          fontSize: 13, color: palette.accent, align: 'left',
        },
      })
    }
  }

  // Page number
  elements.push({
    text: {
      content: `${index + 1} / ${total}`,
      x: 11.8, y: 7.0, w: 1.2, h: 0.3,
      fontSize: 10, color: palette.dark, align: 'right',
    },
  })

  return { layout: plan.layout, elements }
}

// ── Fallback plan ───────────────────────────────────────────────────────────

function fallbackPlan(topic: string): PresentationPlan {
  const title = topic.slice(0, 30)
  return {
    title,
    slides: [
      { title, layout: 'title', goal: '封面', visual: '封面页', points: [topic] },
      { title: '核心内容', layout: 'content', goal: '展开核心要点', visual: '要点列表', points: ['要点一', '要点二', '要点三'] },
      { title: '总结', layout: 'content', goal: '总结与建议', visual: '结论页', points: ['关键结论', '后续行动'] },
    ],
  }
}

// ── Generate plan via LLM if not in scratchpad ──────────────────────────────

const PLAN_SYSTEM_PROMPT = `你是资深演示文稿策划。只输出严格 JSON，不要 markdown 代码块。
返回格式：{"title":"...","slides":[{"title":"...","layout":"title|content|section","goal":"...","visual":"...","points":["..."]}]}
要求：5-10页，首页封面末页结论，每页 points 2-4 条。`

async function generatePlanWithLlm(ctx: Context, topic: string): Promise<PresentationPlan> {
  const scene = inferScene(topic)
  const guide = sceneGuide(scene, 'ppt')
  const userPrompt = `标题：${topic}\n场景偏好：${guide}\n请规划一份完整的 PPT 大纲。`
  try {
    const json = await llmGenerateJson(ctx, PLAN_SYSTEM_PROMPT, userPrompt)
    return json as unknown as PresentationPlan
  } catch {
    return fallbackPlan(topic)
  }
}

// ── Tool definition ─────────────────────────────────────────────────────────

export function apply(ctx: Context): void {
  ctx.tools.register(defineTool({
    name: 'ppt_generate',
    description: '生成完整 PPT 项目（含视觉设计）：基于 ppt_plan 的大纲生成幻灯片，输出可直接下载的 .pptx 文件。使用前必须先调用 ppt_plan 规划大纲。',
    parameters: {
      topic: { type: 'string', required: true, description: 'PPT 主题' },
      theme: { type: 'string', description: '视觉主题：business/tech/warm/minimal（默认 business）' },
      audience: { type: 'string', description: '目标受众（可选）' },
    },
    output: {
      schema: {
        type: 'object',
        additionalProperties: true,
        properties: {
          title: { type: 'string', required: true },
          slideCount: { type: 'integer', required: true },
          filePath: { type: 'string', required: true },
          slides: { type: 'array' },
        },
      },
      render: (_args, value) => [{
        type: 'text',
        text: `已生成 PPT《${value.title}》，共 ${value.slideCount} 页，文件已保存到 ${value.filePath}`,
      }],
    },
    isConcurrencySafe: () => false,
    async execute(args, exec) {
      const topic = args.topic
      if (!topic?.trim()) {
        throw new Error('topic 不能为空')
      }
      const theme = args.theme ?? 'business'
      const office = ctx.office

      // Get plan from scratchpad or generate
      office.emitProgress('running', '读取 PPT 大纲', '正在获取规划方案...')
      let plan = office.getScratchpad('ppt_plan') as unknown as PresentationPlan | undefined
      if (!plan || !plan.slides || plan.slides.length === 0) {
        office.emitProgress('running', '生成 PPT 大纲', '未找到已有大纲，正在自动规划...')
        plan = await generatePlanWithLlm(ctx, topic)
      }

      const palettes = getThemePalettes(theme)
      const palette = palettes[0] ?? getThemePalettes('default')[0]!
      if (!palette) {
        throw new Error('No color palette available')
      }

      // Convert plan to slides
      office.emitProgress('running', '渲染幻灯片', `正在生成 ${plan.slides.length} 页幻灯片...`)
      const slides: Slide[] = plan.slides.map((sp, i) =>
        planToSlide(sp, palette, i, plan.slides.length)
      )

      // Render to PPTX
      office.emitProgress('running', '导出 PPTX', '正在生成 .pptx 文件...')
      const filePath = await renderPptx(ctx, {
        title: plan.title,
        slides,
        palette,
      } as { title: string; slides: Slide[]; palette: Palette })

      return {
        title: plan.title,
        slideCount: slides.length,
        filePath: filePath as string,
        slides: slides.map((s, i) => ({
          index: i + 1,
          layout: s.layout,
          elements: s.elements,
        })) as any,
      }
    },
  }))
}

export const name = 'tool-ppt-generate'
export const inject = ['tools', 'office']
