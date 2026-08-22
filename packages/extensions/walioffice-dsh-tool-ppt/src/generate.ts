/**
 * PPT Generate tool — second step: produce a .pptx file from a plan.
 * Reads the plan from scratchpad (set by ppt_plan) or generates one on the fly.
 * Uses pptxgenjs for rendering.
 * 
 * @module @walioffice/dsh-tool-ppt
 */

import type { Context } from '@deepseek-ai/cordis'
import { defineTool } from '@deepseek-ai/dsh-tools'
import { createOfficeDownload, getThemePalettes, inferScene, llmGenerateJson, resolveOfficeService, sceneGuide } from '@walioffice/dsh-office'
import { fallbackPlan, normalizePresentationPlan, type PresentationPlan, type SlidePlan } from './plan.ts'
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
    valign?: 'top' | 'middle' | 'bottom'
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
  background?: string
  title?: string
  goal?: string
  visual?: string
  points?: string[]
}

// ── Plan → Slide conversion ─────────────────────────────────────────────────

const SLIDE_W = 13.33 // inches (16:9)
const SLIDE_H = 7.5

function hexToRgb(color: string): [number, number, number] | null {
  const value = color.replace('#', '').trim()
  if (!/^[0-9a-f]{6}$/i.test(value)) return null
  return [Number.parseInt(value.slice(0, 2), 16), Number.parseInt(value.slice(2, 4), 16), Number.parseInt(value.slice(4, 6), 16)]
}

function relativeLuminance(color: string): number {
  const rgb = hexToRgb(color)
  if (!rgb) return 1
  const channel = (value: number) => {
    const scaled = value / 255
    return scaled <= 0.03928 ? scaled / 12.92 : Math.pow((scaled + 0.055) / 1.055, 2.4)
  }
  return 0.2126 * channel(rgb[0]) + 0.7152 * channel(rgb[1]) + 0.0722 * channel(rgb[2])
}

function contrastRatio(foreground: string, background: string): number {
  const fg = relativeLuminance(foreground)
  const bg = relativeLuminance(background)
  return (Math.max(fg, bg) + 0.05) / (Math.min(fg, bg) + 0.05)
}

function readableTextColor(preferred: string, background: string): string {
  if (contrastRatio(preferred, background) >= 4.5) return preferred
  return contrastRatio('0F172A', background) >= contrastRatio('F8FAFC', background) ? '0F172A' : 'F8FAFC'
}

function isDarkPalette(palette: Palette): boolean {
  return relativeLuminance(palette.bg) < 0.25
}

function planToSlide(plan: SlidePlan, palette: Palette, index: number, total: number): Slide {
  const elements: SlideElement[] = []
  const isTitle = index === 0 || plan.layout === 'title'
  const isSection = plan.layout === 'section'
  const dark = isDarkPalette(palette)
  const points = plan.points.slice(0, 4)

  if (isTitle) {
    elements.push({ shape: { x: 0, y: 0, w: SLIDE_W, h: SLIDE_H, fill: palette.bg, shapeType: 'rect' } })
    if (dark) {
      elements.push({ shape: { x: 1.05, y: 1.25, w: 10.9, h: 2.15, fill: palette.card, shapeType: 'roundRect' } })
      elements.push({ shape: { x: 1.05, y: 1.55, w: 0.07, h: 1.35, fill: palette.accent, shapeType: 'rect' } })
      elements.push({ text: { content: plan.title, x: 1.5, y: 1.62, w: 9.7, h: 0.58, fontSize: 27, color: palette.primary, bold: true, align: 'left', valign: 'middle' } })
      elements.push({ text: { content: plan.goal || plan.visual || '现代简洁演示文稿', x: 1.5, y: 2.45, w: 9.3, h: 0.32, fontSize: 12, color: palette.dark, align: 'left' } })
    } else {
      elements.push({ shape: { x: 0.75, y: 0.7, w: 11.83, h: 6.1, fill: palette.card, shapeType: 'roundRect' } })
      elements.push({ shape: { x: 0.75, y: 0.7, w: 0.16, h: 6.1, fill: palette.primary, shapeType: 'roundRect' } })
      elements.push({ shape: { x: 10.5, y: 0.95, w: 1.5, h: 1.5, fill: palette.accent, shapeType: 'ellipse' } })
      elements.push({ text: { content: plan.title, x: 1.35, y: 2.25, w: 10.4, h: 0.95, fontSize: 44, color: readableTextColor(palette.dark, palette.card), bold: true, align: 'left' } })
      elements.push({ text: { content: plan.goal || plan.visual || '现代简洁演示文稿', x: 1.4, y: 3.35, w: 9.4, h: 0.55, fontSize: 22, color: readableTextColor('475569', palette.card), align: 'left' } })
    }
    elements.push({ text: { content: `${index + 1} / ${total}`, x: 11.75, y: 7.05, w: 1.05, h: 0.24, fontSize: 9, color: readableTextColor(palette.dark, palette.bg), align: 'right' } })
  } else if (isSection) {
    if (dark) {
      elements.push({ shape: { x: 0, y: 0, w: SLIDE_W, h: SLIDE_H, fill: palette.bg, shapeType: 'rect' } })
      elements.push({ shape: { x: 0, y: 2.72, w: SLIDE_W, h: 0.62, fill: palette.primary, shapeType: 'rect' } })
      elements.push({ text: { content: plan.title, x: 1, y: 2.85, w: 11.33, h: 0.3, fontSize: 20, color: readableTextColor(palette.dark, palette.primary), bold: true, align: 'center', valign: 'middle' } })
      elements.push({ text: { content: plan.goal || '', x: 1.05, y: 3.62, w: 10, h: 0.24, fontSize: 10, color: palette.accent, align: 'center' } })
    } else {
      elements.push({ shape: { x: 0, y: 0, w: SLIDE_W, h: SLIDE_H, fill: palette.primary, shapeType: 'rect' } })
      elements.push({ shape: { x: 8.5, y: -0.5, w: 4.83, h: 4, fill: palette.accent, shapeType: 'ellipse' } })
      elements.push({ text: { content: plan.title, x: 1, y: 3.0, w: 11.33, h: 1.5, fontSize: 40, color: readableTextColor(palette.dark, palette.primary), bold: true, align: 'left' } })
      elements.push({ text: { content: plan.goal || '', x: 1.05, y: 4.6, w: 10, h: 0.5, fontSize: 20, color: readableTextColor(palette.accent, palette.primary), align: 'left' } })
    }
  } else {
    elements.push({ shape: { x: 0, y: 0, w: SLIDE_W, h: SLIDE_H, fill: palette.bg, shapeType: 'rect' } })
    if (dark) {
      elements.push({ shape: { x: 0.8, y: 0.55, w: 0.1, h: 0.68, fill: palette.accent, shapeType: 'rect' } })
      elements.push({ text: { content: plan.title, x: 1.15, y: 0.55, w: 11.1, h: 0.68, fontSize: 28, color: palette.primary, bold: true, align: 'left', valign: 'middle' } })
      elements.push({ shape: { x: 0.8, y: 1.42, w: 11.72, h: 5.28, fill: palette.card, shapeType: 'roundRect' } })
      if (points.length > 0) {
        elements.push({ text: { content: points.map(point => `• ${point}`).join('\n\n'), x: 1.28, y: 2.0, w: 10.8, h: 3.7, fontSize: 18, color: palette.dark, align: 'left', valign: 'top' } })
      }
      if (plan.goal) elements.push({ text: { content: plan.goal, x: 1.28, y: 6.03, w: 10.8, h: 0.3, fontSize: 13, color: palette.accent, align: 'left' } })
    } else {
      elements.push({ shape: { x: 0.75, y: 0.7, w: 11.83, h: 6.1, fill: palette.card, shapeType: 'roundRect' } })
      elements.push({ shape: { x: 0.75, y: 0.7, w: 11.83, h: 0.08, fill: palette.primary, shapeType: 'rect' } })
      elements.push({ text: { content: plan.title, x: 1.1, y: 0.95, w: 11, h: 0.7, fontSize: 32, color: readableTextColor(palette.dark, palette.card), bold: true, align: 'left' } })
    const innerFill = palette.bg
    const innerText = readableTextColor(palette.dark, innerFill)
    const cardText = (point: string, x: number, y: number, w: number, h: number, index: number, fontSize = 16) => {
      elements.push({ shape: { x, y, w, h, fill: innerFill, shapeType: 'roundRect' } })
      elements.push({ shape: { x: x + 0.22, y: y + 0.22, w: 0.38, h: 0.38, fill: palette.accent, shapeType: 'ellipse' } })
      elements.push({ text: { content: String(index + 1).padStart(2, '0'), x: x + 0.22, y: y + 0.25, w: 0.38, h: 0.24, fontSize: 9, color: readableTextColor(palette.dark, palette.accent), bold: true, align: 'center', valign: 'middle' } })
      elements.push({ text: { content: point, x: x + 0.78, y: y + 0.23, w: w - 1.05, h: h - 0.42, fontSize, color: innerText, align: 'left', valign: 'middle' } })
    }
    if (points.length === 4) {
      cardText(points[0]!, 1.2, 1.95, 5.25, 1.7, 0)
      cardText(points[1]!, 6.85, 1.95, 5.25, 1.7, 1)
      cardText(points[2]!, 1.2, 3.95, 5.25, 1.7, 2)
      cardText(points[3]!, 6.85, 3.95, 5.25, 1.7, 3)
    } else if (points.length === 3) {
      cardText(points[0]!, 1.2, 2.05, 3.4, 2.75, 0, 15)
      cardText(points[1]!, 4.95, 2.05, 3.4, 2.75, 1, 15)
      cardText(points[2]!, 8.7, 2.05, 3.4, 2.75, 2, 15)
    } else if (points.length === 2) {
      cardText(points[0]!, 1.2, 2.15, 5.25, 2.65, 0, 17)
      cardText(points[1]!, 6.85, 2.15, 5.25, 2.65, 1, 17)
    } else if (points.length === 1) {
      cardText(points[0]!, 1.2, 2.15, 10.9, 2.65, 0, 19)
    }
    if (plan.goal) {
      elements.push({ text: { content: plan.goal, x: 1.2, y: 6.1, w: 10.9, h: 0.35, fontSize: 12, color: readableTextColor(palette.accent, palette.card), align: 'left' } })
    }
    }
  }

  // Page number
  if (!isTitle) elements.push({ text: { content: `${index + 1} / ${total}`, x: 11.8, y: 7.0, w: 1.2, h: 0.3, fontSize: 10, color: readableTextColor(palette.dark, palette.bg), align: 'right' } })

  return { layout: isTitle ? 'title' : isSection ? 'section' : 'content', background: isSection ? palette.primary : palette.bg, elements, title: plan.title, goal: plan.goal, visual: plan.visual, points: plan.points }
}

// ── Fallback plan ───────────────────────────────────────────────────────────

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
    return normalizePresentationPlan(json, topic)
  } catch {
    return fallbackPlan(topic)
  }
}

// ── Tool definition ─────────────────────────────────────────────────────────

export function apply(ctx: Context): void {
  ctx.tools.register(defineTool({
    name: 'ppt_generate',
    description: '生成可直接下载的完整 PPT（.pptx），适用于演示文稿、汇报、路演、培训课件和发布会材料。优先读取 ppt_plan 的大纲；如果用户直接要求生成 PPT 或没有现成大纲，本工具会自动补齐规划后继续生成，不会因缺少前置调用而失败。',
    parameters: {
      topic: { type: 'string', required: true, description: 'PPT 主题' },
      theme: { type: 'string', description: '视觉主题：tech/business/warm/minimal（默认 tech，深色科技风）' },
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
          download: { type: 'object', additionalProperties: true },
          slides: { type: 'array' },
        },
      },
      render: (_args, value) => [{
        type: 'text',
        text: `已生成 PPT《${value.title}》，共 ${value.slideCount} 页，文件已保存到 ${value.filePath}`,
      }],
      presentationMeta: (_args, value) => ({ kind: 'ppt', ...value }),
    },
    isConcurrencySafe: () => false,
    async execute(args, exec) {
      const topic = args.topic
      if (!topic?.trim()) {
        throw new Error('topic 不能为空')
      }
      const theme = args.theme ?? 'tech'
      const office = resolveOfficeService(ctx)

      // Get plan from scratchpad or generate
      office.emitProgress('running', '读取 PPT 大纲', '正在获取规划方案...')
      let plan: PresentationPlan
      try {
        plan = normalizePresentationPlan(office.getScratchpad('ppt_plan'), topic)
      } catch {
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
        planToSlide(sp, palettes[i % palettes.length] ?? palette, i, plan.slides.length)
      )

      // Render to PPTX
      office.emitProgress('running', '导出 PPTX', '正在生成 .pptx 文件...')
      const filePath = await renderPptx(ctx, {
        title: plan.title,
        slides,
        palette,
      } as { title: string; slides: Slide[]; palette: Palette })
      const download = await createOfficeDownload(filePath as string, 'application/vnd.openxmlformats-officedocument.presentationml.presentation')

      return {
        title: plan.title,
        slideCount: slides.length,
        filePath: filePath as string,
        download,
        slides: slides.map((s, i) => ({
          index: i + 1,
          layout: s.layout,
          background: s.background,
          title: s.title ?? plan.slides[i]?.title,
          goal: s.goal ?? plan.slides[i]?.goal,
          visual: s.visual ?? plan.slides[i]?.visual,
          points: s.points ?? plan.slides[i]?.points,
          elements: s.elements,
        })) as any,
      }
    },
  }))
}

export const name = 'tool-ppt-generate'
export const inject = ['tools', 'office']
