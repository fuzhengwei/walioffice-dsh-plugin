/**
 * Image generation tool: image_prompt
 * Uses Agnes Image 2.1 Flash API for text-to-image and image-to-image.
 * 
 * @module @walioffice/dsh-tool-image
 */

import type { Context } from '@deepseek-ai/cordis'
import { defineTool } from '@deepseek-ai/dsh-tools'
import { llmGenerateJson, resolveOfficeService } from '@walioffice/dsh-office'

// ── Types ───────────────────────────────────────────────────────────────────

interface ImagePromptPlan {
  styles: { title: string; prompt: string }[]
}

interface ImageResult {
  title: string
  description: string
  prompt: string
  image_size: string
  image_ratio: string
  generation_mode: string
  reference_image_count: number
  images: { url: string; style: string }[]
  provider: string
  model: string
}

// ── Config ──────────────────────────────────────────────────────────────────

interface ImageConfig {
  baseUrl: string
  apiKeys: string[]
  model: string
}

function resolveImageConfig(): ImageConfig {
  const baseUrl = process.env.AGNES_IMAGE_BASE_URL || process.env.LLM_IMAGE_BASE_URL || ''
  const apiKeys = (
    process.env.AGNES_IMAGE_API_KEYS?.split(',').map(k => k.trim()).filter(Boolean) ||
    process.env.LLM_IMAGE_API_KEYS?.split(',').map(k => k.trim()).filter(Boolean) ||
    (process.env.LLM_IMAGE_API_KEY ? [process.env.LLM_IMAGE_API_KEY] : [])
  )
  const model = process.env.AGNES_IMAGE_MODEL || 'agnes-image-2.1-flash'

  if (!baseUrl || apiKeys.length === 0) {
    throw new Error(
      '图片生成需要配置环境变量：\n' +
      '  AGNES_IMAGE_BASE_URL (或 LLM_IMAGE_BASE_URL) — API 地址\n' +
      '  AGNES_IMAGE_API_KEYS (或 LLM_IMAGE_API_KEY) — API 密钥\n' +
      '可在 .env 文件或系统环境变量中设置。'
    )
  }

  return { baseUrl, apiKeys, model }
}

// ── Size inference ──────────────────────────────────────────────────────────

function inferImageSize(topic: string): { ratio: string; size: string } {
  const lower = topic.toLowerCase()
  if (/海报|封面|竖版|手机|小红书/.test(topic)) {
    return { ratio: '2:3', size: '1024x1536' }
  }
  if (/头像|logo|方图|icon|头像/.test(topic)) {
    return { ratio: '1:1', size: '1024x1024' }
  }
  return { ratio: '3:2', size: '1536x1024' }
}

// ── Scene guide ─────────────────────────────────────────────────────────────

function inferSceneGuide(topic: string): string {
  if (/产品|包装|场景|展示/.test(topic)) {
    return '商业视觉风格，注重产品质感和场景氛围'
  }
  if (/科技|数据|网络|系统/.test(topic)) {
    return '科技视觉风格，使用深色背景、发光效果、几何元素'
  }
  if (/海报|宣传|营销|活动/.test(topic)) {
    return '传播海报风格，色彩鲜明、视觉冲击力强、信息层次清晰'
  }
  return '高质商用风格，画面精致、构图专业、色彩和谐'
}

// ── API call ────────────────────────────────────────────────────────────────

async function callAgnesImage(
  config: ImageConfig,
  prompt: string,
  size: string,
  referenceImages: string[],
): Promise<string> {
  const url = `${config.baseUrl.replace(/\/$/, '')}/v1/images/generations`
  
  const body: Record<string, unknown> = {
    model: config.model,
    prompt,
    n: 1,
    size,
    response_format: 'url',
  }
  if (referenceImages.length > 0) {
    body.extra_body = { image: referenceImages }
  }

  let lastError = ''
  for (const key of config.apiKeys) {
    try {
      const resp = await fetch(url, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          'Authorization': `Bearer ${key}`,
        },
        body: JSON.stringify(body),
      })
      if (!resp.ok) {
        lastError = `HTTP ${resp.status}: ${await resp.text()}`
        if (resp.status === 401 || resp.status === 403) continue
        if (resp.status === 429) {
          await new Promise(r => setTimeout(r, 2000))
          continue
        }
        continue
      }
      const data = await resp.json() as { data: { url: string }[] }
      if (data.data?.[0]?.url) {
        return data.data[0].url
      }
      lastError = 'Response missing image URL'
    } catch (err) {
      lastError = err instanceof Error ? err.message : String(err)
    }
  }
  throw new Error(`Agnes Image API 调用失败：${lastError}`)
}

// ── System prompt for style planning ────────────────────────────────────────

const STYLE_SYSTEM_PROMPT = `你是视觉提示词专家。只输出严格 JSON，不要 markdown 代码块。
返回格式：
{
  "styles": [
    {"title": "风格名称", "prompt": "英文提示词，描述画面风格、构图、光影、色调等"}
  ]
}
生成 3 种不同风格的提示词，每个 prompt 为英文，描述具体可执行。`

// ── Tool ────────────────────────────────────────────────────────────────────

export function apply(ctx: Context): void {
  ctx.tools.register(defineTool({
    name: 'image_prompt',
    description: '生成高质量图片：基于 Agnes Image 2.1 Flash，支持文生图和图生图（有参考图片时自动切换为图生图模式）。返回可预览的图片链接。',
    parameters: {
      topic: { type: 'string', required: true, description: '图片需求描述' },
      image_urls: {
        type: 'array',
        items: { type: 'string' },
        description: '参考图片 URL（1-2张为 keyframe 模式，3+张为 reference 模式）',
      },
    },
    output: {
      schema: {
        type: 'object',
        additionalProperties: true,
        properties: {
          title: { type: 'string', required: true },
          images: { type: 'array', required: true },
          generationMode: { type: 'string', required: true },
          provider: { type: 'string', required: true },
        },
      },
      render: (_args, value) => [{
        type: 'text',
        text: `已生成 ${value.images.length} 张图片《${value.title}》`,
      }],
      presentationMeta: (_args, value) => ({ kind: 'image', ...value }),
    },
    isConcurrencySafe: () => false,
    async execute(args, exec) {
      const topic = args.topic
      if (!topic?.trim()) {
        throw new Error('topic 不能为空')
      }
      const office = resolveOfficeService(ctx)

      // Resolve reference images
      const referenceImages: string[] = []
      if (args.image_urls && Array.isArray(args.image_urls)) {
        referenceImages.push(...args.image_urls.filter(Boolean))
      }
      // Also check attachments
      const generationMode = referenceImages.length === 0 ? 'text_to_image' : 'image_to_image'

      // Config
      const config = resolveImageConfig()
      const { ratio, size } = inferImageSize(topic)
      const sceneGuide = inferSceneGuide(topic)

      office.emitProgress('running', '规划图片风格', '正在生成多种风格的提示词...')

      // Generate style prompts via LLM
      const styleUserPrompt = [
        `用户需求：${topic}`,
        `视觉风格指导：${sceneGuide}`,
        `图片比例：${ratio}`,
        '',
        '请生成 3 种不同风格的英文提示词。',
      ].join('\n')

      let styles: { title: string; prompt: string }[]
      try {
        const json = await llmGenerateJson(ctx, STYLE_SYSTEM_PROMPT, styleUserPrompt)
        const plan = json as unknown as ImagePromptPlan
        styles = plan.styles?.slice(0, 3) ?? []
        if (styles.length === 0) throw new Error('No styles returned')
      } catch {
        // Fallback
        styles = [{
          title: '默认风格',
          prompt: `${topic}, professional, high quality, detailed, ${ratio} aspect ratio`,
        }]
      }

      // Generate images
      office.emitProgress('running', '生成图片', `正在调用 Agnes Image 生成 ${styles.length} 张图片...`)

      const images: { url: string; style: string }[] = []
      for (const style of styles) {
        try {
          const imageUrl = await callAgnesImage(config, style.prompt, size, referenceImages)
          images.push({ url: imageUrl, style: style.title })
        } catch (err) {
          // Retry once
          await new Promise(r => setTimeout(r, 2000))
          try {
            const imageUrl = await callAgnesImage(config, style.prompt, size, referenceImages)
            images.push({ url: imageUrl, style: style.title })
          } catch {
            // Skip this style
          }
        }
      }

      if (images.length === 0) {
        throw new Error('所有风格的图片生成均失败，请检查 API 配置或稍后重试')
      }

      return {
        title: topic.slice(0, 30),
        images,
        generationMode,
        provider: 'agnes',
        model: config.model,
      }
    },
  }))
}

export const name = 'tool-image'
export const inject = ['tools', 'office']
