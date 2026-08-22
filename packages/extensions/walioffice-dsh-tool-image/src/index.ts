/**
 * Image generation tool: image_prompt
 * Uses Agnes Image 2.1 Flash API for text-to-image and image-to-image.
 * 
 * @module @walioffice/dsh-tool-image
 */

import type { Context } from '@deepseek-ai/cordis'
import { defineTool } from '@deepseek-ai/dsh-tools'
import { llmGenerateJson, resolveOfficeService } from '@walioffice/dsh-office'
import { readFileSync } from 'node:fs'
import { homedir } from 'node:os'
import { join } from 'node:path'

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

interface ImageProviderProfile {
  baseUrl: string
  apiKeys: string[]
  models: string[]
}

const IMAGE_MODEL = 'agnes-image-2.1-flash'

function resolveImageConfig(ctx: Context): ImageConfig {
  const explicitBaseUrl = process.env.AGNES_IMAGE_BASE_URL || process.env.LLM_IMAGE_BASE_URL || ''
  const explicitApiKeys = (
    process.env.AGNES_IMAGE_API_KEYS?.split(',').map(k => k.trim()).filter(Boolean) ||
    process.env.LLM_IMAGE_API_KEYS?.split(',').map(k => k.trim()).filter(Boolean) ||
    (process.env.LLM_IMAGE_API_KEY ? [process.env.LLM_IMAGE_API_KEY] : [])
  )
  const discovered = discoverImageProvider(ctx, IMAGE_MODEL)

  if (!discovered) {
    throw new Error(
      `未在 DSH 模型配置中找到图片模型 ${IMAGE_MODEL}。\n` +
      '请先申请该模型，然后在 ~/.dsh/settings.yaml 的 llm-pi-ai.providers.models 中配置它，再重试。'
    )
  }

  const baseUrl = explicitBaseUrl || discovered.baseUrl
  const apiKeys = explicitApiKeys.length > 0 ? explicitApiKeys : discovered.apiKeys

  if (!baseUrl || apiKeys.length === 0) {
    throw new Error(
      '图片生成需要配置环境变量：\n' +
      '  AGNES_IMAGE_BASE_URL (或 LLM_IMAGE_BASE_URL) — API 地址\n' +
      '  AGNES_IMAGE_API_KEYS (或 LLM_IMAGE_API_KEY) — API 密钥\n' +
      '也可以在 DSH 的 ~/.dsh/settings.yaml 配置图片模型，并在 ~/.dsh/.credentials.yaml 中提供 apiKeyEnv 对应的密钥。'
    )
  }

  return { baseUrl, apiKeys, model: IMAGE_MODEL }
}

function discoverImageProvider(ctx: Context, requestedModel: string): ImageProviderProfile | undefined {
  const settingsService = (ctx as unknown as { get?: (key: string) => unknown }).get?.('settings') as {
    get?: (namespace: string) => unknown
  } | undefined
  const serviceSettings = asRecord(settingsService?.get?.('llm-pi-ai'))
  const fileSettings = readDshSettings()
  const providers = mergeProviders(
    asRecord(fileSettings?.providers),
    asRecord(serviceSettings?.providers),
  )
  if (!providers) return undefined

  const credentials = readDshCredentials()
  const entries = Object.values(providers).map(value => readImageProviderProfile(value, credentials)).filter((item): item is ImageProviderProfile => item !== undefined)
  if (entries.length === 0) return undefined

  const requested = requestedModel.trim()
  if (requested) {
    const exact = entries.find(entry => entry.models.includes(requested))
    if (exact) return exact
  }
  return undefined
}

function readImageProviderProfile(value: unknown, credentials: Record<string, string>): ImageProviderProfile | undefined {
  const profile = asRecord(value)
  if (!profile) return undefined
  const baseUrl = typeof profile.baseURL === 'string'
    ? profile.baseURL.trim()
    : typeof profile.baseUrl === 'string'
      ? profile.baseUrl.trim()
      : ''
  const models = Array.isArray(profile.models)
    ? profile.models.map(item => {
      if (typeof item === 'string') return item.trim()
      return typeof asRecord(item)?.id === 'string' ? String(asRecord(item)?.id).trim() : ''
    }).filter(Boolean)
    : []
  const apiKeyEnv = typeof profile.apiKeyEnv === 'string' ? profile.apiKeyEnv.trim() : ''
  const apiKeysEnv = typeof profile.apiKeysEnv === 'string' ? profile.apiKeysEnv.trim() : ''
  const apiKeys = [...new Set(`${apiKeyEnv},${apiKeysEnv},AGNES_AI_API_KEY`
    .split(',')
    .map(item => item.trim())
    .filter(Boolean)
    .flatMap(name => (process.env[name] ?? credentials[name])?.split(',').map(key => key.trim()).filter(Boolean) ?? []))]
  if (!baseUrl && apiKeys.length === 0 && models.length === 0) return undefined
  return { baseUrl, apiKeys, models }
}

function asRecord(value: unknown): Record<string, unknown> | undefined {
  return typeof value === 'object' && value !== null && !Array.isArray(value)
    ? value as Record<string, unknown>
    : undefined
}

function mergeProviders(
  fileProviders: Record<string, unknown> | undefined,
  serviceProviders: Record<string, unknown> | undefined,
): Record<string, unknown> | undefined {
  const names = new Set([...Object.keys(fileProviders ?? {}), ...Object.keys(serviceProviders ?? {})])
  if (names.size === 0) return undefined
  return Object.fromEntries([...names].map(name => [
    name,
    { ...asRecord(fileProviders?.[name]), ...asRecord(serviceProviders?.[name]) },
  ]))
}

function readDshSettings(): Record<string, unknown> | undefined {
  try {
    const text = readFileSync(join(process.env.DSH_HOME || join(homedir(), '.dsh'), 'settings.yaml'), 'utf8')
    const providers: Record<string, Record<string, unknown>> = {}
    let inLlm = false
    let inProviders = false
    let current: Record<string, unknown> | undefined

    for (const rawLine of text.split(/\r?\n/)) {
      const line = rawLine.replace(/\s+#.*$/, '')
      const trimmed = line.trim()
      if (!trimmed) continue
      const indent = line.length - line.trimStart().length
      if (indent === 0) {
        inLlm = trimmed === 'llm-pi-ai:'
        inProviders = false
        current = undefined
        continue
      }
      if (!inLlm) continue
      if (indent === 2 && trimmed === 'providers:') {
        inProviders = true
        continue
      }
      if (!inProviders) continue
      if (indent === 4 && trimmed.endsWith(':')) {
        current = {}
        providers[trimmed.slice(0, -1).trim()] = current
        continue
      }
      if (!current) continue
      if (indent === 6 && trimmed.includes(':')) {
        const separator = trimmed.indexOf(':')
        const key = trimmed.slice(0, separator).trim()
        const value = trimmed.slice(separator + 1).trim()
        if (value) current[key] = unquote(value)
        continue
      }
      if (indent >= 8 && trimmed.startsWith('- id:')) {
        const model = trimmed.slice(5).trim()
        const models = Array.isArray(current.models) ? current.models as unknown[] : []
        models.push({ id: unquote(model) })
        current.models = models
      }
    }
    return { providers }
  } catch {
    return undefined
  }
}

function readDshCredentials(): Record<string, string> {
  try {
    const text = readFileSync(join(process.env.DSH_HOME || join(homedir(), '.dsh'), '.credentials.yaml'), 'utf8')
    return Object.fromEntries(text.split(/\r?\n/).flatMap(line => {
      const match = line.match(/^([A-Za-z_][A-Za-z0-9_]*)\s*:\s*(.*)$/)
      return match ? [[match[1]!, unquote(match[2]!)] as const] : []
    }))
  } catch {
    return {}
  }
}

function unquote(value: string): string {
  return value.replace(/^(['"])(.*)\1$/, '$2').trim()
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
  const baseUrl = config.baseUrl.replace(/\/+$/, '')
  const url = `${baseUrl.endsWith('/v1') ? baseUrl : `${baseUrl}/v1`}/images/generations`
  
  const body: Record<string, unknown> = {
    model: config.model,
    prompt,
    n: 1,
    size,
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
      const data = await resp.json() as { data: { url?: string; b64_json?: string }[] }
      if (data.data?.[0]?.url) {
        return data.data[0].url
      }
      if (data.data?.[0]?.b64_json) {
        return `data:image/png;base64,${data.data[0].b64_json}`
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
      const config = resolveImageConfig(ctx)
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
      let lastImageError = ''
      for (const style of styles) {
        try {
          const imageUrl = await callAgnesImage(config, style.prompt, size, referenceImages)
          images.push({ url: imageUrl, style: style.title })
        } catch (err) {
          lastImageError = err instanceof Error ? err.message : String(err)
          // Retry once
          await new Promise(r => setTimeout(r, 2000))
          try {
            const imageUrl = await callAgnesImage(config, style.prompt, size, referenceImages)
            images.push({ url: imageUrl, style: style.title })
          } catch (retryError) {
            lastImageError = retryError instanceof Error ? retryError.message : String(retryError)
            // Skip this style
          }
        }
      }

      if (images.length === 0) {
        throw new Error(`图片生成失败：${lastImageError || '服务未返回图片结果'}`)
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
