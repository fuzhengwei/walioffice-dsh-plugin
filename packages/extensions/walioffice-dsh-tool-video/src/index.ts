/**
 * Video generation tools: video_generate and video_storyboard
 * Uses Agnes Video V2.5 API.
 * 
 * @module @walioffice/dsh-tool-video
 */

import type { Context } from '@deepseek-ai/cordis'
import { defineTool } from '@deepseek-ai/dsh-tools'
import { llmGenerateJson, resolveOfficeService } from '@walioffice/dsh-office'
import { readFileSync } from 'node:fs'
import { homedir } from 'node:os'
import { join } from 'node:path'

// ── Types ───────────────────────────────────────────────────────────────────

interface VideoPlan {
  title: string
  description: string
  prompt: string
  negative_prompt: string
  aspect_ratio: string
  seconds: number
  mode: 'text' | 'keyframe' | 'reference'
}

interface CreateVideoResponse {
  id: string
  status: string
}

interface QueryVideoResponse {
  status: 'pending' | 'processing' | 'succeeded' | 'failed'
  progress: number
  url?: string
  metadata?: Record<string, unknown>
}

interface StoryboardShot {
  index: number
  title: string
  description: string
  prompt: string
  mode: 'text' | 'keyframe' | 'reference'
  seconds: number
  first_frame?: string
  last_frame?: string
  reference_images: string[]
  audio_urls: string[]
  transition?: string
}

interface StoryboardPlan {
  title: string
  description: string
  total_shots: number
  total_seconds: number
  aspect_ratio: string
  shots: StoryboardShot[]
}

// ── Config ──────────────────────────────────────────────────────────────────

interface VideoConfig {
  baseUrl: string
  apiKeys: string[]
  model: string
}

interface VideoProviderProfile {
  baseUrl: string
  apiKeys: string[]
  models: string[]
}

const VIDEO_MODEL = 'agnes-video-2.5'

function resolveVideoConfig(ctx: Context): VideoConfig {
  const baseUrl = process.env.AGNES_VIDEO_BASE_URL || process.env.LLM_VIDEO_BASE_URL || ''
  const explicitApiKeys = (
    process.env.AGNES_VIDEO_API_KEYS?.split(',').map(k => k.trim()).filter(Boolean) ||
    process.env.LLM_VIDEO_API_KEYS?.split(',').map(k => k.trim()).filter(Boolean) ||
    (process.env.LLM_VIDEO_API_KEY ? [process.env.LLM_VIDEO_API_KEY] : [])
  )
  const discovered = discoverVideoProvider(ctx, VIDEO_MODEL)

  if (!discovered) {
    throw new Error(
      `未在 DSH 模型配置中找到视频模型 ${VIDEO_MODEL}。\n` +
      '请先申请该模型，然后在 ~/.dsh/settings.yaml 的 llm-pi-ai.providers.models 中配置它，再重试。'
    )
  }

  const resolvedBaseUrl = baseUrl || discovered.baseUrl
  const apiKeys = explicitApiKeys.length > 0 ? explicitApiKeys : discovered.apiKeys

  if (!resolvedBaseUrl || apiKeys.length === 0) {
    throw new Error(
      '视频生成需要配置环境变量：\n' +
      '  AGNES_VIDEO_BASE_URL (或 LLM_VIDEO_BASE_URL) — API 地址\n' +
      '  AGNES_VIDEO_API_KEYS (或 LLM_VIDEO_API_KEY) — API 密钥\n' +
      '也可以在 DSH 的 ~/.dsh/settings.yaml 配置视频模型，并在 ~/.dsh/.credentials.yaml 中提供 apiKeyEnv 对应的密钥。'
    )
  }

  return { baseUrl: resolvedBaseUrl, apiKeys, model: VIDEO_MODEL }
}

function discoverVideoProvider(ctx: Context, requestedModel: string): VideoProviderProfile | undefined {
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
  const entries = Object.values(providers)
    .map(value => readVideoProviderProfile(value, credentials))
    .filter((item): item is VideoProviderProfile => item !== undefined)
  return entries.find(entry => entry.models.includes(requestedModel.trim()))
}

function readVideoProviderProfile(value: unknown, credentials: Record<string, string>): VideoProviderProfile | undefined {
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

// ── Helpers ─────────────────────────────────────────────────────────────────

function normalizeAspectRatio(ratio: string): string {
  const valid = ['9:16', '1:1', '4:3', '3:4', '21:9', '16:9']
  return valid.includes(ratio) ? ratio : '16:9'
}

function inferSize(ratio: string): { width: number; height: number } {
  const map: Record<string, { width: number; height: number }> = {
    '16:9': { width: 1280, height: 720 },
    '9:16': { width: 720, height: 1280 },
    '1:1': { width: 720, height: 720 },
    '4:3': { width: 960, height: 720 },
    '3:4': { width: 720, height: 960 },
    '21:9': { width: 1680, height: 720 },
  }
  return map[ratio] ?? map['16:9']!
}

function normalizeSeconds(seconds: number): number {
  return Math.max(4, Math.min(12, Math.round(seconds)))
}

function inferMode(imageCount: number, explicit?: string): 'text' | 'keyframe' | 'reference' {
  if (explicit && ['text', 'keyframe', 'reference'].includes(explicit)) {
    return explicit as 'text' | 'keyframe' | 'reference'
  }
  if (imageCount === 0) return 'text'
  if (imageCount <= 2) return 'keyframe'
  return 'reference'
}

// ── API calls ───────────────────────────────────────────────────────────────

async function createVideoTask(
  config: VideoConfig,
  plan: VideoPlan,
  images: string[],
  videos: string[],
): Promise<CreateVideoResponse> {
  const url = `${config.baseUrl.replace(/\/$/, '')}/v1/videos/generations`
  const { width, height } = inferSize(plan.aspect_ratio)

  const body: Record<string, unknown> = {
    model: config.model,
    prompt: plan.prompt,
    negative_prompt: plan.negative_prompt,
    width,
    height,
    seconds: plan.seconds,
    mode: plan.mode,
  }
  if (images.length > 0) {
    body.extra_body = { image: images }
  }
  if (videos.length > 0) {
    body.extra_body = { ...((body.extra_body as Record<string, unknown>) ?? {}), video_ref: videos }
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
        if (resp.status === 401 || resp.status === 403 || resp.status === 429) continue
        continue
      }
      return await resp.json() as CreateVideoResponse
    } catch (err) {
      lastError = err instanceof Error ? err.message : String(err)
    }
  }
  throw new Error(`Agnes Video API 创建任务失败：${lastError}`)
}

async function pollVideoTask(
  config: VideoConfig,
  taskId: string,
  maxAttempts = 120,
): Promise<QueryVideoResponse> {
  const url = `${config.baseUrl.replace(/\/$/, '')}/v1/videos/generations/${taskId}`
  
  for (let i = 0; i < maxAttempts; i++) {
    for (const key of config.apiKeys) {
      try {
        const resp = await fetch(url, {
          headers: { 'Authorization': `Bearer ${key}` },
        })
        if (resp.ok) {
          const data = await resp.json() as QueryVideoResponse
          if (data.status === 'succeeded') return data
          if (data.status === 'failed') throw new Error('视频生成失败')
        }
      } catch {
        // try next key
      }
    }
    await new Promise(r => setTimeout(r, 5000)) // poll every 5s
  }
  throw new Error('视频生成超时')
}

// ── video_generate tool ─────────────────────────────────────────────────────

export function apply(ctx: Context): void {
  ctx.tools.register(defineTool({
    name: 'video_generate',
    description: '生成视频：基于 Agnes Video V2.5，支持 text/keyframe/reference 三种模式（0图→text，1-2图→keyframe，3+图→reference），支持从会话历史产物提取图片。复杂视频建议先调用 video_storyboard 分镜。',
    parameters: {
      topic: { type: 'string', required: true, description: '视频需求描述' },
      aspect_ratio: {
        type: 'string',
        description: '宽高比：16:9/9:16/1:1/4:3/3:4/21:9（默认 16:9）',
      },
      seconds: {
        type: 'number',
        description: '视频时长 4-12 秒（默认 5）',
      },
      mode: {
        type: 'string',
        description: '生成模式：text/keyframe/reference（自动推断）',
        enum: ['text', 'keyframe', 'reference'],
      },
      image_urls: {
        type: 'array',
        items: { type: 'string' },
        description: '参考图片 URL',
      },
      video_urls: {
        type: 'array',
        items: { type: 'string' },
        description: '参考视频 URL',
      },
    },
    output: {
      schema: {
        type: 'object',
        additionalProperties: true,
        properties: {
          title: { type: 'string', required: true },
          videoUrl: { type: 'string', required: true },
          duration: { type: 'number', required: true },
          aspectRatio: { type: 'string', required: true },
          mode: { type: 'string', required: true },
          provider: { type: 'string', required: true },
        },
      },
      render: (_args, value) => [{
        type: 'text',
        text: `已生成视频《${value.title}》(${value.duration}秒, ${value.aspectRatio})`,
      }],
      presentationMeta: (_args, value) => ({ kind: 'video', ...value }),
    },
    isConcurrencySafe: () => false,
    async execute(args, exec) {
      const topic = args.topic
      if (!topic?.trim()) {
        throw new Error('topic 不能为空')
      }
      const office = resolveOfficeService(ctx)

      const aspectRatio = normalizeAspectRatio(args.aspect_ratio ?? '16:9')
      const seconds = normalizeSeconds(args.seconds ?? 5)
      const images = (args.image_urls ?? []).filter(Boolean)
      const videos = (args.video_urls ?? []).filter(Boolean)
      const mode = inferMode(images.length, args.mode)

      const config = resolveVideoConfig(ctx)

      office.emitProgress('running', '创建视频任务', `正在创建 ${mode} 模式视频任务...`)

      const plan: VideoPlan = {
        title: topic.slice(0, 24),
        description: topic,
        prompt: topic,
        negative_prompt: 'blurry, low quality, distorted, watermark, text overlay',
        aspect_ratio: aspectRatio,
        seconds,
        mode,
      }

      const task = await createVideoTask(config, plan, images, videos)

      office.emitProgress('running', '等待视频生成', `任务 ${task.id} 正在处理中...`)

      const result = await pollVideoTask(config, task.id)

      if (!result.url) {
        throw new Error('视频生成完成但未返回 URL')
      }

      return {
        title: plan.title,
        videoUrl: result.url,
        duration: seconds,
        aspectRatio,
        mode,
        provider: 'agnes',
        model: config.model,
      }
    },
  }))

  // ── video_storyboard tool ─────────────────────────────────────────────────
  const STORYBOARD_SYSTEM_PROMPT = `你是视频导演。只输出严格 JSON，不要 markdown 代码块。
返回格式：
{
  "title": "视频标题",
  "description": "视频概述",
  "total_shots": 3,
  "total_seconds": 15,
  "aspect_ratio": "16:9",
  "shots": [
    {
      "index": 0,
      "title": "镜头标题",
      "description": "镜头描述",
      "prompt": "English prompt for video generation",
      "mode": "text",
      "seconds": 5,
      "reference_images": [],
      "audio_urls": [],
      "transition": "fade"
    }
  ]
}
要求：
- 每镜头 4-12 秒
- prompt 为英文，描述画面内容、镜头运动、光影氛围
- mode 根据参考图数量自动选择
- 镜头之间有逻辑连贯性`

  ctx.tools.register(defineTool({
    name: 'video_storyboard',
    description: '视频分镜规划：将复杂视频需求拆分为多个镜头，每个镜头含英文提示词、时长、模式。规划完成后可逐镜头调用 video_generate 生成。',
    parameters: {
      topic: { type: 'string', required: true, description: '视频需求描述' },
      aspect_ratio: {
        type: 'string',
        description: '宽高比（默认 16:9）',
      },
      max_shots: {
        type: 'number',
        description: '最大镜头数（默认 3，建议 2-5）',
      },
      seconds_per_shot: {
        type: 'number',
        description: '每镜头秒数（默认 5）',
      },
      image_urls: {
        type: 'array',
        items: { type: 'string' },
        description: '参考图片 URL',
      },
    },
    output: {
      schema: {
        type: 'object',
        additionalProperties: true,
        properties: {
          title: { type: 'string', required: true },
          totalShots: { type: 'integer', required: true },
          totalSeconds: { type: 'integer', required: true },
          aspectRatio: { type: 'string', required: true },
          shots: { type: 'array', required: true },
        },
      },
      render: (_args, value) => [{
        type: 'text',
        text: `已规划视频分镜《${value.title}》：${value.totalShots} 个镜头，共 ${value.totalSeconds} 秒`,
      }],
      presentationMeta: (_args, value) => ({ kind: 'storyboard', ...value }),
    },
    isConcurrencySafe: () => false,
    async execute(args, exec) {
      const topic = args.topic
      if (!topic?.trim()) {
        throw new Error('topic 不能为空')
      }
      const office = resolveOfficeService(ctx)
      const aspectRatio = normalizeAspectRatio(args.aspect_ratio ?? '16:9')
      const maxShots = Math.max(1, Math.min(8, args.max_shots ?? 3))
      const secondsPerShot = normalizeSeconds(args.seconds_per_shot ?? 5)

      office.emitProgress('running', '规划视频分镜', `正在为《${topic}》规划分镜...`)

      const userPrompt = [
        `需求：${topic}`,
        `宽高比：${aspectRatio}`,
        `最大镜头数：${maxShots}`,
        `每镜头时长：${secondsPerShot}秒`,
        '',
        '请规划一份完整的视频分镜方案。',
      ].join('\n')

      const json = await llmGenerateJson(ctx, STORYBOARD_SYSTEM_PROMPT, userPrompt)
      const plan = json as unknown as StoryboardPlan

      if (!plan.shots || !Array.isArray(plan.shots) || plan.shots.length === 0) {
        throw new Error('LLM 返回的分镜数据格式不正确')
      }

      return {
        title: plan.title,
        totalShots: plan.shots.length,
        totalSeconds: plan.shots.reduce((sum, s) => sum + s.seconds, 0),
        aspectRatio,
        shots: plan.shots as any,
      }
    },
  }))
}

export const name = 'tool-video'
export const inject = ['tools', 'office']
