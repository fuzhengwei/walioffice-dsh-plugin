/**
 * Video generation tools: video_generate and video_storyboard
 * Uses Agnes Video V2.0 API.
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
  videoId: string
  status: string
}

interface RawCreateVideoResponse {
  id?: string
  task_id?: string
  video_id?: string
  status?: string
  progress?: number
}

interface QueryVideoResponse {
  id?: string
  video_id?: string
  model?: string
  status: 'queued' | 'pending' | 'processing' | 'in_progress' | 'completed' | 'failed' | 'error' | 'cancelled'
  progress?: number
  seconds?: string
  size?: string
  remixed_from_video_id?: string
  url?: string
  metadata?: Record<string, unknown>
  error?: { message?: string }
  video_url?: string
  data?: { url?: string; video_url?: string }[]
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

const VIDEO_MODEL = 'agnes-video-v2.0'

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

function normalizeSeconds(seconds: number): number {
  return Math.max(4, Math.min(12, Math.round(seconds)))
}

function inferSize(ratio: string): { width: number; height: number } {
  const map: Record<string, { width: number; height: number }> = {
    '16:9': { width: 1152, height: 768 },
    '9:16': { width: 768, height: 1152 },
    '1:1': { width: 960, height: 960 },
    '4:3': { width: 1024, height: 768 },
    '3:4': { width: 768, height: 1024 },
  }
  return map[ratio] ?? map['16:9']!
}

function inferNumFrames(seconds: number): number {
  if (seconds <= 4) return 81
  if (seconds >= 9) return 241
  return 121
}

function inferMode(imageCount: number, videoCount: number, explicit?: string): 'text' | 'keyframe' | 'reference' {
  if (explicit && ['text', 'keyframe', 'reference'].includes(explicit)) {
    return explicit as 'text' | 'keyframe' | 'reference'
  }
  if (videoCount > 0) return 'reference'
  if (imageCount === 0) return 'text'
  if (imageCount <= 2) return 'keyframe'
  return 'reference'
}

function storyboardMode(referenceCount: number): 'text' | 'keyframe' | 'reference' {
  if (referenceCount === 0) return 'text'
  if (referenceCount <= 2) return 'keyframe'
  return 'reference'
}

function fallbackStoryboard(
  topic: string,
  aspectRatio: string,
  maxShots: number,
  secondsPerShot: number,
  referenceImages: string[],
): StoryboardPlan {
  const mode = storyboardMode(referenceImages.length)
  const templates = [
    {
      title: '环境建立',
      description: `展示${topic}发生的环境、主体和整体氛围。`,
      prompt: `Wide establishing shot, introduce the main subject and environment for: ${topic}. Gentle cinematic camera movement, warm natural lighting, polished commercial video style.`,
      transition: 'fade',
    },
    {
      title: '主体动作',
      description: `突出${topic}中的主要动作和画面变化。`,
      prompt: `Medium tracking shot, show the main subject actively moving and interacting within: ${topic}. Smooth camera motion, expressive action, natural lighting, cinematic detail.`,
      transition: 'dissolve',
    },
    {
      title: '温馨收束',
      description: `以${topic}的情绪高潮和温馨结尾收束视频。`,
      prompt: `Close-up and gentle pull-back, capture the warm emotional ending of: ${topic}. Authentic expressions, soft sunlight, comforting atmosphere, high-quality cinematic finish.`,
      transition: 'fade',
    },
  ]
  const shots = templates.slice(0, maxShots).map((template, index) => ({
    index,
    title: template.title,
    description: template.description,
    prompt: template.prompt,
    mode,
    seconds: secondsPerShot,
    reference_images: referenceImages,
    audio_urls: [],
    transition: template.transition,
  }))
  return {
    title: topic.slice(0, 24),
    description: `围绕“${topic}”生成一支连贯的短视频。`,
    total_shots: shots.length,
    total_seconds: shots.length * secondsPerShot,
    aspect_ratio: aspectRatio,
    shots,
  }
}

function normalizeStoryboard(
  value: unknown,
  topic: string,
  aspectRatio: string,
  maxShots: number,
  secondsPerShot: number,
  referenceImages: string[],
): StoryboardPlan {
  const record = asRecord(value)
  const rawShots = Array.isArray(record?.shots) ? record.shots : []
  const mode = storyboardMode(referenceImages.length)
  const shots = rawShots.slice(0, maxShots).map((item, index) => {
    const shot = asRecord(item)
    return {
      index,
      title: typeof shot?.title === 'string' ? shot.title : `镜头 ${index + 1}`,
      description: typeof shot?.description === 'string' ? shot.description : topic,
      prompt: typeof shot?.prompt === 'string' ? shot.prompt : `Cinematic shot showing: ${topic}.`,
      mode,
      seconds: typeof shot?.seconds === 'number' ? normalizeSeconds(shot.seconds) : secondsPerShot,
      reference_images: referenceImages,
      audio_urls: Array.isArray(shot?.audio_urls) ? shot.audio_urls.filter(item => typeof item === 'string') as string[] : [],
      transition: typeof shot?.transition === 'string' ? shot.transition : 'fade',
    }
  })
  if (shots.length === 0) {
    throw new Error('LLM 返回的分镜数据格式不正确')
  }
  return {
    title: typeof record?.title === 'string' ? record.title : topic.slice(0, 24),
    description: typeof record?.description === 'string' ? record.description : `围绕“${topic}”生成一支连贯的短视频。`,
    total_shots: shots.length,
    total_seconds: shots.reduce((sum, shot) => sum + shot.seconds, 0),
    aspect_ratio: aspectRatio,
    shots,
  }
}

function videoApiBaseUrl(baseUrl: string): string {
  const normalized = baseUrl.replace(/\/+$/, '')
  return normalized.endsWith('/v1') ? normalized : `${normalized}/v1`
}

function videoApiRootUrl(baseUrl: string): string {
  return videoApiBaseUrl(baseUrl).replace(/\/v1$/, '')
}

// ── API calls ───────────────────────────────────────────────────────────────

async function createVideoTask(
  config: VideoConfig,
  plan: VideoPlan,
  images: string[],
  videos: string[],
): Promise<CreateVideoResponse> {
  const url = `${videoApiBaseUrl(config.baseUrl)}/videos`
  const { width, height } = inferSize(plan.aspect_ratio)
  const numFrames = inferNumFrames(plan.seconds)

  const body: Record<string, unknown> = {
    model: config.model,
    prompt: plan.prompt,
    negative_prompt: plan.negative_prompt,
    width,
    height,
    num_frames: numFrames,
    frame_rate: 24,
  }

  if (plan.mode === 'keyframe' && images.length > 0) {
    body.extra_body = { image: images, mode: 'keyframes' }
  } else if (images.length > 1) {
    body.extra_body = { image: images }
  } else if (images[0]) {
    body.image = images[0]
  }
  if (videos.length > 0) {
    body.extra_body = {
      ...((body.extra_body as Record<string, unknown>) ?? {}),
      video_ref: videos,
    }
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
      const data = await resp.json() as RawCreateVideoResponse
      const taskId = data.task_id || data.id || data.video_id
      const videoId = data.video_id || data.id || data.task_id
      if (!taskId || !videoId) throw new Error('创建任务响应中缺少任务 ID 或视频 ID')
      return { id: taskId, videoId, status: data.status || 'queued' }
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
  const url = `${videoApiRootUrl(config.baseUrl)}/agnesapi?video_id=${encodeURIComponent(taskId)}&model_name=${encodeURIComponent(config.model)}`
  let lastError = ''

  for (let i = 0; i < maxAttempts; i++) {
    for (const key of config.apiKeys) {
      let resp: Response
      try {
        resp = await fetch(url, {
          headers: { 'Authorization': `Bearer ${key}` },
        })
      } catch (err) {
        lastError = err instanceof Error ? err.message : String(err)
        continue
      }
      if (!resp.ok) {
        lastError = `HTTP ${resp.status}: ${await resp.text()}`
        continue
      }

      let data: QueryVideoResponse
      try {
        data = await resp.json() as QueryVideoResponse
      } catch (err) {
        lastError = err instanceof Error ? err.message : String(err)
        continue
      }
      if (data.status === 'completed') {
        const url = data.url || data.video_url || data.data?.[0]?.url || data.data?.[0]?.video_url
        return url ? { ...data, url } : data
      }
      if (data.status === 'failed' || data.status === 'error' || data.status === 'cancelled') {
        throw new Error(data.error?.message || '视频生成失败')
      }
    }
    await new Promise(r => setTimeout(r, 1500))
  }
  throw new Error(`视频生成超时${lastError ? `：${lastError}` : ''}`)
}

// ── video_generate tool ─────────────────────────────────────────────────────

export function apply(ctx: Context): void {
  ctx.tools.register(defineTool({
    name: 'video_generate',
    description: '生成可预览的 AI 视频。只有用户明确要求视频、短片、短视频、宣传片、广告、动画、片头、转场、让图片动起来或图生视频时调用；支持 text/keyframe/reference 模式并可复用会话图片，复杂多镜头需求先调用 video_storyboard。',
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
      const mode = inferMode(images.length, videos.length, args.mode)

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

      const result = await pollVideoTask(config, task.videoId)

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
    description: '规划复杂视频分镜：用户要求多场景、多镜头、故事线、宣传片或完整短片时先调用本工具，再按镜头调用 video_generate；输出每镜头的英文提示词、时长、生成模式和参考图分配，不直接生成视频。',
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
      const referenceImages = (args.image_urls ?? []).filter(Boolean)

      office.emitProgress('running', '规划视频分镜', `正在为《${topic}》规划分镜...`)

      const userPrompt = [
        `需求：${topic}`,
        `宽高比：${aspectRatio}`,
        `最大镜头数：${maxShots}`,
        `每镜头时长：${secondsPerShot}秒`,
        '',
        '请规划一份完整的视频分镜方案。',
      ].join('\n')

      let plan: StoryboardPlan
      try {
        const json = await llmGenerateJson(ctx, STORYBOARD_SYSTEM_PROMPT, userPrompt)
        plan = normalizeStoryboard(json, topic, aspectRatio, maxShots, secondsPerShot, referenceImages)
      } catch (error) {
        office.emitProgress('running', '使用默认分镜', 'LLM 未返回可解析 JSON，已切换为内置分镜方案继续生成。')
        plan = fallbackStoryboard(topic, aspectRatio, maxShots, secondsPerShot, referenceImages)
        void error
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
