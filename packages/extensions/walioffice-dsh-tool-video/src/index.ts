/**
 * Video generation tools: video_generate and video_storyboard
 * Uses Agnes Video V2.5 API.
 * 
 * @module @walioffice/dsh-tool-video
 */

import type { Context } from '@deepseek-ai/cordis'
import { defineTool } from '@deepseek-ai/dsh-tools'
import { llmGenerateJson, resolveOfficeService } from '@walioffice/dsh-office'

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

function resolveVideoConfig(): VideoConfig {
  const baseUrl = process.env.AGNES_VIDEO_BASE_URL || process.env.LLM_VIDEO_BASE_URL || ''
  const apiKeys = (
    process.env.AGNES_VIDEO_API_KEYS?.split(',').map(k => k.trim()).filter(Boolean) ||
    process.env.LLM_VIDEO_API_KEYS?.split(',').map(k => k.trim()).filter(Boolean) ||
    (process.env.LLM_VIDEO_API_KEY ? [process.env.LLM_VIDEO_API_KEY] : [])
  )
  const model = process.env.AGNES_VIDEO_MODEL || 'agnes-video-v2.5'

  if (!baseUrl || apiKeys.length === 0) {
    throw new Error(
      '视频生成需要配置环境变量：\n' +
      '  AGNES_VIDEO_BASE_URL (或 LLM_VIDEO_BASE_URL) — API 地址\n' +
      '  AGNES_VIDEO_API_KEYS (或 LLM_VIDEO_API_KEY) — API 密钥\n' +
      '可在 .env 文件或系统环境变量中设置。'
    )
  }

  return { baseUrl, apiKeys, model }
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

      const config = resolveVideoConfig()

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
