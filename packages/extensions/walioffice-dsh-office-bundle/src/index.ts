/**
 * WaLiOffice DSH Plugin — unified entry point.
 * 
 * Registers the office service + all office productivity tools as a single plugin.
 * 
 * Tools registered (10):
 * - ppt_plan / ppt_generate — PPT 大纲规划 + 完整生成
 * - doc_generate / md_generate — Word 文档 + Markdown 文档
 * - sheet_generate — Excel 表格
 * - chart_generate — ECharts 图表
 * - drawio_generate — draw.io 图表
 * - image_prompt — AI 图片生成
 * - wali_video_generate / video_storyboard — AI 视频生成 + 分镜规划
 * 
 * Usage in cordis.yml:
 *   plugins:
 *     - name: '@walioffice/dsh-office-bundle'
 * 
 * Environment variables for external APIs:
 * - AGNES_IMAGE_BASE_URL / AGNES_IMAGE_API_KEYS — 图片生成 API
 * - AGNES_VIDEO_BASE_URL / AGNES_VIDEO_API_KEYS — 视频生成 API
 * 
 * @module @walioffice/dsh-office-bundle
 */

import type { Context } from '@deepseek-ai/cordis'
import '@walioffice/dsh-office' // for Context augmentation

// Import office service registration
import * as officePlugin from '@walioffice/dsh-office'

// Import tool plugins
import * as pptPlugin from '@walioffice/dsh-tool-ppt'
import * as docPlugin from '@walioffice/dsh-tool-doc'
import * as sheetPlugin from '@walioffice/dsh-tool-sheet'
import * as chartPlugin from '@walioffice/dsh-tool-chart'
import * as drawioPlugin from '@walioffice/dsh-tool-drawio'
import * as imagePlugin from '@walioffice/dsh-tool-image'
import * as videoPlugin from '@walioffice/dsh-tool-video'

export const name = 'walioffice'
export const inject = ['tools', 'llm']

export function apply(ctx: Context): void {
  // 1. Register the office service first
  ctx.plugin(officePlugin)

  // 2. Register all tools through Cordis so their inject declarations are honored
  ctx.plugin(pptPlugin)
  ctx.plugin(docPlugin)
  ctx.plugin(sheetPlugin)
  ctx.plugin(chartPlugin)
  ctx.plugin(drawioPlugin)
  ctx.plugin(imagePlugin)
  ctx.plugin(videoPlugin)

  console.log('[WaLiOffice] Registered 10 office tools: ppt_plan, ppt_generate, doc_generate, md_generate, sheet_generate, chart_generate, drawio_generate, image_prompt, wali_video_generate, video_storyboard')
}
