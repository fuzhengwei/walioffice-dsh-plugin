/**
 * WaLiOffice DSH Plugin Bundle
 * 
 * Loads all WaLiOffice office productivity tools as DSH plugins.
 * This is the one-stop entry point — add to cordis.yml to enable all tools.
 * 
 * Tools registered:
 * - ppt_plan / ppt_generate — PPT 大纲规划 + 完整生成
 * - doc_generate / md_generate — Word 文档 + Markdown 文档
 * - sheet_generate — Excel 表格
 * - chart_generate — ECharts 图表
 * - drawio_generate — draw.io 图表
 * - image_prompt — AI 图片生成
 * - video_generate / video_storyboard — AI 视频生成 + 分镜规划
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

// Import tool plugins
import { apply as applyPpt, name as namePpt, inject as injectPpt } from '@walioffice/dsh-tool-ppt'
import { apply as applyDoc, name as nameDoc, inject as injectDoc } from '@walioffice/dsh-tool-doc'
import { apply as applySheet, name as nameSheet, inject as injectSheet } from '@walioffice/dsh-tool-sheet'
import { apply as applyChart, name as nameChart, inject as injectChart } from '@walioffice/dsh-tool-chart'
import { apply as applyDrawio, name as nameDrawio, inject as injectDrawio } from '@walioffice/dsh-tool-drawio'
import { apply as applyImage, name as nameImage, inject as injectImage } from '@walioffice/dsh-tool-image'
import { apply as applyVideo, name as nameVideo, inject as injectVideo } from '@walioffice/dsh-tool-video'

export const name = 'walioffice-bundle'
export const inject = [
  'tools',
  'llm',
  ...new Set([
    ...injectPpt, ...injectDoc, ...injectSheet,
    ...injectChart, ...injectDrawio, ...injectImage, ...injectVideo,
  ]),
]

export function apply(ctx: Context): void {
  // The 'office' service is provided by @walioffice/dsh-office plugin,
  // which is inserted before this bundle's tool plugins in cordis.patch.yml.

  // Register all tools
  applyPpt(ctx)
  applyDoc(ctx)
  applySheet(ctx)
  applyChart(ctx)
  applyDrawio(ctx)
  applyImage(ctx)
  applyVideo(ctx)

  // Log registration
  console.log('[WaLiOffice] Registered 9 office tools: ppt_plan, ppt_generate, doc_generate, md_generate, sheet_generate, chart_generate, drawio_generate, image_prompt, video_generate, video_storyboard')
}
