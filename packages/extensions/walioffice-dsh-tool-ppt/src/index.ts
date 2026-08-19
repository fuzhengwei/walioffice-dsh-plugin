/**
 * Entry point for @walioffice/dsh-tool-ppt
 * Registers both ppt_plan and ppt_generate tools.
 * 
 * @module @walioffice/dsh-tool-ppt
 */

import type { Context } from '@deepseek-ai/cordis'
import { apply as applyPlan } from './plan.ts'
import { apply as applyGenerate } from './generate.ts'

export const name = 'tool-ppt'
export const inject = ['tools', 'office']

export function apply(ctx: Context): void {
  applyPlan(ctx)
  applyGenerate(ctx)
}

// Re-export types
export type { SlidePlan, PresentationPlan } from './plan.ts'
export type { SlideElement, Slide } from './generate.ts'
