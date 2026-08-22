import type { Context } from '@deepseek-ai/cordis'
import * as bundlePlugin from '@walioffice/dsh-office-bundle'

export const name = 'walioffice'
export const inject = ['tools', 'llm']

export function apply(ctx: Context): void {
  ctx.plugin(bundlePlugin)
}
