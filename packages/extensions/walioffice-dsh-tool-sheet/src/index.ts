/**
 * Sheet generation tool: sheet_generate
 * Produces structured spreadsheet data and exports to .xlsx
 * 
 * @module @walioffice/dsh-tool-sheet
 */

import type { Context } from '@deepseek-ai/cordis'
import { defineTool } from '@deepseek-ai/dsh-tools'
import { inferScene, llmGenerateJson, resolveOfficeService, sceneGuide } from '@walioffice/dsh-office'
import { renderXlsx } from '@walioffice/dsh-office-render-xlsx'

// ── Types ───────────────────────────────────────────────────────────────────

export interface SheetTable {
  title: string
  headers: string[]
  rows: (string | number)[][]
  summary: string
}

export interface SheetOutput {
  tables: SheetTable[]
}

// ── System prompt ───────────────────────────────────────────────────────────

const SYSTEM_PROMPT = `你是数据分析师。只输出严格 JSON，不要 markdown 代码块。
返回格式：
{
  "tables": [
    {
      "title": "表名",
      "headers": ["列1","列2","列3"],
      "rows": [["值1","值2","值3"]],
      "summary": "表格说明"
    }
  ]
}
要求：
- 每表至少 4 列 6 行，数据真实具体（不要占位符）
- 列名专业、可直接用于分析或执行
- 可分多表（明细+汇总）
- 如用户没给足信息，自动补合理数据`

export function apply(ctx: Context): void {
  ctx.tools.register(defineTool({
    name: 'sheet_generate',
    description: '生成结构化表格（可导出 Excel）：根据主题生成多表数据，适合数据分析、排期、预算、指标明细等场景。',
    parameters: {
      topic: { type: 'string', required: true, description: '表格主题/用户需求' },
      sheets: { type: 'integer', description: '表格数量（默认 1）' },
    },
    output: {
      schema: {
        type: 'object',
        additionalProperties: true,
        properties: {
          tableCount: { type: 'integer', required: true },
          totalRows: { type: 'integer', required: true },
          filePath: { type: 'string', required: true },
          tables: { type: 'array' },
        },
      },
      render: (_args, value) => [{
        type: 'text',
        text: `已生成 ${value.tableCount} 个表格，共 ${value.totalRows} 行数据，文件已保存到 ${value.filePath}`,
      }],
      presentationMeta: (_args, value) => ({ kind: 'sheet', ...value }),
    },
    isConcurrencySafe: () => false,
    async execute(args, exec) {
      const topic = args.topic
      if (!topic?.trim()) {
        throw new Error('topic 不能为空')
      }
      const sheetCount = args.sheets ?? 1
      const scene = inferScene(topic)
      const guide = sceneGuide(scene, 'sheet')
      const office = resolveOfficeService(ctx)

      office.emitProgress('running', '生成表格', `正在为《${topic}》生成数据...`)

      const userPrompt = [
        `需求：${topic}`,
        `表格数量：${sheetCount}`,
        `场景偏好：${guide}`,
        '',
        '请生成专业的结构化表格数据，只返回 JSON。',
      ].filter(Boolean).join('\n')

      const json = await llmGenerateJson(ctx, SYSTEM_PROMPT, userPrompt)
      const output = json as unknown as SheetOutput

      if (!output.tables || !Array.isArray(output.tables) || output.tables.length === 0) {
        throw new Error('LLM 返回的表格数据格式不正确')
      }

      // Render to .xlsx
      office.emitProgress('running', '导出 Excel', '正在生成 .xlsx 文件...')
      const filePath = await renderXlsx(ctx, output)

      const totalRows = output.tables.reduce((sum, t) => sum + t.rows.length, 0)

      return {
        tableCount: output.tables.length,
        totalRows,
        filePath,
        tables: output.tables.map(t => ({
          title: t.title,
          headers: t.headers,
          rows: t.rows.slice(0, 12),
          rowCount: t.rows.length,
          summary: t.summary,
        })),
      }
    },
  }))
}

export const name = 'tool-sheet'
export const inject = ['tools', 'office']
