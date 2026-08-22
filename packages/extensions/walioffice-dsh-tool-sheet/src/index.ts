/**
 * Sheet generation tool: sheet_generate
 * Produces structured spreadsheet data and exports to .xlsx
 * 
 * @module @walioffice/dsh-tool-sheet
 */

import type { Context } from '@deepseek-ai/cordis'
import { defineTool } from '@deepseek-ai/dsh-tools'
import { createOfficeDownload, inferScene, llmGenerateJson, resolveOfficeService, sceneGuide } from '@walioffice/dsh-office'
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

function fallbackSheet(topic: string, sheetCount: number): SheetOutput {
  const tables: SheetTable[] = []
  for (let index = 0; index < Math.max(1, Math.min(6, sheetCount)); index += 1) {
    const title = index === 0 ? `${topic.slice(0, 18)}执行清单` : `${topic.slice(0, 16)}汇总表${index + 1}`
    tables.push({
      title,
      headers: ['编号', '工作项', '负责人', '计划日期', '状态', '备注'],
      rows: [
        ['1', `${topic}目标确认`, '项目负责人', '第1周', '待开始', '明确范围与验收标准'],
        ['2', `${topic}资料准备`, '业务负责人', '第1周', '待开始', '整理现有资料与数据'],
        ['3', `${topic}方案设计`, '方案负责人', '第2周', '进行中', '输出初版方案与任务拆解'],
        ['4', `${topic}协作执行`, '执行团队', '第3周', '未开始', '按优先级推进关键任务'],
        ['5', `${topic}阶段检查`, '质量负责人', '第4周', '未开始', '检查进度、质量与风险'],
        ['6', `${topic}复盘优化`, '项目负责人', '第5周', '未开始', '沉淀经验并调整后续计划'],
      ],
      summary: `围绕“${topic}”整理的结构化执行数据，可用于跟踪任务、负责人、时间和状态。`,
    })
  }
  return { tables }
}

function normalizeSheetOutput(value: unknown, topic: string, sheetCount: number): SheetOutput {
  if (typeof value !== 'object' || value === null || !Array.isArray((value as { tables?: unknown }).tables)) {
    throw new Error('LLM 返回的表格数据格式不正确')
  }
  const tables = (value as { tables: unknown[] }).tables.filter(item => typeof item === 'object' && item !== null).map((item, index) => {
    const table = item as Partial<SheetTable>
    const headers = Array.isArray(table.headers) ? table.headers.filter(item => typeof item === 'string') as string[] : []
    const rows = Array.isArray(table.rows) ? table.rows.filter(row => Array.isArray(row)).map(row => (row as unknown[]).map(cell => typeof cell === 'number' || typeof cell === 'string' ? cell : String(cell))) : []
    return {
      title: typeof table.title === 'string' ? table.title : `${topic.slice(0, 18)}表${index + 1}`,
      headers: headers.length > 0 ? headers : ['项目', '说明'],
      rows: rows.length > 0 ? rows : [['暂无数据', `围绕${topic}补充明细`]],
      summary: typeof table.summary === 'string' ? table.summary : `围绕“${topic}”整理的表格数据。`,
    }
  })
  if (tables.length === 0) throw new Error('LLM 返回的表格数据格式不正确')
  return { tables: tables.slice(0, Math.max(1, Math.min(6, sheetCount))) }
}

export function apply(ctx: Context): void {
  ctx.tools.register(defineTool({
    name: 'sheet_generate',
    description: '生成可下载的结构化 Excel 表格（.xlsx）。用户提到 Excel、xlsx、表格、数据明细、数据分析、排期、预算、指标、台账、清单或 CSV 整理时调用；自动设计字段、表头、示例数据和多张工作表。若用户只想看趋势/占比图，不要调用本工具，优先使用 chart_generate。',
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
          download: { type: 'object', additionalProperties: true },
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

      let output: SheetOutput
      try {
        const json = await llmGenerateJson(ctx, SYSTEM_PROMPT, userPrompt)
        output = normalizeSheetOutput(json, topic, sheetCount)
      } catch (err) {
        office.emitProgress('running', '使用表格模板兜底', 'LLM 未返回可解析内容，已使用内置表格结构继续生成。')
        output = fallbackSheet(topic, sheetCount)
        void err
      }

      // Render to .xlsx
      office.emitProgress('running', '导出 Excel', '正在生成 .xlsx 文件...')
      const filePath = await renderXlsx(ctx, output)
      const download = await createOfficeDownload(filePath, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

      const totalRows = output.tables.reduce((sum, t) => sum + t.rows.length, 0)

      return {
        tableCount: output.tables.length,
        totalRows,
        filePath,
        download,
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
