/**
 * Document generation tools: doc_generate (Word) and md_generate (Markdown).
 * 
 * @module @walioffice/dsh-tool-doc
 */

import type { Context } from '@deepseek-ai/cordis'
import { defineTool } from '@deepseek-ai/dsh-tools'
import { llmGenerateJson, sceneGuide, inferScene } from '@walioffice/dsh-office'
import { renderDocx } from '@walioffice/dsh-office-render-docx'
import { resolve } from 'node:path'
import { mkdir, writeFile } from 'node:fs/promises'

// ── Types ───────────────────────────────────────────────────────────────────

export interface DocSection {
  heading: string
  heading_level: number
  paragraphs: string[]
  bullets: string[]
  table?: DocTable
}

export interface DocTable {
  headers: string[]
  rows: string[][]
}

export interface DocOutput {
  title: string
  sections: DocSection[]
  format: string
  markdown: string
}

// ── Scene-specific format guides ────────────────────────────────────────────

function formatGuide(format: string, scene: ReturnType<typeof inferScene>): string {
  const guides: Record<string, string> = {
    report: '输出正式报告格式，包含摘要、背景分析、核心内容、结论和建议。',
    plan: '输出方案/计划格式，包含目标、策略、执行步骤、资源需求和风险预案。',
    summary: '输出总结格式，包含关键结论、核心数据、亮点和不足、下一步建议。',
    article: '输出文章格式，包含引言、正文分段论述、结语。',
    prd: '输出 PRD 格式，包含背景、用户场景、功能需求、流程说明、验收标准和优先级。',
  }
  return guides[format] ?? guides.report ?? '输出正式报告格式。'
}

// ── Markdown conversion ─────────────────────────────────────────────────────

function sectionsToMarkdown(doc: DocOutput): string {
  const lines: string[] = [`# ${doc.title}`, '']
  for (const section of doc.sections) {
    const prefix = '#'.repeat(Math.min(section.heading_level + 1, 6))
    lines.push(`${prefix} ${section.heading}`, '')
    if (section.paragraphs.length > 0) {
      lines.push(...section.paragraphs, '')
    }
    if (section.bullets.length > 0) {
      lines.push(...section.bullets.map(b => `- ${b}`), '')
    }
    if (section.table) {
      lines.push(`| ${section.table.headers.join(' | ')} |`)
      lines.push(`| ${section.table.headers.map(() => '---').join(' | ')} |`)
      for (const row of section.table.rows) {
        lines.push(`| ${row.join(' | ')} |`)
      }
      lines.push('')
    }
  }
  return lines.join('\n')
}

// ── doc_generate tool ───────────────────────────────────────────────────────

const DOC_SYSTEM_PROMPT = `你是资深文档撰写专家。只输出严格 JSON，不要 markdown 代码块。
返回格式：
{
  "title": "文档标题",
  "sections": [
    {
      "heading": "章节标题",
      "heading_level": 1,
      "paragraphs": ["段落内容"],
      "bullets": ["要点"],
      "table": { "headers": ["列1","列2"], "rows": [["值1","值2"]] }
    }
  ],
  "format": "report"
}
要求：
- 至少 5 个章节，每章 2-3 段或 3-5 个要点
- 至少 2 章包含 table
- 段落 50-150 字，用 **粗体** 和 *斜体* 增强可读性
- 不要输出多余字段`

export function apply(ctx: Context): void {
  // ── doc_generate ──────────────────────────────────────────────────────────
  ctx.tools.register(defineTool({
    name: 'doc_generate',
    description: '生成结构化 Word 文档（.docx）：支持报告/计划/总结/文章/PRD 格式，包含章节、段落、要点和表格。',
    parameters: {
      topic: { type: 'string', required: true, description: '文档主题/用户需求' },
      audience: { type: 'string', description: '目标读者（可选）' },
      format: {
        type: 'string',
        description: '文档格式：report/plan/summary/article/prd（默认 report）',
        enum: ['report', 'plan', 'summary', 'article', 'prd'],
      },
    },
    output: {
      schema: {
        type: 'object',
        additionalProperties: true,
        properties: {
          title: { type: 'string', required: true },
          sectionCount: { type: 'integer', required: true },
          filePath: { type: 'string', required: true },
          markdown: { type: 'string' },
        },
      },
      render: (_args, value) => [{
        type: 'text',
        text: `已生成 Word 文档《${value.title}》，共 ${value.sectionCount} 章，文件已保存到 ${value.filePath}`,
      }],
    },
    isConcurrencySafe: () => false,
    async execute(args, exec) {
      const topic = args.topic
      if (!topic?.trim()) {
        throw new Error('topic 不能为空')
      }
      const format = args.format ?? 'report'
      const audience = args.audience ?? ''
      const scene = inferScene(topic)
      const guide = sceneGuide(scene, 'doc')
      const fmtGuide = formatGuide(format, scene)
      const office = ctx.office

      office.emitProgress('running', '生成 Word 文档', `正在为《${topic}》生成内容...`)

      const userPrompt = [
        `需求：${topic}`,
        audience ? `目标读者：${audience}` : '',
        `文档格式：${format}`,
        `格式要求：${fmtGuide}`,
        `场景偏好：${guide}`,
        '',
        '请生成一份完整的结构化文档，只返回 JSON。',
      ].filter(Boolean).join('\n')

      let doc: DocOutput
      try {
        const json = await llmGenerateJson(ctx, DOC_SYSTEM_PROMPT, userPrompt)
        doc = json as unknown as DocOutput
        if (!doc.title || !Array.isArray(doc.sections)) {
          throw new Error('LLM 返回的文档格式不正确')
        }
      } catch (err) {
        // Fallback
        doc = {
          title: topic.slice(0, 32),
          format,
          sections: [
            {
              heading: '需求原文',
              heading_level: 1,
              paragraphs: [topic],
              bullets: [],
            },
            {
              heading: '待补充章节',
              heading_level: 1,
              paragraphs: [`文档生成过程中出现错误：${err instanceof Error ? err.message : String(err)}`],
              bullets: ['请补充核心内容', '请补充结构化说明'],
            },
          ],
          markdown: '',
        }
      }

      // Generate markdown
      doc.markdown = sectionsToMarkdown(doc)

      // Render to .docx
      office.emitProgress('running', '导出 Word', '正在生成 .docx 文件...')
      const filePath = await renderDocx(ctx, doc)

      return {
        title: doc.title,
        sectionCount: doc.sections.length,
        filePath,
        markdown: doc.markdown,
      }
    },
  }))

  // ── md_generate ───────────────────────────────────────────────────────────
  const MD_SYSTEM_PROMPT = `你是资深技术写作者。只输出严格 JSON，不要 markdown 代码块。
返回格式：
{
  "title": "文档标题",
  "markdown": "# 标题\\n\\n## 小节\\n- 要点",
  "summary": "一句话说明内容价值"
}
要求：
- markdown 必须是完整、可直接保存为 .md 文件的正文
- 使用标准 Markdown 语法，至少包含 4 个二级标题
- 适当使用列表、表格、引用、任务列表或代码块提升可读性
- 内容要具体，不要只写提纲，也不要输出"待补充"
- 不要输出 JSON 之外的任何解释`

  ctx.tools.register(defineTool({
    name: 'md_generate',
    description: '生成 Markdown 文档：适合知识库、README、说明文档、会议纪要、调研整理、操作手册等纯文本结构化内容，可下载 .md 文件。',
    parameters: {
      topic: { type: 'string', required: true, description: 'Markdown 文档主题/用户需求' },
      style: {
        type: 'string',
        description: '文档风格：knowledge_base/readme/guide/notes/research（默认 knowledge_base）',
        enum: ['knowledge_base', 'readme', 'guide', 'notes', 'research'],
      },
      audience: { type: 'string', description: '目标读者（可选）' },
    },
    output: {
      schema: {
        type: 'object',
        additionalProperties: true,
        properties: {
          title: { type: 'string', required: true },
          filePath: { type: 'string', required: true },
          markdown: { type: 'string', required: true },
        },
      },
      render: (_args, value) => [{
        type: 'text',
        text: `已生成 Markdown 文档《${value.title}》，文件已保存到 ${value.filePath}`,
      }],
    },
    isConcurrencySafe: () => false,
    async execute(args, exec) {
      const topic = args.topic
      if (!topic?.trim()) {
        throw new Error('topic 不能为空')
      }
      const style = args.style ?? 'knowledge_base'
      const audience = args.audience ?? ''
      const scene = inferScene(topic)
      const guide = sceneGuide(scene, 'md')
      const office = ctx.office

      const styleGuides: Record<string, string> = {
        readme: '输出 README 风格文档，优先包含简介、核心能力、快速开始、使用步骤、目录结构、示例、注意事项。',
        guide: '输出操作指南，优先包含适用场景、前置条件、步骤说明、关键截图说明位、常见问题和注意事项。',
        notes: '输出会议纪要或整理笔记，优先包含会议背景、关键结论、待办事项、责任人、时间点和后续跟进。',
        research: '输出调研整理文档，优先包含背景、信息来源、关键信息摘要、对比、结论和建议。',
        knowledge_base: '输出知识库风格文档，优先包含概览、核心说明、要点列表、示例、FAQ 和补充说明。',
      }

      office.emitProgress('running', '生成 Markdown 文档', `正在整理《${topic}》的内容...`)

      const userPrompt = [
        `请根据以下需求生成一份适合保存为 Markdown 文件的正式内容。`,
        `要求：${styleGuides[style] ?? styleGuides.knowledge_base}`,
        `场景偏好：${guide}`,
        audience ? `目标读者：${audience}` : '',
        `用户需求：${topic}`,
      ].filter(Boolean).join('\n')

      let result: { title: string; markdown: string; summary?: string }
      try {
        const json = await llmGenerateJson(ctx, MD_SYSTEM_PROMPT, userPrompt)
        result = json as unknown as typeof result
      } catch (err) {
        result = {
          title: topic.slice(0, 32),
          markdown: `# ${topic}\n\n## 待补充\n\n- 请补充核心内容\n- 请补充结构化说明\n\n> 当前为降级草稿：${err instanceof Error ? err.message : String(err)}`,
        }
      }

      // Write .md file
      const outputDir = resolve(process.cwd(), 'output')
      await mkdir(outputDir, { recursive: true })
      const filename = `${result.title.replace(/[^\w\u4e00-\u9fff]/g, '_')}_${Date.now()}.md`
      const filePath = resolve(outputDir, filename)
      await writeFile(filePath, result.markdown, 'utf-8')

      return {
        title: result.title,
        filePath,
        markdown: result.markdown,
      }
    },
  }))
}

export const name = 'tool-doc'
export const inject = ['tools', 'office']
