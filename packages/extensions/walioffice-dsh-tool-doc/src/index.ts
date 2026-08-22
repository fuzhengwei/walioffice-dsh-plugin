/**
 * Document generation tools: doc_generate (Word) and md_generate (Markdown).
 * 
 * @module @walioffice/dsh-tool-doc
 */

import type { Context } from '@deepseek-ai/cordis'
import { defineTool } from '@deepseek-ai/dsh-tools'
import { createOfficeDownload, inferScene, llmGenerateJson, resolveOfficeService, sceneGuide } from '@walioffice/dsh-office'
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

interface DocPreviewSection {
  heading: string
  headingLevel: number
  paragraphs: string[]
  bullets: string[]
  table?: DocTable
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

function compactTitle(topic: string): string {
  const normalized = topic.replace(/\s+/g, ' ').trim()
  const head = normalized.split(/[：:。！？]/)[0]?.trim() || normalized
  return head.slice(0, 36) || 'WaLiOffice 文档'
}

function cleanDocText(value: unknown): string {
  return typeof value === 'string' ? value.replace(/\s+/g, ' ').trim() : ''
}

function isTopicEcho(text: string, topicText: string): boolean {
  if (!text || !topicText) return false
  if (text === topicText) return true
  return topicText.length >= 24 && text.length >= topicText.length && text.startsWith(topicText)
}

function normalizeDocOutput(value: unknown, topic: string, format: string): DocOutput {
  if (typeof value !== 'object' || value === null) throw new Error('LLM 返回的文档格式不正确')
  const source = value as { title?: unknown; sections?: unknown }
  if (!Array.isArray(source.sections)) throw new Error('LLM 返回的文档格式不正确')

  const topicText = cleanDocText(topic)
  const sections = source.sections.map((item, index) => {
    const section = item && typeof item === 'object' ? item as Record<string, unknown> : {}
    const paragraphs = Array.isArray(section.paragraphs)
      ? section.paragraphs.map(cleanDocText).filter(text => text && !isTopicEcho(text, topicText)).filter((text, textIndex, all) => all.indexOf(text) === textIndex).slice(0, 4)
      : []
    const bullets = Array.isArray(section.bullets)
      ? section.bullets.map(cleanDocText).filter(Boolean).filter((text, textIndex, all) => all.indexOf(text) === textIndex).slice(0, 8)
      : []
    const rawTable = section.table && typeof section.table === 'object' ? section.table as Record<string, unknown> : undefined
    const headers = Array.isArray(rawTable?.headers) ? rawTable.headers.map(cleanDocText).filter(Boolean) : []
    const rows = Array.isArray(rawTable?.rows)
      ? rawTable.rows.filter(row => Array.isArray(row)).map(row => (row as unknown[]).map(cleanDocText)).filter(row => row.length > 0)
      : []
    return {
      heading: cleanDocText(section.heading) || `第 ${index + 1} 部分`,
      heading_level: typeof section.heading_level === 'number' ? Math.max(1, Math.min(3, Math.round(section.heading_level))) : 1,
      paragraphs,
      bullets,
      ...(headers.length > 0 && rows.length > 0 ? { table: { headers, rows } } : {}),
    }
  }).filter(section => section.paragraphs.length > 0 || section.bullets.length > 0 || section.table)

  if (sections.length < 3) throw new Error('LLM 返回的文档章节不足')
  const rawTitle = cleanDocText(source.title)
  const title = rawTitle && rawTitle.length <= 48 && !rawTitle.includes(topicText) ? rawTitle : compactTitle(topic)
  return { title, sections, format, markdown: '' }
}

function fallbackDoc(topic: string, format: string, audience: string): DocOutput {
  const audienceText = audience || '相关管理者、执行人员与协作团队'
  const subject = compactTitle(topic)
  return {
    title: subject,
    format,
    markdown: '',
    sections: [
      { heading: '摘要', heading_level: 1, paragraphs: [`本文围绕“${subject}”展开，面向${audienceText}，系统梳理背景、目标、重点内容、实施安排与保障措施，为后续讨论、执行和复盘提供统一依据。`], bullets: ['明确主题范围与核心目标', '形成可执行的工作框架', '为后续协作和评估提供依据'] },
      { heading: '背景与目标', heading_level: 1, paragraphs: [`当前需要围绕“${subject}”建立清晰、完整且可落地的工作方案。通过统一目标、边界和协作方式，可以减少信息偏差，提升执行效率，并让阶段性成果能够被持续追踪和复用。`], bullets: ['明确建设背景与现实需求', '统一参与方对目标和范围的理解', '定义可观察、可复盘的阶段成果'] },
      { heading: '核心内容', heading_level: 1, paragraphs: [`围绕主题推进时，应优先处理影响结果的关键事项，并将复杂工作拆分为清晰的任务单元。每项任务都需要明确负责人、输入条件、交付结果和检查方式，确保方案能够从讨论顺利进入执行。`], bullets: ['梳理关键任务与优先级', '明确角色分工和协作边界', '沉淀过程资料和交付标准'], table: { headers: ['工作模块', '重点内容', '交付结果'], rows: [['目标定义', '统一目标、范围与优先级', '目标说明与范围清单'], ['方案设计', '拆分任务、流程和资源', '执行方案与任务列表'], ['结果评估', '跟踪进度、质量和效果', '阶段复盘与改进建议']] } },
      { heading: '实施安排', heading_level: 1, paragraphs: ['实施阶段建议采用“准备—执行—检查—优化”的闭环方式推进。先完成基础信息和资源准备，再按优先级执行重点任务，并通过阶段检查及时识别偏差，最后将有效经验沉淀为后续工作的标准做法。'], bullets: ['准备阶段：确认人员、资料、资源和时间安排', '执行阶段：按任务清单推进并记录过程', '检查阶段：对照目标检查质量和进度', '优化阶段：根据结果调整方案和资源'] },
      { heading: '风险与保障', heading_level: 1, paragraphs: [`为保证“${subject}”相关工作稳定推进，需要提前识别信息不足、资源变动、协作延迟和质量偏差等风险，并设置明确的沟通、检查和升级机制。所有重要结论应保留记录，便于追踪责任和持续改进。`], bullets: ['建立定期同步和异常反馈机制', '为关键任务设置备选方案', '对重要交付物执行复核和版本管理'], table: { headers: ['风险项', '影响', '应对措施'], rows: [['需求变化', '范围和计划出现偏差', '建立变更记录与评审机制'], ['资源不足', '关键任务延期', '提前识别依赖并准备替代资源'], ['质量波动', '交付结果不稳定', '设置检查清单和验收标准']] } },
      { heading: '结论与下一步', heading_level: 1, paragraphs: [`综合来看，“${subject}”应以明确目标为起点，以任务拆解和过程协作为抓手，以阶段检查和结果复盘形成闭环。下一步建议先确认范围和负责人，再完成首轮任务排期，并根据实际反馈持续优化执行方案。`], bullets: ['确认最终目标、范围和优先级', '落实负责人、时间节点与交付标准', '安排首轮检查并形成复盘记录'] },
    ],
  }
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

function normalizeMarkdownOutput(value: unknown, topic: string): { title: string; markdown: string; summary?: string } {
  if (typeof value !== 'object' || value === null) throw new Error('LLM 返回的 Markdown 格式不正确')
  const source = value as { title?: unknown; markdown?: unknown; summary?: unknown }
  const markdown = typeof source.markdown === 'string' ? source.markdown.trim() : ''
  if (!markdown || markdown.length < 80) throw new Error('LLM 返回的 Markdown 内容不足')
  const rawTitle = cleanDocText(source.title)
  const title = rawTitle && rawTitle.length <= 48 && !rawTitle.includes(cleanDocText(topic)) ? rawTitle : compactTitle(topic)
  return {
    title,
    markdown,
    ...(typeof source.summary === 'string' && source.summary.trim() ? { summary: source.summary.trim() } : {}),
  }
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
          download: { type: 'object' },
          format: { type: 'string' },
          markdown: { type: 'string' },
          sections: { type: 'array' },
        },
      },
      render: (_args, value) => [{
        type: 'text',
        text: `已生成 Word 文档《${value.title}》，共 ${value.sectionCount} 章，文件已保存到 ${value.filePath}`,
      }],
      presentationMeta: (_args, value) => ({ kind: 'doc', ...value }),
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
      const office = resolveOfficeService(ctx)

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
        doc = normalizeDocOutput(json, topic, format)
      } catch (err) {
        office.emitProgress('running', '使用文档模板兜底', 'LLM 未返回可解析内容，已使用内置文档结构继续生成。')
        doc = fallbackDoc(topic, format, audience)
        void err
      }

      // Generate markdown
      doc.markdown = sectionsToMarkdown(doc)

      // Render to .docx
      office.emitProgress('running', '导出 Word', '正在生成 .docx 文件...')
      const filePath = await renderDocx(ctx, doc)
      const download = await createOfficeDownload(filePath, 'application/vnd.openxmlformats-officedocument.wordprocessingml.document')

      return {
        title: doc.title,
        sectionCount: doc.sections.length,
        filePath,
        download,
        format,
        markdown: doc.markdown,
        sections: doc.sections.map((section): DocPreviewSection => ({
          heading: section.heading,
          headingLevel: section.heading_level,
          paragraphs: section.paragraphs.slice(0, 3),
          bullets: section.bullets.slice(0, 6),
          ...(section.table ? {
            table: {
              headers: section.table.headers,
              rows: section.table.rows.slice(0, 8),
            },
          } : {}),
        })) as any,
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
    description: '仅生成 Markdown（.md）文件，适合知识库、README、说明文档、会议纪要、调研整理、操作手册；不要用于 Word 或 .docx，Word 必须调用 doc_generate。',
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
          download: { type: 'object' },
          markdown: { type: 'string', required: true },
        },
      },
      render: (_args, value) => [{
        type: 'text',
        text: `已生成 Markdown 文档《${value.title}》，文件已保存到 ${value.filePath}`,
      }],
      presentationMeta: (_args, value) => ({ kind: 'markdown', ...value }),
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
      const office = resolveOfficeService(ctx)

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
        result = normalizeMarkdownOutput(json, topic)
      } catch (err) {
        const fallback = fallbackDoc(topic, 'article', audience)
        result = {
          title: fallback.title,
          markdown: sectionsToMarkdown(fallback),
          summary: '已使用 WaLiOffice 内置结构完成文档内容生成。',
        }
        void err
      }

      // Write .md file
      const outputDir = resolve(process.cwd(), 'output')
      await mkdir(outputDir, { recursive: true })
      const filename = `${result.title.replace(/[^\w\u4e00-\u9fff]/g, '_')}_${Date.now()}.md`
      const filePath = resolve(outputDir, filename)
      await writeFile(filePath, result.markdown, 'utf-8')
      const download = await createOfficeDownload(filePath, 'text/markdown;charset=utf-8')

      return {
        title: result.title,
        filePath,
        download,
        markdown: result.markdown,
      }
    },
  }))
}

export const name = 'tool-doc'
export const inject = ['tools', 'office']
