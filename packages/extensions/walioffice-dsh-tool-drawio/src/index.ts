/**
 * Draw.io diagram generation tool: drawio_generate
 * Produces draw.io XML that can be opened in draw.io or embedded in the UI.
 * 
 * @module @walioffice/dsh-tool-drawio
 */

import type { Context } from '@deepseek-ai/cordis'
import { defineTool } from '@deepseek-ai/dsh-tools'
import { llmGenerateJson, llmGenerateText, sceneGuide, inferScene } from '@walioffice/dsh-office'

// ── System prompt ───────────────────────────────────────────────────────────

const SYSTEM_PROMPT = `你是流程图/架构图专家。只输出 draw.io XML，不要 markdown 代码块。
要求：
- 使用 <mxGraphModel> 根节点
- 合理布局坐标，节点不重叠
- 箭头正确连接节点
- 颜色区分不同类型节点
- 中文标签
- 输出格式：直接以 <mxGraphModel> 开头，以 </mxGraphModel> 结尾`

export function apply(ctx: Context): void {
  ctx.tools.register(defineTool({
    name: 'drawio_generate',
    description: '生成 draw.io 可编辑图表（流程图/架构图/泳道图/拓扑图/ER图/思维导图），输出 draw.io XML，可在右侧直接渲染和编辑。',
    parameters: {
      topic: { type: 'string', required: true, description: '图表主题/用户需求' },
      diagram_type: {
        type: 'string',
        description: '图表类型：flowchart/architecture/swimlane/topology/er/mindmap（默认 flowchart）',
        enum: ['flowchart', 'architecture', 'swimlane', 'topology', 'er', 'mindmap'],
      },
    },
    output: {
      schema: {
        type: 'object',
        additionalProperties: true,
        properties: {
          title: { type: 'string', required: true },
          diagramType: { type: 'string', required: true },
          xml: { type: 'string', required: true },
        },
      },
      render: (_args, value) => [{
        type: 'text',
        text: `已生成 draw.io 图表《${value.title}》（${value.diagramType}）`,
      }],
    },
    isConcurrencySafe: () => false,
    async execute(args, exec) {
      const topic = args.topic
      if (!topic?.trim()) {
        throw new Error('topic 不能为空')
      }
      const diagramType = args.diagram_type ?? 'flowchart'
      const scene = inferScene(topic)
      const guide = sceneGuide(scene, 'drawio')
      const office = ctx.office

      office.emitProgress('running', '生成图表', `正在为《${topic}》生成 ${diagramType} 图...`)

      const userPrompt = [
        `需求：${topic}`,
        `图表类型：${diagramType}`,
        `场景偏好：${guide}`,
        '',
        '请生成一个清晰、专业的 draw.io 图表。',
      ].filter(Boolean).join('\n')

      // For drawio, we use the raw LLM text output (not JSON)
      const content = await llmGenerateText(ctx, SYSTEM_PROMPT, userPrompt)

      // Clean markdown fences
      let xml = content.trim()
      xml = xml.replace(/^```(?:xml)?\s*/i, '').replace(/\s*```$/i, '')

      // Validate
      if (!xml.includes('<mxGraphModel') && !xml.includes('<mxfile')) {
        throw new Error('LLM 输出不包含有效的 draw.io XML')
      }

      // Extract title from topic
      const title = topic.slice(0, 30)

      return {
        title,
        diagramType,
        xml,
      }
    },
  }))
}

export const name = 'tool-drawio'
export const inject = ['tools', 'office']
