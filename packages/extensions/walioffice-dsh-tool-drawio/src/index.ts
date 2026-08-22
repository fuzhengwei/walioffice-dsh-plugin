/**
 * Draw.io diagram generation tool: drawio_generate
 * Produces draw.io XML that can be opened in draw.io or embedded in the UI.
 * 
 * @module @walioffice/dsh-tool-drawio
 */

import type { Context } from '@deepseek-ai/cordis'
import { defineTool } from '@deepseek-ai/dsh-tools'
import { createTextOfficeDownload, inferScene, llmGenerateJson, llmGenerateText, resolveOfficeService, sceneGuide } from '@walioffice/dsh-office'

// ── System prompt ───────────────────────────────────────────────────────────

const SYSTEM_PROMPT = `你是 draw.io 图表设计专家。只输出 draw.io XML（mxGraphModel 或 mxfile 格式），不要 markdown 代码块，不要解释。

XML 格式示例：
<mxGraphModel dx="800" dy="600" grid="1" gridSize="10" guides="1" tooltips="1" connect="1" arrows="1" fold="1" page="1" pageScale="1" pageWidth="850" pageHeight="600" math="0" shadow="0">
  <root>
    <mxCell id="0"/>
    <mxCell id="1" parent="0"/>
    <mxCell id="2" value="节点1" style="rounded=1;whiteSpace=wrap;html=1;fillColor=#dae8fc;strokeColor=#6c8ebf;" vertex="1" parent="1">
      <mxGeometry x="100" y="100" width="120" height="60" as="geometry"/>
    </mxCell>
  </root>
</mxGraphModel>

要求：
- 使用合理的布局坐标，节点不重叠
- 用箭头连接表示关系
- 使用不同颜色和样式区分节点类型
- 使用中文标签
- 图表完整、清晰，适合继续编辑
- flowchart 体现步骤流转、条件分支和结果
- architecture 体现层次、模块边界、依赖关系和数据流
- swimlane 体现角色、职责和跨角色流转
- topology 体现节点、网络连接和部署关系
- er 体现实体、字段和关系
- mindmap 体现主题、分支和层级
- 如果需求信息不足，请补足合理的模块、阶段、角色或系统组件
- 最终只返回完整 XML，从 <mxGraphModel> 或 <mxfile> 开始，到对应闭合标签结束`

function escapeXml(value: string): string {
  return value
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;')
}

function compactTitle(topic: string): string {
  return topic.replace(/\s+/g, ' ').trim().slice(0, 40) || 'WaLiOffice 图表'
}

function fallbackDrawioXml(topic: string, diagramType: string): string {
  const title = escapeXml(compactTitle(topic))
  const presets: Record<string, string[]> = {
    architecture: ['用户与入口', '业务应用层', '核心服务层', '数据与基础设施'],
    swimlane: ['需求方', '业务团队', '技术团队', '交付结果'],
    topology: ['客户端', '网关/入口', '服务节点', '数据存储'],
    er: ['用户实体', '业务实体', '记录实体', '结果实体'],
    mindmap: [compactTitle(topic), '核心目标', '关键模块', '落地路径'],
    flowchart: ['需求输入', '分析与规划', '方案执行', '结果反馈'],
  }
  const labels = presets[diagramType] ?? presets.flowchart!
  const colors = ['#dae8fc', '#d5e8d4', '#fff2cc', '#f8cecc']
  const strokes = ['#6c8ebf', '#82b366', '#d6b656', '#b85450']
  const nodes = labels.map((label, index) => {
    const x = 90 + index * 210
    const y = diagramType === 'mindmap' && index > 0 ? (index % 2 ? 90 : 250) : 170
    const width = index === 0 && diagramType === 'mindmap' ? 170 : 145
    return `<mxCell id="node-${index + 1}" value="${escapeXml(label)}" style="rounded=1;whiteSpace=wrap;html=1;fillColor=${colors[index % colors.length]};strokeColor=${strokes[index % strokes.length]};fontColor=#1f2937;" vertex="1" parent="1"><mxGeometry x="${x}" y="${y}" width="${width}" height="64" as="geometry"/></mxCell>`
  }).join('')
  const edges = labels.slice(1).map((_label, index) => {
    const source = diagramType === 'mindmap' ? 'node-1' : `node-${index + 1}`
    const target = `node-${index + 2}`
    return `<mxCell id="edge-${index + 1}" value="" style="edgeStyle=orthogonalEdgeStyle;rounded=1;orthogonalLoop=1;jetSize=auto;html=1;endArrow=block;strokeColor=#64748b;" edge="1" parent="1" source="${source}" target="${target}"><mxGeometry relative="1" as="geometry"/></mxCell>`
  }).join('')
  return `<mxfile host="embed.diagrams.net"><diagram name="${title}"><mxGraphModel dx="1200" dy="700" grid="1" gridSize="10" guides="1" tooltips="1" connect="1" arrows="1" fold="1" page="1" pageScale="1" pageWidth="1169" pageHeight="827" math="0" shadow="0"><root><mxCell id="0"/><mxCell id="1" parent="0"/><mxCell id="title" value="${title}" style="text;html=1;strokeColor=none;fillColor=none;fontSize=18;fontStyle=1;fontColor=#334155;" vertex="1" parent="1"><mxGeometry x="90" y="35" width="700" height="40" as="geometry"/></mxCell>${nodes}${edges}</root></mxGraphModel></diagram></mxfile>`
}

function decodeXmlEntities(value: string): string {
  return value
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
}

function extractDrawioXml(content: string): string | null {
  const source = decodeXmlEntities(content.trim())
    .replace(/^```(?:xml|drawio)?\s*/i, '')
    .replace(/\s*```$/i, '')
    .trim()
  const fileMatch = source.match(/<mxfile\b[\s\S]*?<\/mxfile\s*>/i)
  if (fileMatch?.[0]) return fileMatch[0].trim()
  const modelMatch = source.match(/<mxGraphModel\b[\s\S]*?<\/mxGraphModel\s*>/i)
  return modelMatch?.[0]?.trim() ?? null
}

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
          download: { type: 'object' },
        },
      },
      render: (_args, value) => [{
        type: 'text',
        text: `已生成 draw.io 图表《${value.title}》（${value.diagramType}）`,
      }],
      presentationMeta: (_args, value) => ({ kind: 'drawio', ...value }),
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
      const office = resolveOfficeService(ctx)

      office.emitProgress('running', '生成图表', `正在为《${topic}》生成 ${diagramType} 图...`)

      const userPrompt = [
        `需求：${topic}`,
        `图表类型：${diagramType}`,
        `场景偏好：${guide}`,
        '',
        '请生成一个清晰、专业的 draw.io 图表。',
      ].filter(Boolean).join('\n')

      // For drawio, we use the raw LLM text output (not JSON)
      let xml: string | null = null
      try {
        const content = await llmGenerateText(ctx, SYSTEM_PROMPT, userPrompt)
        xml = extractDrawioXml(content)
      } catch {
        xml = null
      }
      if (!xml) {
        office.emitProgress('running', '使用图表模板兜底', '模型未返回完整 XML，已使用内置可编辑图表继续生成。')
        xml = fallbackDrawioXml(topic, diagramType)
      }

      // Extract title from topic
      const title = topic.slice(0, 30)

      return {
        title,
        diagramType,
        xml,
        download: createTextOfficeDownload(`${title || 'drawio'}.drawio`, xml, 'application/xml;charset=utf-8'),
      }
    },
  }))
}

export const name = 'tool-drawio'
export const inject = ['tools', 'office']
