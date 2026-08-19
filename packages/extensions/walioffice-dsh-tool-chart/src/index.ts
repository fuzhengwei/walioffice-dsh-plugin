/**
 * Chart generation tool: chart_generate
 * Produces ECharts-compatible JSON data for frontend rendering.
 * 
 * @module @walioffice/dsh-tool-chart
 */

import type { Context } from '@deepseek-ai/cordis'
import { defineTool } from '@deepseek-ai/dsh-tools'
import { llmGenerateJson } from '@walioffice/dsh-office'

// ── Types ───────────────────────────────────────────────────────────────────

export interface ChartOutput {
  title: string
  summary: string
  chart_type: string
  labels: string[]
  values: number[]
  series_name: string
}

// ── System prompt ───────────────────────────────────────────────────────────

const SYSTEM_PROMPT = `你是数据分析专家。只输出严格 JSON，不要 markdown 代码块。
返回格式：
{
  "title": "图表标题",
  "summary": "一句话说明图表洞察",
  "chart_type": "bar",
  "labels": ["标签1","标签2","标签3"],
  "values": [100, 200, 150],
  "series_name": "系列名"
}
要求：
- labels 和 values 长度必须一致
- values 为纯数字，不含单位
- 用户给数据时贴合业务场景补合理示例
- chart_type 从 line/bar/pie/gauge/funnel/scatter 中选择`

export function apply(ctx: Context): void {
  ctx.tools.register(defineTool({
    name: 'chart_generate',
    description: '生成 ECharts 图表数据：支持趋势(line)、对比(bar)、占比(pie)、排名(bar)、漏斗(funnel)、仪表盘(gauge)、散点(scatter)，适合普通对话中的数据可视化。',
    parameters: {
      topic: { type: 'string', required: true, description: '图表主题/用户需求' },
      chart_type: {
        type: 'string',
        description: '图表类型：line/bar/pie/gauge/funnel/scatter（默认 bar）',
        enum: ['line', 'bar', 'pie', 'gauge', 'funnel', 'scatter'],
      },
    },
    output: {
      schema: {
        type: 'object',
        additionalProperties: true,
        properties: {
          title: { type: 'string', required: true },
          chartType: { type: 'string', required: true },
          labels: { type: 'array', items: { type: 'string' }, required: true },
          values: { type: 'array', items: { type: 'number' }, required: true },
          seriesName: { type: 'string', required: true },
          summary: { type: 'string' },
          echartsOption: { type: 'json', required: true },
        },
      },
      render: (_args, value) => [{
        type: 'text',
        text: `已生成图表《${value.title}》（${value.chartType}），共 ${value.labels.length} 个数据点`,
      }],
    },
    isConcurrencySafe: () => false,
    async execute(args, exec) {
      const topic = args.topic
      if (!topic?.trim()) {
        throw new Error('topic 不能为空')
      }
      const chartType = args.chart_type ?? 'bar'
      const office = ctx.office

      office.emitProgress('running', '生成图表', `正在为《${topic}》生成数据...`)

      const userPrompt = [
        `需求：${topic}`,
        `图表类型：${chartType}`,
        '',
        '请生成专业的图表数据，只返回 JSON。',
      ].filter(Boolean).join('\n')

      const json = await llmGenerateJson(ctx, SYSTEM_PROMPT, userPrompt)
      const output = json as unknown as ChartOutput

      if (!output.labels || !output.values || output.labels.length !== output.values.length) {
        throw new Error('LLM 返回的图表数据格式不正确：labels 和 values 长度不一致')
      }

      // Build ECharts option
      const echartsOption = buildEchartsOption(output, chartType)

      return {
        title: output.title,
        chartType: output.chart_type ?? chartType,
        labels: output.labels,
        values: output.values,
        seriesName: output.series_name ?? '数据',
        summary: output.summary,
        echartsOption: echartsOption as any,
      }
    },
  }))
}

// ── ECharts option builder ──────────────────────────────────────────────────

function buildEchartsOption(output: ChartOutput, chartType: string): Record<string, unknown> {
  const base = {
    title: { text: output.title, left: 'center' },
    tooltip: { trigger: 'axis' },
    legend: { bottom: 0 },
  }

  switch (chartType) {
    case 'pie':
      return {
        ...base,
        tooltip: { trigger: 'item' },
        series: [{
          type: 'pie',
          radius: '60%',
          data: output.labels.map((label, i) => ({ name: label, value: output.values[i] })),
        }],
      }
    case 'funnel':
      return {
        ...base,
        tooltip: { trigger: 'item' },
        series: [{
          type: 'funnel',
          data: output.labels.map((label, i) => ({ name: label, value: output.values[i] })),
        }],
      }
    case 'gauge':
      return {
        ...base,
        series: [{
          type: 'gauge',
          data: [{ value: output.values[0] ?? 0, name: output.labels[0] ?? '' }],
        }],
      }
    case 'scatter':
      return {
        ...base,
        xAxis: { type: 'value' },
        yAxis: { type: 'value' },
        series: [{
          type: 'scatter',
          name: output.series_name,
          data: output.labels.map((_, i) => [i, output.values[i]]),
        }],
      }
    case 'line':
      return {
        ...base,
        xAxis: { type: 'category', data: output.labels },
        yAxis: { type: 'value' },
        series: [{
          type: 'line',
          name: output.series_name,
          data: output.values,
          smooth: true,
        }],
      }
    default: // bar
      return {
        ...base,
        xAxis: { type: 'category', data: output.labels },
        yAxis: { type: 'value' },
        series: [{
          type: 'bar',
          name: output.series_name,
          data: output.values,
        }],
      }
  }
}

export const name = 'tool-chart'
export const inject = ['tools', 'office']
