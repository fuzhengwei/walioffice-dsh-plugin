/**
 * Chart generation tool: chart_generate
 * Produces ECharts-compatible JSON data for frontend rendering.
 * 
 * @module @walioffice/dsh-tool-chart
 */

import type { Context } from '@deepseek-ai/cordis'
import { defineTool } from '@deepseek-ai/dsh-tools'
import { llmGenerateJson, resolveOfficeService } from '@walioffice/dsh-office'

// ── Types ───────────────────────────────────────────────────────────────────

export interface ChartOutput {
  title: string
  summary: string
  chart_type: string
  labels: string[]
  values: number[]
  series_name: string
}

type ChartType = 'line' | 'bar' | 'pie' | 'gauge' | 'funnel' | 'scatter'

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
- 用户指定 chart_type 时必须优先遵循
- pie/funnel/gauge 的 values 必须为非负数
- 用户给数据时贴合业务场景补合理示例
- chart_type 从 line/bar/pie/gauge/funnel/scatter 中选择`

export function apply(ctx: Context): void {
  ctx.tools.register(defineTool({
    name: 'chart_generate',
    description: '生成可直接渲染的 ECharts 图表数据。用户提到图表、可视化、趋势图、折线图、柱状图、饼图、占比/分布、排名、漏斗、仪表盘、散点或“用 ECharts 展示”时调用；没有真实数据时自动生成贴合主题的示例数据，LLM 返回异常时使用内置模板继续交付。',
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
      presentationMeta: (_args, value) => ({ kind: 'chart', ...value }),
    },
    isConcurrencySafe: () => false,
    async execute(args, exec) {
      const topic = args.topic
      if (!topic?.trim()) {
        throw new Error('topic 不能为空')
      }
      const chartType = normalizeChartType(args.chart_type, inferChartType(topic))
      const office = resolveOfficeService(ctx)

      office.emitProgress('running', '生成图表', `正在为《${topic}》生成数据...`)

      const userPrompt = [
        `需求：${topic}`,
        `图表类型：${chartType}`,
        '',
        '请生成专业的图表数据，只返回 JSON。',
      ].filter(Boolean).join('\n')

      let output: ChartOutput
      try {
        const json = await llmGenerateJson(ctx, SYSTEM_PROMPT, userPrompt)
        output = normalizeChartOutput(json, topic, chartType)
      } catch (error) {
        office.emitProgress('running', '使用图表模板兜底', 'LLM 未返回可解析内容，已使用内置图表数据继续生成。')
        output = fallbackChart(topic, chartType)
        void error
      }

      // Build ECharts option
      const echartsOption = buildEchartsOption(output, normalizeChartType(output.chart_type, chartType))

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

function normalizeChartType(value: unknown, fallback: ChartType): ChartType {
  return value === 'line' || value === 'bar' || value === 'pie' || value === 'gauge' || value === 'funnel' || value === 'scatter'
    ? value
    : fallback
}

function inferChartType(topic: string): ChartType {
  if (/占比|比例|分布|构成|份额/.test(topic)) return 'pie'
  if (/趋势|变化|走势|增长/.test(topic)) return 'line'
  if (/漏斗/.test(topic)) return 'funnel'
  if (/仪表盘|仪表|完成率|达成率/.test(topic)) return 'gauge'
  return 'bar'
}

function normalizeChartOutput(value: unknown, topic: string, chartType: ChartType): ChartOutput {
  if (typeof value !== 'object' || value === null) {
    throw new Error('LLM 返回的图表数据不是对象')
  }

  const raw = value as {
    title?: unknown
    summary?: unknown
    chart_type?: unknown
    chartType?: unknown
    labels?: unknown
    values?: unknown
    series_name?: unknown
    seriesName?: unknown
  }
  const labels = Array.isArray(raw.labels)
    ? raw.labels.filter(item => typeof item === 'string' || typeof item === 'number').map(String)
    : []
  const values = Array.isArray(raw.values)
    ? raw.values.map(item => typeof item === 'number' ? item : typeof item === 'string' && item.trim() ? Number(item) : Number.NaN)
    : []

  const normalizedChartType = normalizeChartType(raw.chart_type ?? raw.chartType, chartType)
  const requiresNonNegative = normalizedChartType === 'pie' || normalizedChartType === 'funnel' || normalizedChartType === 'gauge'
  if (labels.length === 0 || labels.length !== values.length || values.some(value => !Number.isFinite(value)) || (requiresNonNegative && values.some(value => value < 0))) {
    throw new Error('LLM 返回的图表数据格式不正确：labels 和 values 长度或数值不合法')
  }

  return {
    title: typeof raw.title === 'string' && raw.title.trim() ? raw.title : topic,
    summary: typeof raw.summary === 'string' ? raw.summary : `围绕“${topic}”生成的${chartType}图表。`,
    chart_type: normalizedChartType,
    labels,
    values,
    series_name: typeof raw.series_name === 'string'
      ? raw.series_name
      : typeof raw.seriesName === 'string' ? raw.seriesName : '数据',
  }
}

function fallbackChart(topic: string, chartType: ChartType): ChartOutput {
  const isLoginTopic = /登录|登陆|认证|登录状态|login/i.test(topic)
  const isAiTopic = /AI技术|人工智能|大语言模型|计算机视觉|自然语言处理|机器学习|强化学习|知识图谱|多模态\s*AI|AI\s*Agent/i.test(topic)
  const labels = isLoginTopic
    ? ['登录成功', '密码错误', '账号锁定', '验证码错误', '账号不存在']
    : isAiTopic
      ? ['大语言模型', '计算机视觉', '自然语言处理', '机器学习', '强化学习', '知识图谱', '多模态AI', 'AI Agent']
      : ['类别一', '类别二', '类别三', '类别四']
  const values = isLoginTopic ? [68, 12, 5, 9, 6] : isAiTopic ? [25, 18, 15, 14, 8, 7, 8, 5] : [100, 80, 60, 40]

  return {
    title: isLoginTopic ? '登录状态分布' : isAiTopic ? 'AI技术领域占比分布' : topic,
    summary: isLoginTopic
      ? '登录成功占比最高，失败原因主要集中在密码错误和验证码错误。'
      : isAiTopic
        ? '大语言模型和计算机视觉占比较高，其他 AI 技术方向保持多元分布。'
        : `围绕“${topic}”整理的示例数据，可继续替换为真实业务数据。`,
    chart_type: chartType,
    labels,
    values,
    series_name: '数量（%）',
  }
}

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
