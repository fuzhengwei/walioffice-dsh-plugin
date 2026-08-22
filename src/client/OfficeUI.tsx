import { useEffect, useMemo, useState, type CSSProperties } from 'react'
import type { PropsRuntime } from '@deepseek-ai/dsh-client-ui-slots'
import type { ToolCallViewProps } from '@deepseek-ai/dsh-client-ui-tool/client'
import type {} from '@deepseek-ai/dsh-client-ui-conversation/client'
import {
  OFFICE_ARTIFACT_EVENT,
  OFFICE_MODES,
  OFFICE_PANEL_EVENT,
  artifactFromTool,
  modeDefinition,
  openOfficePanel,
  publishArtifact,
  type OfficeArtifact,
  type OfficeArtifactMeta,
  type OfficeMode,
} from './office-state.ts'

const ARTIFACT_STORAGE = 'walioffice:artifacts:v1'
const MODE_STORAGE = 'walioffice:mode:v1'
const MAX_ARTIFACTS = 24
type OfficeOpenProps = { openOfficeDetails: () => void }
const NATIVE_DETAILS_STYLES = `
.wo-panel{position:relative;z-index:auto;inset:auto;top:auto;right:auto;bottom:auto;width:100%;height:100%;min-width:0;min-height:0;border:0;border-left:1px solid var(--dsw-alias-border-l2,#e5e7eb);border-radius:0;box-shadow:none}
@media(max-width:760px){.wo-panel{position:relative;inset:auto;width:100%;height:100%}}
`
const RICH_PREVIEW_STYLES = `
.wo-rich{padding:0 12px 12px}.wo-meta-grid{display:grid;gap:8px}.wo-kpis{display:grid;grid-template-columns:repeat(3,minmax(0,1fr));gap:8px;margin:0 12px 12px}.wo-kpi{border:1px solid var(--dsw-alias-border-l1,#e5e7eb);border-radius:12px;padding:10px;background:#fff}.wo-kpi strong{display:block;font-size:16px}.wo-kpi span{display:block;font-size:11px;color:var(--dsw-alias-label-tertiary,#6b7280);margin-top:4px}.wo-chart-bars{display:flex;flex-direction:column;gap:8px}.wo-bar-row{display:grid;grid-template-columns:72px minmax(0,1fr) 48px;gap:8px;align-items:center}.wo-bar-row label,.wo-bar-row em{font-size:11px;color:var(--dsw-alias-label-secondary,#4b5563);font-style:normal}.wo-bar-track{height:10px;border-radius:999px;background:#e5eefc;overflow:hidden}.wo-bar-fill{height:100%;border-radius:999px;background:linear-gradient(90deg,#7c3aed,#2563eb)}.wo-chart-svg{width:100%;height:auto;border:1px solid var(--dsw-alias-border-l1,#e5e7eb);border-radius:12px;background:#fff}.wo-pie{width:180px;height:180px;border-radius:999px;margin:0 auto;background:conic-gradient(#7c3aed 0deg,#2563eb 120deg,#06b6d4 240deg,#ec4899 360deg)}.wo-legend{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:6px;margin-top:10px}.wo-legend-item{display:flex;align-items:center;gap:6px;font-size:11px}.wo-dot{width:10px;height:10px;border-radius:999px;display:inline-block}.wo-md{display:flex;flex-direction:column;gap:8px}.wo-md h1,.wo-md h2,.wo-md h3,.wo-md p,.wo-md ul{margin:0}.wo-md h1{font-size:16px}.wo-md h2{font-size:14px}.wo-md h3{font-size:13px}.wo-md p,.wo-md li{font-size:12px;line-height:1.65;color:var(--dsw-alias-label-secondary,#4b5563)}.wo-md ul{padding-left:18px}.wo-table-wrap{overflow:auto;border:1px solid var(--dsw-alias-border-l1,#e5e7eb);border-radius:12px;background:#fff}.wo-table{width:100%;border-collapse:collapse;font-size:12px}.wo-table th,.wo-table td{padding:8px 10px;border-bottom:1px solid var(--dsw-alias-border-l1,#e5e7eb);text-align:left;white-space:nowrap}.wo-table th{background:#f8fafc;color:#334155}.wo-slides{display:flex;flex-direction:column;gap:8px}.wo-slide{border:1px solid var(--dsw-alias-border-l1,#e5e7eb);border-radius:12px;padding:10px;background:#fff}.wo-slide strong{display:block;font-size:12px}.wo-slide span,.wo-slide p{display:block;font-size:11px;color:var(--dsw-alias-label-tertiary,#6b7280);margin:4px 0 0}.wo-media-grid{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:10px}.wo-media-grid img,.wo-video{width:100%;border-radius:12px;border:1px solid var(--dsw-alias-border-l1,#e5e7eb);background:#fff}.wo-xml{margin:0;padding:12px;border-radius:12px;background:var(--dsw-alias-markdown-code-block,#f6f7f9);font:11px/1.5 ui-monospace,SFMono-Regular,Menlo,monospace;white-space:pre-wrap;word-break:break-word}.wo-storyboard{display:flex;flex-direction:column;gap:8px}.wo-story{border:1px solid var(--dsw-alias-border-l1,#e5e7eb);border-radius:12px;padding:10px;background:#fff}.wo-story strong{font-size:12px}.wo-story p,.wo-story code{display:block;font-size:11px;color:var(--dsw-alias-label-secondary,#4b5563);margin-top:4px}.wo-preview-fallback{margin-top:12px}.wo-preview-fallback summary{cursor:pointer;font-size:12px;color:#2563eb}.wo-error-box{margin:12px;padding:12px;border-radius:12px;background:#ecfdf5;color:#166534;font-size:12px;line-height:1.6}.wo-error-raw{margin:12px;padding:12px;border-radius:12px;background:#f8fafc;color:#475569;font:11px/1.55 ui-monospace,SFMono-Regular,Menlo,monospace;white-space:pre-wrap;word-break:break-word}@media(max-width:760px){.wo-kpis,.wo-media-grid,.wo-legend{grid-template-columns:1fr}}
`

const DOCUMENT_PREVIEW_STYLES = `
.wo-preview-headline{display:flex;align-items:center;justify-content:space-between;gap:12px;flex-wrap:wrap}
.wo-preview-headline small{font-size:11px;color:var(--dsw-alias-label-tertiary,#6b7280)}
.wo-file-head{display:flex;align-items:center;justify-content:space-between;gap:8px}
.wo-file-action,.wo-sheet-tab{border:1px solid #bbf7d0;background:#fff;color:#15803d;border-radius:999px;padding:5px 10px;font-size:11px;cursor:pointer}
.wo-file-action:hover,.wo-sheet-tab:hover{background:#f0fdf4}
.wo-meta-fold,.wo-doc-fold{margin:12px;border:1px solid var(--dsw-alias-border-l1,#e5e7eb);border-radius:14px;background:#f8fafc;overflow:hidden}
.wo-meta-fold summary,.wo-doc-fold summary{list-style:none;cursor:pointer;display:flex;align-items:center;justify-content:space-between;gap:10px;padding:12px 14px}
.wo-meta-fold summary::-webkit-details-marker,.wo-doc-fold summary::-webkit-details-marker{display:none}
.wo-meta-fold summary span,.wo-doc-fold summary span{font-size:12px;font-weight:600;color:#0f172a}
.wo-meta-fold summary small,.wo-doc-fold summary small{font-size:11px;color:#64748b}
.wo-meta-fold summary::after,.wo-doc-fold summary::after{content:'展开';font-size:11px;color:#2563eb;flex:0 0 auto}
.wo-meta-fold[open] summary::after,.wo-doc-fold[open] summary::after{content:'收起'}
.wo-meta-fold-body,.wo-doc-fold-body{padding:0 0 12px}
.wo-doc-outline{display:flex;flex-wrap:wrap;gap:8px;margin:0 12px 12px}
.wo-chip{border:1px solid #dbeafe;background:#f8fbff;color:#2563eb;border-radius:999px;padding:6px 10px;font-size:11px}
.wo-doc-surface{padding:12px 12px 0}
.wo-doc-page{border:1px solid #e5e7eb;border-radius:18px;background:linear-gradient(180deg,#fff,#fcfcfd);padding:16px;box-shadow:0 8px 30px rgba(15,23,42,.05)}
.wo-doc-hero{padding-bottom:14px;border-bottom:1px solid #e5e7eb;margin-bottom:14px}
.wo-doc-badge{display:inline-flex;align-items:center;border-radius:999px;background:#ecfdf5;color:#15803d;padding:5px 10px;font-size:11px;font-weight:600}
.wo-doc-hero h1{margin:10px 0 6px;font-size:20px;line-height:1.35;color:#0f172a}
.wo-doc-hero p{margin:0;font-size:12px;line-height:1.7;color:#64748b}
.wo-doc-section{padding:14px 0;border-top:1px dashed #e2e8f0}
.wo-doc-section:first-of-type{border-top:0;padding-top:0}
.wo-doc-section h2,.wo-doc-section h3{margin:0 0 10px;color:#0f172a}
.wo-doc-section h2{font-size:16px}
.wo-doc-section h3{font-size:14px}
.wo-doc-section p{margin:0 0 10px;font-size:12px;line-height:1.75;color:#334155}
.wo-doc-bullets{margin:0;padding-left:18px}
.wo-doc-bullets li{margin:6px 0;font-size:12px;line-height:1.7;color:#334155}
.wo-doc-note{margin:12px 12px 0;padding:12px;border-radius:14px;background:#f8fafc;color:#475569;font-size:12px;line-height:1.7}
.wo-kpis-compact{margin:0 12px 12px}
.wo-sheet-tabs{display:flex;gap:8px;flex-wrap:wrap;margin:0 12px 12px}
.wo-sheet-tab[data-active]{background:#059669;color:#fff;border-color:#059669}
.wo-sheet-card{margin:0 12px 12px;padding:14px;border:1px solid #d1fae5;border-radius:16px;background:linear-gradient(180deg,#f0fdf4,#fff)}
.wo-sheet-card strong{display:block;font-size:14px;color:#065f46}
.wo-sheet-card p{margin:6px 0 0;font-size:12px;line-height:1.7;color:#4b5563}
.wo-sheet-meta{display:flex;gap:8px;flex-wrap:wrap;margin-top:10px}
.wo-sheet-pill{display:inline-flex;align-items:center;border-radius:999px;background:#fff;color:#059669;border:1px solid #a7f3d0;padding:5px 10px;font-size:11px}
.wo-sheet-empty{margin:0 12px 12px;padding:14px;border:1px dashed #cbd5e1;border-radius:14px;color:#64748b;font-size:12px;line-height:1.7;background:#fff}
.wo-table th:first-child,.wo-table td:first-child{position:sticky;left:0;background:inherit}
.wo-table th:first-child{background:#f8fafc}
@media(max-width:760px){
  .wo-preview-headline,.wo-file-head{align-items:flex-start;flex-direction:column}
  .wo-doc-page{padding:14px}
  .wo-doc-surface{padding:10px 10px 0}
  .wo-meta-fold,.wo-doc-fold{margin:10px}
}
`

const LAYOUT_OVERRIDE_STYLES = `
.wo-dock{width:calc(100% - var(--dsh-composer-side-clearance,24px)*2)!important;max-width:var(--dsh-composer-card-max-width,920px)!important;min-width:min(100%,680px)!important;min-height:52px!important;align-self:center!important;margin:0 auto 10px!important;padding:10px 12px!important}
.wo-dock-brand{flex:0 0 auto}
.wo-mode-list{flex:1 1 auto;min-width:0}
.wo-artifacts-button{flex:0 0 auto}
.wo-panel-tabs{padding:8px 14px;gap:10px}
.wo-panel-tabs button{padding:0 14px;height:34px}
.wo-preview-stage,.wo-files-view{min-height:0;flex:1;display:flex;flex-direction:column}
.wo-preview-stage{background:color-mix(in srgb,var(--dsw-alias-bg-base,#fff) 96%,#eff6ff)}
.wo-preview-switcher{display:flex;gap:8px;overflow:auto;padding:10px 14px 0;scrollbar-width:none}
.wo-preview-switcher::-webkit-scrollbar{display:none}
.wo-switch-chip{border:1px solid var(--dsw-alias-border-l1,#e5e7eb);background:#fff;color:var(--dsw-alias-label-secondary,#4b5563);border-radius:999px;padding:7px 12px;font-size:12px;white-space:nowrap;cursor:pointer}
.wo-switch-chip[data-active]{border-color:#93c5fd;background:#eff6ff;color:#1d4ed8}
.wo-preview-stage .wo-preview{padding-top:8px}
.wo-files-view .wo-artifact-list{padding-top:14px;border-bottom:0}
.wo-artifact-list{padding:10px;overflow:auto}
.wo-preview-card{margin:0 auto;width:100%;max-width:100%}
.wo-preview-card[data-doc-preview]{display:flex;flex-direction:column}
.wo-panel-tabs button[data-active]{background:#eff6ff;color:#1d4ed8;box-shadow:inset 0 -2px 0 #3b82f6}
.wo-artifact-row{padding:11px 10px}
.wo-artifact-copy strong{font-size:13px}
.wo-artifact-copy small{font-size:11px}
.wo-mode-button{padding:0 12px;flex:0 0 auto}
.wo-panel{width:min(880px,calc(100vw - 40px))!important;max-width:none!important}
@media(max-width:760px){
  .wo-dock{width:calc(100% - 16px)!important;min-width:0!important;min-height:52px!important;padding:8px 10px!important}
  .wo-panel{width:calc(100vw - 16px)!important}
  .wo-preview-switcher{padding:8px 12px 0}
}
`

const DOCK_STABLE_HEIGHT_STYLES = `
.wo-dock{min-height:52px}
`

export function OfficeDock({ useInput, inputActions, openOfficeDetails }: PropsRuntime<'conversation.input.dock'> & OfficeOpenProps): JSX.Element {
  const input = useInput(state => state)
  const [mode, setMode] = useState<OfficeMode>(() => readMode())

  const selectMode = (nextMode: OfficeMode): void => {
    setMode(nextMode)
    safeStorageSet(MODE_STORAGE, nextMode)
    const definition = modeDefinition(nextMode)
    if (input.draft.trim() === '') inputActions.setDraft(definition.prompt)
  }

  return <>
    <OfficeStyles />
    <div className="wo-dock" aria-label="WaLiOffice 办公工具栏">
      <div className="wo-dock-brand">
        <span className="wo-brand-mark">W</span>
        <span>智能办公</span>
      </div>
      <div className="wo-mode-list">
        {OFFICE_MODES.map(item => <button
          key={item.id}
          type="button"
          className="wo-mode-button"
          data-active={mode === item.id ? 'true' : undefined}
          style={{ '--wo-mode-color': item.color } as CSSProperties}
          onClick={() => selectMode(item.id)}
          title={item.label}
        >
          <span className="wo-mode-dot" />
          {item.shortLabel}
        </button>)}
      </div>
      <button type="button" className="wo-artifacts-button" onClick={() => { openOfficeDetails(); openOfficePanel() }}>
        <span>▤</span> 产物
      </button>
    </div>
  </>
}

export function OfficeToolView({ toolName, callId, block, openFile, inspect, openOfficeDetails }: ToolCallViewProps & OfficeOpenProps): JSX.Element {
  const artifact = useMemo(() => artifactFromTool(toolName, callId, block), [toolName, callId, block])
  const running = !('kind' in block)
  const isError = artifact?.isError ?? false
  const mode = modeDefinition(artifact?.mode ?? modeForToolName(toolName))
  const argsRaw = 'kind' in block ? block.call?.argsRaw ?? '' : block.argsRaw
  const topic = readTopic(argsRaw)

  useEffect(() => {
    if (artifact) publishArtifact(artifact)
  }, [artifact])

  return <>
    <OfficeStyles />
    <article className="wo-tool-card" data-error={isError ? 'true' : undefined}>
      <div className="wo-tool-head">
        <span className="wo-tool-icon" style={{ background: mode.color }}>{toolGlyph(toolName)}</span>
        <div className="wo-tool-heading">
          <strong>{toolTitle(toolName)}</strong>
          <span>{topic || artifact?.summary || '正在准备办公产物'}</span>
        </div>
        <span className="wo-status" data-running={running ? 'true' : undefined} data-error={isError ? 'true' : undefined}>
          {running ? '生成中' : isError ? '失败' : '已完成'}
        </span>
      </div>
      {artifact && <div className="wo-tool-body">
        <p>{artifact.summary}</p>
        {artifact.filePath && <code>{artifact.filePath}</code>}
      </div>}
      <div className="wo-tool-actions">
        {artifact && <button type="button" onClick={() => { openOfficeDetails(); openOfficePanel(artifact.id) }}>查看产物</button>}
        {artifact?.filePath && <button type="button" onClick={() => openFile(artifact.filePath!)}>打开文件</button>}
        {inspect && <button type="button" onClick={inspect}>执行详情</button>}
      </div>
    </article>
  </>
}

export function OfficeDetailsPanel({ closeDetails }: { closeDetails: () => void }): JSX.Element {
  const [tab, setTab] = useState<'preview' | 'files'>('preview')
  const [artifacts, setArtifacts] = useState<OfficeArtifact[]>(() => readArtifacts())
  const [selectedId, setSelectedId] = useState<string | undefined>(() => readArtifacts()[0]?.id)

  useEffect(() => {
    const onPanel = (event: Event): void => {
      const detail = (event as CustomEvent<{ open?: boolean; artifactId?: string }>).detail
      if (detail?.artifactId) {
        setSelectedId(detail.artifactId)
        setTab('preview')
      }
    }
    const onArtifact = (event: Event): void => {
      const artifact = (event as CustomEvent<OfficeArtifact>).detail
      if (!artifact) return
      setArtifacts(current => {
        const next = [artifact, ...current.filter(item => item.id !== artifact.id)].slice(0, MAX_ARTIFACTS)
        safeStorageSet(ARTIFACT_STORAGE, JSON.stringify(next))
        return next
      })
      setSelectedId(artifact.id)
    }
    window.addEventListener(OFFICE_PANEL_EVENT, onPanel)
    window.addEventListener(OFFICE_ARTIFACT_EVENT, onArtifact)
    return () => {
      window.removeEventListener(OFFICE_PANEL_EVENT, onPanel)
      window.removeEventListener(OFFICE_ARTIFACT_EVENT, onArtifact)
    }
  }, [])

  const selected = artifacts.find(item => item.id === selectedId) ?? artifacts[0]

  return <>
    <OfficeStyles />
    <aside className="wo-panel" aria-label="WaLiOffice 智能办公助手">
      <header className="wo-panel-header">
        <div className="wo-panel-title">
          <span className="wo-logo">W</span>
          <div><strong>智能办公助手</strong><span>打开即用，专注办公创作</span></div>
        </div>
        <button type="button" className="wo-close" onClick={closeDetails} aria-label="关闭办公详情栏">×</button>
      </header>
      <nav className="wo-panel-tabs">
        <button type="button" data-active={tab === 'preview' ? 'true' : undefined} onClick={() => setTab('preview')}>产物汇总 <span>{artifacts.length}</span></button>
        <button type="button" data-active={tab === 'files' ? 'true' : undefined} onClick={() => setTab('files')}>我的文件 <span>{artifacts.length}</span></button>
      </nav>
      {tab === 'preview'
        ? <div className="wo-preview-stage">
          {artifacts.length > 1 && <div className="wo-preview-switcher">
            {artifacts.map(item => <button
              key={item.id}
              type="button"
              className="wo-switch-chip"
              data-active={item.id === selected?.id ? 'true' : undefined}
              onClick={() => setSelectedId(item.id)}
            >
              {modeGlyph(item.mode)} {item.title}
            </button>)}
          </div>}
          <div className="wo-preview">
            {selected ? <ArtifactPreview artifact={selected} /> : <EmptyArtifacts />}
          </div>
        </div>
        : <div className="wo-files-view">
          <div className="wo-artifact-list">
            {artifacts.length === 0
              ? <EmptyArtifacts />
              : artifacts.map(item => <ArtifactRow key={item.id} artifact={item} selected={item.id === selected?.id} onSelect={() => { setSelectedId(item.id); setTab('preview') }} />)}
          </div>
        </div>}
    </aside>
  </>
}

function ArtifactRow({ artifact, selected, onSelect }: { artifact: OfficeArtifact; selected: boolean; onSelect: () => void }): JSX.Element {
  const mode = modeDefinition(artifact.mode)
  return <button type="button" className="wo-artifact-row" data-selected={selected ? 'true' : undefined} onClick={onSelect}>
    <span className="wo-artifact-icon" style={{ background: mode.color }}>{modeGlyph(artifact.mode)}</span>
    <span className="wo-artifact-copy"><strong>{artifact.title}</strong><small>{artifact.summary}</small></span>
    <time>{formatTime(artifact.createdAt)}</time>
  </button>
}

function ArtifactPreview({ artifact }: { artifact: OfficeArtifact }): JSX.Element {
  if (artifact.isError) {
    return <div className="wo-preview-card">
      <div className="wo-error-box">{artifact.output || '执行失败，未返回更多错误细节。'}</div>
      <pre className="wo-error-raw">{artifact.output || 'Error: unknown'}</pre>
    </div>
  }

  return <div className="wo-preview-card" data-doc-preview={artifact.meta?.kind === 'doc' ? 'true' : undefined}>
    <ArtifactRichPreview artifact={artifact} />
  </div>
}

function ArtifactRichPreview({ artifact }: { artifact: OfficeArtifact }): JSX.Element {
  const meta = artifact.meta
  if (!meta) return <div className="wo-empty"><strong>暂无可用预览</strong><p>请重新生成文件以查看预览内容。</p></div>
  switch (meta.kind) {
    case 'chart': return <ChartPreview meta={meta} />
    case 'doc': return <DocPreview meta={meta} />
    case 'markdown': return <MarkdownPreview markdown={meta.markdown} />
    case 'sheet': return <SheetPreview meta={meta} />
    case 'ppt': return <PptPreview meta={meta} />
    case 'image': return <ImagePreview meta={meta} />
    case 'video': return <VideoPreview meta={meta} />
    case 'storyboard': return <StoryboardPreview meta={meta} />
    case 'drawio': return <DrawioPreview meta={meta} />
    default: return <pre>{artifact.output || '工具未返回可展示的文本内容。'}</pre>
  }
}

function DocPreview({ meta }: { meta: Extract<OfficeArtifactMeta, { kind: 'doc' }> }): JSX.Element {
  const sections = meta.sections?.length ? meta.sections : sectionsFromMarkdown(meta.markdown)
  const lead = sections[0]
  return <div className="wo-rich">
    <div className="wo-doc-surface"><div className="wo-doc-page">
      <div className="wo-doc-hero"><span className="wo-doc-badge">Word 预览</span><h1>{meta.title}</h1><p>{docLeadText(lead)}</p></div>
      {sections.slice(0, 5).map((section, index) => <section key={`${section.heading}-${index}`} className="wo-doc-section">
        {!HIDDEN_DOC_SECTION_HEADINGS.has(section.heading.trim()) && (section.headingLevel > 1 ? <h3>{section.heading}</h3> : <h2>{section.heading}</h2>)}
        {section.paragraphs.slice(0, 2).map((paragraph, paragraphIndex) => <p key={paragraphIndex}>{stripInlineMarkdown(paragraph)}</p>)}
        {section.bullets.length > 0 && <ul className="wo-doc-bullets">{section.bullets.slice(0, 5).map((bullet, bulletIndex) => <li key={bulletIndex}>{stripInlineMarkdown(bullet)}</li>)}</ul>}
        {section.table && <div className="wo-table-wrap"><table className="wo-table"><thead><tr>{section.table.headers.map(header => <th key={header}>{stripInlineMarkdown(header)}</th>)}</tr></thead><tbody>{section.table.rows.slice(0, 6).map((row, rowIndex) => <tr key={rowIndex}>{row.map((cell, cellIndex) => <td key={`${rowIndex}-${cellIndex}`}>{stripInlineMarkdown(cell)}</td>)}</tr>)}</tbody></table></div>}
      </section>)}
    </div></div>
  </div>
}

function ChartPreview({ meta }: { meta: Extract<OfficeArtifactMeta, { kind: 'chart' }> }): JSX.Element {
  const values = meta.values.length ? meta.values : [0]
  const max = Math.max(...values, 1)
  const palette = ['#7c3aed', '#2563eb', '#06b6d4', '#ec4899', '#10b981', '#f59e0b']
  if (meta.chartType === 'pie') {
    const total = values.reduce((sum, value) => sum + value, 0) || 1
    let current = 0
    const stops = values.map((value, index) => {
      const start = current
      current += value / total * 360
      return `${palette[index % palette.length]} ${start}deg ${current}deg`
    }).join(',')
    return <div className="wo-rich"><div className="wo-pie" style={{ background: `conic-gradient(${stops})` }} /><div className="wo-legend">{meta.labels.map((label, index) => <div key={label} className="wo-legend-item"><span className="wo-dot" style={{ background: palette[index % palette.length] }} />{label} · {meta.values[index]}</div>)}</div></div>
  }
  if (meta.chartType === 'line' || meta.chartType === 'scatter') {
    const points = meta.values.map((value, index) => `${index * (280 / Math.max(meta.values.length - 1, 1)) + 20},${160 - value / max * 120}`)
    return <div className="wo-rich"><svg viewBox="0 0 320 180" className="wo-chart-svg"><line x1="20" y1="160" x2="300" y2="160" stroke="#cbd5e1" /><line x1="20" y1="20" x2="20" y2="160" stroke="#cbd5e1" />{meta.chartType === 'line' && <polyline fill="none" stroke="#2563eb" strokeWidth="3" points={points.join(' ')} />}{points.map((point, index) => { const [cx, cy] = point.split(',').map(Number); return <g key={point}><circle cx={cx} cy={cy} r="4" fill={palette[index % palette.length]} />{meta.chartType === 'scatter' && <circle cx={cx} cy={cy} r="7" fill={`${palette[index % palette.length]}33`} />}</g> })}</svg><div className="wo-legend">{meta.labels.map((label, index) => <div key={label} className="wo-legend-item"><span className="wo-dot" style={{ background: palette[index % palette.length] }} />{label} · {meta.values[index]}</div>)}</div></div>
  }
  if (meta.chartType === 'gauge') {
    const value = meta.values[0] ?? 0
    const percent = Math.max(0, Math.min(100, value))
    return <div className="wo-rich"><div className="wo-kpis"><div className="wo-kpi"><strong>{value}</strong><span>{meta.seriesName}</span></div><div className="wo-kpi"><strong>{percent}%</strong><span>仪表盘进度</span></div><div className="wo-kpi"><strong>{meta.labels[0] ?? '当前值'}</strong><span>标签</span></div></div><div className="wo-bar-track"><div className="wo-bar-fill" style={{ width: `${percent}%` }} /></div></div>
  }
  return <div className="wo-rich"><div className="wo-chart-bars">{meta.labels.map((label, index) => {
    const currentValue = meta.values[index] ?? 0
    return <div key={label} className="wo-bar-row"><label>{label}</label><div className="wo-bar-track"><div className="wo-bar-fill" style={{ width: `${currentValue / max * 100}%`, background: meta.chartType === 'funnel' ? 'linear-gradient(90deg,#ec4899,#f59e0b)' : undefined }} /></div><em>{currentValue}</em></div>
  })}</div></div>
}

function MarkdownPreview({ markdown }: { markdown: string }): JSX.Element {
  const lines = markdown.split('\n')
  const items: JSX.Element[] = []
  let list: string[] = []
  const flush = (): void => {
    if (list.length) items.push(<ul key={`list-${items.length}`}>{list.map(item => <li key={item}>{item}</li>)}</ul>)
    list = []
  }
  for (const raw of lines) {
    const line = raw.trim()
    if (!line) { flush(); continue }
    if (line.startsWith('- ')) { list.push(line.slice(2)); continue }
    flush()
    if (line.startsWith('### ')) items.push(<h3 key={`${items.length}-${line}`}>{line.slice(4)}</h3>)
    else if (line.startsWith('## ')) items.push(<h2 key={`${items.length}-${line}`}>{line.slice(3)}</h2>)
    else if (line.startsWith('# ')) items.push(<h1 key={`${items.length}-${line}`}>{line.slice(2)}</h1>)
    else items.push(<p key={`${items.length}-${line}`}>{line}</p>)
  }
  flush()
  return <div className="wo-rich wo-md">{items.slice(0, 40)}</div>
}

function SheetPreview({ meta }: { meta: Extract<OfficeArtifactMeta, { kind: 'sheet' }> }): JSX.Element {
  const tables = meta.tables ?? []
  const [activeIndex, setActiveIndex] = useState(0)
  const safeIndex = Math.min(activeIndex, Math.max(tables.length - 1, 0))
  const table = tables[safeIndex]
  return <div className="wo-rich">
    {(meta.tableCount || meta.totalRows) && <div className="wo-kpis"><div className="wo-kpi"><strong>{meta.tableCount ?? tables.length}</strong><span>工作表</span></div><div className="wo-kpi"><strong>{meta.totalRows ?? 0}</strong><span>总行数</span></div><div className="wo-kpi"><strong>{table?.headers.length ?? 0}</strong><span>列数</span></div></div>}
    {tables.length > 1 && <div className="wo-sheet-tabs">{tables.map((item, index) => <button key={`${item.title}-${index}`} type="button" className="wo-sheet-tab" data-active={index === safeIndex ? 'true' : undefined} onClick={() => setActiveIndex(index)}>{item.title}</button>)}</div>}
    {table && <div className="wo-sheet-card"><strong>{table.title}</strong>{table.summary && <p>{table.summary}</p>}<div className="wo-sheet-meta"><span className="wo-sheet-pill">预览 {table.rows?.length ?? 0} 行</span>{typeof table.rowCount === 'number' && <span className="wo-sheet-pill">总计 {table.rowCount} 行</span>}{table.headers.length > 0 && <span className="wo-sheet-pill">{table.headers.length} 列字段</span>}</div></div>}
    {table?.rows && table.rows.length > 0
      ? <div className="wo-table-wrap"><table className="wo-table"><thead><tr>{table.headers.map(header => <th key={header}>{header}</th>)}</tr></thead><tbody>{table.rows.slice(0, 12).map((row, rowIndex) => <tr key={rowIndex}>{row.map((cell, cellIndex) => <td key={`${rowIndex}-${cellIndex}`}>{String(cell)}</td>)}</tr>)}</tbody></table></div>
      : <div className="wo-sheet-empty">当前结果未附带可视化预览行，但已生成 Excel 文件。您可以切换工作表，或直接下载产物查看完整内容。</div>}
  </div>
}

function PptPreview({ meta }: { meta: Extract<OfficeArtifactMeta, { kind: 'ppt' }> }): JSX.Element {
  return <div className="wo-rich">{meta.slideCount > 0 && <div className="wo-kpis"><div className="wo-kpi"><strong>{meta.slideCount}</strong><span>幻灯片页数</span></div><div className="wo-kpi"><strong>{meta.slides?.[0]?.layout ?? '-'}</strong><span>首页布局</span></div><div className="wo-kpi"><strong>{meta.filePath ? 'PPTX' : '大纲'}</strong><span>产物类型</span></div></div>}<div className="wo-slides">{meta.slides?.slice(0, 6).map((slide, index) => <div key={index} className="wo-slide"><strong>第 {slide.index ?? index + 1} 页</strong><span>{slide.title ?? slide.layout ?? '未命名页面'}</span>{slide.goal && <p>{slide.goal}</p>}{slide.points && slide.points.length > 0 && <span>{slide.points.slice(0, 3).join(' · ')}</span>}</div>)}</div></div>
}

function ImagePreview({ meta }: { meta: Extract<OfficeArtifactMeta, { kind: 'image' }> }): JSX.Element {
  return <div className="wo-rich"><div className="wo-kpis"><div className="wo-kpi"><strong>{meta.images.length}</strong><span>图片数量</span></div><div className="wo-kpi"><strong>{meta.generationMode ?? '-'}</strong><span>生成模式</span></div><div className="wo-kpi"><strong>{meta.provider ?? '-'}</strong><span>服务商</span></div></div><div className="wo-media-grid">{meta.images.map(image => <img key={image.url} src={image.url} alt={image.style} title={image.style} />)}</div></div>
}

function VideoPreview({ meta }: { meta: Extract<OfficeArtifactMeta, { kind: 'video' }> }): JSX.Element {
  return <div className="wo-rich"><div className="wo-kpis"><div className="wo-kpi"><strong>{meta.duration ?? '-'}</strong><span>时长（秒）</span></div><div className="wo-kpi"><strong>{meta.aspectRatio ?? '-'}</strong><span>宽高比</span></div><div className="wo-kpi"><strong>{meta.mode ?? '-'}</strong><span>模式</span></div></div><video className="wo-video" controls src={meta.videoUrl} /></div>
}

function StoryboardPreview({ meta }: { meta: Extract<OfficeArtifactMeta, { kind: 'storyboard' }> }): JSX.Element {
  return <div className="wo-rich"><div className="wo-kpis"><div className="wo-kpi"><strong>{meta.totalShots}</strong><span>镜头数</span></div><div className="wo-kpi"><strong>{meta.totalSeconds}</strong><span>总时长</span></div><div className="wo-kpi"><strong>{meta.aspectRatio ?? '-'}</strong><span>宽高比</span></div></div><div className="wo-storyboard">{meta.shots.slice(0, 8).map((shot, index) => <div key={index} className="wo-story"><strong>{shot.index ?? index + 1}. {shot.title ?? '镜头'}</strong>{shot.description && <p>{shot.description}</p>}{shot.prompt && <code>{shot.prompt}</code>}</div>)}</div></div>
}

function DrawioPreview({ meta }: { meta: Extract<OfficeArtifactMeta, { kind: 'drawio' }> }): JSX.Element {
  return <div className="wo-rich"><div className="wo-kpis"><div className="wo-kpi"><strong>{meta.diagramType}</strong><span>图表类型</span></div><div className="wo-kpi"><strong>{meta.title}</strong><span>标题</span></div><div className="wo-kpi"><strong>{Math.min(meta.xml.length, 9999)}</strong><span>XML 长度</span></div></div><pre className="wo-xml">{meta.xml.slice(0, 2400)}</pre></div>
}

function EmptyArtifacts(): JSX.Element {
  return <div className="wo-empty"><span>▤</span><strong>暂无办公产物</strong><p>在对话中生成 Word、Excel、PPT 等内容后，会自动汇总到这里。</p></div>
}

function OfficeStyles(): JSX.Element {
  return <style>{STYLES + NATIVE_DETAILS_STYLES + RICH_PREVIEW_STYLES + DOCUMENT_PREVIEW_STYLES + LAYOUT_OVERRIDE_STYLES + DOCK_STABLE_HEIGHT_STYLES}</style>
}

const HIDDEN_DOC_SECTION_HEADINGS = new Set(['需求原文', '待补充章节'])

function docFormatLabel(format: string | undefined): string {
  switch (format) {
    case 'prd': return 'PRD'
    case 'plan': return '方案'
    case 'summary': return '总结'
    case 'article': return '文章'
    default: return '报告'
  }
}

function docLeadText(section: { paragraphs: string[]; bullets: string[] } | undefined): string {
  if (!section) return '已根据结构化内容生成文档预览，可在右侧快速浏览章节布局与重点。'
  return stripInlineMarkdown(section.paragraphs[0] ?? section.bullets[0] ?? '已根据结构化内容生成文档预览，可在右侧快速浏览章节布局与重点。')
}

function sectionsFromMarkdown(markdown: string): { heading: string; headingLevel: number; paragraphs: string[]; bullets: string[]; table?: { headers: string[]; rows: string[][] } }[] {
  const sections: { heading: string; headingLevel: number; paragraphs: string[]; bullets: string[]; table?: { headers: string[]; rows: string[][] } }[] = []
  let current: { heading: string; headingLevel: number; paragraphs: string[]; bullets: string[]; table?: { headers: string[]; rows: string[][] } } | null = null
  for (const rawLine of markdown.split('\n')) {
    const line = rawLine.trim()
    if (!line) continue
    if (line.startsWith('# ')) continue
    if (line.startsWith('## ') || line.startsWith('### ')) {
      if (current) sections.push(current)
      current = {
        heading: line.replace(/^#+\s+/, ''),
        headingLevel: line.startsWith('### ') ? 3 : 2,
        paragraphs: [],
        bullets: [],
      }
      continue
    }
    if (!current) {
      current = { heading: '内容概览', headingLevel: 2, paragraphs: [], bullets: [] }
    }
    if (line.startsWith('- ')) current.bullets.push(line.slice(2))
    else current.paragraphs.push(line)
  }
  if (current) sections.push(current)
  return sections
}

function stripInlineMarkdown(value: string): string {
  return value.replace(/\*\*(.*?)\*\*/g, '$1').replace(/\*(.*?)\*/g, '$1').replace(/`(.*?)`/g, '$1').trim()
}

function readMode(): OfficeMode {
  const value = safeStorageGet(MODE_STORAGE)
  return OFFICE_MODES.some(item => item.id === value) ? value as OfficeMode : 'all'
}

function readArtifacts(): OfficeArtifact[] {
  const value = safeStorageGet(ARTIFACT_STORAGE)
  if (!value) return []
  try {
    const parsed = JSON.parse(value)
    return Array.isArray(parsed) ? parsed.slice(0, MAX_ARTIFACTS) as OfficeArtifact[] : []
  } catch {
    return []
  }
}

function safeStorageGet(key: string): string | null {
  try { return window.localStorage.getItem(key) } catch { return null }
}

function safeStorageSet(key: string, value: string): void {
  try { window.localStorage.setItem(key, value) } catch { /* best effort */ }
}

function readTopic(argsRaw: string): string {
  try {
    const value = JSON.parse(argsRaw)
    return typeof value?.topic === 'string' ? value.topic : ''
  } catch {
    return ''
  }
}

function modeForToolName(toolName: string): OfficeMode {
  if (toolName.startsWith('ppt_')) return 'ppt'
  if (toolName === 'doc_generate' || toolName === 'md_generate') return 'doc'
  if (toolName === 'sheet_generate') return 'sheet'
  if (toolName === 'chart_generate') return 'chart'
  if (toolName === 'drawio_generate') return 'drawio'
  if (toolName === 'image_prompt') return 'image'
  if (toolName.startsWith('video_')) return 'video'
  return 'all'
}

function toolTitle(toolName: string): string {
  const titles: Record<string, string> = {
    ppt_plan: 'PPT 大纲规划', ppt_generate: 'PPT 演示生成', doc_generate: 'Word 文档生成', md_generate: 'Markdown 文档生成',
    sheet_generate: 'Excel 表格生成', chart_generate: '数据图表生成', drawio_generate: 'Draw.io 图表生成', image_prompt: 'AI 图片生成',
    video_generate: 'AI 视频生成', video_storyboard: '视频分镜规划',
  }
  return titles[toolName] ?? toolName
}

function toolGlyph(toolName: string): string {
  if (toolName.includes('ppt')) return 'P'
  if (toolName.includes('doc') || toolName.includes('md')) return 'W'
  if (toolName.includes('sheet')) return 'X'
  if (toolName.includes('chart')) return '⌁'
  if (toolName.includes('drawio')) return '◇'
  if (toolName.includes('image')) return '▧'
  if (toolName.includes('video')) return '▶'
  return 'W'
}

function modeGlyph(mode: OfficeMode): string {
  return ({ all: 'W', doc: 'W', sheet: 'X', ppt: 'P', chart: '⌁', drawio: '◇', image: '▧', video: '▶' })[mode]
}

function modeDescription(mode: OfficeMode): string {
  return ({ all: '自动选择工具', doc: '报告、方案与说明书', sheet: '数据、预算与分析表', ppt: '汇报、路演与培训课件', chart: '柱状、折线与饼图', drawio: '流程、架构与关系图', image: '封面、插图与视觉素材', video: '分镜与短视频生成' })[mode]
}

function formatTime(value: number): string {
  if (!Number.isFinite(value)) return ''
  return new Date(value).toLocaleString('zh-CN', { month: '2-digit', day: '2-digit', hour: '2-digit', minute: '2-digit' })
}

const STYLES = `
.wo-dock,.wo-panel,.wo-launcher,.wo-tool-card{font-family:Inter,"PingFang SC","Microsoft YaHei",sans-serif;box-sizing:border-box}.wo-dock *,.wo-panel *,.wo-launcher *,.wo-tool-card *{box-sizing:border-box}.wo-dock{pointer-events:auto;width:calc(100% - var(--dsh-composer-side-clearance,24px)*2);max-width:var(--dsh-composer-card-max-width,920px);margin:0 auto 8px;padding:8px 10px;border:1px solid var(--dsw-alias-border-l1,#e5e7eb);border-radius:16px;background:color-mix(in srgb,var(--dsw-alias-bg-base,#fff) 94%,transparent);box-shadow:0 8px 30px rgba(15,23,42,.08);display:flex;align-items:center;gap:10px;overflow:hidden}.wo-dock-brand{display:flex;align-items:center;gap:7px;font-size:12px;font-weight:700;color:var(--dsw-alias-label-primary,#111827);white-space:nowrap}.wo-brand-mark{display:grid;place-items:center;width:24px;height:24px;border-radius:8px;color:#fff;background:linear-gradient(145deg,#1d4ed8,#06b6d4);font-weight:900}.wo-mode-list{display:flex;align-items:center;gap:5px;min-width:0;overflow-x:auto;scrollbar-width:none}.wo-mode-list::-webkit-scrollbar{display:none}.wo-mode-button,.wo-artifacts-button{border:1px solid var(--dsw-alias-border-l1,#e5e7eb);background:var(--dsw-alias-bg-base,#fff);color:var(--dsw-alias-label-secondary,#4b5563);border-radius:999px;height:30px;padding:0 10px;display:flex;align-items:center;gap:5px;white-space:nowrap;cursor:pointer;font-size:12px}.wo-mode-button:hover,.wo-artifacts-button:hover{background:var(--dsw-alias-interactive-bg-hover,#f3f4f6)}.wo-mode-button[data-active]{border-color:color-mix(in srgb,var(--wo-mode-color) 44%,transparent);color:var(--wo-mode-color);background:color-mix(in srgb,var(--wo-mode-color) 9%,var(--dsw-alias-bg-base,#fff))}.wo-mode-dot{width:7px;height:7px;border-radius:999px;background:var(--wo-mode-color)}.wo-artifacts-button{margin-left:auto;font-weight:650}.wo-launcher{pointer-events:auto;position:fixed;right:18px;top:92px;z-index:72;border:none;border-radius:16px;background:var(--dsw-alias-bg-base,#fff);color:var(--dsw-alias-label-primary,#111827);box-shadow:0 12px 36px rgba(15,23,42,.18);padding:8px 12px 8px 8px;display:flex;align-items:center;gap:7px;cursor:pointer;font-weight:700}.wo-panel{pointer-events:auto;position:fixed;z-index:70;top:12px;right:12px;bottom:12px;width:min(520px,calc(100vw - 72px));border:1px solid var(--dsw-alias-border-l1,#e5e7eb);border-radius:22px;background:var(--dsw-alias-bg-base,#fff);box-shadow:0 24px 70px rgba(15,23,42,.24);display:flex;flex-direction:column;overflow:hidden;color:var(--dsw-alias-label-primary,#111827)}.wo-panel-header{height:72px;padding:12px 16px;border-bottom:1px solid var(--dsw-alias-border-l1,#e5e7eb);display:flex;align-items:center;justify-content:space-between;flex:none}.wo-panel-title{display:flex;align-items:center;gap:11px}.wo-logo{width:42px;height:42px;border-radius:14px;display:grid;place-items:center;color:#fff;background:linear-gradient(145deg,#1d4ed8,#06b6d4);font-size:20px;font-weight:900;box-shadow:0 8px 22px rgba(37,99,235,.28)}.wo-panel-title div{display:flex;flex-direction:column;gap:3px}.wo-panel-title strong{font-size:17px}.wo-panel-title span{font-size:12px;color:var(--dsw-alias-label-tertiary,#6b7280)}.wo-close{width:34px;height:34px;border:0;border-radius:999px;background:transparent;color:var(--dsw-alias-label-secondary,#4b5563);font-size:24px;cursor:pointer}.wo-close:hover{background:var(--dsw-alias-interactive-bg-hover,#f3f4f6)}.wo-panel-tabs{height:48px;padding:7px 14px;border-bottom:1px solid var(--dsw-alias-border-l1,#e5e7eb);display:flex;gap:8px;flex:none}.wo-panel-tabs button{border:0;border-radius:10px;background:transparent;color:var(--dsw-alias-label-secondary,#4b5563);padding:0 13px;font-weight:650;cursor:pointer}.wo-panel-tabs button[data-active]{background:#eff6ff;color:#1d4ed8}.wo-panel-tabs span{display:inline-grid;place-items:center;min-width:19px;height:19px;margin-left:4px;padding:0 5px;border-radius:999px;background:#dbeafe;font-size:11px}.wo-panel-content{padding:16px;overflow:auto}.wo-hero-card{border-radius:18px;padding:20px;color:#fff;background:linear-gradient(135deg,#0f172a 0%,#1d4ed8 64%,#06b6d4 140%);box-shadow:0 14px 30px rgba(37,99,235,.2)}.wo-hero-card>span{font-size:11px;text-transform:uppercase;letter-spacing:.14em;opacity:.75}.wo-hero-card h2{margin:7px 0 6px;font-size:22px}.wo-hero-card p{margin:0;font-size:13px;line-height:1.7;opacity:.86}.wo-tool-grid{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:10px;margin-top:14px}.wo-tool-tile{min-height:92px;border:1px solid var(--dsw-alias-border-l1,#e5e7eb);border-radius:16px;background:var(--dsw-alias-bg-base,#fff);padding:13px;text-align:left;display:grid;grid-template-columns:34px 1fr;grid-template-rows:auto auto;column-gap:10px;cursor:pointer}.wo-tool-tile:hover{border-color:#93c5fd;background:#f8fbff;transform:translateY(-1px)}.wo-tile-icon{grid-row:1/3;width:34px;height:34px;border-radius:10px;color:#fff;display:grid;place-items:center;font-weight:850}.wo-tool-tile strong{font-size:13px;align-self:end}.wo-tool-tile small{font-size:11px;color:var(--dsw-alias-label-tertiary,#6b7280);margin-top:3px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}.wo-recent-section{margin-top:18px}.wo-section-heading{display:flex;align-items:center;justify-content:space-between;margin-bottom:8px}.wo-section-heading strong{font-size:14px}.wo-section-heading button{border:0;background:transparent;color:#2563eb;cursor:pointer;font-size:12px}.wo-artifact-layout{min-height:0;flex:1;display:grid;grid-template-rows:minmax(150px,42%) minmax(0,1fr)}.wo-artifact-list{padding:10px;overflow:auto;border-bottom:1px solid var(--dsw-alias-border-l1,#e5e7eb)}.wo-artifact-row{width:100%;border:0;border-radius:13px;background:transparent;padding:9px;display:grid;grid-template-columns:34px minmax(0,1fr) auto;align-items:center;gap:9px;text-align:left;cursor:pointer;color:inherit}.wo-artifact-row:hover,.wo-artifact-row[data-selected]{background:var(--dsw-alias-interactive-bg-hover,#f3f4f6)}.wo-artifact-row[data-selected]{box-shadow:inset 3px 0 #2563eb}.wo-artifact-icon{width:34px;height:34px;border-radius:10px;color:#fff;display:grid;place-items:center;font-weight:850}.wo-artifact-copy{min-width:0;display:flex;flex-direction:column;gap:3px}.wo-artifact-copy strong{font-size:12px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}.wo-artifact-copy small{font-size:11px;color:var(--dsw-alias-label-tertiary,#6b7280);white-space:nowrap;overflow:hidden;text-overflow:ellipsis}.wo-artifact-row time{font-size:10px;color:var(--dsw-alias-label-caption,#9ca3af)}.wo-preview{min-height:0;padding:14px;overflow:auto;background:color-mix(in srgb,var(--dsw-alias-bg-base,#fff) 96%,#eff6ff)}.wo-preview-card{border:1px solid var(--dsw-alias-border-l1,#e5e7eb);border-radius:16px;background:var(--dsw-alias-bg-base,#fff);overflow:hidden}.wo-preview-head{padding:14px;display:flex;gap:10px;align-items:center;border-bottom:1px solid var(--dsw-alias-border-l1,#e5e7eb)}.wo-preview-icon{width:38px;height:38px;border-radius:11px;color:#fff;display:grid;place-items:center;font-weight:900}.wo-preview-head div{display:flex;flex-direction:column;gap:3px;min-width:0}.wo-preview-head strong{font-size:13px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}.wo-preview-head span{font-size:11px;color:var(--dsw-alias-label-tertiary,#6b7280)}.wo-preview-summary{margin:12px;padding:12px;border-radius:12px;background:#f0fdf4;color:#166534;font-size:12px;line-height:1.55}.wo-file-card{margin:12px;padding:10px 12px;border:1px solid #bbf7d0;border-radius:12px;background:#f0fdf4;display:flex;flex-direction:column;gap:5px}.wo-file-card span{font-size:11px;color:#15803d}.wo-file-card code{font-size:11px;white-space:normal;word-break:break-all}.wo-preview pre{margin:12px;max-height:260px;overflow:auto;border-radius:12px;padding:12px;background:var(--dsw-alias-markdown-code-block,#f6f7f9);color:var(--dsw-alias-label-secondary,#4b5563);font:11px/1.6 ui-monospace,SFMono-Regular,Menlo,monospace;white-space:pre-wrap;word-break:break-word}.wo-empty{min-height:120px;display:flex;flex-direction:column;align-items:center;justify-content:center;text-align:center;color:var(--dsw-alias-label-tertiary,#6b7280);padding:18px}.wo-empty>span{font-size:24px}.wo-empty strong{margin-top:7px;color:var(--dsw-alias-label-secondary,#4b5563);font-size:13px}.wo-empty p{max-width:280px;margin:5px 0 0;font-size:11px;line-height:1.6}.wo-tool-card{margin:6px 0 8px 4px;border:1px solid #bfdbfe;border-radius:14px;background:linear-gradient(180deg,#f8fbff,#fff);overflow:hidden;color:var(--dsw-alias-label-primary,#111827)}.wo-tool-card[data-error]{border-color:#fecaca;background:#fffafa}.wo-tool-head{padding:10px 12px;display:flex;align-items:center;gap:9px}.wo-tool-icon{width:28px;height:28px;border-radius:9px;color:#fff;display:grid;place-items:center;font-weight:850;font-size:12px}.wo-tool-heading{min-width:0;flex:1;display:flex;flex-direction:column;gap:2px}.wo-tool-heading strong{font-size:12px}.wo-tool-heading span{font-size:11px;color:var(--dsw-alias-label-tertiary,#6b7280);white-space:nowrap;overflow:hidden;text-overflow:ellipsis}.wo-status{border-radius:999px;padding:3px 8px;background:#dcfce7;color:#15803d;font-size:10px}.wo-status[data-running]{background:#dbeafe;color:#1d4ed8}.wo-status[data-error]{background:#fee2e2;color:#b91c1c}.wo-tool-body{border-top:1px solid #dbeafe;padding:9px 12px;font-size:11px;color:var(--dsw-alias-label-secondary,#4b5563)}.wo-tool-body p{margin:0;line-height:1.55}.wo-tool-body code{display:block;margin-top:6px;padding:6px 8px;border-radius:8px;background:#eff6ff;word-break:break-all}.wo-tool-actions{padding:0 12px 10px;display:flex;gap:6px}.wo-tool-actions button{border:1px solid #bfdbfe;border-radius:8px;background:#fff;color:#1d4ed8;padding:4px 8px;font-size:10px;cursor:pointer}.wo-tool-actions button:hover{background:#eff6ff}@media(max-width:760px){.wo-panel{inset:8px;width:auto}.wo-tool-grid{grid-template-columns:1fr}.wo-dock-brand{display:none}.wo-mode-button{padding:0 8px}.wo-panel-content{padding:12px}}`
