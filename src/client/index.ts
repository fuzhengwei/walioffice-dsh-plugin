import type { ClientContext } from '@deepseek-ai/dsh-client-runtime/client'
import type {} from '@deepseek-ai/dsh-client-ui-layout/client'
import type {} from '@deepseek-ai/dsh-client-ui-conversation/client'
import type {} from '@deepseek-ai/dsh-client-ui-tool/client'
import { OfficeDetailsPanel, OfficeDock, OfficeToolView } from './OfficeUI.tsx'
import { OFFICE_PANEL_EVENT, OFFICE_TOOLS } from './office-state.ts'

export const inject = ['slots', 'layout']

export function apply(ctx: ClientContext): void {
  const layout = ctx.layout
  const slots = ctx.slots

  slots.inject('details', () => slots.register(
    {
      name: 'details',
      priority: -10,
      inject: () => ({ closeDetails: () => layout.closeDetails() }),
    },
    OfficeDetailsPanel,
  ))

  ctx.effect(() => {
    const onPanel = (event: Event): void => {
      const detail = (event as CustomEvent<{ open?: boolean }>).detail
      if (detail?.open === false) layout.closeDetails()
      else {
        layout.openDetails()
        window.requestAnimationFrame(() => widenOfficeDetails())
      }
    }
    window.addEventListener(OFFICE_PANEL_EVENT, onPanel)
    return () => window.removeEventListener(OFFICE_PANEL_EVENT, onPanel)
  }, 'walioffice: native details panel controller')

  slots.inject('conversation.input.dock', () => slots.register(
    {
      name: 'conversation.input.dock',
      id: 'walioffice-dock',
      order: 80,
      inject: () => ({ openOfficeDetails: () => layout.openDetails() }),
    },
    OfficeDock,
  ))

  slots.inject('tool.call.toolview', function* () {
    for (const toolName of OFFICE_TOOLS) {
      yield slots.register(
        {
          name: 'tool.call.toolview',
          key: toolName,
          inject: () => ({ openOfficeDetails: () => layout.openDetails() }),
        },
        OfficeToolView,
      )
    }
  })
}

function widenOfficeDetails(): void {
  const panel = document.querySelector<HTMLElement>('.wo-panel')
  if (!panel) return

  let frame: HTMLElement | null = panel
  while (frame && !frame.style.gridTemplateColumns) frame = frame.parentElement
  if (!frame) return

  const viewport = frame.getBoundingClientRect().width
  const columns = getComputedStyle(frame).gridTemplateColumns.split(/\s+/)
  const sidebar = Number.parseFloat(columns[0] ?? '')
  if (!Number.isFinite(viewport) || viewport <= 0 || !Number.isFinite(sidebar)) return

  const details = Math.min(640, Math.max(520, Math.floor(viewport - sidebar - 560)))
  if (details <= 520) return

  const center = Math.max(0, viewport - sidebar - details)
  frame.style.setProperty(
    'grid-template-columns',
    `${sidebar}px minmax(0, ${center}px) ${details}px`,
    'important',
  )
}
