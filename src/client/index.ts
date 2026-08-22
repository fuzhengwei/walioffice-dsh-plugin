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
      else layout.openDetails()
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
