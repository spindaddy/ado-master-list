import {
  BrowserWindow,
  Menu,
  type ContextMenuParams,
  type MenuItemConstructorOptions,
  type WebContents
} from 'electron'

export function attachEditContextMenu(wc: WebContents): void {
  wc.on('context-menu', (_e, params: ContextMenuParams) => {
    const items: MenuItemConstructorOptions[] = []
    if (params.isEditable) {
      items.push(
        { role: 'undo', enabled: params.editFlags.canUndo },
        { role: 'redo', enabled: params.editFlags.canRedo },
        { type: 'separator' },
        { role: 'cut', enabled: params.editFlags.canCut },
        { role: 'copy', enabled: params.editFlags.canCopy },
        { role: 'paste', enabled: params.editFlags.canPaste },
        { role: 'delete', enabled: params.editFlags.canDelete },
        { type: 'separator' },
        { role: 'selectAll', enabled: params.editFlags.canSelectAll }
      )
    } else if (params.selectionText?.trim()) {
      items.push({ role: 'copy' })
    }
    if (items.length === 0) return
    Menu.buildFromTemplate(items).popup({
      window: BrowserWindow.fromWebContents(wc) ?? undefined
    })
  })
}

export function installAppMenu(): void {
  const isMac = process.platform === 'darwin'
  const template: MenuItemConstructorOptions[] = [
    ...(isMac
      ? ([
          {
            role: 'appMenu'
          }
        ] as MenuItemConstructorOptions[])
      : []),
    {
      label: 'Edit',
      submenu: [
        { role: 'undo' },
        { role: 'redo' },
        { type: 'separator' },
        { role: 'cut' },
        { role: 'copy' },
        { role: 'paste' },
        { role: 'pasteAndMatchStyle' },
        { role: 'delete' },
        { role: 'selectAll' }
      ]
    },
    {
      label: 'View',
      submenu: [
        { role: 'reload' },
        { role: 'toggleDevTools' },
        { type: 'separator' },
        { role: 'resetZoom' },
        { role: 'zoomIn' },
        { role: 'zoomOut' },
        { type: 'separator' },
        { role: 'togglefullscreen' }
      ]
    },
    {
      role: 'windowMenu'
    }
  ]
  Menu.setApplicationMenu(Menu.buildFromTemplate(template))
}
