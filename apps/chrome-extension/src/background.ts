import { type Message } from './common/message'

const debugLog = (...args: unknown[]) => {
  if (import.meta.env.DEV) console.log(...args)
}

debugLog('[CDC background] Service worker loaded')

const sharedDocumentUrlPatterns: string[] = [
  'https://*.feishu.cn/*',
  'https://*.feishu.net/*',
  'https://*.larksuite.com/*',
  'https://*.feishu-pre.net/*',
  'https://*.larkoffice.com/*',
  'https://*.larkenterprise.com/*',
]

const sharedDocumentHosts = [
  'feishu.cn',
  'feishu.net',
  'larksuite.com',
  'feishu-pre.net',
  'larkoffice.com',
  'larkenterprise.com',
]

const isSharedDocumentUrl = (url?: string): boolean => {
  if (!url) return true

  try {
    const { protocol, hostname } = new URL(url)
    return (
      protocol === 'https:' &&
      sharedDocumentHosts.some(
        host => hostname === host || hostname.endsWith(`.${host}`),
      )
    )
  } catch {
    return false
  }
}

const isExpectedHostPermissionError = (error: unknown): boolean =>
  error instanceof Error &&
  error.message.includes('Cannot access contents of the page')

enum MenuItemId {
  DOWNLOAD_DOCX_AS_MARKDOWN = 'download_docx_as_markdown',
  COPY_DOCX_AS_MARKDOWN = 'copy_docx_as_markdown',
  VIEW_DOCX_AS_MARKDOWN = 'view_docx_as_markdown',
  DOWNLOAD_DOCX_AS_HTML = 'download_docx_as_html',
  DOWNLOAD_BITABLE_AS_MD = 'download_bitable_as_md',
  DOWNLOAD_BITABLE_AS_HTML = 'download_bitable_as_html',
}

chrome.runtime.onInstalled.addListener(() => {
  chrome.contextMenus.create({
    id: MenuItemId.DOWNLOAD_DOCX_AS_MARKDOWN,
    title: chrome.i18n.getMessage('download_docx_as_markdown'),
    documentUrlPatterns: sharedDocumentUrlPatterns,
    contexts: ['page', 'editable'],
  })

  chrome.contextMenus.create({
    id: MenuItemId.COPY_DOCX_AS_MARKDOWN,
    title: chrome.i18n.getMessage('copy_docx_as_markdown'),
    documentUrlPatterns: sharedDocumentUrlPatterns,
    contexts: ['page', 'editable'],
  })

  chrome.contextMenus.create({
    id: MenuItemId.VIEW_DOCX_AS_MARKDOWN,
    title: chrome.i18n.getMessage('view_docx_as_markdown'),
    documentUrlPatterns: sharedDocumentUrlPatterns,
    contexts: ['page', 'editable'],
  })

  chrome.contextMenus.create({
    id: MenuItemId.DOWNLOAD_DOCX_AS_HTML,
    title: chrome.i18n.getMessage('download_docx_as_html'),
    documentUrlPatterns: sharedDocumentUrlPatterns,
    contexts: ['page', 'editable'],
  })

  chrome.contextMenus.create({
    id: MenuItemId.DOWNLOAD_BITABLE_AS_MD,
    title: chrome.i18n.getMessage('download_bitable_as_md'),
    documentUrlPatterns: sharedDocumentUrlPatterns,
    contexts: ['page', 'editable'],
  })

  chrome.contextMenus.create({
    id: MenuItemId.DOWNLOAD_BITABLE_AS_HTML,
    title: chrome.i18n.getMessage('download_bitable_as_html'),
    documentUrlPatterns: sharedDocumentUrlPatterns,
    contexts: ['page', 'editable'],
  })
})

const executeScriptByFlag = async (flag: string | number, tabId: number) => {
  debugLog('[CDC background] executeScriptByFlag, flag:', flag, 'tabId:', tabId)
  switch (flag) {
    case MenuItemId.DOWNLOAD_DOCX_AS_MARKDOWN:
      await chrome.scripting.executeScript({
        files: ['bundles/scripts/download-lark-docx-as-markdown.js'],
        target: { tabId },
        world: 'MAIN',
      })
      break
    case MenuItemId.COPY_DOCX_AS_MARKDOWN:
      await chrome.scripting.executeScript({
        files: ['bundles/scripts/copy-lark-docx-as-markdown.js'],
        target: { tabId },
        world: 'MAIN',
      })
      break
    case MenuItemId.VIEW_DOCX_AS_MARKDOWN:
      await chrome.scripting.executeScript({
        files: ['bundles/scripts/view-lark-docx-as-markdown.js'],
        target: { tabId },
        world: 'MAIN',
      })
      break
    case MenuItemId.DOWNLOAD_DOCX_AS_HTML:
      await chrome.scripting.executeScript({
        files: ['bundles/scripts/download-lark-docx-as-html.js'],
        target: { tabId },
        world: 'MAIN',
      })
      break
    case MenuItemId.DOWNLOAD_BITABLE_AS_MD:
      await chrome.scripting.executeScript({
        files: [
          'bundles/scripts/bitable-clientvars-cache.js',
          'bundles/scripts/bitable-export.js',
        ],
        target: { tabId },
        world: 'MAIN',
      })
      break
    case MenuItemId.DOWNLOAD_BITABLE_AS_HTML:
      await chrome.scripting.executeScript({
        files: [
          'bundles/scripts/bitable-clientvars-cache.js',
          'bundles/scripts/bitable-export.js',
        ],
        target: { tabId },
        world: 'MAIN',
      })
      break
    default:
      break
  }
}

const executeScriptSafely = async (
  flag: string | number,
  tabId: number,
  tabUrl?: string,
) => {
  if (!isSharedDocumentUrl(tabUrl)) {
    debugLog('[CDC background] Skip unsupported tab url:', tabUrl)
    return
  }

  try {
    await executeScriptByFlag(flag, tabId)
  } catch (error) {
    if (isExpectedHostPermissionError(error)) {
      debugLog('[CDC background] Skip tab without host access:', tabUrl, error)
      return
    }

    console.error(error)
  }
}

chrome.contextMenus.onClicked.addListener(({ menuItemId }, tab) => {
  debugLog(
    '[CDC background] contextMenu clicked, menuItemId:',
    menuItemId,
    'tabId:',
    tab?.id,
  )
  if (tab?.id !== undefined) {
    executeScriptSafely(menuItemId, tab.id, tab.url)
  }
})

chrome.runtime.onMessage.addListener((_message, sender, sendResponse) => {
  const message = _message as Message
  debugLog(
    '[CDC background] received message, flag:',
    message.flag,
    'from:',
    sender.tab?.id,
  )

  const executeScript = async () => {
    const activeTabs = await chrome.tabs.query({
      currentWindow: true,
      active: true,
    })

    const activeTabId = activeTabs.at(0)?.id
    const activeTabUrl = activeTabs.at(0)?.url
    debugLog(
      '[CDC background] active tab:',
      activeTabId,
      'total active tabs:',
      activeTabs.length,
    )

    if (activeTabs.length === 1 && activeTabId !== undefined) {
      await executeScriptSafely(message.flag, activeTabId, activeTabUrl)
    }
  }

  executeScript().then(sendResponse).catch(console.error)

  return true
})
