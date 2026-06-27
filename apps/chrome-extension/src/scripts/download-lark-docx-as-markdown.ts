import i18next from 'i18next'

if (import.meta.env.DEV) {
  console.log(
    '[CDC] download-md script loaded, href:',
    location.href,
    'readyState:',
    document.readyState,
  )
  window.addEventListener('error', e => {
    console.error(
      '[CDC] Global error:',
      e.error || e.message,
      e.filename,
      e.lineno,
    )
  })
  window.addEventListener('unhandledrejection', e => {
    console.error('[CDC] Unhandled rejection:', e.reason)
  })
}

// Persistent debug panel at bottom of screen
function createDebugPanel(): (msg: string) => void {
  if (!import.meta.env.DEV || !location.search.includes('cdc_debug=1')) {
    return () => {}
  }

  try {
    let panel = document.getElementById(
      'cdc-debug-panel',
    ) as HTMLDivElement | null
    if (!panel) {
      panel = document.createElement('div')
      panel.id = 'cdc-debug-panel'
      panel.style.cssText =
        'position:fixed;bottom:0;left:0;right:0;max-height:40vh;overflow-y:auto;z-index:2147483646;background:#1e1e1e;color:#d4d4d4;padding:12px 16px;font:12px Menlo,Consolas,monospace;box-shadow:0 -2px 12px rgba(0,0,0,.4);white-space:pre-wrap;word-break:break-all;text-align:left;'
      const header = document.createElement('div')
      header.style.cssText =
        'display:flex;justify-content:space-between;align-items:center;margin-bottom:8px;color:#569cd6;font-weight:bold;'
      header.innerHTML = '<span>[CDC] Debug Panel</span>'
      const closeBtn = document.createElement('button')
      closeBtn.textContent = '× Close'
      closeBtn.style.cssText =
        'background:none;border:1px solid #555;color:#ccc;font-size:12px;padding:2px 8px;cursor:pointer;border-radius:3px;'
      closeBtn.onclick = () => panel?.remove()
      header.appendChild(closeBtn)
      panel.appendChild(header)
      document.documentElement.appendChild(panel)
    }
    return (msg: string) => {
      const line = document.createElement('div')
      line.style.marginBottom = '3px'
      line.textContent = msg
      panel.appendChild(line)
      panel.scrollTop = panel.scrollHeight
    }
  } catch {
    return () => {}
  }
}
const debugLog = createDebugPanel()

import { Toast, Docx, docx, doc, type mdast } from '@dolphin/lark'
import { Minute, OneHundred, Second, waitFor } from '@dolphin/common'
import { fileSave, supported } from 'browser-fs-access'
import { fs } from '@zip.js/zip.js'
import normalizeFileName from 'filenamify/browser'
import { cluster } from 'radash'
import { CommonTranslationKey, en, Namespace, zh } from '../common/i18n'
import { confirm } from '../common/notification'
import { legacyFileSave } from '../common/legacy'
import { reportBug } from '../common/issue'
import {
  transformMentionUsers,
  UniqueFileName,
  withSignal,
  transformTableBySettings,
} from '../common/utils'
import { getSettings, Grid } from '../common/settings'
import { DownloadMethod, SettingKey } from '@/common/settings'
import {
  bitableToMarkdown,
  extractStandaloneBitableTable,
  extractStandaloneBitableTableFromWebApi,
  isStandaloneBitablePage,
} from './bitable-export'

const uniqueFileName = new UniqueFileName()

const DOWNLOAD_ABORTED = 'Download aborted'

const enum TranslationKey {
  CONTENT_LOADING = 'content_loading',
  UNKNOWN_ERROR = 'unknown_error',
  NOT_SUPPORT = 'not_support',
  NOT_SUPPORT_DOC_1_0 = 'not_support_doc_1_0',
  DOWNLOADING_FILE = 'downloading_file',
  FAILED_TO_DOWNLOAD = 'failed_to_download',
  DOWNLOAD_PROGRESS = 'download_progress',
  DOWNLOAD_COMPLETE = 'download_complete',
  STILL_SAVING = 'still_saving',
  IMAGE = 'image',
  FILE = 'file',
  CANCEL = 'cancel',
  SCROLL_DOCUMENT = 'scroll_document',
  BITABLE_OPENAPI_REQUIRED = 'bitable_openapi_required',
}

enum ToastKey {
  DOWNLOADING = 'downloading',
  REPORT_BUG = 'report_bug',
}

i18next
  .init({
    lng: docx.language,
    resources: {
      en: {
        translation: {
          [TranslationKey.CONTENT_LOADING]:
            'Part of the content is still loading and cannot be downloaded at the moment. Please wait for loading to complete and retry',
          [TranslationKey.UNKNOWN_ERROR]: 'Unknown error during download',
          [TranslationKey.NOT_SUPPORT]:
            'This is not a lark document page and cannot be downloaded as Markdown',
          [TranslationKey.NOT_SUPPORT_DOC_1_0]:
            'This is a old version lark document page and cannot be downloaded as Markdown',
          [TranslationKey.DOWNLOADING_FILE]:
            'Download {{name}} in: {{progress}}% (please do not refresh or close the page)',
          [TranslationKey.FAILED_TO_DOWNLOAD]: 'Failed to download {{name}}',
          [TranslationKey.STILL_SAVING]:
            'Still saving (please do not refresh or close the page)',
          [TranslationKey.DOWNLOAD_PROGRESS]:
            '{{name}} download progress: {{progress}} %',
          [TranslationKey.DOWNLOAD_COMPLETE]: 'Download complete',
          [TranslationKey.IMAGE]: 'Image',
          [TranslationKey.FILE]: 'File',
          [TranslationKey.CANCEL]: 'Cancel',
          [TranslationKey.SCROLL_DOCUMENT]: 'Scrolling to load document',
          [TranslationKey.BITABLE_OPENAPI_REQUIRED]:
            'Standalone Bitable data is not available in the page DOM. OpenAPI support is required to export it completely.',
        },
        ...en,
      },
      zh: {
        translation: {
          [TranslationKey.CONTENT_LOADING]:
            '部分内容仍在加载中，暂时无法下载。请等待加载完成后重试',
          [TranslationKey.UNKNOWN_ERROR]: '下载过程中出现未知错误',
          [TranslationKey.NOT_SUPPORT]:
            '这不是一个飞书文档页面，无法下载为 Markdown',
          [TranslationKey.NOT_SUPPORT_DOC_1_0]:
            '这是一个旧版飞书文档页面，无法下载为 Markdown',
          [TranslationKey.DOWNLOADING_FILE]:
            '下载 {{name}} 中：{{progress}}%（请不要刷新或关闭页面）',
          [TranslationKey.FAILED_TO_DOWNLOAD]: '下载 {{name}} 失败',
          [TranslationKey.STILL_SAVING]: '仍在保存中（请不要刷新或关闭页面）',
          [TranslationKey.DOWNLOAD_PROGRESS]: '{{name}}下载进度：{{progress}}%',
          [TranslationKey.DOWNLOAD_COMPLETE]: '下载完成',
          [TranslationKey.IMAGE]: '图片',
          [TranslationKey.FILE]: '文件',
          [TranslationKey.CANCEL]: '取消',
          [TranslationKey.SCROLL_DOCUMENT]: '滚动中，以便加载文档',
          [TranslationKey.BITABLE_OPENAPI_REQUIRED]:
            '独立多维表格数据未暴露在页面 DOM 中，需要接入飞书 OpenAPI 才能完整导出。',
        },
        ...zh,
      },
    },
  })
  .catch(console.error)

interface ProgressOptions {
  onProgress?: (progress: number) => void
  onComplete?: () => void
}

async function toBlob(
  response: Response,
  options: ProgressOptions = {},
): Promise<Blob> {
  if (!response.ok) {
    throw new Error(`HTTP error! status: ${response.status.toFixed()}`)
  }

  if (!response.body) {
    throw new Error('This request has no response body.')
  }

  const { onProgress, onComplete } = options

  const reader = response.body.getReader()
  const contentLength = parseInt(
    response.headers.get('Content-Length') ?? '0',
    10,
  )

  let receivedLength = 0
  const chunks = []

  let _done = false
  while (!_done) {
    const { done, value } = await reader.read()

    _done = done

    if (done) {
      onComplete?.()

      break
    }

    chunks.push(value)
    receivedLength += value.length

    onProgress?.(receivedLength / contentLength)
  }

  const blob = new Blob(chunks)

  return blob
}

const downloadImage = async (
  image: mdast.Image,
  options: {
    signal?: AbortSignal
    useUUID?: boolean
    markdownFileName?: string
  } = {},
): Promise<DownloadResult | null> => {
  if (!image.data) return null

  const { signal, useUUID = false, markdownFileName = '' } = options

  const { name: originName, fetchSources, fetchBlob } = image.data

  const result = await withSignal(
    async isAborted => {
      try {
        // whiteboard
        if (fetchBlob) {
          if (isAborted()) {
            return null
          }

          const content = await fetchBlob()
          if (!content) return null

          const baseName = markdownFileName
            ? `${markdownFileName}-diagram.png`
            : 'diagram.png'
          const name = useUUID
            ? uniqueFileName.generateWithUUID(baseName)
            : uniqueFileName.generate(baseName)
          const filename = `images/${name}`

          image.url = filename

          return {
            filename,
            content,
          }
        }

        // image
        if (originName && fetchSources) {
          if (isAborted()) {
            return null
          }
          const sources = await fetchSources()
          if (!sources) return null

          const baseName = markdownFileName
            ? `${markdownFileName}-${originName}`
            : originName
          const name = useUUID
            ? uniqueFileName.generateWithUUID(baseName)
            : uniqueFileName.generate(baseName)
          const filename = `images/${name}`

          const { src } = sources
          if (isAborted()) {
            return null
          }
          const response = await fetch(src, {
            signal,
            credentials: 'include',
          })

          try {
            if (isAborted()) {
              return null
            }
            const blob = await toBlob(response, {
              onProgress: progress => {
                if (isAborted()) {
                  Toast.remove(filename)

                  return
                }

                Toast.loading({
                  content: i18next.t(TranslationKey.DOWNLOADING_FILE, {
                    name,
                    progress: Math.floor(progress * OneHundred),
                  }),
                  keepAlive: true,
                  key: filename,
                })
              },
            })

            image.url = filename

            return {
              filename,
              content: blob,
            }
          } finally {
            Toast.remove(filename)
          }
        }

        return null
      } catch (error) {
        const isAbortError =
          isAborted() ||
          (error instanceof DOMException && error.name === 'AbortError')

        if (!isAbortError) {
          Toast.error({
            content: i18next.t(TranslationKey.FAILED_TO_DOWNLOAD, {
              name: originName,
            }),
            actionText: i18next.t(CommonTranslationKey.CONFIRM_REPORT_BUG, {
              ns: Namespace.COMMON,
            }),
            onActionClick: () => {
              reportBug(error)
            },
          })
        }

        return null
      }
    },
    { signal },
  )

  return result
}

const downloadFile = async (
  file: mdast.Link,
  options: {
    signal?: AbortSignal
    useUUID?: boolean
    markdownFileName?: string
  } = {},
): Promise<DownloadResult | null> => {
  if (!file.data?.name || !file.data.fetchFile) return null

  const { signal, useUUID = false, markdownFileName = '' } = options

  const { name, fetchFile } = file.data

  let controller = new AbortController()

  const cancel = () => {
    controller.abort()
  }

  const result = await withSignal(
    async () => {
      try {
        const baseName = markdownFileName ? `${markdownFileName}-${name}` : name
        const filename = `files/${
          useUUID
            ? uniqueFileName.generateWithUUID(baseName)
            : uniqueFileName.generate(baseName)
        }`

        const response = await fetchFile({ signal: controller.signal })
        try {
          const blob = await toBlob(response, {
            onProgress: progress => {
              Toast.loading({
                content: i18next.t(TranslationKey.DOWNLOADING_FILE, {
                  name,
                  progress: Math.floor(progress * OneHundred),
                }),
                keepAlive: true,
                key: filename,
                actionText: i18next.t(TranslationKey.CANCEL),
                onActionClick: cancel,
              })
            },
          })

          file.url = filename

          return {
            filename,
            content: blob,
          }
        } finally {
          Toast.remove(filename)
        }
      } catch (error) {
        if (error instanceof DOMException && error.name === 'AbortError') {
          return null
        }

        Toast.error({
          content: i18next.t(TranslationKey.FAILED_TO_DOWNLOAD, {
            name,
          }),
          actionText: i18next.t(CommonTranslationKey.CONFIRM_REPORT_BUG, {
            ns: Namespace.COMMON,
          }),
          onActionClick: () => {
            reportBug(error)
          },
        })

        return null
      }
    },
    { signal, onAbort: cancel },
  )

  // @ts-expect-error remove reference
  controller = null

  return result
}

interface DownloadResult {
  filename: string
  content: Blob
}

type File = mdast.Image | mdast.Link

const downloadFiles = async (
  files: File[],
  options: ProgressOptions & {
    /**
     * @default 3
     */
    batchSize?: number
    signal?: AbortSignal
    useUUID?: boolean
    markdownFileName?: string
  } = {},
): Promise<DownloadResult[]> => {
  const {
    onProgress,
    onComplete,
    batchSize = 3,
    signal,
    useUUID = false,
    markdownFileName = '',
  } = options

  let completeEventCalled = false
  const onCompleteOnce = () => {
    if (!completeEventCalled) {
      completeEventCalled = true
      onComplete?.()
    }
  }

  const results = await withSignal(
    async isAborted => {
      const _results: DownloadResult[] = []

      const totalSize = files.length
      let downloadedSize = 0

      for (const batch of cluster(files, batchSize)) {
        if (isAborted()) {
          break
        }

        await Promise.allSettled(
          batch.map(async file => {
            if (isAborted()) {
              return
            }

            try {
              const result =
                file.type === 'image'
                  ? await downloadImage(file, {
                      signal,
                      useUUID,
                      markdownFileName,
                    })
                  : await downloadFile(file, {
                      signal,
                      useUUID,
                      markdownFileName,
                    })

              if (result) {
                _results.push(result)
              }
            } finally {
              downloadedSize++

              if (!isAborted()) {
                onProgress?.(downloadedSize / totalSize)
              }
            }
          }),
        )
      }

      onCompleteOnce()

      return _results
    },
    {
      signal,
      onAbort: onCompleteOnce,
    },
  )

  return results ?? []
}

interface PrepareResult {
  isReady: boolean
  recoverScrollTop?: () => void
}

const prepare = async (): Promise<PrepareResult> => {
  const checkIsReady = () => docx.isReady({ checkWhiteboard: true })

  let recoverScrollTop

  if (!checkIsReady()) {
    const initialScrollTop = docx.container?.scrollTop ?? 0
    recoverScrollTop = () => {
      docx.scrollTo({
        top: initialScrollTop,
        behavior: 'instant',
      })
    }

    let top = 0

    docx.scrollTo({
      top,
      behavior: 'instant',
    })

    const maxTryTimes = OneHundred
    let tryTimes = 0

    Toast.loading({
      content: i18next.t(TranslationKey.SCROLL_DOCUMENT),
      keepAlive: true,
      key: TranslationKey.SCROLL_DOCUMENT,
      actionText: i18next.t(TranslationKey.CANCEL),
      onActionClick: () => {
        tryTimes = maxTryTimes
      },
    })

    while (!checkIsReady() && tryTimes <= maxTryTimes) {
      docx.scrollTo({
        top,
        behavior: 'smooth',
      })

      await waitFor(0.4 * Second)

      tryTimes++

      top = docx.container?.scrollHeight ?? 0
    }

    Toast.remove(TranslationKey.SCROLL_DOCUMENT)
  }

  return {
    isReady: checkIsReady(),
    recoverScrollTop,
  }
}

const downloadStandaloneBitableAsMarkdown = async (): Promise<void> => {
  const settings = await getSettings([SettingKey.DownloadMethod])
  const table =
    (await extractStandaloneBitableTableFromWebApi()) ??
    extractStandaloneBitableTable()

  if (!table) {
    Toast.warning({
      content: i18next.t(TranslationKey.BITABLE_OPENAPI_REQUIRED),
    })
    throw new Error(DOWNLOAD_ABORTED)
  }

  const filename = `${normalizeFileName(table.title.slice(0, OneHundred))}.md`
  const toBlob = (): Blob =>
    new Blob([bitableToMarkdown(table)], { type: 'text/markdown' })

  if (
    settings[SettingKey.DownloadMethod] === DownloadMethod.ShowSaveFilePicker &&
    supported
  ) {
    if (!navigator.userActivation.isActive) {
      const confirmed = await confirm()
      if (!confirmed) throw new Error(DOWNLOAD_ABORTED)
    }

    await fileSave(toBlob(), {
      fileName: filename,
      extensions: ['.md'],
    })
  } else {
    legacyFileSave(toBlob(), { fileName: filename })
  }
}

const findLegacyEditorBody = (): HTMLElement | null =>
  document.getElementById('innerdocbody') ??
  document.querySelector<HTMLElement>('.innerdocbody')

const findLegacyScrollContainer = (
  editorBody: HTMLElement | null,
): HTMLElement => {
  let el: HTMLElement | null = editorBody?.parentElement ?? null
  while (el) {
    const style = window.getComputedStyle(el)
    if (
      (style.overflowY === 'auto' || style.overflowY === 'scroll') &&
      el.scrollHeight > el.clientHeight
    ) {
      return el
    }
    el = el.parentElement
  }

  return (document.scrollingElement as HTMLElement) || document.documentElement
}

const legacyLineKey = (
  line: HTMLElement,
  index: number,
  scrollTop: number,
): string => {
  const imageKeys = getLegacyLineImageSignature(line)
  const text = normalizeLegacyLineText(line.textContent ?? '')

  const classLineId = Array.from(line.classList).find(className =>
    className.startsWith('lineguid-'),
  )
  if (classLineId) return `${classLineId}:${text}:${imageKeys}`

  const explicitId =
    line.id ||
    line.dataset['lineId'] ||
    line.dataset['lineGuid'] ||
    line.dataset['guid'] ||
    line.getAttribute('data-line-id') ||
    line.getAttribute('data-lineguid')
  if (explicitId) return `${explicitId}:${text}:${imageKeys}`

  if (!imageKeys && text.length >= 16) {
    return `fallback-content:${text}`
  }

  const top = Math.round(line.getBoundingClientRect().top + scrollTop)

  return `fallback:${top}:${index}:${text}:${imageKeys}`
}

const normalizeLegacyLineText = (text: string): string =>
  text
    .replace(/[\u200B\u200C\u200D\uFEFF]/g, '')
    .replace(/\s+/g, ' ')
    .trim()

const getLegacyLineImageSignature = (line: HTMLElement): string =>
  Array.from(line.querySelectorAll<HTMLImageElement>('img'))
    .map(img => getLegacyImageSource(img) || img.currentSrc || img.src || '')
    .join('|')

const hasLegacyLineContent = (line: HTMLElement): boolean => {
  const text = line.textContent
    ?.replace(/[\u200B\u200C\u200D\uFEFF]/g, '')
    .trim()
  if (text) return true

  return line.querySelector('img, canvas, svg, table') !== null
}

const hasOwnLegacyVisualContent = (line: HTMLElement): boolean => {
  for (const child of Array.from(line.children)) {
    const tag = child.tagName.toLowerCase()
    if (tag === 'img' || tag === 'canvas' || tag === 'svg' || tag === 'table') {
      return true
    }
  }

  return false
}

const isLegacyCompactTableLine = (line: HTMLElement): boolean => {
  const rawText = line.textContent ?? ''
  if (!/[\u200B\u200C\u200D\uFEFF]/.test(rawText)) return false

  const cells = rawText
    .split(/[\u200B\u200C\u200D\uFEFF]+/)
    .map(normalizeLegacyLineText)
    .filter(Boolean)

  return cells.length >= 4
}

const getLegacyRenderedLineCandidates = (
  editorBody: HTMLElement,
): HTMLElement[] => {
  const candidates = new Set<HTMLElement>()

  editorBody
    .querySelectorAll<HTMLElement>(
      [
        '.ace-line',
        '[class*="lineguid-"]',
        '[data-line-id]',
        '[data-lineguid]',
        '[data-line]',
      ].join(', '),
    )
    .forEach(line => candidates.add(line))

  // Some old Feishu rows are rendered as direct child nodes without .ace-line.
  // Only include direct children; broad descendant list selectors also match
  // Feishu smart-link/table UI and leak unrelated card/outline text.
  for (const child of Array.from(editorBody.children) as HTMLElement[]) {
    if (child.querySelector('.ace-line')) continue
    if (hasLegacyLineContent(child)) candidates.add(child)
  }

  const sorted = Array.from(candidates).sort((a, b) => {
    if (a === b) return 0
    return a.compareDocumentPosition(b) & Node.DOCUMENT_POSITION_FOLLOWING
      ? -1
      : 1
  })

  return sorted.filter(candidate => {
    const containingParentHandlesChildren = sorted.some(
      other =>
        other !== candidate &&
        other.contains(candidate) &&
        (isLegacyCompactTableLine(other) || hasOwnLegacyVisualContent(other)),
    )
    if (containingParentHandlesChildren) return false

    const containedCandidates = sorted.filter(
      other => other !== candidate && candidate.contains(other),
    )
    const hasContainedCandidate = containedCandidates.length > 0
    const isHardLine =
      candidate.classList.contains('ace-line') ||
      Array.from(candidate.classList).some(className =>
        className.startsWith('lineguid-'),
      ) ||
      candidate.hasAttribute('data-line-id') ||
      candidate.hasAttribute('data-lineguid') ||
      candidate.hasAttribute('data-line')
    if (isHardLine && !hasContainedCandidate) return true
    if (isHardLine && hasContainedCandidate) {
      return (
        isLegacyCompactTableLine(candidate) ||
        hasOwnLegacyVisualContent(candidate)
      )
    }

    return !hasContainedCandidate
  })
}

const normalizeLegacyImageSource = (src: string): string => {
  const trimmed = src.trim()
  if (
    !trimmed ||
    trimmed.startsWith('data:') ||
    trimmed.startsWith('blob:') ||
    trimmed.startsWith('http://') ||
    trimmed.startsWith('https://')
  ) {
    return trimmed
  }

  try {
    return new URL(trimmed, location.href).href
  } catch {
    return trimmed
  }
}

const isPlaceholderLegacyImageSource = (src: string): boolean => {
  const normalized = src.trim()
  if (
    !normalized ||
    normalized === 'about:blank' ||
    normalized === '//about:blank'
  ) {
    return true
  }
  if (
    /transparent|spacer|placeholder|blank|doc-image-loading|loading-rotate|loading-dash/i.test(
      normalized,
    )
  ) {
    return true
  }

  return normalized.startsWith('data:image/') && normalized.length < 200
}

const firstLegacySrcsetCandidate = (srcset: string | null): string =>
  srcset?.split(',')[0]?.trim().split(/\s+/)[0] ?? ''

const getLegacyImageSource = (img: HTMLImageElement): string => {
  const candidates = [
    img.getAttribute('data-src'),
    img.getAttribute('data-original'),
    img.getAttribute('data-lazy-src'),
    img.getAttribute('data-actualsrc'),
    img.getAttribute('data-origin-src'),
    img.getAttribute('data-url'),
    firstLegacySrcsetCandidate(img.getAttribute('data-srcset')),
    firstLegacySrcsetCandidate(img.getAttribute('srcset')),
    img.currentSrc,
    img.getAttribute('src'),
    img.src,
  ]
    .filter((src): src is string => !!src)
    .map(normalizeLegacyImageSource)

  return (
    candidates.find(src => !isPlaceholderLegacyImageSource(src)) ??
    candidates[0] ??
    ''
  )
}

const hasPendingLegacyImage = (editorBody: HTMLElement | null): boolean => {
  if (!editorBody) return false

  const hasPendingImg = Array.from(editorBody.querySelectorAll('img')).some(
    img => {
      const src = getLegacyImageSource(img)
      return !src || isPlaceholderLegacyImageSource(src)
    },
  )
  if (hasPendingImg) return true

  return Array.from(editorBody.querySelectorAll('svg')).some(svg =>
    /doc-image-loading|loading-rotate|loading-dash/i.test(svg.outerHTML),
  )
}

const waitForLegacyImages = async (
  editorBody: HTMLElement | null,
  timeoutMs = 800,
): Promise<void> => {
  const start = Date.now()
  while (hasPendingLegacyImage(editorBody) && Date.now() - start < timeoutMs) {
    await waitFor(0.2 * Second)
  }
}

const hydrateLegacySnapshotClone = (
  source: HTMLElement,
  clone: HTMLElement,
): void => {
  const sourceImages = Array.from(source.querySelectorAll('img'))
  const cloneImages = Array.from(clone.querySelectorAll('img'))
  sourceImages.forEach((sourceImage, index) => {
    const cloneImage = cloneImages[index]
    if (!cloneImage) return

    const src = getLegacyImageSource(sourceImage)
    if (!src) return

    cloneImage.setAttribute('src', src)
    cloneImage.removeAttribute('srcset')
  })

  const sourceCanvases = Array.from(source.querySelectorAll('canvas'))
  const cloneCanvases = Array.from(clone.querySelectorAll('canvas'))
  sourceCanvases.forEach((sourceCanvas, index) => {
    const cloneCanvas = cloneCanvases[index]
    if (!cloneCanvas) return

    try {
      if (sourceCanvas.width <= 1 || sourceCanvas.height <= 1) return
      const src = sourceCanvas.toDataURL('image/png')
      if (!src || isPlaceholderLegacyImageSource(src)) return

      const image = document.createElement('img')
      image.src = src
      image.alt = sourceCanvas.getAttribute('aria-label') ?? 'drawing'
      image.width = sourceCanvas.width
      image.height = sourceCanvas.height
      image.style.cssText = sourceCanvas.getAttribute('style') ?? ''
      cloneCanvas.replaceWith(image)
    } catch {
      // Canvas can be tainted by cross-origin content; leave the cloned node in place.
    }
  })
}

const collectLegacyRenderedLines = (
  editorBody: HTMLElement | null,
  scrollTarget: HTMLElement,
  snapshots: Map<string, HTMLElement>,
): number => {
  if (!editorBody) return 0

  const lines = getLegacyRenderedLineCandidates(editorBody).filter(line => {
    if (line.hidden || line.getAttribute('aria-hidden') === 'true') return false
    if (line.closest('.adit-virtual-scroll-placeholder')) return false
    if (!hasLegacyLineContent(line)) return false
    return true
  })

  for (const [index, line] of lines.entries()) {
    const key = legacyLineKey(line, index, scrollTarget.scrollTop)
    if (snapshots.has(key)) continue

    const clone = line.cloneNode(true) as HTMLElement
    clone.removeAttribute('style')
    clone.dataset['cdcLegacySnapshotLine'] = 'true'
    hydrateLegacySnapshotClone(line, clone)
    snapshots.set(key, clone)
  }

  return lines.length
}

const formatLegacyMarkdown = (markdown: string): string => {
  const lines = markdown.split('\n')
  const formatted: string[] = []
  let blankCount = 0
  let inFence = false

  for (const line of lines) {
    if (/^\s*(```|~~~)/.test(line)) {
      inFence = !inFence
      formatted.push(line)
      blankCount = 0
      continue
    }

    const normalizedLine = inFence ? line : line.replace(/[ \t]+$/g, '')

    if (!inFence && normalizedLine.trim() === '') {
      blankCount++
      if (blankCount <= 1) formatted.push('')
      continue
    }

    blankCount = 0
    formatted.push(normalizedLine)
  }

  return `${formatted.join('\n').trimEnd()}\n`
}

const createLegacySnapshotContainer = (
  lines: Iterable<HTMLElement>,
  sourceEditorBody: HTMLElement | null,
): HTMLElement => {
  const snapshot = document.createElement('div')
  snapshot.className = 'innerdocbody cdc-legacy-snapshot'
  snapshot.dataset['cdcLegacySnapshot'] = 'true'
  snapshot.style.cssText =
    'position:fixed;left:-100000px;top:0;width:900px;opacity:0;pointer-events:none;z-index:-1;'

  const sourceWidth = sourceEditorBody?.getBoundingClientRect().width
  if (sourceWidth && sourceWidth > 0) {
    snapshot.style.width = `${Math.round(sourceWidth)}px`
  }

  for (const line of lines) {
    snapshot.appendChild(line)
  }

  document.documentElement.appendChild(snapshot)
  return snapshot
}

const downloadLegacyDocAsMarkdown = async (): Promise<void> => {
  if (import.meta.env.DEV) {
    console.log('[download-md] downloadLegacyDocAsMarkdown called')
  }
  debugLog(`--- Legacy Doc Parser Debug ---`)

  // 直接使用默认设置，跳过getSettings避免MAIN/ISOLATED world消息通信超时问题
  const settings = {
    [SettingKey.DownloadMethod]: supported
      ? DownloadMethod.ShowSaveFilePicker
      : DownloadMethod.Direct,
    [SettingKey.TextHighlight]: true,
  }

  // 第一步：先滚动页面加载所有虚拟滚动内容（Etherpad也用虚拟滚动！）
  debugLog(`Step 1: Scroll to load all content...`)
  const editorBody = findLegacyEditorBody()
  debugLog(
    `editorBody found: ${!!editorBody}, tag=${editorBody?.tagName}, class=${editorBody?.className?.substring(0, 60)}`,
  )

  // Choose the correct scroll target (window OR editor container, whichever actually scrolls)
  const scrollTarget = findLegacyScrollContainer(editorBody)
  const viewportHeight =
    Math.min(window.innerHeight, scrollTarget.clientHeight) * 0.45
  const initialScroll = scrollTarget.scrollTop
  const lineSnapshots = new Map<string, HTMLElement>()

  let lastTotalLines = 0
  let stableCount = 0
  const maxScrolls = 180
  let reachedBottom = false
  let snapshotContainer: HTMLElement | null = null

  Toast.loading({
    content: i18next.t(TranslationKey.SCROLL_DOCUMENT),
    keepAlive: true,
    key: TranslationKey.SCROLL_DOCUMENT,
  })

  debugLog(
    `Scrolling target: scrollHeight=${scrollTarget.scrollHeight}, clientHeight=${scrollTarget.clientHeight}, viewportStep=${Math.round(viewportHeight)}`,
  )

  scrollTarget.scrollTo({ top: 0, behavior: 'instant' as ScrollBehavior })
  await waitFor(0.6 * Second)
  await waitForLegacyImages(editorBody)
  collectLegacyRenderedLines(editorBody, scrollTarget, lineSnapshots)
  debugLog(
    `  initial collect: top=${Math.round(scrollTarget.scrollTop)}, ace-lines=${document.querySelectorAll('.ace-line').length}, collected=${lineSnapshots.size}`,
  )

  // Gradually scroll down in viewport-sized increments, DO NOT JUMP TO BOTTOM — virtual scroll only renders content you scroll past!
  for (let i = 0; i < maxScrolls; i++) {
    // Scroll down one viewport at a time
    scrollTarget.scrollBy({
      top: viewportHeight,
      behavior: 'instant' as ScrollBehavior,
    })
    await waitFor(0.4 * Second)
    await waitForLegacyImages(editorBody)

    // Check if we reached the bottom
    const currentScrollTop = scrollTarget.scrollTop
    const currentScrollHeight = scrollTarget.scrollHeight
    const atBottom =
      currentScrollTop + scrollTarget.clientHeight >= currentScrollHeight - 20

    const currentLines = collectLegacyRenderedLines(
      editorBody,
      scrollTarget,
      lineSnapshots,
    )
    const placeholders = document.querySelectorAll(
      '.adit-virtual-scroll-placeholder',
    ).length
    debugLog(
      `  scroll ${i + 1}/${maxScrolls}: top=${Math.round(currentScrollTop)}, totalHeight=${currentScrollHeight}, atBottom=${atBottom}, placeholders=${placeholders}, ace-lines=${currentLines}, collected=${lineSnapshots.size}`,
    )

    if (atBottom) {
      if (!reachedBottom) {
        debugLog(
          `  → Reached bottom for the first time, waiting extra for final content to render...`,
        )
        reachedBottom = true
        // Scroll back up a little and down again to trigger any remaining lazy content
        scrollTarget.scrollBy({
          top: -viewportHeight * 2,
          behavior: 'instant' as ScrollBehavior,
        })
        await waitFor(0.3 * Second)
        await waitForLegacyImages(editorBody, 800)
        collectLegacyRenderedLines(editorBody, scrollTarget, lineSnapshots)
        scrollTarget.scrollTo({
          top: currentScrollHeight,
          behavior: 'instant' as ScrollBehavior,
        })
        await waitFor(1 * Second)
        await waitForLegacyImages(editorBody)
        collectLegacyRenderedLines(editorBody, scrollTarget, lineSnapshots)
      }
    }

    if (lineSnapshots.size === lastTotalLines) {
      stableCount++
      if (reachedBottom && stableCount >= 5) break // 到达底部后连续5次行数不变，说明加载完成
    } else {
      stableCount = 0
      lastTotalLines = lineSnapshots.size
    }
  }

  // Wait extra for images to load
  debugLog(`  → Waiting 2 seconds for images to load...`)
  await waitFor(2 * Second)
  await waitForLegacyImages(editorBody, 2000)
  collectLegacyRenderedLines(editorBody, scrollTarget, lineSnapshots)

  debugLog(
    `  → Reverse sweep to catch virtualized rows skipped by downward scroll...`,
  )
  const reverseStart = scrollTarget.scrollTop
  const reverseStep = Math.max(120, viewportHeight)
  for (let top = reverseStart; top > 0; top -= reverseStep) {
    scrollTarget.scrollTo({
      top,
      behavior: 'instant' as ScrollBehavior,
    })
    await waitFor(0.25 * Second)
    await waitForLegacyImages(editorBody, 300)
    collectLegacyRenderedLines(editorBody, scrollTarget, lineSnapshots)
  }
  scrollTarget.scrollTo({ top: 0, behavior: 'instant' as ScrollBehavior })
  await waitFor(0.25 * Second)
  collectLegacyRenderedLines(editorBody, scrollTarget, lineSnapshots)

  // Scroll back to top
  scrollTarget.scrollTo({
    top: initialScroll,
    behavior: 'instant' as ScrollBehavior,
  })
  await waitFor(0.5 * Second)
  Toast.remove(TranslationKey.SCROLL_DOCUMENT)
  debugLog(
    `Scroll complete. Current ace-lines: ${document.querySelectorAll('.ace-line').length}, collected lines: ${lineSnapshots.size}, total images: ${document.querySelectorAll('img').length}`,
  )

  if (lineSnapshots.size > 0) {
    snapshotContainer = createLegacySnapshotContainer(
      lineSnapshots.values(),
      editorBody,
    )
    debugLog(
      `Snapshot container created: childElementCount=${snapshotContainer.childElementCount}, textContent length=${snapshotContainer.textContent?.trim().length ?? 0}`,
    )
  }

  // 初始化旧版文档解析器
  debugLog(`Step 2: Initialize parser...`)
  doc.init({
    highlight: settings[SettingKey.TextHighlight],
    container: snapshotContainer ?? undefined,
  })

  debugLog(`doc.isReady: ${doc.isReady}`)
  debugLog(`doc.blocksCount: ${doc.blocksCount}`)
  debugLog(`doc.imagesCount: ${doc.imagesCount}`)
  debugLog(`doc.pageTitle: ${doc.pageTitle}`)
  const c = doc.container
  if (c) {
    debugLog(
      `container: <${c.tagName.toLowerCase()}> class="${c.className?.substring(0, 80)}"`,
    )
    debugLog(`  childElementCount: ${c.childElementCount}`)
    debugLog(`  textContent length: ${c.textContent?.trim().length ?? 0}`)
    // Show first few children structure
    const kids = Array.from(c.children).slice(0, 5)
    debugLog(`  first ${kids.length} children:`)
    kids.forEach((k, i) => {
      debugLog(
        `    [${i}] <${k.tagName.toLowerCase()}> class="${(k as HTMLElement).className?.substring(0, 60)}" text="${k.textContent?.trim().substring(0, 50)}"`,
      )
    })
  } else {
    debugLog(`❌ container is NULL!`)
  }

  if (!doc.isReady) {
    debugLog(`❌ doc.isReady is false, aborting`)
    snapshotContainer?.remove()
    Toast.warning({
      content: i18next.t(TranslationKey.CONTENT_LOADING),
    })
    throw new Error(DOWNLOAD_ABORTED)
  }

  debugLog(`Calling doc.intoMarkdownAST()...`)
  const { root, images } = (() => {
    try {
      return doc.intoMarkdownAST({
        highlight: settings[SettingKey.TextHighlight],
      })
    } finally {
      snapshotContainer?.remove()
    }
  })()

  debugLog(`root.children.length: ${root.children.length}`)
  debugLog(`images.length: ${images.length}`)
  if (root.children.length > 0) {
    const firstBlocks = root.children.slice(0, 5)
    debugLog(`first ${firstBlocks.length} root children types:`)
    firstBlocks.forEach((b, i) => {
      debugLog(`  [${i}] type=${b.type}`)
    })
  }

  // Generate quick markdown preview
  try {
    const mdPreview = Docx.stringify(root).substring(0, 500)
    debugLog(`--- Markdown Preview (first 500 chars) ---`)
    debugLog(mdPreview.replace(/\n/g, '⏎\n'))
  } catch (e) {
    debugLog(`❌ Failed to stringify markdown: ${e}`)
  }

  const recommendName = doc.pageTitle
    ? normalizeFileName(doc.pageTitle.slice(0, OneHundred))
    : 'doc'
  const isZip = images.length > 0
  const ext = isZip ? '.zip' : '.md'
  const filename = `${recommendName}${ext}`
  debugLog(`output filename: ${filename}`)
  debugLog(`--- End Debug ---`)

  const toBlob = async () => {
    Toast.loading({
      content: i18next.t(TranslationKey.STILL_SAVING),
      keepAlive: true,
      key: ToastKey.DOWNLOADING,
    })

    const singleFileContent = () => {
      const markdown = formatLegacyMarkdown(Docx.stringify(root))
      return new Blob([markdown])
    }

    const zipFileContent = async () => {
      const zipFs = new fs.FS()

      const imgs = images.filter(image => image.data?.fetchSources)

      const results = await Promise.all([
        downloadFiles(imgs, {
          batchSize: 15,
          onProgress: progress => {
            Toast.loading({
              content: i18next.t(TranslationKey.DOWNLOAD_PROGRESS, {
                name: i18next.t(TranslationKey.IMAGE),
                progress: Math.floor(progress * OneHundred),
              }),
              keepAlive: true,
              key: TranslationKey.IMAGE,
            })
          },
          onComplete: () => {
            Toast.remove(TranslationKey.IMAGE)
          },
        }),
      ])
      results.flat(1).forEach(({ filename, content }) => {
        zipFs.addBlob(filename, content)
      })

      const markdown = formatLegacyMarkdown(Docx.stringify(root))
      zipFs.addText(`${recommendName}.md`, markdown)

      return await zipFs.exportBlob()
    }

    const content = isZip ? await zipFileContent() : singleFileContent()
    return content
  }

  if (
    settings[SettingKey.DownloadMethod] === DownloadMethod.ShowSaveFilePicker &&
    supported
  ) {
    if (!navigator.userActivation.isActive) {
      const confirmed = await confirm()
      if (!confirmed) {
        throw new Error(DOWNLOAD_ABORTED)
      }
    }

    await fileSave(toBlob(), {
      fileName: filename,
      extensions: [ext],
    })
  } else {
    const blob = await toBlob()
    legacyFileSave(blob, {
      fileName: filename,
    })
  }
}

const main = async (options: { signal?: AbortSignal } = {}) => {
  const { signal } = options

  debugLog(`main() called`)
  debugLog(`href: ${location.href}`)
  debugLog(`pathname: ${location.pathname}`)

  const isLegacyDocPath =
    location.pathname.startsWith('/docs/') ||
    location.pathname.startsWith('/doc/')
  const hasInnerdocbody = !!document.getElementById('innerdocbody')
  const globalWindow = window as Window & {
    editor?: unknown
    PageMain?: unknown
  }

  debugLog(`isLegacyDocPath (URL check): ${isLegacyDocPath}`)
  debugLog(`hasInnerdocbody (#innerdocbody exists): ${hasInnerdocbody}`)
  debugLog(`docx.isDocx: ${docx.isDocx}`)
  debugLog(`docx.isDoc: ${docx.isDoc}`)
  debugLog(`window.editor: ${typeof globalWindow.editor !== 'undefined'}`)
  debugLog(`window.PageMain: ${typeof globalWindow.PageMain !== 'undefined'}`)
  debugLog(`#mainBox: ${!!document.querySelector('#mainBox')}`)
  debugLog(`.ace_editor: ${!!document.querySelector('.ace_editor')}`)
  debugLog(`.outerdocbody: ${!!document.querySelector('.outerdocbody')}`)

  if (import.meta.env.DEV) {
    console.log(
      '[download-md] main called, path:',
      location.pathname,
      'isLegacyDocPath:',
      isLegacyDocPath,
      'hasInnerdocbody:',
      hasInnerdocbody,
      'docx.isDocx:',
      docx.isDocx,
      'docx.isDoc:',
      docx.isDoc,
    )
  }

  if (!docx.isDocx && isStandaloneBitablePage()) {
    debugLog(`→ Detected standalone bitable page`)
    await downloadStandaloneBitableAsMarkdown()
    return
  }

  if (docx.isDoc) {
    debugLog(`→ Detected legacy/etherpad doc, using doc parser`)
    await downloadLegacyDocAsMarkdown()
    return
  }

  if (!docx.isDocx) {
    debugLog(`❌ NOT_SUPPORTED: not docx, not doc, not bitable`)
    debugLog(`URL path: ${location.pathname}`)
    debugLog(`Body classes: ${document.body.className}`)
    if (import.meta.env.DEV) {
      const errDiv = document.createElement('div')
      errDiv.style.cssText =
        'position:fixed;top:40px;left:0;right:0;z-index:2147483645;background:#c62828;color:#fff;padding:12px 16px;font:14px sans-serif;text-align:center;'
      errDiv.textContent =
        '[CDC] Error: This page type is not recognized. Check debug panel at bottom.'
      document.documentElement.appendChild(errDiv)
    }

    Toast.warning({ content: i18next.t(TranslationKey.NOT_SUPPORT) })

    throw new Error(DOWNLOAD_ABORTED)
  }

  debugLog(`→ Detected Docx page, using docx transformer`)

  const { isReady, recoverScrollTop } = await prepare()

  if (!isReady) {
    Toast.warning({
      content: i18next.t(TranslationKey.CONTENT_LOADING),
    })

    throw new Error(DOWNLOAD_ABORTED)
  }

  const settings = await getSettings([
    SettingKey.DownloadMethod,
    SettingKey.Table,
    SettingKey.Grid,
    SettingKey.TextHighlight,
    SettingKey.DownloadFileWithUniqueName,
  ])

  const { root, images, files, tableWithParents, mentionUsers } =
    docx.intoMarkdownAST({
      whiteboard: true,
      diagram: true,
      file: true,
      highlight: settings[SettingKey.TextHighlight],
      flatGrid: settings[SettingKey.Grid] === Grid.Flatten,
    })

  await transformMentionUsers(mentionUsers)

  const recommendName = docx.pageTitle
    ? normalizeFileName(docx.pageTitle.slice(0, OneHundred))
    : 'doc'
  const isZip = images.length > 0 || files.length > 0
  const ext = isZip ? '.zip' : '.md'
  const filename = `${recommendName}${ext}`

  const toBlob = async () => {
    Toast.loading({
      content: i18next.t(TranslationKey.STILL_SAVING),
      keepAlive: true,
      key: ToastKey.DOWNLOADING,
    })

    const singleFileContent = () => {
      transformTableBySettings(tableWithParents, settings)

      const markdown = Docx.stringify(root)

      return new Blob([markdown])
    }

    const zipFileContent = async () => {
      const zipFs = new fs.FS()

      const imgs = images.filter(image => image.data?.fetchSources)
      const diagrams = images.filter(image => image.data?.fetchBlob)

      const results = await Promise.all([
        downloadFiles(imgs, {
          batchSize: 15,
          onProgress: progress => {
            Toast.loading({
              content: i18next.t(TranslationKey.DOWNLOAD_PROGRESS, {
                name: i18next.t(TranslationKey.IMAGE),
                progress: Math.floor(progress * OneHundred),
              }),
              keepAlive: true,
              key: TranslationKey.IMAGE,
            })
          },
          onComplete: () => {
            Toast.remove(TranslationKey.IMAGE)
          },
          signal,
          useUUID: settings[SettingKey.DownloadFileWithUniqueName],
          markdownFileName: recommendName,
        }),
        // Diagrams must be downloaded one by one
        downloadFiles(diagrams, {
          batchSize: 1,
          signal,
          useUUID: settings[SettingKey.DownloadFileWithUniqueName],
          markdownFileName: recommendName,
        }),
        downloadFiles(files, {
          onProgress: progress => {
            Toast.loading({
              content: i18next.t(TranslationKey.DOWNLOAD_PROGRESS, {
                name: i18next.t(TranslationKey.FILE),
                progress: Math.floor(progress * OneHundred),
              }),
              keepAlive: true,
              key: TranslationKey.FILE,
            })
          },
          onComplete: () => {
            Toast.remove(TranslationKey.FILE)
          },
          signal,
          useUUID: settings[SettingKey.DownloadFileWithUniqueName],
          markdownFileName: recommendName,
        }),
      ])
      results.flat(1).forEach(({ filename, content }) => {
        zipFs.addBlob(filename, content)
      })

      transformTableBySettings(tableWithParents, settings)

      const markdown = Docx.stringify(root)

      zipFs.addText(`${recommendName}.md`, markdown)

      return await zipFs.exportBlob()
    }

    const content = isZip ? await zipFileContent() : singleFileContent()

    recoverScrollTop?.()

    return content
  }

  if (
    settings[SettingKey.DownloadMethod] === DownloadMethod.ShowSaveFilePicker &&
    supported
  ) {
    if (!navigator.userActivation.isActive) {
      const confirmed = await confirm()
      if (!confirmed) {
        throw new Error(DOWNLOAD_ABORTED)
      }
    }

    await fileSave(toBlob(), {
      fileName: filename,
      extensions: [ext],
    })
  } else {
    const blob = await toBlob()

    legacyFileSave(blob, {
      fileName: filename,
    })
  }
}

let controller = new AbortController()
main({
  signal: controller.signal,
})
  .then(() => {
    Toast.success({
      content: i18next.t(TranslationKey.DOWNLOAD_COMPLETE),
    })
  })
  .catch((error: unknown) => {
    const aborted =
      error instanceof Error &&
      (error.name === 'AbortError' || error.message === DOWNLOAD_ABORTED)

    if (aborted) {
      controller.abort()
    } else {
      Toast.error({
        key: ToastKey.REPORT_BUG,
        content: String(error),
        actionText: i18next.t(CommonTranslationKey.CONFIRM_REPORT_BUG, {
          ns: Namespace.COMMON,
        }),
        duration: Minute,
        onActionClick: () => {
          reportBug(error)

          Toast.remove(ToastKey.REPORT_BUG)
        },
      })
    }
  })
  .finally(() => {
    Toast.remove(ToastKey.DOWNLOADING)

    // @ts-expect-error remove reference
    controller = null
  })
