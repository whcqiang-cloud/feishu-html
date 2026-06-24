import i18next from 'i18next'
import { Toast, docx, type mdast } from '@dolphin/lark'
import { OneHundred, Second, waitFor } from '@dolphin/common'
import { fileSave, supported } from 'browser-fs-access'
import normalizeFileName from 'filenamify/browser'
import { cluster } from 'radash'
import { CommonTranslationKey, en, Namespace, zh } from '../common/i18n'
import { confirm } from '../common/notification'
import { legacyFileSave } from '../common/legacy'
import { reportBug } from '../common/issue'
import {
  transformMentionUsers,
  withSignal,
  transformTableBySettings,
} from '../common/utils'
import { getSettings, Grid } from '../common/settings'
import { DownloadMethod, SettingKey } from '@/common/settings'
import { toHast } from 'mdast-util-to-hast'
import { toHtml } from 'hast-util-to-html'
import { wrapIntoFullHtml, type AttachmentInfo } from './html-templates'
import {
  bitableToHtml,
  extractStandaloneBitableTable,
  extractStandaloneBitableTableFromWebApi,
  isStandaloneBitablePage,
} from './bitable-export'

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
            'This is not a lark document page and cannot be downloaded as HTML',
          [TranslationKey.NOT_SUPPORT_DOC_1_0]:
            'This is a old version lark document page and cannot be downloaded as HTML',
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
            '这不是一个飞书文档页面，无法下载为 HTML',
          [TranslationKey.NOT_SUPPORT_DOC_1_0]:
            '这是一个旧版飞书文档页面，无法下载为 HTML',
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

async function responseToBlob(
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
  const chunks: Uint8Array[] = []

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

  return new Blob(chunks as BlobPart[])
}

function blobToDataURL(blob: Blob): Promise<string> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader()
    reader.onloadend = () => {
      resolve(reader.result as string)
    }
    reader.onerror = reject
    reader.readAsDataURL(blob)
  })
}

/**
 * Compress image using Canvas API, always output as JPEG so quality parameter takes effect.
 */
function compressImage(blob: Blob, quality: number): Promise<Blob> {
  return new Promise(resolve => {
    const img = new Image()
    img.onload = () => {
      const canvas = document.createElement('canvas')
      canvas.width = img.naturalWidth
      canvas.height = img.naturalHeight

      const ctx = canvas.getContext('2d')
      if (!ctx) {
        resolve(blob)
        return
      }

      ctx.drawImage(img, 0, 0)

      canvas.toBlob(
        compressedBlob => {
          if (compressedBlob) {
            resolve(compressedBlob)
          } else {
            resolve(blob)
          }
        },
        'image/jpeg',
        quality,
      )

      URL.revokeObjectURL(img.src)
    }
    img.onerror = () => {
      resolve(blob)
    }
    img.src = URL.createObjectURL(blob)
  })
}

const downloadAndInlineImage = async (
  image: mdast.Image,
  options: {
    signal?: AbortSignal
    maxInlineSizeKb?: number
    compressEnabled?: boolean
    compressQuality?: number
  } = {},
): Promise<boolean> => {
  if (!image.data) return false

  const {
    signal,
    maxInlineSizeKb = 10240,
    compressEnabled = true,
    compressQuality = 0.6,
  } = options

  const { name: originName, fetchSources, fetchBlob } = image.data

  const result = await withSignal(
    async isAborted => {
      try {
        let blob: Blob | null = null

        // whiteboard / diagram
        if (fetchBlob) {
          if (isAborted()) return false
          blob = await fetchBlob()
          if (!blob) return false
        }
        // image
        else if (originName && fetchSources) {
          if (isAborted()) return false
          const sources = await fetchSources()
          if (!sources) return false

          const { src } = sources
          if (isAborted()) return false

          const response = await fetch(src, { signal })
          blob = await responseToBlob(response, {
            onProgress: progress => {
              if (isAborted()) return
              Toast.loading({
                content: i18next.t(TranslationKey.DOWNLOADING_FILE, {
                  name: originName,
                  progress: Math.floor(progress * OneHundred),
                }),
                keepAlive: true,
                key: originName,
              })
            },
          })
        } else {
          return false
        }

        Toast.remove(originName ?? '')

        let finalBlob: Blob = blob
        const sizeKb = blob.size / 1024

        // Compress image if compression is enabled (all images, not just oversized ones)
        if (compressEnabled) {
          const displayName = originName ?? 'unknown'
          Toast.loading({
            content: i18next.t(TranslationKey.DOWNLOADING_FILE, {
              name: displayName,
              progress: Math.floor(50),
            }),
            keepAlive: true,
            key: `${displayName}_compress`,
          })

          finalBlob = await compressImage(blob, compressQuality)
          const compressedSizeKb = finalBlob.size / 1024

          Toast.remove(`${displayName}_compress`)

          // Still too large even after compression
          if (compressedSizeKb > maxInlineSizeKb) {
            image.url = ''
            image.alt = `[image too large: ${displayName} (${Math.round(compressedSizeKb).toString()}KB after compression)]`
            return true
          }
        } else if (sizeKb > maxInlineSizeKb) {
          image.url = ''
          const displayName = originName ?? 'unknown'
          image.alt = `[image too large: ${displayName} (${Math.round(sizeKb).toString()}KB)]`
          return true
        }

        const dataUrl = await blobToDataURL(finalBlob)
        image.url = dataUrl
        return true
      } catch (error) {
        const isAbortError =
          isAborted() ||
          (error instanceof DOMException && error.name === 'AbortError')

        if (!isAbortError) {
          Toast.error({
            content: i18next.t(TranslationKey.FAILED_TO_DOWNLOAD, {
              name: originName ?? 'unknown',
            }),
            actionText: i18next.t(CommonTranslationKey.CONFIRM_REPORT_BUG, {
              ns: Namespace.COMMON,
            }),
            onActionClick: () => {
              reportBug(error)
            },
          })
        }

        return false
      }
    },
    { signal },
  )

  return result ?? false
}

const downloadAndInlineImages = async (
  images: mdast.Image[],
  options: {
    batchSize?: number
    signal?: AbortSignal
    maxInlineSizeKb?: number
    compressEnabled?: boolean
    compressQuality?: number
  } = {},
): Promise<void> => {
  const {
    batchSize = 3,
    signal,
    maxInlineSizeKb,
    compressEnabled,
    compressQuality,
  } = options

  const imgs = images.filter(image => image.data?.fetchSources)
  const diagrams = images.filter(image => image.data?.fetchBlob)

  const allImages = [...imgs, ...diagrams]

  for (const batch of cluster(allImages, batchSize)) {
    await Promise.allSettled(
      batch.map(image =>
        downloadAndInlineImage(image, {
          signal,
          maxInlineSizeKb,
          compressEnabled,
          compressQuality,
        }),
      ),
    )
  }
}

function collectAttachmentInfo(files: mdast.Link[]): AttachmentInfo[] {
  return files
    .filter((file): file is mdast.Link & { data: { name: string } } =>
      Boolean(file.data?.name),
    )
    .map(file => ({
      name: file.data.name,
      token: file.url || '',
    }))
}

/**
 * Convert inlineMath and math nodes to html nodes so that
 * mdast-util-to-hast (which has no handler for math nodes) preserves them.
 * Uses KaTeX-compatible delimiters: \(...\) for inline, \[...\] for display.
 */
function convertMathNodes(root: mdast.Root): void {
  function walk(nodes: mdast.Nodes[]): void {
    for (let i = 0; i < nodes.length; i++) {
      const node = nodes[i]

      // inlineMath and math are from mdast-util-math, not in the base mdast.Nodes union
      // We use a string check to avoid TypeScript narrowing to 'never'
      const nodeType = (node as { type: string }).type

      if (nodeType === 'inlineMath') {
        const value = (node as { value?: string }).value ?? ''
        nodes[i] = {
          type: 'html',
          value: `<span class="math inline">\\(${value}\\)</span>`,
        } as mdast.Html
      } else if (nodeType === 'math') {
        const value = (node as { value?: string }).value ?? ''
        nodes[i] = {
          type: 'html',
          value: `<div class="math display">\\[${value}\\]</div>`,
        } as mdast.Html
      } else if ('children' in node && Array.isArray(node.children)) {
        walk(node.children as mdast.Nodes[])
      }
    }
  }

  walk(root.children)
}

function mdastToHtml(root: mdast.Root): string {
  // Preprocess math nodes before toHast (which would discard them)
  convertMathNodes(root)

  const hastRoot = toHast(root, {
    allowDangerousHtml: true,
  })

  return toHtml(hastRoot, {
    allowDangerousHtml: true,
  })
}

interface PrepareResult {
  isReady: boolean
  recoverScrollTop?: () => void
}

const prepare = async (): Promise<PrepareResult> => {
  const checkIsReady = () => docx.isReady({ checkWhiteboard: true })

  let recoverScrollTop: (() => void) | undefined

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

const downloadStandaloneBitableAsHtml = async (): Promise<void> => {
  const settings = await getSettings([
    SettingKey.DownloadMethod,
    SettingKey.HtmlIncludeStyles,
    SettingKey.HtmlPrintFriendly,
  ])

  const table =
    (await extractStandaloneBitableTableFromWebApi()) ??
    extractStandaloneBitableTable()
  if (!table) {
    Toast.warning({
      content: i18next.t(TranslationKey.BITABLE_OPENAPI_REQUIRED),
    })
    throw new Error(DOWNLOAD_ABORTED)
  }

  const filename = `${normalizeFileName(table.title.slice(0, OneHundred))}.html`
  const toBlob = (): Blob => {
    const html = wrapIntoFullHtml({
      pageTitle: table.title,
      bodyHtml: bitableToHtml(table),
      attachments: [],
      includeStyles: settings[SettingKey.HtmlIncludeStyles],
      printFriendly: settings[SettingKey.HtmlPrintFriendly],
    })

    return new Blob([html], { type: 'text/html' })
  }

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
      extensions: ['.html'],
    })
  } else {
    legacyFileSave(toBlob(), { fileName: filename })
  }
}

const main = async (options: { signal?: AbortSignal } = {}) => {
  const { signal } = options

  if (!docx.isDocx && isStandaloneBitablePage()) {
    await downloadStandaloneBitableAsHtml()
    return
  }

  if (docx.isDoc) {
    Toast.warning({ content: i18next.t(TranslationKey.NOT_SUPPORT_DOC_1_0) })
    throw new Error(DOWNLOAD_ABORTED)
  }

  if (!docx.isDocx) {
    Toast.warning({ content: i18next.t(TranslationKey.NOT_SUPPORT) })
    throw new Error(DOWNLOAD_ABORTED)
  }

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
    SettingKey.HtmlImageInline,
    SettingKey.HtmlMaxInlineSizeKb,
    SettingKey.HtmlImageCompressEnabled,
    SettingKey.HtmlImageCompressQuality,
    SettingKey.HtmlIncludeStyles,
    SettingKey.HtmlPrintFriendly,
  ])

  const { root, images, files, tableWithParents, mentionUsers } =
    docx.intoMarkdownAST({
      whiteboard: true,
      diagram: true,
      file: true,
      bitable: true,
      highlight: settings[SettingKey.TextHighlight],
      flatGrid: settings[SettingKey.Grid] === Grid.Flatten,
    })

  await transformMentionUsers(mentionUsers)

  const recommendName = docx.pageTitle
    ? normalizeFileName(docx.pageTitle.slice(0, OneHundred))
    : 'doc'
  const filename = `${recommendName}.html`

  const generateHtml = async (): Promise<Blob> => {
    Toast.loading({
      content: i18next.t(TranslationKey.STILL_SAVING),
      keepAlive: true,
      key: ToastKey.DOWNLOADING,
    })

    // Step 1: Download and inline images (must happen before toHast)
    if (settings[SettingKey.HtmlImageInline]) {
      Toast.loading({
        content: i18next.t(TranslationKey.DOWNLOAD_PROGRESS, {
          name: i18next.t(TranslationKey.IMAGE),
          progress: 0,
        }),
        keepAlive: true,
        key: TranslationKey.IMAGE,
      })

      await downloadAndInlineImages(images, {
        batchSize: 3,
        signal,
        maxInlineSizeKb: settings[SettingKey.HtmlMaxInlineSizeKb],
        compressEnabled: settings[SettingKey.HtmlImageCompressEnabled],
        compressQuality: settings[SettingKey.HtmlImageCompressQuality],
      })

      Toast.remove(TranslationKey.IMAGE)
    }

    // Step 2: Transform tables to HTML
    transformTableBySettings(tableWithParents, settings)

    // Step 3: Collect attachment info (no download, just name + token)
    const attachments = collectAttachmentInfo(files)

    // Step 4: Convert mdast to HTML string
    const bodyHtml = mdastToHtml(root)

    // Step 5: Wrap into full HTML document
    const fullHtml = wrapIntoFullHtml({
      pageTitle: docx.pageTitle ?? 'Document',
      bodyHtml,
      attachments,
      includeStyles: settings[SettingKey.HtmlIncludeStyles],
      printFriendly: settings[SettingKey.HtmlPrintFriendly],
    })

    recoverScrollTop?.()

    return new Blob([fullHtml], { type: 'text/html' })
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

    await fileSave(generateHtml(), {
      fileName: filename,
      extensions: ['.html'],
    })
  } else {
    const blob = await generateHtml()

    legacyFileSave(blob, {
      fileName: filename,
    })
  }
}

const controller = new AbortController()
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
      return
    }

    console.error(error)

    Toast.error({
      content: i18next.t(TranslationKey.UNKNOWN_ERROR),
      actionText: i18next.t(CommonTranslationKey.CONFIRM_REPORT_BUG, {
        ns: Namespace.COMMON,
      }),
      onActionClick: () => {
        reportBug(error)
      },
    })
  })
