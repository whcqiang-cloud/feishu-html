import type * as mdast from 'mdast'
import type { TableWithParent } from './docx'
import { isDoc, isDocx } from './env'
import {
  isBlockquoteContent,
  isParent,
  isPhrasingContent,
  isRootContent,
  isTableCell,
} from './utils/mdast'

// ---------------------------------------------------------------------------
// Shared image collector — passed through the DOM traversal so every parsed
// image node is also registered in the result set.
// ---------------------------------------------------------------------------

interface ImageCollector {
  push(image: mdast.Image): void
}

class ArrayImageCollector implements ImageCollector {
  private _images: mdast.Image[] = []

  push(image: mdast.Image) {
    this._images.push(image)
  }

  get images(): mdast.Image[] {
    return this._images
  }
}

// ---------------------------------------------------------------------------
// Result type — structurally identical to TransformResult<mdast.Root> so
// downstream scripts (copy / view / download) can use the same pipeline.
// ---------------------------------------------------------------------------

export type { TableWithParent } from './docx'

export interface DocTransformResult {
  root: mdast.Root
  images: mdast.Image[]
  files: mdast.Link[]
  tableWithParents: TableWithParent[]
  mentionUsers: mdast.InlineCode[]
}

export interface DocParseOptions {
  highlight?: boolean
  container?: HTMLElement
}

// ---------------------------------------------------------------------------
// DOM selectors for old-version (Doc 1.0) editor content area
// ---------------------------------------------------------------------------

/**
 * Try to find the main content container of an old-version Feishu document.
 * Tries known selectors first, then falls back to finding the element
 * that contains the most block-level content.
 */
function findEditorContainer(): HTMLElement | null {
  // First try known selectors for new/old Feishu docs (not limited to #mainBox)
  const selectors = [
    // Etherpad editor core — CLASS selector first! innerdocbody is a CLASS not ID in new Etherpad docs
    '.innerdocbody',
    '#innerdocbody',
    '.ace_editor',
    // Etherpad-based new doc (2024+)
    '.outerdocbody-inner',
    '.etherpad-client-container',
    '.etherpad-container-wrapper',
    '.outerdocbody',
    '.suite-body',
    '.doc-container',
    '.editor-content',
    // Classic Doc 1.0
    '.edit-area',
    '.doc-editor',
    '#doc-editor',
    '#editor-container',
    '.lce-editor-root',
    '.editor-container',
    '.docx-content',
    '.left-content',
    '.right-content',
    // Inside mainBox
    '#mainBox .innerdocbody',
    '#mainBox .outerdocbody',
    '#mainBox .suite-body',
    '#mainBox .doc-container',
    '#mainBox .etherpad-client-container',
    '#mainBox .ace_editor',
    '#mainBox .edit-area',
    '#mainBox .doc-editor',
    '#mainBox',
    // Generic
    'article',
    'main',
    '[role="main"]',
  ]

  // Content check: look for any meaningful content including etherpad line divs
  const contentSelector =
    'h1,h2,h3,h4,h5,h6,p,img,table,ul,ol,pre,div[class*="ace-line"],div[class*="etherpad"],div[class*="line"]'

  for (const sel of selectors) {
    const el = document.querySelector<HTMLElement>(sel)
    if (el && el.querySelector(contentSelector)) {
      console.log(`[doc parser] Found container via selector: ${sel}`)
      return el
    }
  }

  // Fallback: find the element with the highest density of block content
  let bestEl: HTMLElement | null = null
  let bestScore = 0

  const allElements = document.querySelectorAll<HTMLElement>(
    'article, main, div, section',
  )
  for (const el of allElements) {
    // Skip tiny elements and non-content containers
    if (el.offsetHeight < 200 || el.offsetWidth < 200) continue
    if (
      el.closest(
        'nav, header, footer, aside, .toolbar, .sidebar, .menu, .comments, .sidebar-panel, .navigation, .header, .footer',
      )
    )
      continue

    const blocks = el.querySelectorAll(
      'h1,h2,h3,h4,h5,h6,p,img,table,ul,ol,pre,blockquote,li,div[class*="ace-line"]',
    ).length
    const textLen = el.textContent?.trim().length ?? 0

    // Need at least some content
    if (blocks === 0 && textLen < 50) continue

    // Score based on number of blocks and text length
    const score = blocks * 10 + Math.min(textLen, 5000)
    if (score > bestScore) {
      bestScore = score
      bestEl = el
    }
  }

  if (bestEl) {
    console.log(
      `[doc parser] Found container via fallback, tag=${bestEl.tagName}, class=${bestEl.className?.substring(0, 80)}, score=${bestScore}`,
    )
  } else {
    console.warn(
      '[doc parser] No content container found! Dumping DOM body classes:',
      document.body.className,
    )
  }

  return bestEl
}

/**
 * Get the page title from an old-version document.
 * Typically the title is in the page <title> tag or a specific header element.
 */
function getPageTitle(): string {
  // Try document title first
  if (document.title && document.title !== '飞书文档') {
    return cleanText(document.title)
  }

  // Try common title elements in old Doc editor
  const titleSelectors = [
    '.doc-title',
    '.editor-title',
    'h1.doc-title',
    '[class*="title"]',
  ]

  for (const sel of titleSelectors) {
    const el = document.querySelector<HTMLElement>(sel)
    if (el?.textContent) {
      const t = cleanText(el.textContent)
      if (t.length > 0) return t
    }
  }

  // Last resort: first non-empty line in editor is the title
  const innerdocbody =
    document.getElementById('innerdocbody') ||
    document.querySelector<HTMLElement>('.innerdocbody')
  if (innerdocbody) {
    for (const child of Array.from(innerdocbody.children) as HTMLElement[]) {
      if (isVirtualScrollPlaceholder(child)) continue
      const t = cleanText(child.textContent || '')
      if (t.length > 0) return t
    }
  }

  return ''
}

// ---------------------------------------------------------------------------
// DOM → mdast conversion helpers
// ---------------------------------------------------------------------------

/**
 * Strip HTML comments, zero-width characters, and collapse excessive whitespace.
 * Etherpad inserts many U+200B zero-width spaces for cursor positioning.
 */
const ZERO_WIDTH_CHARS =
  /[\u200B-\u200F\uFEFF\u2028-\u202E\u2060-\u2064\u2066-\u2069\u00A0\u00AD]/g

function cleanText(text: string): string {
  return text
    .replace(/<!--[\s\S]*?-->/g, '')
    .replace(ZERO_WIDTH_CHARS, '')
    .replace(/\r\n/g, '\n')
    .replace(/\t/g, '  ')
    .replace(/[ \t]+/g, ' ')
    .trim()
}

/**
 * Infer a filename with extension from an image URL.
 * If the URL doesn't have an extension, defaults to .png.
 */
function getImageFilename(src: string, fallback = 'image'): string {
  try {
    const dataType = (/^data:image\/([a-z0-9.+-]+)[;,]/i.exec(src))?.[1]
    if (dataType) {
      const ext = dataType.includes('svg')
        ? 'svg'
        : dataType.replace('jpeg', 'jpg')
      return `${fallback}.${ext}`
    }

    const urlWithoutQuery = src.split('?')[0].split('#')[0]
    const pathname = new URL(urlWithoutQuery, window.location.origin).pathname
    const segments = pathname.split('/').filter(Boolean)
    const lastSegment = segments[segments.length - 1] || ''

    if (/\.(png|jpe?g|gif|webp|svg|bmp|tiff?)$/i.test(lastSegment)) {
      return lastSegment || `${fallback}.png`
    }

    if (src.includes('format=png') || /\.png(\?|$)/i.test(src))
      return `${fallback}.png`
    if (
      src.includes('format=jpeg') ||
      src.includes('format=jpg') ||
      /\.jpe?g(\?|$)/i.test(src)
    )
      return `${fallback}.jpg`
    if (src.includes('format=webp') || /\.webp(\?|$)/i.test(src))
      return `${fallback}.webp`
    if (src.includes('format=gif') || /\.gif(\?|$)/i.test(src))
      return `${fallback}.gif`

    return `${fallback}.png`
  } catch {
    return `${fallback}.png`
  }
}

function normalizeImageSource(src: string): string {
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
    return new URL(trimmed, window.location.href).href
  } catch {
    return trimmed
  }
}

function isPlaceholderImageSource(src: string): boolean {
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

function srcsetFirstCandidate(srcset: string | null): string {
  return srcset?.split(',')[0]?.trim().split(/\s+/)[0] ?? ''
}

function getImageSource(imgEl: HTMLImageElement): string {
  const candidates = [
    imgEl.getAttribute('data-src'),
    imgEl.getAttribute('data-original'),
    imgEl.getAttribute('data-lazy-src'),
    imgEl.getAttribute('data-actualsrc'),
    imgEl.getAttribute('data-origin-src'),
    imgEl.getAttribute('data-url'),
    srcsetFirstCandidate(imgEl.getAttribute('data-srcset')),
    srcsetFirstCandidate(imgEl.getAttribute('srcset')),
    imgEl.currentSrc,
    imgEl.getAttribute('src'),
    imgEl.src,
  ]
    .filter((src): src is string => !!src)
    .map(normalizeImageSource)

  return (
    candidates.find(src => !isPlaceholderImageSource(src)) ??
    candidates[0] ??
    ''
  )
}

function imageNodeFromSource(
  src: string,
  alt: string,
  fallbackName: string,
): mdast.Image | null {
  if (!src || isPlaceholderImageSource(src)) return null

  return {
    type: 'image',
    url: src,
    alt: cleanText(alt),
    data: {
      name: getImageFilename(src, fallbackName),
      fetchSources: async () => ({ originSrc: src, src }),
    },
  }
}

function svgToDataUrl(svgEl: SVGElement): string | null {
  try {
    if (isLoadingSvg(svgEl)) return null
    const source = new XMLSerializer().serializeToString(svgEl)
    return `data:image/svg+xml;charset=utf-8,${encodeURIComponent(source)}`
  } catch {
    return null
  }
}

function isLoadingSvg(svgEl: SVGElement): boolean {
  const source = svgEl.outerHTML
  return /doc-image-loading|loading-rotate|loading-dash|loading/i.test(source)
}

function svgNumber(value: string | null): number {
  if (!value) return 0
  const parsed = parseFloat(value)
  return Number.isFinite(parsed) ? parsed : 0
}

function svgViewBoxSize(svgEl: SVGElement): { width: number; height: number } {
  const viewBox = svgEl.getAttribute('viewBox')?.trim().split(/\s+/) ?? []
  if (viewBox.length === 4) {
    return {
      width: svgNumber(viewBox[2]),
      height: svgNumber(viewBox[3]),
    }
  }

  return {
    width: svgNumber(svgEl.getAttribute('width')),
    height: svgNumber(svgEl.getAttribute('height')),
  }
}

function isMeaningfulSvg(svgEl: SVGElement): boolean {
  if (isLoadingSvg(svgEl)) return false

  const source = svgEl.outerHTML
  const { width, height } = svgViewBoxSize(svgEl)
  const textContent = cleanText(svgEl.textContent ?? '')
  const hasEmbeddedContent =
    svgEl.querySelector('foreignObject, image') !== null ||
    textContent.length > 0
  const vectorNodeCount = svgEl.querySelectorAll(
    'path, rect, polygon, polyline, ellipse, circle, line, use',
  ).length

  return (
    hasEmbeddedContent ||
    vectorNodeCount > 3 ||
    source.length > 1000 ||
    width >= 80 ||
    height >= 80
  )
}

function canvasToDataUrl(canvas: HTMLCanvasElement): string | null {
  try {
    const width = canvas.width || canvas.clientWidth
    const height = canvas.height || canvas.clientHeight
    if (width <= 1 || height <= 1) return null
    return canvas.toDataURL('image/png')
  } catch {
    return null
  }
}

const LEGACY_DRAWING_SELECTOR = [
  '[class*="diagram" i]',
  '[class*="whiteboard" i]',
  '[class*="drawing" i]',
  '[class*="draw" i]',
  '[class*="flowchart" i]',
  '[class*="mind" i]',
  '[data-block-type*="diagram" i]',
  '[data-block-type*="whiteboard" i]',
  '[data-node-type*="diagram" i]',
  '[data-node-type*="whiteboard" i]',
].join(',')

function elementToSvgSnapshot(el: HTMLElement): string | null {
  const rect = el.getBoundingClientRect()
  const width = Math.ceil(rect.width || el.scrollWidth || el.offsetWidth)
  const height = Math.ceil(rect.height || el.scrollHeight || el.offsetHeight)
  if (width <= 10 || height <= 10) return null

  const html = new XMLSerializer().serializeToString(el.cloneNode(true))
  const svg = [
    `<svg xmlns="http://www.w3.org/2000/svg" width="${width}" height="${height}">`,
    `<foreignObject width="100%" height="100%">${html}</foreignObject>`,
    '</svg>',
  ].join('')

  return `data:image/svg+xml;charset=utf-8,${encodeURIComponent(svg)}`
}

function parseLegacyVisualBlock(
  el: HTMLElement,
  collector: ImageCollector,
): mdast.Paragraph | null {
  const tag = el.tagName.toLowerCase()
  const hasRealImage = Array.from(el.querySelectorAll('img')).some(
    img => !isPlaceholderImageSource(getImageSource(img)),
  )
  if (hasRealImage) return null

  const canvas =
    tag === 'canvas'
      ? (el as HTMLCanvasElement)
      : el.querySelector<HTMLCanvasElement>('canvas')
  const canvasDataUrl = canvas ? canvasToDataUrl(canvas) : null
  const svg =
    tag === 'svg'
      ? (el as unknown as SVGElement)
      : Array.from(el.querySelectorAll('svg')).find(svgEl =>
          isMeaningfulSvg(svgEl),
        )
  const svgDataUrl =
    svg && isMeaningfulSvg(svg)
      ? svgToDataUrl(svg)
      : null
  const isMeaningfulSvgElement =
    tag === 'svg' && isMeaningfulSvg(el as unknown as SVGElement)
  const drawingEl =
    el.matches(LEGACY_DRAWING_SELECTOR) ||
    isMeaningfulSvgElement ||
    tag === 'canvas'
      ? el
      : el.querySelector<HTMLElement>(LEGACY_DRAWING_SELECTOR)
  const domSnapshot =
    !canvasDataUrl && !svgDataUrl && drawingEl
      ? elementToSvgSnapshot(drawingEl)
      : null
  const src = canvasDataUrl ?? svgDataUrl ?? domSnapshot
  if (!src) return null

  const image = imageNodeFromSource(
    src,
    cleanText(el.textContent ?? ''),
    'drawing',
  )
  if (!image) return null

  collector.push(image)
  return {
    type: 'paragraph',
    children: [image],
  }
}

/**
 * Elements that are "transparent" containers for phrasing content — they don't
 * add semantic meaning themselves, just wrap content. We should descend into
 * them rather than flattening to text.
 */
const TRANSPARENT_INLINE_TAGS = new Set([
  'div',
  'span',
  'section',
  'article',
  'font',
  'small',
  'sub',
  'sup',
  'mark',
  'ins',
  'del',
  'abbr',
  'bdi',
  'bdo',
  'cite',
  'data',
  'dfn',
  'kbd',
  'q',
  's',
  'samp',
  'time',
  'var',
  'wbr',
])

/**
 * Convert phrasing (inline) content from a DOM element to mdast phrasing nodes.
 * Images encountered during traversal are registered with the collector.
 */
function mdastToPhrasing(
  el: HTMLElement,
  collector: ImageCollector,
): mdast.PhrasingContent[] {
  const results: mdast.PhrasingContent[] = []

  for (const child of el.childNodes) {
    if (child.nodeType === Node.TEXT_NODE) {
      const text = cleanText(child.textContent ?? '')
      if (text) {
        results.push({ type: 'text', value: text })
      }
    } else if (child.nodeType === Node.ELEMENT_NODE) {
      const childEl = child as HTMLElement
      const tag = childEl.tagName.toLowerCase()

      // For transparent container elements, descend directly into their children
      // instead of going through processInlineElement (which would flatten them
      // to text and lose images/links/formatting).
      if (TRANSPARENT_INLINE_TAGS.has(tag)) {
        results.push(...mdastToPhrasing(childEl, collector))
        continue
      }

      const node = processInlineElement(childEl, collector)
      if (node) {
        results.push(node)
      }
    }
  }

  return results
}

/**
 * Process a single inline element into an mdast phrasing node.
 * Images are registered with the collector.
 */
function processInlineElement(
  el: HTMLElement,
  collector: ImageCollector,
): mdast.PhrasingContent | null {
  const tag = el.tagName.toLowerCase()

  switch (tag) {
    case 'strong':
    case 'b': {
      const children = mdastToPhrasing(el, collector)
      return children.length > 0 ? { type: 'strong', children } : null
    }
    case 'em':
    case 'i': {
      const children = mdastToPhrasing(el, collector)
      return children.length > 0 ? { type: 'emphasis', children } : null
    }
    case 'u': {
      const children = mdastToPhrasing(el, collector)
      const value = children
        .filter(n => n.type === 'text')
        .map(n => n.value)
        .join('')
      return value
        ? { type: 'html', value: `<u>${escapeHtml(value)}</u>` }
        : null
    }
    case 'del':
    case 's':
    case 'strike': {
      const children = mdastToPhrasing(el, collector)
      return children.length > 0 ? { type: 'delete', children } : null
    }
    case 'code': {
      const parentTag = el.parentElement?.tagName.toLowerCase()
      if (parentTag === 'pre') return null
      const value = cleanText(el.textContent ?? '')
      return value ? { type: 'inlineCode', value } : null
    }
    case 'a': {
      const href = el.getAttribute('href') ?? ''
      const children = mdastToPhrasing(el, collector)
      if (href.startsWith('mention:')) {
        const userId = href.replace('mention:', '')
        const value = children
          .filter(n => n.type === 'text')
          .map(n => n.value)
          .join('')
        if (value) {
          const mention: mdast.InlineCode = {
            type: 'inlineCode',
            value,
            data: { mentionUserId: userId },
          }
          return mention
        }
      }
      return href && children.length > 0
        ? { type: 'link', url: href, children }
        : null
    }
    case 'img': {
      const imgEl = el as HTMLImageElement
      const src = getImageSource(imgEl)
      const alt = imgEl.alt || imgEl.getAttribute('alt') || ''
      const title = imgEl.title || imgEl.getAttribute('title') || ''
      const fallbackName = cleanText(title) || cleanText(alt) || 'image'
      const image = imageNodeFromSource(src, alt, fallbackName)
      if (!image) return null

      collector.push(image)
      return image
    }
    case 'br': {
      return { type: 'text', value: '\n' }
    }
    default: {
      // For unrecognized elements, check if they contain block children;
      // if they do, return null (collectBlocks will handle them). Otherwise
      // treat them as transparent and recurse into their children — but since
      // we can only return a single phrasing node, wrap text content and
      // preserve images/links by returning an html node for complex content.
      const hasBlockChildren = el.querySelector(
        'h1,h2,h3,h4,h5,h6,p,pre,ul,ol,li,table,blockquote,figure,hr,div[class~="ace-line"],div[class~="etherpad-line"],div[class~="line"],div[data-line]',
      )
      if (hasBlockChildren) return null

      const children = mdastToPhrasing(el, collector)
      if (children.length === 0) return null

      // If all children are plain text, merge them
      const allText = children.every(n => n.type === 'text')
      if (allText) {
        return {
          type: 'text',
          value: children.map(n => ('value' in n ? n.value : '')).join(''),
        }
      }

      // For mixed content (text + images + formatting), return an html <span>
      // wrapper to preserve the content structure
      return {
        type: 'html',
        value: el.outerHTML,
      }
    }
  }
}

/**
 * Escape special characters for use inside mdast html node values.
 */
function escapeHtml(str: string): string {
  return str
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;')
}

/**
 * Merge adjacent text nodes and consolidate phrasing content.
 * Mirrors the logic in docx.ts mergePhrasingContents.
 */
function mergePhrasingContents(
  nodes: mdast.PhrasingContent[],
): mdast.PhrasingContent[] {
  if (nodes.length <= 1) return nodes

  const merged: mdast.PhrasingContent[] = [nodes[0]]

  for (let i = 1; i < nodes.length; i++) {
    const current = nodes[i]
    const prev = merged[merged.length - 1]

    if (prev.type === 'text' && current.type === 'text') {
      prev.value += current.value
    } else {
      merged.push(current)
    }
  }

  return merged
}

/**
 * Normalize a list item node based on its first paragraph's content.
 */
function normalizeListItem(
  firstPara: mdast.Paragraph,
  type: 'bullet' | 'ordered' | 'todo',
  checked?: boolean,
  seq?: number | 'auto',
): mdast.ListItem {
  const listItem: mdast.ListItem = {
    type: 'listItem',
    children: [],
    ...(type === 'todo' ? { checked: checked ?? false } : null),
    ...(type === 'ordered'
      ? {
          data: {
            seq:
              typeof seq === 'number'
                ? seq
                : seq === 'auto'
                  ? 'auto'
                  : undefined,
          },
        }
      : null),
  }

  // Move first paragraph's children into the list item directly
  listItem.children.push(...firstPara.children)

  return listItem
}

// ---------------------------------------------------------------------------
// Main DOM traversal: simplified Etherpad-first approach
// ---------------------------------------------------------------------------

/**
 * Check if an element is "empty" (contains no visible text or meaningful content).
 */
function isLineEmpty(el: HTMLElement): boolean {
  const text = cleanText(el.textContent || '')
  if (text.length > 0) return false
  if (
    el.querySelector(
      [
        'img',
        'svg',
        'canvas',
        'video',
        'audio',
        'iframe:not([src*="about:blank"])',
        'table',
        'ul',
        'ol',
        LEGACY_DRAWING_SELECTOR,
      ].join(', '),
    )
  ) {
    return false
  }
  if (el.matches(LEGACY_DRAWING_SELECTOR)) return false
  return true
}

/**
 * Flush the phrasing buffer into a paragraph node in the result array.
 */
function flushPhrasingBuffer(
  buffer: mdast.PhrasingContent[],
  result: mdast.BlockContent[],
): void {
  if (buffer.length === 0) return
  const merged = mergePhrasingContents(buffer)
  const hasContent = merged.some(
    node =>
      (node.type === 'text' && cleanText(node.value).length > 0) ||
      node.type === 'image' ||
      node.type === 'link' ||
      node.type === 'strong' ||
      node.type === 'emphasis' ||
      node.type === 'inlineCode',
  )
  if (hasContent) {
    result.push({
      type: 'paragraph',
      children: merged as mdast.paragraph['children'],
    })
  }
  buffer.length = 0
}

function phrasingNodeText(node: mdast.Nodes): string {
  if ('value' in node && typeof node.value === 'string') return node.value
  if ('children' in node && Array.isArray(node.children)) {
    return node.children
      .map(child => phrasingNodeText(child as mdast.Nodes))
      .join('\n')
  }
  return ''
}

function tableTextTokens(table: mdast.Table): string[] {
  return table.children.flatMap(row =>
    row.children.map(cell => cleanText(phrasingNodeText(cell))),
  )
}

function paragraphTextTokens(paragraph: mdast.Paragraph): string[] {
  return phrasingNodeText(paragraph)
    .split('\n')
    .map(token => cleanText(token))
    .filter(Boolean)
}

function sameTokens(left: string[], right: string[]): boolean {
  if (left.length === 0 || right.length === 0) return false
  if (left.length % right.length !== 0) return false

  return left.every((token, index) => token === right[index % right.length])
}

function removeDuplicateTableParagraphs(
  blocks: mdast.BlockContent[],
): mdast.BlockContent[] {
  const result: mdast.BlockContent[] = []
  let previousTableTokens: string[] | null = null

  for (const block of blocks) {
    if (
      block.type === 'paragraph' &&
      previousTableTokens &&
      sameTokens(paragraphTextTokens(block), previousTableTokens)
    ) {
      continue
    }

    result.push(block)
    previousTableTokens = block.type === 'table' ? tableTextTokens(block) : null
  }

  return result
}

const REPEATED_LEGACY_LINE_MIN_LENGTH = 12

function isPlainTextParagraph(
  paragraph: mdast.Paragraph,
): paragraph is mdast.Paragraph & { children: mdast.Text[] } {
  return paragraph.children.every(child => child.type === 'text')
}

function hasMeaningfulLongLine(lines: string[]): boolean {
  return lines.some(line => line.length >= REPEATED_LEGACY_LINE_MIN_LENGTH)
}

function isSameLineRun(
  lines: string[],
  left: number,
  right: number,
  size: number,
) {
  for (let i = 0; i < size; i++) {
    if (lines[left + i] !== lines[right + i]) return false
  }

  return true
}

function collapseAdjacentRepeatedLineRuns(lines: string[]): string[] {
  const result: string[] = []

  for (let i = 0; i < lines.length; i++) {
    const maxRunSize = Math.floor((lines.length - i) / 2)
    let duplicateRunSize = 0

    for (let size = maxRunSize; size >= 1; size--) {
      const run = lines.slice(i, i + size)
      if (!hasMeaningfulLongLine(run)) continue
      if (isSameLineRun(lines, i, i + size, size)) {
        duplicateRunSize = size
        break
      }
    }

    if (duplicateRunSize > 0) {
      result.push(...lines.slice(i, i + duplicateRunSize))
      i += duplicateRunSize * 2 - 1
      continue
    }

    result.push(lines[i])
  }

  return result
}

function removeRepeatedLegacyParagraphLines(
  blocks: mdast.BlockContent[],
): mdast.BlockContent[] {
  return blocks.map(block => {
    if (block.type !== 'paragraph' || !isPlainTextParagraph(block)) {
      return block
    }

    const text = block.children.map(child => child.value).join('')
    const lines = text
      .split('\n')
      .map(line => cleanText(line))
      .filter(Boolean)
    const collapsedLines = collapseAdjacentRepeatedLineRuns(lines)

    if (collapsedLines.length === lines.length) return block

    return {
      ...block,
      children: [{ type: 'text', value: collapsedLines.join('\n') }],
    }
  })
}

function normalizeLegacyBlocks(
  blocks: mdast.BlockContent[],
): mdast.BlockContent[] {
  return removeRepeatedLegacyParagraphLines(
    removeDuplicateTableParagraphs(blocks),
  )
}

/**
 * Get the actual Etherpad editor body element (where lines are direct children).
 * This always returns the innermost editor container to avoid parsing sidebars/navigation.
 */
function getEditorBody(container: HTMLElement): HTMLElement {
  if (
    container.classList.contains('innerdocbody') ||
    container.id === 'innerdocbody' ||
    container.dataset['cdcLegacySnapshot'] === 'true'
  ) {
    console.log('[doc parser] Using provided editor body container')
    return container
  }

  // Always look for innerdocbody FIRST — this is the real Etherpad editor body
  // where lines are direct children, guaranteed to not include sidebars/menus.
  // IMPORTANT: innerdocbody is a CLASS (.innerdocbody) in new Etherpad docs, NOT an ID!
  let innerdocbody =
    container.querySelector<HTMLElement>('.innerdocbody') ??
    document.querySelector<HTMLElement>('.innerdocbody')
  if (innerdocbody && innerdocbody.offsetHeight > 100) {
    console.log(
      '[doc parser] Found .innerdocbody (class selector) — using as editor body',
    )
    return innerdocbody
  }
  innerdocbody =
    container.querySelector<HTMLElement>('#innerdocbody') ??
    (document.getElementById('innerdocbody'))
  if (innerdocbody && innerdocbody.offsetHeight > 100) {
    console.log(
      '[doc parser] Found #innerdocbody (id selector) — using as editor body',
    )
    return innerdocbody
  }

  // Look for other known editor body selectors within the container or document
  const editorBodySelectors = [
    '.ace_editor .ace_content',
    '.ace_editor',
    '.outerdocbody-inner',
    '.etherpad-client-container .innerdocbody',
    '.outerdocbody',
    '.suite-body',
    '.edit-area',
    '.doc-editor',
    '#doc-editor',
  ]

  for (const sel of editorBodySelectors) {
    const el =
      container.querySelector<HTMLElement>(sel) ??
      document.querySelector<HTMLElement>(sel)
    if (el && el.offsetHeight > 100) {
      console.log(`[doc parser] Found editor body via selector: ${sel}`)
      return el
    }
  }

  // Fallback to the provided container if nothing better found
  console.log('[doc parser] Using provided container as editor body')
  return container
}

/**
 * Check if an element is part of the sidebar/navigation/UI chrome.
 */
function isSidebarElement(el: HTMLElement): boolean {
  if (
    el.closest(
      'nav, header, footer, aside, .sidebar, .navigation, .menu, .toolbar, .comments, .outline, .toc, .catalog, [role="navigation"], [role="banner"], [role="complementary"]',
    )
  ) {
    return true
  }
  // Feishu doc outline/sidebar often has these classes
  if (
    el.classList.contains('catalog') ||
    el.classList.contains('outline') ||
    el.classList.contains('sidebar')
  ) {
    return true
  }
  return false
}

/**
 * Check if an element is a virtual scroll placeholder (Etherpad renders these
 * for off-screen content instead of real DOM). These should be completely ignored.
 */
function isVirtualScrollPlaceholder(el: HTMLElement): boolean {
  return (
    el.classList.contains('adit-virtual-scroll-placeholder') ||
    el.classList.contains('virtual-scroll-placeholder') ||
    el.hasAttribute('data-ignore-mutation') ||
    el.classList.contains('dom-passthrough') ||
    el.classList.contains('dom-pas')
  )
}

/**
 * Find all "line" elements in an Etherpad container.
 * In newer Etherpad versions, .ace-line may be nested inside wrapper divs, not just direct children.
 * Since we already confirmed editorBody is the real .innerdocbody (no sidebar/UI chrome),
 * we can safely query all .ace-line elements recursively in document order.
 */
function findAllLineElements(editorBody: HTMLElement): HTMLElement[] {
  if (editorBody.dataset['cdcLegacySnapshot'] === 'true') {
    const snapshotLines = (
      Array.from(editorBody.children) as HTMLElement[]
    ).filter(child => {
      if (isVirtualScrollPlaceholder(child)) return false
      if (child.hidden || child.getAttribute('aria-hidden') === 'true')
        return false
      if (isSidebarElement(child)) return false
      return !isLineEmpty(child)
    })
    if (snapshotLines.length > 0) {
      console.log(
        `[doc parser] Snapshot mode: using ${snapshotLines.length} collected line elements`,
      )
      return snapshotLines
    }
  }

  // First priority: Find all .ace-line elements, regardless of nesting.
  // This works for both old (direct child) and new (nested in wrapper) Etherpad structures.
  const aceLines = Array.from(
    editorBody.querySelectorAll<HTMLElement>('.ace-line'),
  ).filter(el => {
    // Skip virtual scroll placeholders
    if (isVirtualScrollPlaceholder(el)) return false
    // Skip hidden elements
    if (el.hidden || el.getAttribute('aria-hidden') === 'true') return false
    // Skip lines that are inside sidebar/navigation (should not happen if editorBody is correct)
    if (isSidebarElement(el)) return false
    return true
  })

  if (aceLines.length > 0) {
    console.log(
      `[doc parser] Found ${aceLines.length} .ace-line elements (recursive)`,
    )
    return aceLines
  }

  // Fallback: Get direct children of editorBody
  const directChildren = Array.from(editorBody.children) as HTMLElement[]
  const contentLines = directChildren.filter(child => {
    const tag = child.tagName.toLowerCase()
    if (tag === 'script' || tag === 'style' || tag === 'noscript') return false
    if (child.hidden || child.getAttribute('aria-hidden') === 'true')
      return false
    if (isVirtualScrollPlaceholder(child)) return false
    if (isSidebarElement(child)) return false
    return true
  })

  if (contentLines.length > 0) {
    console.log(
      `[doc parser] Fallback: Found ${contentLines.length} direct child lines`,
    )
    return contentLines
  }

  // Last resort
  console.warn('[doc parser] WARNING: No lines found, returning all child divs')
  return Array.from(editorBody.children).filter(
    child => child.tagName.toLowerCase() === 'div',
  ) as HTMLElement[]
}

/**
 * Internal recursive helper to collect blocks from any element (for nested content
 * like blockquotes, list items, table cells).
 */
function collectBlocksRecursive(
  el: HTMLElement,
  collector: ImageCollector,
  options: { highlight?: boolean } = {},
  result: mdast.BlockContent[] = [],
  phrasingBuffer: mdast.PhrasingContent[] = [],
): mdast.BlockContent[] {
  const tag = el.tagName.toLowerCase()

  // Skip non-content elements
  if (
    tag === 'script' ||
    tag === 'style' ||
    tag === 'noscript' ||
    tag === 'nav' ||
    tag === 'header' ||
    tag === 'footer' ||
    tag === 'button' ||
    tag === 'input' ||
    tag === 'textarea' ||
    tag === 'select' ||
    tag === 'form' ||
    el.getAttribute('aria-hidden') === 'true' ||
    el.classList.contains('toolbar') ||
    el.classList.contains('menu') ||
    el.classList.contains('sidebar') ||
    el.classList.contains('comment')
  ) {
    return result
  }

  // Try to parse as known block
  const node = parseBlockElement(el, collector, options)
  if (node) {
    flushPhrasingBuffer(phrasingBuffer, result)
    result.push(node)
    return result
  }

  // Empty element = paragraph break
  if (isLineEmpty(el)) {
    flushPhrasingBuffer(phrasingBuffer, result)
    return result
  }

  // Check for standard block children
  const hasBlockChildren = el.querySelector(
    'h1,h2,h3,h4,h5,h6,p,pre,ul,ol,li,table,blockquote,figure,hr',
  )

  if (!hasBlockChildren) {
    // Leaf element: extract phrasing content
    const phrasing = mdastToPhrasing(el, collector)
    if (phrasing.length > 0) {
      phrasingBuffer.push(...phrasing)
    }
    // If this is a container element (li, td, blockquote), flush
    if (tag === 'li' || tag === 'td' || tag === 'th' || tag === 'blockquote') {
      flushPhrasingBuffer(phrasingBuffer, result)
    }
    return result
  }

  // Recurse into children
  for (const child of Array.from(el.childNodes)) {
    if (child.nodeType === Node.TEXT_NODE) {
      const text = cleanText(child.textContent || '')
      if (text) {
        phrasingBuffer.push({ type: 'text', value: text })
      }
    } else if (child.nodeType === Node.ELEMENT_NODE) {
      collectBlocksRecursive(
        child as HTMLElement,
        collector,
        options,
        result,
        phrasingBuffer,
      )
    }
  }

  // Flush at container boundaries
  if (
    tag === 'li' ||
    tag === 'td' ||
    tag === 'th' ||
    tag === 'blockquote' ||
    tag === 'article' ||
    tag === 'section' ||
    tag === 'main'
  ) {
    flushPhrasingBuffer(phrasingBuffer, result)
  }

  return result
}

/**
 * Collect all blocks — simplified, line-first approach for Etherpad.
 * This is the top-level entry point.
 */
function collectBlocks(
  container: HTMLElement,
  collector: ImageCollector,
  options: { highlight?: boolean } = {},
): mdast.BlockContent[] {
  const result: mdast.BlockContent[] = []
  const phrasingBuffer: mdast.PhrasingContent[] = []

  // ALWAYS get the real editor body (#innerdocbody) first — this avoids parsing
  // sidebars, navigation, comments, and other UI chrome that exists in outer containers.
  const editorBody = getEditorBody(container)
  console.log(
    `[doc parser] Editor body: <${editorBody.tagName.toLowerCase()}> id="${editorBody.id}" class="${editorBody.className?.substring(0, 80)}"`,
  )
  console.log(
    `[doc parser] Editor body childElementCount: ${editorBody.childElementCount}, textLen: ${editorBody.textContent?.trim().length ?? 0}`,
  )

  // Check if editor body contains standard HTML blocks (for very old Doc 1.0)
  // Use :scope to ONLY check direct children or elements INSIDE editorBody, not outer sidebars
  const hasStandardBlocks = editorBody.querySelector(
    ':scope > h1, :scope > h2, :scope > h3, :scope > h4, :scope > h5, :scope > h6, :scope > p, :scope > pre, :scope > ul, :scope > ol, :scope > table, :scope > blockquote, :scope > figure',
  )

  if (!hasStandardBlocks) {
    // Etherpad-style document: process line by line (direct children of editorBody)
    console.log(
      '[doc parser] Etherpad mode: processing line by line from editorBody direct children',
    )
    const lines = findAllLineElements(editorBody)
    console.log(`[doc parser] Total lines to process: ${lines.length}`)

    for (let i = 0; i < lines.length; i++) {
      const line = lines[i]

      // Skip UI chrome and sidebar elements
      if (isSidebarElement(line)) {
        continue
      }
      if (
        line.classList.contains('toolbar') ||
        line.classList.contains('menu') ||
        line.classList.contains('comment') ||
        line.hidden ||
        line.getAttribute('aria-hidden') === 'true'
      ) {
        continue
      }

      // First try to parse this line as a known block type (heading, list, code, table, etc.)
      const blockNode = parseBlockElement(line, collector, options)
      if (blockNode) {
        flushPhrasingBuffer(phrasingBuffer, result)
        result.push(blockNode)
        continue
      }

      // Check if this line is empty → paragraph separator
      if (isLineEmpty(line)) {
        flushPhrasingBuffer(phrasingBuffer, result)
        continue
      }

      // Regular line: extract phrasing content (text, formatting, images, links)
      const inlineContent = mdastToPhrasing(line, collector)
      if (inlineContent.length > 0) {
        phrasingBuffer.push(...inlineContent)
        // Add a single space between consecutive lines (soft break in Etherpad)
        phrasingBuffer.push({ type: 'text', value: '\n' })
      }
    }

    flushPhrasingBuffer(phrasingBuffer, result)
  } else {
    // Classic Doc 1.0 editor with real HTML block elements
    console.log('[doc parser] Classic Doc 1.0 mode: recursive block collection')
    collectBlocksRecursive(
      editorBody,
      collector,
      options,
      result,
      phrasingBuffer,
    )
    flushPhrasingBuffer(phrasingBuffer, result)
  }

  // Last resort: if no blocks found, extract all text from editorBody
  if (result.length === 0) {
    console.log(
      '[doc parser] No blocks found via line/recursive parsing, extracting all text from editorBody',
    )
    const allPhrasing = mdastToPhrasing(editorBody, collector)
    if (allPhrasing.length > 0) {
      result.push({
        type: 'paragraph',
        children: mergePhrasingContents(
          allPhrasing,
        ) as mdast.paragraph['children'],
      })
    }
  }

  return normalizeLegacyBlocks(result)
}

/**
 * Try to convert a single DOM element to an mdast block node.
 * Returns null if the element is not a recognized block type.
 */
function parseBlockElement(
  el: HTMLElement,
  collector: ImageCollector,
  options: { highlight?: boolean } = {},
): mdast.BlockContent | null {
  const tag = el.tagName.toLowerCase()

  // --- Headings ---
  if (/^h[1-9]$/.test(tag)) {
    const depth = parseInt(tag.charAt(1), 10) as mdast.Heading['depth']
    if (depth > 0 && depth <= 9) {
      const children = mdastToPhrasing(el, collector)
      if (children.length > 0) {
        return {
          type: 'heading',
          depth,
          children: mergePhrasingContents(
            children,
          ) as mdast.heading['children'],
        }
      }
    }
    return null
  }

  // --- Horizontal rule ---
  if (tag === 'hr') {
    return { type: 'thematicBreak' }
  }

  // --- Code blocks ---
  if (tag === 'pre') {
    const codeEl = el.querySelector('code')
    const lang =
      codeEl
        ?.getAttribute('class')
        ?.replace('language-', '')
        .replace('lang-', '') ?? ''
    const value = cleanText(el.textContent ?? '')
    return {
      type: 'code',
      lang,
      value,
    }
  }

  // --- Blockquote ---
  if (
    tag === 'blockquote' ||
    el.classList.contains('quote') ||
    el.classList.contains('quote-container')
  ) {
    const innerBlocks: mdast.BlockContent[] = []
    for (const child of el.children) {
      const node = parseBlockElement(child as HTMLElement, collector, options)
      if (node) {
        innerBlocks.push(node)
      } else {
        collectBlocksRecursive(
          child as HTMLElement,
          collector,
          options,
          innerBlocks,
        )
      }
    }
    if (innerBlocks.length > 0) {
      return { type: 'blockquote', children: innerBlocks }
    }
    return null
  }

  // --- Unordered list ---
  if (tag === 'ul') {
    const items: mdast.ListItem[] = []
    for (const child of el.children) {
      const li = child as HTMLElement
      if (li.tagName.toLowerCase() === 'li') {
        const parsed = parseListItem(li, 'bullet', collector, options)
        if (parsed) items.push(parsed)
      }
    }
    if (items.length > 0) {
      return {
        type: 'list',
        ordered: false,
        children: items,
      }
    }
    return null
  }

  // --- Ordered list ---
  if (tag === 'ol') {
    const items: mdast.ListItem[] = []
    let start = parseInt(el.getAttribute('start') ?? '1', 10)
    if (isNaN(start)) start = 1

    for (const child of el.children) {
      const li = child as HTMLElement
      if (li.tagName.toLowerCase() === 'li') {
        const parsed = parseListItem(li, 'ordered', collector, options, start)
        if (parsed) items.push(parsed)
        start++
      }
    }
    if (items.length > 0) {
      return {
        type: 'list',
        ordered: true,
        start,
        children: items,
      }
    }
    return null
  }

  // --- Tables ---
  if (tag === 'table') {
    return parseTable(el, collector)
  }

  const compactTable = parseLegacyCompactTable(el)
  if (compactTable) {
    return compactTable
  }

  const visualBlock = parseLegacyVisualBlock(el, collector)
  if (visualBlock) {
    return visualBlock
  }

  // --- Standalone images (not inside p/a) ---
  if (
    tag === 'img' &&
    el.closest('p, a, h1, h2, h3, h4, h5, h6, li') === null
  ) {
    const imgEl = el as HTMLImageElement
    const src = getImageSource(imgEl)
    const alt = el.getAttribute('alt') || ''
    const title = el.getAttribute('title') || ''
    const fallbackName = title || alt || 'image'
    const image = imageNodeFromSource(src, alt, fallbackName)
    if (!image) return null

    collector.push(image)

    return {
      type: 'paragraph',
      children: [image],
    }
  }

  // --- Callout / special container ---
  if (
    el.classList.contains('callout') ||
    el.getAttribute('data-node-type') === 'callout'
  ) {
    const innerBlocks: mdast.BlockContent[] = []
    for (const child of el.children) {
      const node = parseBlockElement(child as HTMLElement, collector, options)
      if (node) {
        innerBlocks.push(node)
      } else {
        collectBlocksRecursive(
          child as HTMLElement,
          collector,
          options,
          innerBlocks,
        )
      }
    }
    if (innerBlocks.length > 0) {
      return { type: 'blockquote', children: innerBlocks }
    }
    return null
  }

  // --- Iframe / embedded content ---
  if (tag === 'iframe') {
    const src = el.getAttribute('src') || ''
    // Skip empty/blank iframes
    if (!src || src.includes('about:blank') || src === '//about:blank') {
      return null
    }
    return {
      type: 'html',
      value: `<iframe src="${escapeHtml(src)}"></iframe>`,
    }
  }

  // --- Figure elements (common in etherpad docs) ---
  if (tag === 'figure') {
    const img = el.querySelector('img')
    if (img) {
      const src = getImageSource(img)
      const alt = img.getAttribute('alt') || ''
      const caption = el.querySelector('figcaption')
      const captionText = caption?.textContent?.trim() || ''
      const fallbackName = captionText || alt || 'image'
      const image = imageNodeFromSource(src, captionText || alt, fallbackName)
      if (!image) return null
      collector.push(image)
      return {
        type: 'paragraph',
        children: [image],
      }
    }
    // Try to parse other figure content
    const innerBlocks: mdast.BlockContent[] = []
    collectBlocksRecursive(el, collector, options, innerBlocks)
    return innerBlocks[0] ?? null
  }

  // --- Special etherpad/Feishu block classes ---
  // Check for heading classes (etherpad uses class-based headings, not <h1>-<h6>)
  const className = el.className?.toString() || ''
  if (
    /heading[1-9]|h[1-9]|title-level/i.test(className) ||
    el.getAttribute('data-heading-level')
  ) {
    const levelMatch = /h([1-9])|heading([1-9])/i.exec(className)
    const dataLevel = el.getAttribute('data-heading-level')
    const level = parseInt(
      levelMatch?.[1] || levelMatch?.[2] || dataLevel || '1',
      10,
    )
    const depth = Math.min(Math.max(level, 1), 6) as 1 | 2 | 3 | 4 | 5 | 6
    const children = mdastToPhrasing(el, collector)
    if (children.length > 0) {
      return {
        type: 'heading',
        depth,
        children: mergePhrasingContents(children) as mdast.heading['children'],
      }
    }
  }

  // Check for list items / list containers in etherpad
  if (
    /list-bullet|list-numbered|bullet-list|numbered-list|ace-list/i.test(
      className,
    )
  ) {
    // Treat etherpad lists as unordered for now; we'll parse their items
    const items: mdast.ListItem[] = []
    const childElements = el.querySelectorAll(
      'li, [class*="list-item"], [class*="ace-list-item"]',
    )
    if (childElements.length === 0 && tag === 'div') {
      // No explicit li elements; this div itself might be a list item
      const phrasingChildren = mdastToPhrasing(el, collector)
      if (phrasingChildren.length > 0) {
        items.push({
          type: 'listItem',
          spread: false,
          children: [
            {
              type: 'paragraph',
              children: mergePhrasingContents(
                phrasingChildren,
              ) as mdast.paragraph['children'],
            },
          ],
        })
      }
    } else {
      childElements.forEach(li => {
        const phrasingChildren = mdastToPhrasing(li as HTMLElement, collector)
        if (phrasingChildren.length > 0) {
          items.push({
            type: 'listItem',
            spread: false,
            children: [
              {
                type: 'paragraph',
                children: mergePhrasingContents(
                  phrasingChildren,
                ) as mdast.paragraph['children'],
              },
            ],
          })
        }
      })
    }
    if (items.length > 0) {
      return {
        type: 'list',
        ordered: /numbered|ordered/i.test(className),
        children: items,
      }
    }
  }

  // --- Code blocks ---
  if (
    tag === 'pre' ||
    el.classList.contains('code') ||
    /code-block|preformatted/i.test(className)
  ) {
    const code = cleanText(el.textContent || '')
    if (code) {
      return {
        type: 'code',
        value: code,
      }
    }
  }

  // --- Blockquote / callout ---
  if (tag === 'blockquote' || /quote|callout|blockquote/i.test(className)) {
    const innerBlocks: mdast.BlockContent[] = []
    collectBlocksRecursive(el, collector, options, innerBlocks)
    if (innerBlocks.length > 0) {
      return { type: 'blockquote', children: innerBlocks }
    }
    return null
  }

  // --- HR / separator ---
  if (
    tag === 'hr' ||
    el.classList.contains('hr') ||
    /separator|divider/i.test(className)
  ) {
    return { type: 'thematicBreak' }
  }

  // <br> at block level means a paragraph break; don't create a paragraph for it
  if (tag === 'br') {
    return null
  }

  // --- Standard semantic elements as paragraphs (only <p> tags and contentEditable roots) ---
  // For divs, we don't automatically treat them as paragraphs because etherpad
  // wraps every single line in a div. Instead, we collect phrasing content and
  // merge adjacent text/phrasing nodes into paragraphs in a post-processing step.
  if (tag === 'p') {
    const children = mdastToPhrasing(el, collector)
    if (children.length > 0) {
      return {
        type: 'paragraph',
        children: mergePhrasingContents(
          children,
        ) as mdast.paragraph['children'],
      }
    }
  }

  return null
}

/**
 * Parse a <li> element into a mdast.ListItem.
 */
function parseListItem(
  el: HTMLElement,
  listType: 'bullet' | 'ordered' | 'todo',
  collector: ImageCollector,
  options: { highlight?: boolean } = {},
  start?: number,
): mdast.ListItem | null {
  const children: mdast.Nodes[] = []
  let firstParagraph: mdast.Paragraph | null = null

  // Collect all blocks within this list item
  const innerBlocks: mdast.BlockContent[] = []
  collectBlocksRecursive(el, collector, options, innerBlocks)

  for (const node of innerBlocks) {
    if (node.type === 'paragraph' && !firstParagraph) {
      firstParagraph = node
      // Add the paragraph's children to the list item
      children.push(...node.children)
      continue
    }

    children.push(node as mdast.Nodes)
  }

  // Check for todo checkbox
  let checked: boolean | undefined
  if (listType === 'bullet' || !listType) {
    const isChecked =
      el.getAttribute('data-done') === 'true' ||
      el.classList.contains('checked') ||
      el.querySelector('input[type="checkbox"]:checked') !== null
    const isUnchecked =
      el.getAttribute('data-done') === 'false' ||
      el.querySelector('input[type="checkbox"]:not(:checked)') !== null
    if (isChecked) {
      listType = 'todo'
      checked = true
    } else if (isUnchecked) {
      listType = 'todo'
      checked = false
    }
  }

  if (listType === 'todo') {
    return {
      type: 'listItem',
      checked: checked ?? false,
      children,
    }
  }

  if (listType === 'ordered') {
    return {
      type: 'listItem',
      ...(start !== undefined ? { data: { seq: start } } : null),
      children,
    }
  }

  return {
    type: 'listItem',
    children,
  }
}

const LEGACY_TABLE_SEPARATOR_RE = /[\u200B\u200C\u200D\uFEFF]+/g

function inferLegacyTableColumnCount(cells: string[]): number | null {
  if (cells.length < 2) return null
  if (cells.length === 2) return 2

  const divisors = [6, 5, 4, 3, 2].filter(
    columnCount => cells.length % columnCount === 0,
  )
  if (divisors.length === 0) return null

  if (
    cells.length % 3 === 0 &&
    cells.some(cell => /^\d{4}[./-]\d{1,2}[./-]\d{1,2}$/.test(cell))
  ) {
    return 3
  }

  if (cells.length <= 4 && cells.length % 2 === 0) return 2

  return divisors.find(columnCount => columnCount <= 4) ?? divisors[0]
}

function parseLegacyCompactTable(el: HTMLElement): mdast.Table | null {
  const rawText = el.textContent ?? ''
  if (!LEGACY_TABLE_SEPARATOR_RE.test(rawText)) return null
  LEGACY_TABLE_SEPARATOR_RE.lastIndex = 0

  const cells = rawText
    .split(LEGACY_TABLE_SEPARATOR_RE)
    .map(cell => cleanText(cell))
    .filter(Boolean)
  const columnCount = inferLegacyTableColumnCount(cells)
  if (!columnCount || cells.length < columnCount) return null

  const rows: mdast.TableRow[] = []
  for (let i = 0; i < cells.length; i += columnCount) {
    const rowCells = cells.slice(i, i + columnCount)
    if (rowCells.length !== columnCount) return null
    rows.push({
      type: 'tableRow',
      children: rowCells.map(cell => ({
        type: 'tableCell',
        children: cell ? [{ type: 'text', value: cell }] : [],
      })),
    })
  }

  return rows.length > 0 ? { type: 'table', children: rows } : null
}

/**
 * Parse an HTML <table> element into an mdast.Table node.
 */
function parseTable(
  tableEl: HTMLElement,
  collector: ImageCollector,
): mdast.Table | null {
  const rows: mdast.TableRow[] = []
  const children = tableEl.children

  for (let i = 0; i < children.length; i++) {
    const rowEl = children[i] as HTMLElement
    const rowTag = rowEl.tagName.toLowerCase()

    if (rowTag !== 'tr') continue

    const cells: HTMLElement[] = []
    for (let j = 0; j < rowEl.children.length; j++) {
      cells.push(rowEl.children[j] as HTMLElement)
    }

    // Handle colspan
    const parsedCells: mdast.TableCell[] = []
    for (const cellEl of cells) {
      const colspan = parseInt(cellEl.getAttribute('colspan') ?? '1', 10)
      const rowspan = parseInt(cellEl.getAttribute('rowspan') ?? '1', 10)

      for (let c = 0; c < colspan; c++) {
        // Parse cell content
        const cellChildren: mdast.Nodes[] = []
        let firstPara: mdast.Paragraph | null = null

        const processCellChildren = (parent: HTMLElement) => {
          const cellBlocks: mdast.BlockContent[] = []
          collectBlocksRecursive(parent, collector, {}, cellBlocks)

          for (const node of cellBlocks) {
            if (node.type === 'paragraph' && !firstPara) {
              firstPara = node
              cellChildren.push(...node.children)
            } else {
              cellChildren.push(node as mdast.Nodes)
            }
          }
          // Also handle direct text nodes
          for (const textNode of parent.childNodes) {
            if (textNode.nodeType === Node.TEXT_NODE) {
              const text = cleanText(textNode.textContent ?? '')
              if (text) {
                cellChildren.push({ type: 'text', value: text })
              }
            }
          }
        }

        processCellChildren(cellEl)

        const cell: mdast.TableCell = {
          type: 'tableCell',
          children: mergePhrasingContents(
            cellChildren.filter(isPhrasingContent),
          ),
          ...(rowspan > 1 ? { data: { rowSpan: rowspan } } : {}),
        }

        parsedCells.push(cell)
      }
    }

    if (parsedCells.length > 0) {
      rows.push({
        type: 'tableRow',
        children: parsedCells,
      })
    }
  }

  if (rows.length === 0) return null

  return {
    type: 'table',
    children: rows,
  }
}

// ---------------------------------------------------------------------------
// Doc class — mirrors Docx class interface for old-version documents
// ---------------------------------------------------------------------------

export class Doc {
  private _container: HTMLElement | null = null
  private _blocks: mdast.Nodes[] = []
  private _collector = new ArrayImageCollector()

  /**
   * Initialize the parser by locating the editor container and collecting blocks.
   * Returns true if the container was found and has content.
   */
  init(options: DocParseOptions = {}): boolean {
    console.log('[doc parser] init called, location.href:', location.href)
    console.log(
      '[doc parser] window.editor:',
      !!window.editor,
      'window.PageMain:',
      !!window.PageMain,
      'isDoc:',
      isDoc(),
    )
    this._collector = new ArrayImageCollector()
    this._container = options.container ?? findEditorContainer()
    if (!this._container) {
      console.warn('[doc parser] Could not find editor container')
      this._blocks = []
      return false
    }

    console.log(
      '[doc parser] Found container:',
      this._container.tagName,
      'id=',
      this._container.id,
      'class=',
      this._container.className?.substring(0, 100),
    )
    console.log(
      '[doc parser] Container childElementCount:',
      this._container.childElementCount,
      'textContent length:',
      this._container.textContent?.trim().length ?? 0,
    )

    // Dump first few children for debugging
    const firstKids = Array.from(this._container.children).slice(0, 3)
    firstKids.forEach((k, i) => {
      const kid = k as HTMLElement
      console.log(
        `[doc parser]   child[${i}]: <${kid.tagName.toLowerCase()}> id="${kid.id}" class="${kid.className?.toString().substring(0, 60)}" text="${kid.textContent?.trim().substring(0, 80)}"`,
      )
    })

    // Collect all block elements from the container using simplified line-first approach
    const blocks: mdast.BlockContent[] = collectBlocks(
      this._container,
      this._collector,
      options,
    )

    console.log(
      '[doc parser] Collected blocks:',
      blocks.length,
      'images:',
      this._collector.images.length,
    )

    // Show preview of first few blocks
    if (blocks.length > 0) {
      const previewCount = Math.min(blocks.length, 5)
      console.log(`[doc parser] First ${previewCount} blocks preview:`)
      for (let i = 0; i < previewCount; i++) {
        const block = blocks[i]
        let preview = ''
        if (block.type === 'paragraph') {
          preview = block.children
            .map(c => {
              if (c.type === 'text') return c.value
              if (c.type === 'image') return `[img:${c.alt}]`
              if (c.type === 'link')
                return `[link:${(c.children?.[0] as any)?.value || c.url}]`
              if (c.type === 'strong')
                return `**${(c.children?.[0] as any)?.value || ''}**`
              if (c.type === 'emphasis')
                return `*${(c.children?.[0] as any)?.value || ''}*`
              return `[${c.type}]`
            })
            .join('')
            .substring(0, 100)
        } else if (block.type === 'heading') {
          const text = block.children.map((c: any) => c.value || '').join('')
          preview = `${'#'.repeat(block.depth)} ${text.substring(0, 80)}`
        } else if (block.type === 'list') {
          preview = `[${block.ordered ? 'ordered' : 'bullet'} list with ${block.children.length} items]`
        } else if (block.type === 'code') {
          preview = `\`\`\`${block.lang || ''}\n${block.value.substring(0, 60)}...\n\`\`\``
        } else {
          preview = `[${block.type}]`
        }
        console.log(`[doc parser]   block[${i}] (${block.type}): ${preview}`)
      }
    } else {
      console.warn(
        '[doc parser] WARNING: No blocks collected! Dumping container innerHTML first 500 chars:',
      )
      console.warn(this._container.innerHTML.substring(0, 500))
    }

    this._blocks = blocks
    return blocks.length > 0
  }

  /**
   * Check whether this is an old-version document page.
   * Matches the semantics of `docx.isDoc` — returns true when window.editor exists
   * but window.PageMain does not.
   */
  get isDoc(): boolean {
    return !isDocx() && isDoc()
  }

  /**
   * Check whether the document container was found and has content.
   */
  get isReady(): boolean {
    return this._blocks.length > 0
  }

  /**
   * The scrollable container element, if found.
   */
  get container(): HTMLElement | null {
    return this._container
  }

  /**
   * Number of collected blocks (for debugging).
   */
  get blocksCount(): number {
    return this._blocks.length
  }

  /**
   * Number of collected images (for debugging).
   */
  get imagesCount(): number {
    return this._collector.images.length
  }

  /**
   * The page title extracted from the document.
   */
  get pageTitle(): string {
    return getPageTitle()
  }

  /**
   * Scroll the document container to the specified position.
   * Provides the same interface as Docx.scrollTo() for downstream compatibility.
   */
  scrollTo(options: ScrollToOptions): void {
    if (this._container) {
      const {
        left,
        top = this._container.scrollHeight,
        behavior = 'smooth',
      } = options

      this._container.scrollTo({
        left,
        top: Math.min(top, this._container.scrollHeight),
        behavior,
      })
    }
  }

  /**
   * Transform the document into an mdast Root.
   *
   * Returns the same shape as `Docx.intoMarkdownAST()` so downstream scripts
   * (copy / view / download) can use identical post-processing pipelines.
   */
  intoMarkdownAST(options: DocParseOptions = {}): DocTransformResult {
    // If not yet initialized, do it now
    if (this._blocks.length === 0 && this.isDoc) {
      this.init(options)
    }

    const root: mdast.Root = {
      type: 'root',
      children: this._blocks.filter(isRootContent),
    }

    return {
      root,
      images: this._collector.images,
      files: [],
      tableWithParents: [],
      mentionUsers: [],
    }
  }
}

/**
 * Singleton instance for old-version document parsing.
 * Mirrors the `docx` singleton pattern in docx.ts.
 */
export const doc: Doc = new Doc()
