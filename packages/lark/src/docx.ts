import type * as mdast from 'mdast'
import { chunk } from 'es-toolkit/array'
import {
  toBlob as svgToBlob,
  imageDataToBlob,
  compare,
  isDefined,
  waitForFunction,
  waitFor,
  OneHundred,
  Second,
  checkCanvasDimensions,
} from '@dolphin/common'
import { toMarkdown, type Options } from 'mdast-util-to-markdown'
import { gfmStrikethroughToMarkdown } from 'mdast-util-gfm-strikethrough'
import { gfmTaskListItemToMarkdown } from 'mdast-util-gfm-task-list-item'
import { gfmTableToMarkdown } from 'mdast-util-gfm-table'
import { mathToMarkdown, type InlineMath } from 'mdast-util-math'
import { PageMain, User, isDoc, isDocx } from './env'
import {
  isBlockquoteContent,
  isParent,
  isPhrasingContent,
  isRootContent,
  isTableCell,
  isListItemContent,
} from './utils/mdast'
import { resolveFileDownloadUrl } from './file'
import { isString } from 'es-toolkit/compat'
import { escape } from 'es-toolkit/compat'
import { toCamelCaseKeys } from 'es-toolkit/object'

declare module 'mdast' {
  interface ImageData {
    name?: string
    token?: string
    fetchSources?: () => Promise<ImageSources | null>
    fetchBlob?: () => Promise<Blob | null>
  }

  interface ListItemData {
    seq?: number | 'auto'
  }

  interface LinkData {
    name?: string
    fetchFile?: (init?: RequestInit) => Promise<Response>
  }

  interface TableData {
    type?: BlockType.TABLE | BlockType.GRID
    colWidths?: number[]
    invalid?: boolean
    cellSet?: Record<string, CellData>
  }

  interface TableCellData {
    width?: number
    invalidChildren?: mdast.Nodes[]
    rowSpan?: number
    colSpan?: number
  }

  interface InlineCodeData {
    mentionUserId?: string
    parentBlockRecordId?: string
  }

  interface HtmlData {
    fetchHtml?: () => Promise<string | null>
  }
}

/**
 * @see https://open.feishu.cn/document/client-docs/docs-add-on/06-data-structure/BlockType
 */
export enum BlockType {
  PAGE = 'page',
  BITABLE = 'bitable',
  BASE_REFER = 'base_refer',
  CALLOUT = 'callout',
  CHAT_CARD = 'chat_card',
  CODE = 'code',
  DIAGRAM = 'diagram',
  DIVIDER = 'divider',
  FILE = 'file',
  GRID = 'grid',
  GRID_COLUMN = 'grid_column',
  HEADING1 = 'heading1',
  HEADING2 = 'heading2',
  HEADING3 = 'heading3',
  HEADING4 = 'heading4',
  HEADING5 = 'heading5',
  HEADING6 = 'heading6',
  HEADING7 = 'heading7',
  HEADING8 = 'heading8',
  HEADING9 = 'heading9',
  IFRAME = 'iframe',
  IMAGE = 'image',
  ISV = 'isv',
  MINDNOTE = 'mindnote',
  BULLET = 'bullet',
  ORDERED = 'ordered',
  TODO = 'todo',
  QUOTE = 'quote',
  QUOTE_CONTAINER = 'quote_container',
  SHEET = 'sheet',
  TABLE = 'table',
  CELL = 'table_cell',
  TEXT = 'text',
  VIEW = 'view',
  SYNCED_SOURCE = 'synced_source',
  SYNCED_REFERENCE = 'synced_reference',
  WHITEBOARD = 'whiteboard',
  FALLBACK = 'fallback',
}

interface Attributes {
  fixEnter?: string

  italic?: string
  bold?: string
  strikethrough?: string
  underline?: string

  inlineCode?: string
  equation?: string
  textHighlight?: string
  textHighlightBackground?: string
  'inline-component'?: string

  link?: string
  mentionUserId?: string

  [attrName: string]: unknown
}

interface Operation {
  attributes?: Attributes
  insert: string
}

interface BlockZoneState {
  allText: string
  content: {
    ops: Operation[]
  }
}

interface BlockSnapshot {
  type: BlockType | 'pending'
}

interface Block<T extends Blocks = Blocks> {
  id: number
  type: BlockType
  zoneState?: BlockZoneState
  record?: { id: string }
  snapshot: BlockSnapshot
  children: T[]
}

export interface PageBlock extends Block {
  type: BlockType.PAGE
}

interface DividerBlock extends Block {
  type: BlockType.DIVIDER
}

interface HeadingBlock extends Block<TextBlock> {
  type:
    | BlockType.HEADING1
    | BlockType.HEADING2
    | BlockType.HEADING3
    | BlockType.HEADING4
    | BlockType.HEADING5
    | BlockType.HEADING6
    | BlockType.HEADING7
    | BlockType.HEADING8
    | BlockType.HEADING9
  depth: 1 | 2 | 3 | 4 | 5 | 6 | 7 | 8 | 9
  snapshot: {
    type:
      | BlockType.HEADING1
      | BlockType.HEADING2
      | BlockType.HEADING3
      | BlockType.HEADING4
      | BlockType.HEADING5
      | BlockType.HEADING6
      | BlockType.HEADING7
      | BlockType.HEADING8
      | BlockType.HEADING9
    /**
     * sequence value
     */
    seq?: string
    seq_level?: string
  }
}

interface CodeBlock extends Block<TextBlock> {
  type: BlockType.CODE
  language: string
}

interface QuoteContainerBlock extends Block {
  type: BlockType.QUOTE_CONTAINER
}

interface BulletBlock extends Block {
  type: BlockType.BULLET
}

interface OrderedBlock extends Block<TextBlock> {
  type: BlockType.ORDERED
  snapshot: {
    type: BlockType.ORDERED
    seq: string
  }
}

interface TodoBlock extends Block {
  type: BlockType.TODO
  snapshot: {
    type: BlockType.TODO
    done?: boolean
  }
}

interface TextBlock extends Block {
  type: BlockType.TEXT
}

interface Caption {
  text: {
    initialAttributedTexts: {
      text: { 0: string } | null
    }
  }
}

interface ImageBlockData {
  token: string
  width: number
  height: number
  mimeType: string
  name: string
  caption?: Caption
}

interface ImageSources {
  originSrc: string
  src: string
}

interface ImageBlock extends Block {
  type: BlockType.IMAGE
  snapshot: {
    type: BlockType.IMAGE
    image: ImageBlockData
  }
  imageManager: {
    fetch: (
      image: {
        token: string
        isHD: boolean
        fuzzy: boolean
        width?: number
        height?: number
      },
      options: unknown,
      callback: (sources: ImageSources) => void,
    ) => Promise<void>
  }
}

interface MergeInfo {
  row_span: number
  col_span: number
}

interface ColumnData {
  column_width: number
}

interface CellData {
  merge_info: MergeInfo
}

interface TableBlock extends Block<TableCellBlock> {
  type: BlockType.TABLE
  snapshot: {
    type: BlockType.TABLE
    rows_id: string[]
    columns_id: string[]
    column_set: Record<string, ColumnData>
    cell_set: Record<string, CellData>
  }
}

interface TableCellBlock extends Block {
  type: BlockType.CELL
  cellId: string
}

interface Grid extends Block<GridColumn> {
  type: BlockType.GRID
}

interface GridColumn extends Block {
  type: BlockType.GRID_COLUMN
  snapshot: {
    type: BlockType.GRID_COLUMN
    width_ratio?: number
  }
}

interface Callout extends Block {
  type: BlockType.CALLOUT
}

interface SyncedSource extends Block {
  type: BlockType.SYNCED_SOURCE
}

interface SyncedReferenceInnerBlockManager {
  rootBlockModel?: PageBlock
}

interface SyncedReference extends Block {
  type: BlockType.SYNCED_REFERENCE
  isAllDataReady: boolean
  innerBlockManager?: SyncedReferenceInnerBlockManager
}

interface ImageDataWrapper {
  data: ImageData
  release: () => void
}

interface RatioApp {
  ratioAppProxy?: {
    getOriginImageDataByNodeId: (
      i: number,
      o: [''],
      r: false,
      n: number,
    ) => Promise<ImageDataWrapper | null>
  }
  app?: {
    application: {
      nodeManager: {
        getNodesBounds: () => {
          minX: number
          maxX: number
          minY: number
          maxY: number
        }
      }
    }
    renderManager: {
      getImageOffscreenCanvas: (
        bounds: {
          minX: number
          maxX: number
          minY: number
          maxY: number
        },
        r: number,
        bgColor: string,
      ) => HTMLCanvasElement | null
    }
  }
}

interface WhiteboardBlock {
  isolateEnv: {
    hasRatioApp: () => boolean
    getRatioApp: () => RatioApp
  }
  abilityKit: {
    getRatioApp: () => RatioApp
  }
}

interface Whiteboard extends Block {
  type: BlockType.WHITEBOARD
  whiteboardBlock?: WhiteboardBlock
  snapshot: {
    type: BlockType.WHITEBOARD
    caption?: Caption
  }
}

interface BlockView {
  getSvg: () => SVGElement | null
}

interface BlockManager {
  getBlockViewByBlockId: (blockId: number) => BlockView | null
}

interface DiagramBlock extends Block {
  type: BlockType.DIAGRAM
  blockManager?: BlockManager
  snapshot: {
    type: BlockType.DIAGRAM
  }
}

interface View extends Block<File> {
  type: BlockType.VIEW
}

interface File extends Block {
  type: BlockType.FILE
  snapshot: {
    type: BlockType.FILE
    file: {
      name: string
      token: string
    }
  }
}

enum ISVBlockTypeId {
  /**
   * Text Drawing
   */
  TextDrawing = 'blk_631fefbbae02400430b8f9f4',

  /**
   * Timeline
   */
  Timeline = 'blk_6358a421bca0001c22536e4c',
  /**
   * Other ISV block (type inference)
   */
  _Other = '',
}

interface OtherISVBlock extends Block {
  type: BlockType.ISV
  snapshot: {
    type: BlockType.ISV
    /**
     * ISV block type id
     */
    block_type_id: ISVBlockTypeId._Other
    /**
     * ISV block data
     */
    data: unknown
  }
}

interface TextDrawingBlock extends Block {
  type: BlockType.ISV
  snapshot: {
    type: BlockType.ISV
    /**
     * ISV block type id
     */
    block_type_id: ISVBlockTypeId.TextDrawing
    /**
     * ISV block data
     */
    data: {
      /**
       * Mermaid code
       */
      data: string
    }
  }
}

interface Timeline {
  time: string
  title: string
  text?: string
}

interface TimelineBlock extends Block {
  type: BlockType.ISV
  snapshot: {
    type: BlockType.ISV
    /**
     * ISV block type id
     */
    block_type_id: ISVBlockTypeId.Timeline
    /**
     * ISV block data
     */
    data: {
      /**
       * Mermaid code
       */
      items: Timeline[]
    }
  }
}

type ISVBlocks = TextDrawingBlock | TimelineBlock | OtherISVBlock

interface BitableBlock extends Block {
  type: BlockType.BITABLE | BlockType.BASE_REFER
  snapshot: {
    type: BlockType.BITABLE | BlockType.BASE_REFER
    caption?: Caption
  }
}

interface SheetBlock extends Block {
  type: BlockType.SHEET
  snapshot: {
    type: BlockType.SHEET
    caption?: Caption
  }
  children: []
}

interface NotSupportedBlock extends Block {
  type:
    | BlockType.QUOTE
    | BlockType.CHAT_CARD
    | BlockType.MINDNOTE
    | BlockType.FALLBACK
  children: []
}

type Blocks =
  | PageBlock
  | DividerBlock
  | HeadingBlock
  | CodeBlock
  | QuoteContainerBlock
  | BulletBlock
  | OrderedBlock
  | TodoBlock
  | TextBlock
  | ImageBlock
  | TableBlock
  | TableCellBlock
  | Grid
  | GridColumn
  | Callout
  | SyncedSource
  | SyncedReference
  | Whiteboard
  | DiagramBlock
  | BitableBlock
  | SheetBlock
  | View
  | File
  | IframeBlock
  | ISVBlocks
  | NotSupportedBlock

interface IframeBlock extends Block {
  type: BlockType.IFRAME
  snapshot: {
    type: BlockType.IFRAME
    iframe: Partial<{
      height: number
      component: Partial<{
        url: string
      }>
    }>
  }
}

const iframeToHTML = (iframe: IframeBlock): mdast.Html | null => {
  const { height = 4 * OneHundred, component = {} } = iframe.snapshot.iframe
  const { url } = component

  if (!url) {
    return null
  }

  const html = `<iframe src="${url}" sandbox="allow-scripts allow-same-origin allow-presentation allow-forms allow-popups allow-downloads" allowfullscreen allow="encrypted-media; fullscreen; autoplay" referrerpolicy="strict-origin-when-cross-origin" frameborder="0" style="width: 100%; min-height: ${height.toFixed()}px; border-radius: 8px;"></iframe>`

  return {
    type: 'html',
    value: html,
  }
}

const findRenderedBitableElement = (recordId: string): HTMLElement | null => {
  const selectors = [
    `[data-block-id="${recordId}"]`,
    `[data-record-id="${recordId}"]`,
    `[data-recordid="${recordId}"]`,
    `[data-block-record-id="${recordId}"]`,
    `[data-node-id="${recordId}"]`,
    `[id="${recordId}"]`,
  ]

  for (const selector of selectors) {
    const element = document.querySelector<HTMLElement>(selector)
    const bitableElement =
      element?.closest<HTMLElement>(
        '[data-block-type="bitable"], [class*="bitable"], [class*="Bitable"]',
      ) ?? element

    if (bitableElement) {
      return bitableElement
    }
  }

  return null
}

const extractBitableHtml = (bitableElement: HTMLElement): string => {
  const tableElement = bitableElement.querySelector('table')
  if (tableElement) return tableElement.outerHTML

  const gridElement = bitableElement.querySelector<HTMLElement>(
    '[role="grid"], [role="table"], [class*="grid"], [class*="table"], [class*="Table"]',
  )

  return gridElement?.outerHTML ?? bitableElement.innerHTML
}

const escapeSelectorAttributeValue = (value: string): string =>
  value.replace(/\\/g, '\\\\').replace(/"/g, '\\"')

const sheetContainerSelector =
  '[data-block-type="sheet"], [data-sheet-element="embeddedSheetContainer"], [class*="sheet-widget"], [class*="spreadsheet-widget"], [class*="embed-spreadsheet"], [class*="spreadsheet-wrap"]'

const sheetCanvasSelector =
  'canvas.spreadsheet-canvas, canvas[role="faster"], canvas'

const sheetDataAttributes = (block: SheetBlock): string => {
  const attrs = [
    block.record?.id
      ? `data-sheet-record-id="${escape(block.record.id)}"`
      : null,
    `data-sheet-block-id="${escape(String(block.id))}"`,
  ].filter(isString)

  return attrs.length > 0 ? ` ${attrs.join(' ')}` : ''
}

const findRenderedSheetElement = (block: SheetBlock): HTMLElement | null => {
  const identifiers = [block.record?.id, String(block.id)].filter(isString)

  for (const identifier of identifiers) {
    const escapedIdentifier = escapeSelectorAttributeValue(identifier)
    const selectors = [
      `[data-block-id="${escapedIdentifier}"]`,
      `[data-record-id="${escapedIdentifier}"]`,
      `[data-recordid="${escapedIdentifier}"]`,
      `[data-block-record-id="${escapedIdentifier}"]`,
      `[data-node-id="${escapedIdentifier}"]`,
      `[id="${escapedIdentifier}"]`,
    ]

    for (const selector of selectors) {
      const element = document.querySelector<HTMLElement>(selector)
      const sheetElement =
        element?.closest<HTMLElement>(sheetContainerSelector) ?? element

      if (sheetElement) {
        return sheetElement
      }
    }
  }

  return null
}

const isVisibleElement = (element: HTMLElement): boolean => {
  const style = window.getComputedStyle(element)
  if (style.display === 'none' || style.visibility === 'hidden') return false

  const rect = element.getBoundingClientRect()
  const viewportHeight =
    window.innerHeight || document.documentElement.clientHeight
  const viewportWidth =
    window.innerWidth || document.documentElement.clientWidth

  return (
    rect.width > 0 &&
    rect.height > 0 &&
    rect.bottom > 0 &&
    rect.right > 0 &&
    rect.top < viewportHeight &&
    rect.left < viewportWidth
  )
}

const findVisibleRenderedSheetElement = (): HTMLElement | null => {
  const sheetElements = Array.from(
    document.querySelectorAll<HTMLElement>(sheetContainerSelector),
  )
  const canvasSheetElements = Array.from(
    document.querySelectorAll<HTMLCanvasElement>(sheetCanvasSelector),
  )
    .map(canvas => canvas.closest<HTMLElement>(sheetContainerSelector))
    .filter(isDefined)

  return (
    [...sheetElements, ...canvasSheetElements]
      .filter((element, index, elements) => elements.indexOf(element) === index)
      .filter(isVisibleElement)
      .sort(
        (a, b) =>
          Math.abs(a.getBoundingClientRect().top) -
          Math.abs(b.getBoundingClientRect().top),
      )[0] ?? null
  )
}

const isVisibleCanvas = (canvas: HTMLCanvasElement): boolean => {
  const style = window.getComputedStyle(canvas)

  return (
    canvas.width > 0 &&
    canvas.height > 0 &&
    style.display !== 'none' &&
    style.visibility !== 'hidden'
  )
}

const getVisibleSheetCanvases = (
  sheetElement: HTMLElement,
): HTMLCanvasElement[] =>
  Array.from(
    sheetElement.querySelectorAll<HTMLCanvasElement>(sheetCanvasSelector),
  ).filter(isVisibleCanvas)

const getRenderableSheetCanvases = (
  sheetElement: HTMLElement,
): HTMLCanvasElement[] => getVisibleSheetCanvases(sheetElement)

const snapshotCanvasSignatures = (canvases: HTMLCanvasElement[]): string[] =>
  canvases.map(canvas => {
    try {
      return canvas.toDataURL('image/png')
    } catch {
      return ''
    }
  })

const didCanvasSnapshotsChange = (
  baseline: string[],
  canvases: HTMLCanvasElement[],
): boolean => {
  const signatures = snapshotCanvasSignatures(canvases)

  return signatures.some(
    (signature, index) => signature !== '' && signature !== baseline[index],
  )
}

const canvasSnapshotDimensions = (canvas: HTMLCanvasElement) => {
  const rect = canvas.getBoundingClientRect()
  const width = Math.round(rect.width || canvas.clientWidth || canvas.width)
  const height = Math.round(rect.height || canvas.clientHeight || canvas.height)

  return { rect, width, height }
}

const sheetSnapshotToHtml = (
  dataUrl: string,
  width: number,
  height: number,
): string | null => {
  if (dataUrl === 'data:,' || width <= 0 || height <= 0) return null

  return `<img class="sheet-snapshot" src="${dataUrl}" alt="Embedded sheet snapshot" width="${width.toFixed()}" height="${height.toFixed()}">`
}

const composeSheetCanvases = (canvases: HTMLCanvasElement[]): string | null => {
  if (canvases.length === 0) return null

  if (canvases.length === 1) {
    const { width, height } = canvasSnapshotDimensions(canvases[0])
    try {
      return sheetSnapshotToHtml(
        canvases[0].toDataURL('image/png'),
        width,
        height,
      )
    } catch {
      return null
    }
  }

  const canvasRects = canvases.map(canvas => ({
    canvas,
    ...canvasSnapshotDimensions(canvas),
  }))
  const left = Math.min(...canvasRects.map(({ rect }) => rect.left || 0))
  const top = Math.min(...canvasRects.map(({ rect }) => rect.top || 0))
  const right = Math.max(
    ...canvasRects.map(({ rect, width }) => (rect.left || 0) + width),
  )
  const bottom = Math.max(
    ...canvasRects.map(({ rect, height }) => (rect.top || 0) + height),
  )
  const width = Math.round(right - left)
  const height = Math.round(bottom - top)
  if (width <= 0 || height <= 0) return null

  const scale = Math.max(
    1,
    ...canvasRects.map(({ canvas, width: cssWidth }) =>
      cssWidth > 0 ? canvas.width / cssWidth : 1,
    ),
  )

  const output = document.createElement('canvas')
  output.width = Math.max(1, Math.round(width * scale))
  output.height = Math.max(1, Math.round(height * scale))

  const context = output.getContext('2d')
  if (!context) return null

  for (const {
    canvas,
    rect,
    width: cssWidth,
    height: cssHeight,
  } of canvasRects) {
    context.drawImage(
      canvas,
      Math.round(((rect.left || 0) - left) * scale),
      Math.round(((rect.top || 0) - top) * scale),
      Math.round(cssWidth * scale),
      Math.round(cssHeight * scale),
    )
  }

  try {
    return sheetSnapshotToHtml(output.toDataURL('image/png'), width, height)
  } catch {
    return null
  }
}

const isScrollableElement = (element: HTMLElement): boolean =>
  element.scrollWidth > element.clientWidth + 1 ||
  element.scrollHeight > element.clientHeight + 1

const findSheetScrollContainer = (
  sheetElement: HTMLElement,
): HTMLElement | null => {
  const canvases = getVisibleSheetCanvases(sheetElement)

  for (const canvas of canvases) {
    let element: HTMLElement | null = canvas.parentElement

    while (element && sheetElement.contains(element)) {
      if (isScrollableElement(element)) return element
      if (element === sheetElement) break
      element = element.parentElement
    }
  }

  return (
    [
      sheetElement,
      ...Array.from(sheetElement.querySelectorAll<HTMLElement>('*')),
    ]
      .filter(isScrollableElement)
      .sort(
        (a, b) =>
          a.scrollWidth * a.scrollHeight - b.scrollWidth * b.scrollHeight,
      )[0] ?? null
  )
}

const buildScrollStops = (
  viewportSize: number,
  scrollSize: number,
): number[] => {
  const maxScroll = Math.max(0, scrollSize - viewportSize)
  if (maxScroll === 0) return [0]

  const step = Math.max(1, viewportSize)
  const stops: number[] = []

  for (let position = 0; position < maxScroll; position += step) {
    stops.push(position)
  }

  stops.push(maxScroll)

  return Array.from(new Set(stops))
}

const setScrollPosition = (
  element: HTMLElement,
  left: number,
  top: number,
): void => {
  element.scrollLeft = left
  element.scrollTop = top
  element.dispatchEvent(new Event('scroll', { bubbles: true }))
}

const visibleCanvasBounds = (
  canvas: HTMLCanvasElement,
): { left: number; top: number; right: number; bottom: number } | null => {
  const context = canvas.getContext('2d', { willReadFrequently: true })
  if (!context) return null

  const { data, width, height } = context.getImageData(
    0,
    0,
    canvas.width,
    canvas.height,
  )
  let left = width
  let top = height
  let right = 0
  let bottom = 0

  for (let y = 0; y < height; y++) {
    for (let x = 0; x < width; x++) {
      const alpha = data[(y * width + x) * 4 + 3]
      if (alpha <= 8) continue

      left = Math.min(left, x)
      top = Math.min(top, y)
      right = Math.max(right, x + 1)
      bottom = Math.max(bottom, y + 1)
    }
  }

  if (right <= left || bottom <= top) return null

  return { left, top, right, bottom }
}

const croppedSheetSnapshotToHtml = (
  canvas: HTMLCanvasElement,
  scale: number,
): string | null => {
  try {
    const bounds = visibleCanvasBounds(canvas)
    if (!bounds) return null

    const width = bounds.right - bounds.left
    const height = bounds.bottom - bounds.top
    if (width <= 0 || height <= 0) return null

    const croppedCanvas = document.createElement('canvas')
    croppedCanvas.width = width
    croppedCanvas.height = height

    const context = croppedCanvas.getContext('2d')
    if (!context) return null

    context.drawImage(
      canvas,
      bounds.left,
      bounds.top,
      width,
      height,
      0,
      0,
      width,
      height,
    )

    return sheetSnapshotToHtml(
      croppedCanvas.toDataURL('image/png'),
      width / scale,
      height / scale,
    )
  } catch {
    return null
  }
}

const stitchScrollableSheetCanvases = async (
  sheetElement: HTMLElement,
): Promise<string | null> => {
  const scrollContainer = findSheetScrollContainer(sheetElement)
  if (!scrollContainer) return null

  const hasHorizontalScroll =
    scrollContainer.scrollWidth > scrollContainer.clientWidth + 1
  const hasVerticalScroll =
    scrollContainer.scrollHeight > scrollContainer.clientHeight + 1
  if (!hasHorizontalScroll && !hasVerticalScroll) return null

  const originalLeft = scrollContainer.scrollLeft
  const originalTop = scrollContainer.scrollTop
  const initialCanvases = getVisibleSheetCanvases(sheetElement)
  if (initialCanvases.length === 0) return null

  const initialSignatures = snapshotCanvasSignatures(initialCanvases)
  const scrollLeftStops = buildScrollStops(
    scrollContainer.clientWidth,
    scrollContainer.scrollWidth,
  )
  const scrollTopStops = buildScrollStops(
    scrollContainer.clientHeight,
    scrollContainer.scrollHeight,
  )

  const initialScale = Math.max(
    1,
    ...initialCanvases.map(canvas => {
      const { width } = canvasSnapshotDimensions(canvas)
      return width > 0 ? canvas.width / width : 1
    }),
  )
  const outputWidth = Math.max(
    1,
    Math.round(scrollContainer.scrollWidth * initialScale),
  )
  const outputHeight = Math.max(
    1,
    Math.round(scrollContainer.scrollHeight * initialScale),
  )
  const output = document.createElement('canvas')
  output.width = outputWidth
  output.height = outputHeight

  const context = output.getContext('2d', { willReadFrequently: true })
  if (!context) return null

  try {
    const containerRect = scrollContainer.getBoundingClientRect()
    const initialCanvasRects = initialCanvases.map(canvas => {
      const { rect } = canvasSnapshotDimensions(canvas)
      return rect
    })
    const originLeft = Math.min(
      ...initialCanvasRects.map(rect => rect.left - containerRect.left),
    )
    const originTop = Math.min(
      ...initialCanvasRects.map(rect => rect.top - containerRect.top),
    )
    let didScrollRevealNewPixels = false

    for (const top of scrollTopStops) {
      for (const left of scrollLeftStops) {
        setScrollPosition(scrollContainer, left, top)
        await waitFor(0.25 * Second)

        const canvases = getVisibleSheetCanvases(sheetElement)
        if (canvases.length === 0) continue
        if (
          left !== originalLeft ||
          top !== originalTop ||
          canvases.length !== initialCanvases.length
        ) {
          didScrollRevealNewPixels ||= didCanvasSnapshotsChange(
            initialSignatures,
            canvases,
          )
        }

        for (const canvas of canvases) {
          const { rect, width, height } = canvasSnapshotDimensions(canvas)
          if (width <= 0 || height <= 0) continue

          const scale = width > 0 ? canvas.width / width : initialScale
          context.drawImage(
            canvas,
            Math.round(
              (left + rect.left - containerRect.left - originLeft) * scale,
            ),
            Math.round(
              (top + rect.top - containerRect.top - originTop) * scale,
            ),
            Math.round(width * scale),
            Math.round(height * scale),
          )
        }
      }
    }

    if (!didScrollRevealNewPixels) return null

    const croppedSnapshot = croppedSheetSnapshotToHtml(output, initialScale)
    if (croppedSnapshot) return croppedSnapshot

    return sheetSnapshotToHtml(
      output.toDataURL('image/png'),
      scrollContainer.scrollWidth,
      scrollContainer.scrollHeight,
    )
  } catch {
    return null
  } finally {
    setScrollPosition(scrollContainer, originalLeft, originalTop)
  }
}

const extractSheetHtml = (sheetElement: HTMLElement): string | null => {
  const tableElement = sheetElement.querySelector('table')
  if (tableElement) return tableElement.outerHTML

  const gridElement = sheetElement.querySelector<HTMLElement>(
    '[role="grid"]:not(canvas), [role="table"]:not(canvas), table[role="grid"], table[role="table"]',
  )
  if (gridElement) return gridElement.outerHTML

  const canvasSnapshot = composeSheetCanvases(
    getRenderableSheetCanvases(sheetElement),
  )
  if (canvasSnapshot) return canvasSnapshot

  const fallbackGridElement = sheetElement.querySelector<HTMLElement>(
    '[class*="grid"]:not(canvas), [class*="table"]:not(canvas), [class*="Table"]:not(canvas), [class*="Spreadsheet"]:not(canvas), [class*="spreadsheet"]:not(canvas)',
  )

  return fallbackGridElement?.outerHTML ?? null
}

const extractSheetHtmlAsync = async (
  sheetElement: HTMLElement,
): Promise<string | null> => {
  const tableElement = sheetElement.querySelector('table')
  if (tableElement) return tableElement.outerHTML

  const gridElement = sheetElement.querySelector<HTMLElement>(
    '[role="grid"]:not(canvas), [role="table"]:not(canvas), table[role="grid"], table[role="table"]',
  )
  if (gridElement) return gridElement.outerHTML

  return (
    (await stitchScrollableSheetCanvases(sheetElement)) ??
    extractSheetHtml(sheetElement)
  )
}

const hasRenderableSheetContent = (sheetElement: HTMLElement): boolean =>
  Boolean(
    sheetElement.querySelector(
      'table, [role="grid"]:not(canvas), [role="table"]:not(canvas)',
    ) ?? getVisibleSheetCanvases(sheetElement)[0],
  )

const sheetToHtml = (
  block: SheetBlock,
  sheetElement: HTMLElement | null,
): string => {
  const caption = evaluateAlt(block.snapshot.caption)
  const captionHtml = caption
    ? `<figcaption>${escape(caption)}</figcaption>`
    : ''
  const contentHtml =
    (sheetElement ? extractSheetHtml(sheetElement) : null) ??
    `<p class="sheet-missing">Sheet content is not loaded in the current page.</p>`

  return `<figure class="sheet"${sheetDataAttributes(block)}>${captionHtml}<div class="sheet-wrapper">${contentHtml}</div></figure>`
}

const sheetToHtmlAsync = async (
  block: SheetBlock,
  sheetElement: HTMLElement | null,
): Promise<string> => {
  const caption = evaluateAlt(block.snapshot.caption)
  const captionHtml = caption
    ? `<figcaption>${escape(caption)}</figcaption>`
    : ''
  const contentHtml =
    (sheetElement ? await extractSheetHtmlAsync(sheetElement) : null) ??
    `<p class="sheet-missing">Sheet content is not loaded in the current page.</p>`

  return `<figure class="sheet"${sheetDataAttributes(block)}>${captionHtml}<div class="sheet-wrapper">${contentHtml}</div></figure>`
}

/**
 * @description Removes an enter from the end of this string if it exists.
 */
const trimEndEnter = (input: string) =>
  input.length > 0 && input.endsWith('\n') ? input.slice(0, -1) : input

const chunkBy = <T>(
  items: T[],
  isEqual: (current: T, next: T) => boolean,
): T[][] => {
  const chunks: T[][] = []
  let index = 0

  while (index < items.length) {
    let nextIndex = index + 1
    while (
      nextIndex < items.length &&
      isEqual(items[index], items[nextIndex])
    ) {
      nextIndex++
    }

    chunks.push(items.slice(index, nextIndex))

    index = nextIndex
  }

  return chunks
}

export const mergeListItems = <T extends mdast.Nodes>(
  nodes: T[],
): (mdast.List | T)[] =>
  chunkBy(nodes, (current, next) => {
    const listItemType = (listItem: mdast.ListItem) => {
      if (typeof listItem.checked === 'boolean') {
        return BlockType.TODO
      }

      if (
        typeof listItem.data?.seq === 'number' ||
        listItem.data?.seq === 'auto'
      ) {
        return BlockType.ORDERED
      }

      return BlockType.BULLET
    }

    const isEqualOrderedListItem = (
      node: mdast.ListItem,
      other: mdast.ListItem,
    ) => {
      const seq = node.data?.seq
      const otherSeq = other.data?.seq

      if (!seq || !otherSeq) return false

      if (seq === 'auto') {
        return otherSeq === 'auto'
      }

      return otherSeq === 'auto' || seq + 1 === otherSeq
    }

    const isEqualListItem = (node: mdast.ListItem, other: mdast.ListItem) => {
      const type = listItemType(node)
      const otherType = listItemType(other)

      if (type === otherType) {
        return type === BlockType.ORDERED
          ? isEqualOrderedListItem(node, other)
          : true
      }

      return false
    }

    return (
      current.type === 'listItem' &&
      next.type === 'listItem' &&
      isEqualListItem(current, next)
    )
  }).map(nodes => {
    const node = nodes[0]

    if (node.type === 'listItem') {
      const list: mdast.List = {
        type: 'list',
        ...(typeof node.data?.seq === 'number'
          ? {
              ordered: true,
              start: node.data.seq,
            }
          : null),
        children: nodes as mdast.ListItem[],
      }
      return list
    }

    return node
  })

export const mergePhrasingContents = (
  nodes: mdast.PhrasingContent[],
): mdast.PhrasingContent[] =>
  chunkBy(nodes, (current, next) => {
    if (current.type === 'link' && next.type === 'link') {
      return current.url === next.url
    }

    if (
      current.type === 'emphasis' ||
      current.type === 'strong' ||
      current.type === 'delete' ||
      current.type === 'text' ||
      (current.type === 'inlineCode' && !current.data?.mentionUserId)
    ) {
      return current.type === next.type
    }

    return false
  })
    .map(nodes => {
      const node = nodes.reduce((pre, cur) => {
        if ('children' in pre && 'children' in cur) {
          return {
            ...pre,
            ...cur,
            children: pre.children.concat(cur.children),
          }
        }

        if ('value' in pre && 'value' in cur) {
          return {
            ...pre,
            ...cur,
            value: pre.value.concat(cur.value),
          }
        }

        return pre
      })

      if ('children' in node) {
        node.children = mergePhrasingContents(node.children)
      }

      return node
    })
    .flatMap((current, index, merged) => {
      const next = merged.at(index + 1)

      return next && current.type === next.type
        ? [current, { type: 'text', value: ' ' } satisfies mdast.Text]
        : [current]
    })

export interface transformOperationsToPhrasingContentsOptions {
  highlight?: boolean
}

export const transformOperationsToPhrasingContents = (
  ops: Operation[],
  options: transformOperationsToPhrasingContentsOptions = {},
): { contents: mdast.PhrasingContent[]; mentionUsers: mdast.InlineCode[] } => {
  const mentionUsers: mdast.InlineCode[] = []

  const operations = ops
    .filter(operation => {
      if (
        isDefined(operation.attributes) &&
        isDefined(operation.attributes.fixEnter)
      ) {
        return false
      }

      if (!isDefined(operation.attributes) && operation.insert === '\n') {
        return false
      }

      return true
    })
    .map(op => {
      if (isDefined(op.attributes) && op.attributes['inline-component']) {
        try {
          const inlineComponent = JSON.parse(
            op.attributes['inline-component'],
          ) as
            | {
                type: 'mention_doc'
                data: {
                  raw_url: string
                  title: string
                }
              }
            | {
                type: 'user'
                data: {
                  uid: string
                }
              }
            | {
                type: 'string'
                data: unknown
              }
          if (inlineComponent.type === 'mention_doc') {
            return {
              attributes: {
                ...op.attributes,
                link: inlineComponent.data.raw_url,
              },
              insert: op.insert + inlineComponent.data.title,
            } as Operation
          } else if (inlineComponent.type === 'user') {
            return {
              attributes: {
                ...op.attributes,
                mentionUserId: inlineComponent.data.uid,
              },
              insert: '',
            }
          }

          return op
        } catch {
          return op
        }
      }

      return op
    })

  let indexToMarks = operations.map(({ attributes = {} }) => {
    type SupportAttrName = 'italic' | 'bold' | 'strikethrough' | 'link'

    const isSupportAttr = (attr: string): attr is SupportAttrName =>
      attr === 'italic' ||
      attr === 'bold' ||
      attr === 'strikethrough' ||
      attr === 'link'

    const attrNameToNodeType = (
      attr: SupportAttrName,
    ): 'emphasis' | 'strong' | 'delete' | 'link' => {
      switch (attr) {
        case 'italic':
          return 'emphasis'
        case 'bold':
          return 'strong'
        case 'strikethrough':
          return 'delete'
        case 'link':
          return 'link'
        default:
          return undefined as never
      }
    }

    const marks = Object.keys(attributes)
      .filter(isSupportAttr)
      .map(attrNameToNodeType)

    return marks
  })

  indexToMarks = indexToMarks.map((marks, index) => {
    const markToPriority = new Map(marks.map(mark => [mark, 0]))

    marks.forEach(mark => {
      let priority = 0
      let start = index
      while (start >= 0 && indexToMarks[start].includes(mark)) {
        priority += operations[start].insert.length
        start--
      }
      let end = index + 1
      while (end < indexToMarks.length && indexToMarks[end].includes(mark)) {
        priority += operations[end].insert.length
        end++
      }
      markToPriority.set(mark, priority)
    })

    return marks.sort((a, b) =>
      compare(markToPriority.get(a) ?? 0, markToPriority.get(b) ?? 0),
    )
  })

  const createLiteral = (
    op: Operation,
  ): mdast.Text | mdast.InlineCode | InlineMath | mdast.Html => {
    const { attributes, insert } = op
    const {
      inlineCode,
      equation,
      textHighlight,
      textHighlightBackground,
      mentionUserId,
      underline,
    } = attributes ?? {}

    if (mentionUserId) {
      const mentionUser: mdast.InlineCode = {
        type: 'inlineCode',
        value: insert,
        data: {
          mentionUserId,
        },
      }

      mentionUsers.push(mentionUser)

      return mentionUser
    }

    if (inlineCode) {
      return {
        type: 'inlineCode',
        value: insert,
      }
    }

    if (equation && equation.length > 0) {
      return {
        type: 'inlineMath',
        value: trimEndEnter(equation),
      }
    }

    if (options.highlight && (textHighlight || textHighlightBackground)) {
      const highlighted = `<span style="color: ${textHighlight ?? 'inherit'}; background-color: ${textHighlightBackground ?? 'inherit'}">${escape(insert)}</span>`

      return {
        type: 'html',
        value: underline ? `<u>${highlighted}</u>` : highlighted,
      }
    }

    if (underline) {
      return {
        type: 'html',
        value: `<u>${escape(insert)}</u>`,
      }
    }

    return {
      type: 'text',
      value: insert,
    }
  }

  const nodes = indexToMarks.map((marks, index) => {
    const op = operations[index]

    let node: mdast.PhrasingContent = createLiteral(op)
    for (const mark of marks) {
      node =
        mark === 'link'
          ? {
              type: mark,
              url: decodeURIComponent(op.attributes?.link ?? ''),
              children: [node],
            }
          : {
              type: mark,
              children: [node],
            }
    }

    return node
  })

  const contents = mergePhrasingContents(nodes)

  return {
    contents,
    mentionUsers,
  }
}

const fetchImageSources = (imageBlock: ImageBlock) =>
  new Promise<ImageSources>((resolve, reject) => {
    const {
      imageManager,
      snapshot: {
        image: { token },
      },
    } = imageBlock

    imageManager
      .fetch({ token, isHD: true, fuzzy: false }, {}, resolve)
      .catch(reject)
  })

/**
 * @description Whether the whiteboard's scene has finished loading and can be
 * captured. Merely having `whiteboardBlock` defined is not enough: when the
 * block is captured before its nodes are laid out, `getNodesBounds()` returns a
 * degenerate box, producing a blank image with abnormal dimensions.
 */
const isWhiteboardContentReady = (whiteboard: Whiteboard): boolean => {
  const block = whiteboard.whiteboardBlock
  if (!block) return false

  const ratioApp = block.abilityKit.getRatioApp()
  if (ratioApp.app) {
    const { minX, maxX, minY, maxY } =
      ratioApp.app.application.nodeManager.getNodesBounds()

    // The app path can render once its scene has non-empty bounds.
    if (maxX - minX > 0 && maxY - minY > 0) return true
  }

  // Otherwise fall back to the isolateEnv path's own readiness signal.
  return (
    block.isolateEnv.hasRatioApp() &&
    block.isolateEnv.getRatioApp().ratioAppProxy !== undefined
  )
}

interface WhiteboardCaptureOptions {
  padding: number
  ratio: number
  backgroundColor: string
}

/**
 * @description Capture via the new `abilityKit` app path. Returns `null` when it
 * cannot produce a valid image (no app, empty scene bounds, or no canvas) so the
 * caller can fall back to the isolateEnv path.
 */
const whiteboardAppToBlob = async (
  block: WhiteboardBlock,
  { padding, ratio, backgroundColor }: WhiteboardCaptureOptions,
): Promise<Blob | null> => {
  const ratioApp = block.abilityKit.getRatioApp()
  if (!ratioApp.app) return null

  const bounds = ratioApp.app.application.nodeManager.getNodesBounds()

  // A degenerate box means the scene is not laid out yet; bail so we don't emit
  // a blank image with padding-only dimensions.
  if (bounds.maxX - bounds.minX <= 0 || bounds.maxY - bounds.minY <= 0) {
    return null
  }

  bounds.maxX += padding
  bounds.minX -= padding
  bounds.maxY += padding
  bounds.minY -= padding

  const canvas = ratioApp.app.renderManager.getImageOffscreenCanvas(
    bounds,
    ratio,
    backgroundColor,
  )

  if (!canvas) return null

  return await new Promise<Blob | null>(resolve => {
    checkCanvasDimensions(canvas)

    canvas.toBlob(resolve)
  })
}

/**
 * @description Capture via the legacy `isolateEnv` path. Used as a fallback when
 * the app path is unavailable or its scene bounds are still empty.
 */
const whiteboardIsolateEnvToBlob = async (
  block: WhiteboardBlock,
  { padding, ratio }: WhiteboardCaptureOptions,
): Promise<Blob | null> => {
  const ratioApp = block.isolateEnv.getRatioApp()

  const imageDataWrapper =
    await ratioApp.ratioAppProxy?.getOriginImageDataByNodeId(
      padding,
      [''],
      false,
      ratio,
    )

  if (!imageDataWrapper) return null

  return await imageDataToBlob(imageDataWrapper.data, {
    onDispose: imageDataWrapper.release,
  })
}

const whiteboardToBlob = async (
  whiteboard: Whiteboard,
): Promise<Blob | null> => {
  const block = whiteboard.whiteboardBlock
  if (!block) return null

  const options: WhiteboardCaptureOptions = {
    padding: 24,
    ratio: window.devicePixelRatio,
    backgroundColor: '#ffffff',
  }

  // Prefer the app path; fall back to isolateEnv when it can't produce a valid
  // image (e.g. app exists but its scene bounds are still empty).
  return (
    (await whiteboardAppToBlob(block, options)) ??
    (await whiteboardIsolateEnvToBlob(block, options))
  )
}

const diagramToSVGElement = (diagram: DiagramBlock): SVGElement | null => {
  if (!diagram.blockManager) return null

  const blockView = diagram.blockManager.getBlockViewByBlockId(diagram.id)
  if (!blockView) return null

  const svgElement = blockView.getSvg()
  if (!svgElement) return null

  return svgElement
}

const generateMermaidTimeline = (items: Timeline[]): string => {
  let chart = 'timeline\n'

  items.forEach(item => {
    const cleanTitle = (item.title || '').replace(/:/g, '：')
    const time = item.time || ''

    if (item.text) {
      const cleanText = item.text.replace(/\n/g, '<br>')
      chart += `    ${time} : ${cleanTitle} : ${cleanText}\n`
    } else {
      chart += `    ${time} : ${cleanTitle}\n`
    }
  })

  return chart
}

const evaluateAlt = (caption?: Caption) =>
  trimEndEnter(caption?.text.initialAttributedTexts.text?.[0] ?? '')

type Mutate<T extends Block> = T extends PageBlock
  ? mdast.Root
  : T extends DividerBlock
    ? mdast.ThematicBreak
    : T extends HeadingBlock
      ? mdast.Heading
      : T extends CodeBlock
        ? mdast.Code
        : T extends QuoteContainerBlock | Callout
          ? mdast.Blockquote
          : T extends BulletBlock | OrderedBlock | TodoBlock
            ? mdast.ListItem
            : T extends TextBlock
              ? mdast.Text
              : T extends TableBlock | Grid
                ? mdast.Table
                : T extends TableCellBlock | GridColumn
                  ? mdast.TableCell
                  : T extends Whiteboard | DiagramBlock
                    ? mdast.Image
                    : T extends View
                      ? mdast.Paragraph
                      : T extends File
                        ? mdast.Link
                        : T extends IframeBlock
                          ? mdast.Html
                          : T extends BitableBlock
                            ? mdast.Html
                            : T extends SheetBlock
                              ? mdast.Html
                              : T extends TextDrawingBlock | TimelineBlock
                                ? mdast.Code
                                : null

interface TransformerOptions {
  /**
   * Enable convert whiteboard to image.
   * @default false
   */
  whiteboard?: boolean
  /**
   * Enable convert diagram to image.
   * @default false
   */
  diagram?: boolean
  /**
   * Enable convert file to resource link.
   * @default false
   */
  file?: boolean
  /**
   * Enable convert text highlight to html.
   * @default false
   */
  highlight?: boolean
  /**
   * Enable flat grid.
   * @default false
   */
  flatGrid?: boolean
  /**
   * Enable convert bitable to HTML table.
   * @default false
   */
  bitable?: boolean
  /**
   * Locate block with record id.
   */
  locateBlockWithRecordId?: (recordId: string) => Promise<boolean>
}

export interface TableWithParent {
  inner: mdast.Table
  parent: mdast.Parent | null
}

interface TransformResult<T> {
  root: T
  images: mdast.Image[]
  tableWithParents: TableWithParent[]
  files: mdast.Link[]
  mentionUsers: mdast.InlineCode[]
}

export class Transformer {
  private parent: mdast.Parent | null = null
  private images: mdast.Image[] = []
  private mentionUsers: mdast.InlineCode[] = []
  private tableWithParents: TableWithParent[] = []
  /**
   * Resource link to file.
   */
  private files: mdast.Link[] = []
  /**
   * heading sequence state
   */
  private sequences: (string | undefined)[] = []

  constructor(
    public options: TransformerOptions = {
      whiteboard: false,
      diagram: false,
      file: false,
      highlight: false,
      flatGrid: false,
    },
  ) {}

  private normalizeImage(image: mdast.Image): mdast.Image | mdast.Paragraph {
    return this.parent?.type === 'tableCell'
      ? image
      : { type: 'paragraph', children: [image] }
  }

  private transformParentBlock<T extends Blocks>(
    block: T,
    evaluateNode: (block: T) => Mutate<T>,
    transformChildren: (
      children: mdast.Nodes[],
    ) => Mutate<T> extends mdast.Parent ? Mutate<T>['children'] : never,
  ) {
    const previousParent = this.parent

    const currentParent = evaluateNode(block)
    if (!currentParent || !isParent(currentParent)) {
      return currentParent
    }
    this.parent = currentParent

    const flatChildren = (children: Blocks[]): Blocks[] =>
      children
        .map(child => {
          if (child.type === BlockType.GRID && this.options.flatGrid) {
            return flatChildren(
              child.children.map(column => column.children).flat(1),
            )
          }

          if (
            child.type === BlockType.HEADING1 ||
            child.type === BlockType.HEADING2 ||
            child.type === BlockType.HEADING3 ||
            child.type === BlockType.HEADING4 ||
            child.type === BlockType.HEADING5 ||
            child.type === BlockType.HEADING6 ||
            child.type === BlockType.HEADING7 ||
            child.type === BlockType.HEADING8 ||
            child.type === BlockType.HEADING9 ||
            child.type === BlockType.TEXT
          ) {
            return [child, ...flatChildren(child.children)]
          }

          if (child.type === BlockType.SYNCED_SOURCE) {
            return flatChildren(child.children)
          }

          if (child.type === BlockType.SYNCED_REFERENCE) {
            return flatChildren(
              child.innerBlockManager?.rootBlockModel?.children ??
                child.children,
            )
          }

          return child
        })
        .flat(1)

    currentParent.children = transformChildren(
      flatChildren(block.children).map(this._transform).filter(isDefined),
    )

    this.parent = previousParent

    return currentParent
  }

  private _transform = (block: Blocks): mdast.Nodes | null => {
    const createChildrenFromOps = () => {
      const { contents, mentionUsers } = transformOperationsToPhrasingContents(
        block.zoneState?.content.ops ?? [],
        { highlight: this.options.highlight },
      )

      mentionUsers.forEach(user => {
        if (user.data) {
          user.data.parentBlockRecordId = block.record?.id
        }
      })

      this.mentionUsers = this.mentionUsers.concat(mentionUsers)

      return contents
    }

    switch (block.type) {
      case BlockType.PAGE: {
        return this.transformParentBlock(
          block,
          () => ({
            type: 'root',
            children: [],
          }),
          nodes => mergeListItems(nodes).filter(isRootContent),
        )
      }
      case BlockType.DIVIDER: {
        const thematicBreak: mdast.ThematicBreak = {
          type: 'thematicBreak',
        }
        return thematicBreak
      }
      case BlockType.HEADING1:
      case BlockType.HEADING2:
      case BlockType.HEADING3:
      case BlockType.HEADING4:
      case BlockType.HEADING5:
      case BlockType.HEADING6: {
        const depth = Number(block.type.at(-1)) as mdast.Heading['depth']

        const heading: mdast.Heading = {
          type: 'heading',
          depth,
          children: createChildrenFromOps(),
        }

        if (typeof block.snapshot.seq === 'string') {
          // reset sequences state
          this.sequences = this.sequences.slice(0, depth)

          // automatic incremental sequence number
          if (block.snapshot.seq === 'auto') {
            const previousSequenceSibling = this.sequences[depth - 1] ?? '0'
            this.sequences[depth - 1] = String(
              parseInt(previousSequenceSibling, 10) + 1,
            )
          } else {
            this.sequences[depth - 1] = block.snapshot.seq
          }

          const sequences =
            block.snapshot.seq_level === 'auto'
              ? this.sequences.slice(0, depth).filter(isString)
              : [block.snapshot.seq]

          heading.children.unshift({
            type: 'text',
            value: sequences.join('.') + (sequences.length === 1 ? '. ' : ' '),
          })
        }

        return heading
      }
      case BlockType.CODE: {
        const code: mdast.Code = {
          type: 'code',
          lang: block.language.toLocaleLowerCase(),
          value: trimEndEnter(block.zoneState?.allText ?? ''),
        }
        return code
      }
      case BlockType.QUOTE_CONTAINER:
      case BlockType.CALLOUT: {
        return this.transformParentBlock(
          block,
          () => ({
            type: 'blockquote',
            children: [],
          }),
          nodes => mergeListItems(nodes).filter(isBlockquoteContent),
        )
      }
      case BlockType.BULLET:
      case BlockType.ORDERED:
      case BlockType.TODO: {
        const paragraph: mdast.Paragraph = {
          type: 'paragraph',
          children: createChildrenFromOps(),
        }
        return this.transformParentBlock(
          block,
          () => ({
            type: 'listItem',
            children: [],
            ...(block.type === BlockType.TODO
              ? { checked: Boolean(block.snapshot.done) }
              : null),
            ...(block.type === BlockType.ORDERED
              ? {
                  data: {
                    seq: /[0-9]+/.test(block.snapshot.seq)
                      ? Number(block.snapshot.seq)
                      : 'auto',
                  },
                }
              : null),
          }),
          nodes => [
            paragraph,
            ...mergeListItems(nodes).filter(isListItemContent),
          ],
        )
      }
      case BlockType.TEXT:
      case BlockType.HEADING7:
      case BlockType.HEADING8:
      case BlockType.HEADING9: {
        const paragraph: mdast.Paragraph = {
          type: 'paragraph',
          children: createChildrenFromOps(),
        }
        return paragraph
      }
      case BlockType.IMAGE: {
        const imageBlockToImage = (block: ImageBlock) => {
          const { caption, name, token } = block.snapshot.image
          const image: mdast.Image = {
            type: 'image',
            url: '',
            alt: evaluateAlt(caption),
            data: {
              name,
              token,
              fetchSources: () => fetchImageSources(block),
            },
          }
          return image
        }

        const image: mdast.Image = imageBlockToImage(block)

        this.images.push(image)

        return this.normalizeImage(image)
      }
      case BlockType.WHITEBOARD: {
        if (!this.options.whiteboard) return null

        const whiteboardToImage = (whiteboard: Whiteboard): mdast.Image => {
          const image: mdast.Image = {
            type: 'image',
            url: '',
            alt: evaluateAlt(whiteboard.snapshot.caption),
            data: {
              fetchBlob: async () => {
                try {
                  const {
                    locateBlockWithRecordId = () => Promise.resolve(false),
                  } = this.options

                  await waitForFunction(
                    () =>
                      locateBlockWithRecordId(whiteboard.record?.id ?? '').then(
                        isSuccess =>
                          isSuccess && isWhiteboardContentReady(whiteboard),
                      ),
                    {
                      // Heavy whiteboards may take a while to lay out after the
                      // block is located; keep polling so we don't capture an
                      // empty scene (a blank image with padding-only dimensions).
                      timeout: 10 * Second,
                    },
                  )
                } catch (error) {
                  console.error(error)
                }

                return await whiteboardToBlob(whiteboard)
              },
            },
          }
          return image
        }

        const image: mdast.Image = whiteboardToImage(block)

        this.images.push(image)

        return this.normalizeImage(image)
      }
      case BlockType.DIAGRAM: {
        if (!this.options.diagram) return null

        const diagramToImage = (diagram: DiagramBlock): mdast.Image => {
          const image: mdast.Image = {
            type: 'image',
            url: '',
            data: {
              fetchBlob: async () => {
                try {
                  const {
                    locateBlockWithRecordId = () => Promise.resolve(false),
                  } = this.options

                  await waitForFunction(
                    () =>
                      locateBlockWithRecordId(diagram.record?.id ?? '').then(
                        isSuccess => isSuccess,
                      ),
                    {
                      timeout: 3 * Second,
                    },
                  )
                } catch (error) {
                  console.error(error)
                }

                const svgElement = diagramToSVGElement(diagram)
                if (!svgElement) return null

                return await svgToBlob(svgElement)
              },
            },
          }
          return image
        }

        const image: mdast.Image = diagramToImage(block)

        this.images.push(image)

        return this.normalizeImage(image)
      }
      case BlockType.TABLE:
      case BlockType.GRID: {
        let table: mdast.Table = {
          type: 'table',
          children: [],
          data: {
            type: block.type,
            ...(block.type === BlockType.TABLE
              ? { cellSet: block.snapshot.cell_set }
              : {}),
          },
        }

        table = this.transformParentBlock(
          block,
          () => table,
          nodes => {
            const tableCells = nodes.filter(isTableCell)

            const widthCells = tableCells.filter(
              (cell): cell is mdast.TableCell & { data: { width: number } } =>
                typeof cell.data?.width === 'number',
            )
            const colWidths =
              block.type === BlockType.GRID
                ? widthCells.length === tableCells.length
                  ? widthCells.map(cell => cell.data.width)
                  : undefined
                : block.snapshot.columns_id.map(
                    id => block.snapshot.column_set[id].column_width,
                  )

            table.data = {
              ...table.data,
              type: block.type,
              ...(colWidths ? { colWidths } : {}),
              invalid: tableCells.some(cell => cell.data?.invalidChildren),
            }

            return (
              block.type === BlockType.GRID
                ? [tableCells]
                : chunk(tableCells, block.snapshot.columns_id.length)
            ).map(tableCells => ({
              type: 'tableRow',
              children: tableCells,
            }))
          },
        )

        this.tableWithParents.push({
          inner: table,
          parent: this.parent,
        })

        return table
      }
      case BlockType.CELL:
      case BlockType.GRID_COLUMN: {
        const cell: mdast.TableCell = {
          type: 'tableCell',
          children: [],
          ...(block.type === BlockType.GRID_COLUMN
            ? { data: { width: block.snapshot.width_ratio } }
            : {
                data: {
                  ...toCamelCaseKeys(
                    (this.parent as mdast.Table).data?.cellSet?.[block.cellId]
                      ?.merge_info,
                  ),
                },
              }),
        }

        return this.transformParentBlock(
          block,
          () => cell,
          nodes => {
            const mergedNodes = mergeListItems(nodes)
            const normalizedNodes: mdast.Nodes[] = []

            for (let i = 0; i < mergedNodes.length; i++) {
              const node = mergedNodes[i]
              const nextNode = mergedNodes.at(i + 1)

              if (node.type === 'paragraph') {
                normalizedNodes.push(...node.children)
              } else {
                normalizedNodes.push(node as mdast.PhrasingContent)
              }

              if (
                nextNode &&
                node.type === 'paragraph' &&
                nextNode.type === 'paragraph'
              ) {
                normalizedNodes.push({ type: 'html', value: '<br />' })
              }
            }

            if (normalizedNodes.every(isPhrasingContent)) {
              return normalizedNodes
            }

            cell.data = {
              ...cell.data,
              invalidChildren: normalizedNodes,
            }

            return normalizedNodes.filter(isPhrasingContent)
          },
        )
      }
      case BlockType.VIEW: {
        if (!this.options.file) return null

        const paragraph: mdast.Paragraph = this.transformParentBlock(
          block,
          () => ({
            type: 'paragraph',
            children: [],
          }),
          nodes => nodes.filter(isPhrasingContent),
        )
        return paragraph
      }
      case BlockType.FILE: {
        if (!this.options.file) return null

        const { name, token } = block.snapshot.file

        const link: mdast.Link = {
          type: 'link',
          url: '',
          children: [{ type: 'text', value: name }],
          data: {
            name,
            fetchFile: (init?: RequestInit) =>
              fetch(
                resolveFileDownloadUrl({
                  token,
                  recordId: block.record?.id ?? '',
                }),
                {
                  method: 'Get',
                  credentials: 'include',
                  ...init,
                },
              ),
          },
        }

        this.files.push(link)

        return link
      }
      case BlockType.IFRAME: {
        return iframeToHTML(block)
      }
      case BlockType.ISV: {
        if (block.snapshot.block_type_id === ISVBlockTypeId.TextDrawing) {
          const code: mdast.Code = {
            type: 'code',
            lang: 'mermaid',
            value: block.snapshot.data.data,
          }

          return code
        } else if (block.snapshot.block_type_id === ISVBlockTypeId.Timeline) {
          const code: mdast.Code = {
            type: 'code',
            lang: 'mermaid',
            value: generateMermaidTimeline(block.snapshot.data.items),
          }

          return code
        }

        return null
      }
      case BlockType.BITABLE:
      case BlockType.BASE_REFER: {
        if (!this.options.bitable) return null

        const recordId = block.record?.id
        if (!recordId) return null

        const bitableElement = findRenderedBitableElement(recordId)
        const tableHtml = bitableElement
          ? extractBitableHtml(bitableElement)
          : ''

        const caption = evaluateAlt(block.snapshot.caption)
        const captionHtml = caption ? `<figcaption>${caption}</figcaption>` : ''
        const contentHtml =
          tableHtml ||
          `<p class="bitable-missing">Bitable content is not loaded in the current page.</p>`

        const html: mdast.Html = {
          type: 'html',
          value: `<figure class="bitable">${captionHtml}<div class="bitable-wrapper">${contentHtml}</div></figure>`,
        }

        return html
      }
      case BlockType.SHEET: {
        if (!this.options.bitable) return null

        const { locateBlockWithRecordId = () => Promise.resolve(false) } =
          this.options
        const sheetElement = findRenderedSheetElement(block)

        const html: mdast.Html = {
          type: 'html',
          value: sheetToHtml(block, sheetElement),
          data: {
            fetchHtml: async () => {
              let locatedBlock = false
              if (block.record?.id) {
                locatedBlock = await locateBlockWithRecordId(block.record.id)
              }

              try {
                await waitForFunction(
                  () => {
                    const exactElement = findRenderedSheetElement(block)
                    if (
                      exactElement &&
                      hasRenderableSheetContent(exactElement)
                    ) {
                      return true
                    }

                    const visibleElement = findVisibleRenderedSheetElement()

                    return (
                      locatedBlock &&
                      visibleElement !== null &&
                      hasRenderableSheetContent(visibleElement)
                    )
                  },
                  { timeout: 3 * Second },
                )
              } catch {
                // Keep the placeholder if the sheet is still not rendered.
              }

              return await sheetToHtmlAsync(
                block,
                findRenderedSheetElement(block) ??
                  (locatedBlock ? findVisibleRenderedSheetElement() : null),
              )
            },
          },
        }

        return html
      }
      default:
        return null
    }
  }

  transform<T extends Blocks>(block: T): TransformResult<Mutate<T>> {
    const node = this._transform(block) as Mutate<T>

    const result: TransformResult<Mutate<T>> = {
      root: node,
      images: this.images,
      tableWithParents: this.tableWithParents,
      files: this.files,
      mentionUsers: this.mentionUsers,
    }

    this.parent = null

    this.images = []
    this.tableWithParents = []
    this.files = []
    this.mentionUsers = []

    this.sequences = []

    return result
  }
}

export class Docx {
  static stringify(root: mdast.Root, options?: Options): string {
    return toMarkdown(root, {
      ...options,
      extensions: [
        gfmStrikethroughToMarkdown(),
        gfmTaskListItemToMarkdown(),
        gfmTableToMarkdown(),
        mathToMarkdown({
          singleDollarTextMath: false,
        }),
        ...(options?.extensions ?? []),
      ],
    })
  }

  static async locateBlockWithRecordId(recordId: string): Promise<boolean> {
    try {
      if (!PageMain) {
        return false
      }

      return await PageMain.locateBlockWithRecordIdImpl(recordId)
    } catch (error) {
      console.error(error)
    }

    return false
  }

  get isDocx(): boolean {
    return isDocx()
  }

  get isDoc(): boolean {
    return !isDocx() && isDoc()
  }

  get rootBlock(): PageBlock | null {
    if (!PageMain) {
      return null
    }

    return PageMain.blockManager.rootBlockModel
  }

  get language(): 'zh' | 'en' {
    return User?.language === 'zh' ? 'zh' : 'en'
  }

  get pageTitle(): string | undefined {
    if (!this.rootBlock?.zoneState) return undefined

    return trimEndEnter(this.rootBlock.zoneState.allText)
  }

  get container(): HTMLDivElement | null {
    const container = document.querySelector<HTMLDivElement>(
      '#mainBox .bear-web-x-container',
    )

    return container
  }

  isReady(
    options: {
      /**
       * @default false
       */
      checkWhiteboard?: boolean
    } = {},
  ): boolean {
    const { checkWhiteboard = false } = options

    return (
      !!this.rootBlock &&
      this.rootBlock.children.every(block => {
        const prerequisite = block.snapshot.type !== 'pending'

        const isWhiteboard = (block: Blocks): boolean =>
          block.type === BlockType.WHITEBOARD ||
          (block.type === BlockType.FALLBACK &&
            block.snapshot.type === BlockType.WHITEBOARD)

        const isSyncedReferenceReady = (block: Blocks): boolean =>
          block.type !== BlockType.SYNCED_REFERENCE || block.isAllDataReady

        if (checkWhiteboard && isWhiteboard(block)) {
          return prerequisite && block.type !== BlockType.FALLBACK
        }

        return prerequisite && isSyncedReferenceReady(block)
      })
    )
  }

  scrollTo(options: ScrollToOptions): void {
    const container = this.container
    if (container) {
      const {
        left,
        top = container.scrollHeight,
        behavior = 'smooth',
      } = options

      container.scrollTo({
        left,
        top: Math.min(top, container.scrollHeight),
        behavior,
      })
    }
  }

  intoMarkdownAST(
    transformerOptions: TransformerOptions = {},
  ): TransformResult<mdast.Root> {
    if (!this.rootBlock) {
      return {
        root: { type: 'root', children: [] },
        images: [],
        tableWithParents: [],
        files: [],
        mentionUsers: [],
      }
    }

    const transformer = new Transformer({
      locateBlockWithRecordId: recordId =>
        Docx.locateBlockWithRecordId(recordId),
      ...transformerOptions,
    })

    return transformer.transform(this.rootBlock)
  }
}

export const docx: Docx = new Docx()
