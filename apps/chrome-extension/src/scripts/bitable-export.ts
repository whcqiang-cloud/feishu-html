export interface BitableTable {
  title: string
  rows: string[][]
}

const BITABLE_URL_PATTERN = /\/(base|bitable)\//i

export function isStandaloneBitablePage(): boolean {
  return BITABLE_URL_PATTERN.test(window.location.href)
}

function normalizeText(value: string): string {
  return value.replace(/\s+/g, ' ').trim()
}

function uniqueRows(rows: string[][]): string[][] {
  const seen = new Set<string>()
  return rows.filter(row => {
    const key = row.join('\u0000')
    if (seen.has(key)) return false
    seen.add(key)
    return row.some(Boolean)
  })
}

function extractFromNativeTable(table: HTMLTableElement): string[][] {
  return uniqueRows(
    Array.from(table.querySelectorAll('tr')).map(row =>
      Array.from(row.querySelectorAll('th,td')).map(cell =>
        normalizeText(cell.textContent),
      ),
    ),
  )
}

function extractFromRoleGrid(grid: HTMLElement): string[][] {
  const rows = Array.from(
    grid.querySelectorAll<HTMLElement>('[role="row"]'),
  ).map(row =>
    Array.from(
      row.querySelectorAll<HTMLElement>(
        '[role="columnheader"], [role="rowheader"], [role="gridcell"], [role="cell"]',
      ),
    ).map(cell => normalizeText(cell.innerText || cell.textContent || '')),
  )

  return uniqueRows(rows)
}

function getPageTitle(): string {
  const title =
    document.querySelector<HTMLElement>(
      '[data-testid*="title"], [class*="title"], [class*="Title"]',
    )?.innerText ?? document.title
  const normalizedTitle = normalizeText(title.replace(/[-|].*$/, ''))

  return normalizedTitle.length > 0 ? normalizedTitle : 'bitable'
}

export function extractStandaloneBitableTable(): BitableTable | null {
  const table = document.querySelector<HTMLTableElement>('table')
  const grid = document.querySelector<HTMLElement>(
    '[role="grid"], [role="table"]',
  )
  const rows = table
    ? extractFromNativeTable(table)
    : extractFromRoleGrid(grid ?? document.body)

  if (rows.length === 0) return null

  return {
    title: getPageTitle(),
    rows,
  }
}

interface BitableClientVars {
  base?: string
  table?: string
}

interface BitableField {
  id?: string
  name?: string
  property?: {
    options?: { id?: string; name?: string }[]
  }
}

interface BitableTableData {
  fieldMap?: Record<string, BitableField>
  viewMap?: Record<
    string,
    {
      property?: {
        fields?: string[]
        colInfos?: Record<string, { hidden?: boolean }>
      }
    }
  >
  currentView?: string
  groupList?: { recordIDList?: string[] }[]
  recordMap?: Record<string, unknown>
}

function parseStandaloneBitableUrl(): {
  token: string
  tableId: string
  viewId: string
} | null {
  const token = /\/(?:base|bitable)\/([^/?#]+)/i.exec(
    window.location.pathname,
  )?.[1]
  const params = new URLSearchParams(window.location.search)
  const tableId = params.get('table')
  const viewId = params.get('view')

  if (!token || !tableId || !viewId) return null

  return { token, tableId, viewId }
}

async function decodeGzipBase64(value: string): Promise<string> {
  const binary = atob(value)
  const bytes = Uint8Array.from(binary, char => char.charCodeAt(0))
  const stream = new Blob([bytes])
    .stream()
    .pipeThrough(new DecompressionStream('gzip'))

  return new Response(stream).text()
}

async function decodeClientVarsField<T>(
  value: string | undefined,
): Promise<T | null> {
  if (!value) return null
  const text = value.startsWith('H4sI') ? await decodeGzipBase64(value) : value

  return JSON.parse(text) as T
}

function optionMapForField(
  field: BitableField | undefined,
): Map<string, string> {
  return new Map(
    field?.property?.options?.map(option => [
      option.id ?? '',
      option.name ?? '',
    ]) ?? [],
  )
}

function stringifyCellValue(
  value: unknown,
  options: Map<string, string>,
): string {
  if (value == null) return ''
  if (typeof value === 'string') return options.get(value) ?? value
  if (typeof value === 'number' || typeof value === 'boolean')
    return String(value)
  if (Array.isArray(value)) {
    return value
      .map(item => stringifyCellValue(item, options))
      .filter(Boolean)
      .join(', ')
  }
  if (typeof value === 'object') {
    const record = value as Record<string, unknown>
    const preferred =
      record['text'] ??
      record['name'] ??
      record['value'] ??
      record['title'] ??
      record['id']
    if (preferred !== undefined && preferred !== value) {
      return stringifyCellValue(preferred, options)
    }

    return Object.values(record)
      .map(item => stringifyCellValue(item, options))
      .filter(Boolean)
      .join(', ')
  }

  return ''
}

function getRecordFields(record: unknown): Record<string, unknown> {
  if (!record || typeof record !== 'object') return {}
  const value = record as Record<string, unknown>

  return (
    (value['fields'] as Record<string, unknown> | undefined) ??
    (value['fieldValues'] as Record<string, unknown> | undefined) ??
    (value['cellValues'] as Record<string, unknown> | undefined) ??
    value
  )
}

function tableDataToRows(table: BitableTableData): string[][] {
  const view = table.viewMap?.[table.currentView ?? '']
  const allFieldIds =
    view?.property?.fields ?? Object.keys(table.fieldMap ?? {})
  const fieldIds = allFieldIds.filter(
    fieldId => !view?.property?.colInfos?.[fieldId]?.hidden,
  )
  const recordIds =
    table.groupList?.flatMap(group => group.recordIDList ?? []) ??
    Object.keys(table.recordMap ?? {})

  const header = fieldIds.map(
    fieldId => table.fieldMap?.[fieldId]?.name ?? fieldId,
  )
  const body = recordIds.map(recordId => {
    const fields = getRecordFields(table.recordMap?.[recordId])

    return fieldIds.map(fieldId =>
      stringifyCellValue(
        fields[fieldId],
        optionMapForField(table.fieldMap?.[fieldId]),
      ),
    )
  })

  return uniqueRows([header, ...body])
}

async function fetchClientVars(parsed: {
  token: string
  tableId: string
  viewId: string
}): Promise<{ data?: BitableClientVars } | null> {
  const cached = sessionStorage.getItem('__feishu_html_bitable_clientvars__')
  if (cached) {
    return JSON.parse(cached) as { data?: BitableClientVars }
  }

  const { token, tableId, viewId } = parsed
  const response = await fetch(
    `/space/api/v1/bitable/${token}/clientvars?tableID=${tableId}&viewID=${viewId}&recordLimit=5000&ondemandLimit=5000&needBase=true&viewLazyLoad=true&ondemandVer=2&openType=0&noMissCS=true&optimizationFlag=1&removeFmlExtra=true`,
    { credentials: 'include' },
  )
  if (!response.ok) return null

  return (await response.json()) as { data?: BitableClientVars }
}

export async function extractStandaloneBitableTableFromWebApi(): Promise<BitableTable | null> {
  const parsed = parseStandaloneBitableUrl()
  if (!parsed) return null

  const json = await fetchClientVars(parsed)
  if (!json) return null
  const base = await decodeClientVarsField<{ name?: string }>(json.data?.base)
  const table = await decodeClientVarsField<BitableTableData>(json.data?.table)
  if (!table) return null

  const rows = tableDataToRows(table)
  if (rows.length <= 1) return null

  return {
    title: base?.name ?? getPageTitle(),
    rows,
  }
}

function escapeHtml(value: string): string {
  return value
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
}

function normalizeColumnCount(rows: string[][]): string[][] {
  const colCount = Math.max(...rows.map(row => row.length))
  return rows.map(row => [
    ...row,
    ...Array.from({ length: colCount - row.length }, () => ''),
  ])
}

export function bitableToHtml(table: BitableTable): string {
  const rows = normalizeColumnCount(table.rows)
  const [header = [], ...bodyRows] = rows

  return `<table><thead><tr>${header
    .map(cell => `<th>${escapeHtml(cell)}</th>`)
    .join('')}</tr></thead><tbody>${bodyRows
    .map(
      row =>
        `<tr>${row.map(cell => `<td>${escapeHtml(cell)}</td>`).join('')}</tr>`,
    )
    .join('')}</tbody></table>`
}

function escapeMarkdownCell(value: string): string {
  return value.replace(/\|/g, '\\|').replace(/\n/g, '<br>')
}

export function bitableToMarkdown(table: BitableTable): string {
  const rows = normalizeColumnCount(table.rows)
  const [header = [], ...bodyRows] = rows
  if (header.length === 0) return ''

  const headerLine = `| ${header.map(escapeMarkdownCell).join(' | ')} |`
  const separatorLine = `| ${header.map(() => '---').join(' | ')} |`
  const bodyLines = bodyRows.map(
    row => `| ${row.map(escapeMarkdownCell).join(' | ')} |`,
  )

  return [`# ${table.title}`, '', headerLine, separatorLine, ...bodyLines].join(
    '\n',
  )
}
