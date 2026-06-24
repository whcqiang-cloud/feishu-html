export {}

const CLIENTVARS_CACHE_KEY = '__feishu_html_bitable_clientvars__'
const originalFetch: typeof window.fetch = window.fetch.bind(window)

function isClientvarsUrl(input: string | undefined): boolean {
  return Boolean(
    input?.includes('/space/api/v1/bitable/') && input.includes('/clientvars'),
  )
}

function cacheClientvars(url: string | undefined, text: string): void {
  if (!isClientvarsUrl(url)) return

  try {
    const json = JSON.parse(text) as { data?: { base?: unknown; table?: unknown } }
    if (json.data?.base || json.data?.table) {
      sessionStorage.setItem(CLIENTVARS_CACHE_KEY, text)
    }
  } catch {
    // Ignore non-JSON responses from unrelated requests.
  }
}

window.fetch = async (...args) => {
  const input = args[0]
  const url =
    typeof input === 'string'
      ? input
      : input instanceof Request
        ? input.url
        : input.toString()
  const response = await originalFetch(...args)

  if (isClientvarsUrl(url)) {
    response
      .clone()
      .text()
      .then(text => {
        cacheClientvars(url, text)
      })
      .catch(() => {
        // Ignore cache failures.
      })
  }

  return response
}

const OriginalXHR: typeof window.XMLHttpRequest = window.XMLHttpRequest
window.XMLHttpRequest = function XMLHttpRequestClientvarsCacheProxy() {
  const xhr = new OriginalXHR()
  let url = ''
  const originalOpen = xhr.open.bind(xhr)

  xhr.open = function openClientvarsCacheProxy(
    openMethod: string,
    openUrl: string | URL,
    async?: boolean,
    username?: string | null,
    password?: string | null,
  ): void {
    url = String(openUrl)
    if (async === undefined) {
      ;(originalOpen as (...args: unknown[]) => void)(openMethod, openUrl)
      return
    }
    ;(originalOpen as (...args: unknown[]) => void)(
      openMethod,
      openUrl,
      async,
      username,
      password,
    )
  }

  xhr.addEventListener('loadend', () => {
    if (isClientvarsUrl(url)) {
      cacheClientvars(url, xhr.responseText || '')
    }
  })

  return xhr
} as unknown as typeof XMLHttpRequest
