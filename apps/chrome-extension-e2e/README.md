# Chrome Extension E2E (Playwright)

This workspace provides a debug-friendly E2E environment for
`@dolphin/chrome-extension` by loading the unpacked extension into Chromium via
Playwright.

## What it covers

- Builds the extension before tests (`global-setup.ts`).
- Copies build output into `.cache/extension`.
- Runs live Feishu page + extension in Chromium and validates copy markdown.

## Commands

From repository root:

```bash
pnpm run test:e2e:install
pnpm run test:e2e
pnpm run test:e2e:debug
```

Run live assertions:

```bash
pnpm run test:e2e --grep @live
```

## Built-in live cases

- `https://my.feishu.cn/wiki/Ez2WwNvB2iMjd9kXMw3cfbqDnTe` => equals `源内容`
- `https://my.feishu.cn/wiki/Pi5ww1AdKilUGrkyfgrc791unQ8` => equals `源内容` (同步块引用内容)
- `https://my.feishu.cn/wiki/X9tGwEQHgiodeqkIVSmcwqJynOh` => contains `<u>下划线样式</u>`
- HTML download: `https://my.feishu.cn/wiki/Ez2WwNvB2iMjd9kXMw3cfbqDnTe` => contains `<!DOCTYPE html>`
- HTML download: `https://my.feishu.cn/docx/NG8AdUZq4ogKvox4fAXcoztnnke` => contains `<table><colgroup>` (Grid=ToHTML)

Run in headed mode for local debugging:

```bash
pnpm run test:e2e:debug --grep @live
```

## Optional environment variables

- `CDC_E2E_SKIP_BUILD=1`: skip rebuilding extension (uses existing `dist`).
- `CDC_E2E_HEADLESS`: browser mode switch for `test:e2e` (`1` for headless, `0` for headed).  
  Default: `1`
- `CDC_E2E_USER_DATA_DIR`: persistent Chromium profile for keeping login.
  Default: `apps/chrome-extension-e2e/.cache/user-data-live`

## Notes

- The current spec focuses on live copy markdown and HTML download validation against Feishu.
- URL and expected content are defined in:
  - `tests/live-copy-markdown.e2e.test.ts` — copy markdown
  - `tests/download-html.e2e.test.ts` — download HTML
  - `tests/download-bitable.e2e.test.ts` — standalone Bitable (placeholder, requires real URL)
