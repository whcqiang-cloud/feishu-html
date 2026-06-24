import fs from 'node:fs/promises'
import os from 'node:os'
import path from 'node:path'
import { fileURLToPath } from 'node:url'
import { chromium, expect, test } from '@playwright/test'
import { resolveLiveCopyConfig } from '../src/env.js'
import {
  type Settings,
  SettingKey,
  Grid,
} from '../../chrome-extension/src/common/settings.js'

const dirname = path.dirname(fileURLToPath(import.meta.url))
const workspaceDir = path.resolve(dirname, '..')
const extensionPath = path.join(workspaceDir, '.cache/extension')
const { userDataDir: fixedUserDataDir, headless } = resolveLiveCopyConfig()

interface DownloadHtmlCase {
  name: string
  url: string
  expectedText: string
  match: 'equals' | 'contains'
  settings?: Partial<Settings>
}

const downloadHtmlCases: DownloadHtmlCase[] = [
  {
    name: 'basic-content',
    url: 'https://my.feishu.cn/wiki/Ez2WwNvB2iMjd9kXMw3cfbqDnTe',
    expectedText: '<!DOCTYPE html>',
    match: 'contains',
  },
  {
    name: 'Grid & Synced Reference',
    url: 'https://my.feishu.cn/docx/NG8AdUZq4ogKvox4fAXcoztnnke',
    match: 'contains',
    expectedText: '<table><colgroup>',
    settings: {
      [SettingKey.Grid]: Grid.ToHTML,
    },
  },
]

const createUserDataDir = async (): Promise<{
  path: string
  shouldCleanup: boolean
}> => {
  if (fixedUserDataDir) {
    await fs.mkdir(fixedUserDataDir, { recursive: true })
    return {
      path: fixedUserDataDir,
      shouldCleanup: false,
    }
  }

  return {
    path: await fs.mkdtemp(path.join(os.tmpdir(), 'cdc-extension-e2e-html-')),
    shouldCleanup: true,
  }
}

for (const downloadCase of downloadHtmlCases) {
  test(`@live download html contains expected text [${downloadCase.name}]`, async () => {
    const userDataDir = await createUserDataDir()
    const context = await chromium.launchPersistentContext(userDataDir.path, {
      channel: 'chromium',
      headless,
      args: [
        '--disable-popup-blocking',
        `--disable-extensions-except=${extensionPath}`,
        `--load-extension=${extensionPath}`,
      ],
    })

    try {
      await context.grantPermissions(['clipboard-read', 'clipboard-write'], {
        origin: new URL(downloadCase.url).origin,
      })

      let serviceWorker = context.serviceWorkers().at(0)
      serviceWorker ??= await context.waitForEvent('serviceworker')
      expect(serviceWorker.url()).toContain('chrome-extension://')

      if (downloadCase.settings) {
        await serviceWorker.evaluate(async settings => {
          // @ts-expect-error chrome is not typed in e2e test environment
          // eslint-disable-next-line @typescript-eslint/no-unsafe-member-access, @typescript-eslint/no-unsafe-call
          await chrome.storage.sync.set(settings)
        }, downloadCase.settings)
      }

      const page = await context.newPage()
      await page.goto(downloadCase.url, {
        waitUntil: 'domcontentloaded',
      })
      await page.bringToFront()

      const downloadHtmlButton = page.locator(
        '[data-CDC-button-type="download-html"]',
      )
      await expect(downloadHtmlButton).toBeVisible({
        timeout: 2 * 60 * 1000,
      })

      await downloadHtmlButton.click()

      // Wait for download to complete
      const downloadEvent = await page.waitForEvent('download')
      const downloadPath = await downloadEvent.path()

      expect(downloadPath).toBeTruthy()

      const content = await fs.readFile(downloadPath, 'utf-8')

      if (downloadCase.match === 'equals') {
        expect(content).toBe(downloadCase.expectedText)
      } else {
        expect(content).toContain(downloadCase.expectedText)
      }
    } finally {
      await context.close()
      if (userDataDir.shouldCleanup) {
        await fs.rm(userDataDir.path, { recursive: true, force: true })
      }
    }
  })
}
