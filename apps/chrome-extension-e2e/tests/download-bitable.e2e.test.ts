import fs from 'node:fs/promises'
import os from 'node:os'
import path from 'node:path'
import { fileURLToPath } from 'node:url'
import { chromium, expect, test } from '@playwright/test'
import { resolveLiveCopyConfig } from '../src/env.js'

const dirname = path.dirname(fileURLToPath(import.meta.url))
const workspaceDir = path.resolve(dirname, '..')
const extensionPath = path.join(workspaceDir, '.cache/extension')
const { userDataDir: fixedUserDataDir, headless } = resolveLiveCopyConfig()

interface BitableDownloadCase {
  name: string
  url: string
  expectedText: string
  match: 'contains'
}

// URL for standalone Bitable page
// Format: https://nio.feishu.cn/base/{token}?table={tableId}&view={viewId}
const bitableCases: BitableDownloadCase[] = [
  {
    name: 'standalone-bitable-html',
    url: '', // TODO: fill with actual Bitable URL
    expectedText: '<!DOCTYPE html>',
    match: 'contains',
  },
  {
    name: 'standalone-bitable-md',
    url: '', // TODO: fill with actual Bitable URL
    expectedText: '| ',
    match: 'contains',
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
    path: await fs.mkdtemp(
      path.join(os.tmpdir(), 'cdc-extension-e2e-bitable-'),
    ),
    shouldCleanup: true,
  }
}

for (const bitableCase of bitableCases) {
  test.skip(
    `@live download standalone bitable [${bitableCase.name}] - requires manual URL`,
    () => {},
  )
}
