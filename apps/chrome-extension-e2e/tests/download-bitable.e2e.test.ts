import { test } from '@playwright/test'

// URL for standalone Bitable page
// Format: https://nio.feishu.cn/base/{token}?table={tableId}&view={viewId}
const bitableCases = [
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

for (const bitableCase of bitableCases) {
  // eslint-disable-next-line @typescript-eslint/no-empty-function
  test.skip(`@live download standalone bitable [${bitableCase.name}] - requires manual URL`, () => {})
}
