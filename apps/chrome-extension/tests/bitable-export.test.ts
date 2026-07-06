import { describe, expect, it, beforeEach } from 'vitest'
import {
  bitableToHtml,
  extractStandaloneBitableTableFromWebApi,
  isStandaloneBitablePage,
} from '../src/scripts/bitable-export'

describe('standalone bitable export', () => {
  beforeEach(() => {
    sessionStorage.clear()
    const happyDOM = (
      window as unknown as { happyDOM: { setURL: (url: string) => void } }
    ).happyDOM
    happyDOM.setURL('http://localhost:3000/base/baseToken?table=tbl1&view=v1')
  })

  it('wraps exported HTML in the bitable container styles', () => {
    const html = bitableToHtml({
      title: 'Validation Rules',
      rows: [
        ['Use Case', 'Rule'],
        ['CS件/DB件', '遍历SV'],
      ],
    })

    expect(html).toContain('<figure class="bitable">')
    expect(html).toContain('<div class="bitable-wrapper">')
    expect(html).toContain('<figcaption>Validation Rules</figcaption>')
  })

  it('keeps fields that are present in records but missing from the current view field list', async () => {
    expect(window.location.pathname).toBe('/base/baseToken')
    expect(isStandaloneBitablePage()).toBe(true)

    sessionStorage.setItem(
      '__feishu_html_bitable_clientvars__',
      JSON.stringify({
        data: {
          base: JSON.stringify({ name: 'Validation Rules' }),
          table: JSON.stringify({
            fieldMap: {
              f1: { name: 'Use Case' },
              f2: { name: 'Problem' },
              f3: { name: 'Hidden Field' },
            },
            viewMap: {
              v1: {
                property: {
                  fields: ['f1'],
                  colInfos: {
                    f3: { hidden: true },
                  },
                },
              },
            },
            currentView: 'v1',
            recordMap: {
              r1: {
                fields: {
                  f1: 'CS件/DB件',
                  f2: 'right side content',
                  f3: 'hidden content',
                },
              },
            },
          }),
        },
      }),
    )

    const table = await extractStandaloneBitableTableFromWebApi()

    expect(table?.title).toBe('Validation Rules')
    expect(table?.rows).toEqual([
      ['Use Case', 'Problem'],
      ['CS件/DB件', 'right side content'],
    ])
  })
})
