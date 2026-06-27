import { afterEach, describe, expect, test } from 'vitest'
import { Doc } from '../src/doc'
import { Docx } from '../src/docx'

afterEach(() => {
  document.body.innerHTML = ''
})

describe('Doc legacy parser', () => {
  test('parses the provided snapshot container instead of the live editor body', () => {
    const liveEditorBody = document.createElement('div')
    liveEditorBody.className = 'innerdocbody'
    liveEditorBody.innerHTML =
      '<div class="ace-line lineguid-live">stale live line</div>'
    document.body.appendChild(liveEditorBody)

    const snapshot = document.createElement('div')
    snapshot.className = 'innerdocbody cdc-legacy-snapshot'
    snapshot.dataset['cdcLegacySnapshot'] = 'true'
    snapshot.innerHTML = [
      '<div class="ace-line lineguid-top">top collected line</div>',
      '<div class="ace-line lineguid-bottom">bottom collected line</div>',
    ].join('')

    const legacyDoc = new Doc()

    expect(legacyDoc.init({ container: snapshot })).toBe(true)
    const markdownText = legacyDoc
      .intoMarkdownAST()
      .root.children.map(node => JSON.stringify(node))
      .join('\n')

    expect(markdownText).toContain('top collected line')
    expect(markdownText).toContain('bottom collected line')
    expect(markdownText).not.toContain('stale live line')
  })

  test('converts compact zero-width separated legacy table lines to mdast tables', () => {
    const snapshot = document.createElement('div')
    snapshot.className = 'innerdocbody cdc-legacy-snapshot'
    snapshot.dataset['cdcLegacySnapshot'] = 'true'
    snapshot.innerHTML = [
      '<div class="ace-line lineguid-table wrapper">时间\u200B变更内容\u200B变更人\u200B2022.03.20\u200B初版编写\u200B高雪林</div>',
      '<div class="ace-line lineguid-table-text-1">时间\n变更内容\n变更人\n2022.03.20\n初版编写\n高雪林</div>',
      '<div class="ace-line lineguid-table-text-2">时间\n变更内容\n变更人\n2022.03.20\n初版编写\n高雪林</div>',
    ].join('')

    const legacyDoc = new Doc()

    expect(legacyDoc.init({ container: snapshot })).toBe(true)
    const rootChildren = legacyDoc.intoMarkdownAST().root.children
    const firstNode = rootChildren[0]

    expect(firstNode?.type).toBe('table')
    expect(JSON.stringify(firstNode)).toContain('2022.03.20')
    expect(rootChildren).toHaveLength(1)
  })

  test('prefers real lazy image sources over transparent placeholders', async () => {
    const realSrc = 'https://example.com/real.png'
    const placeholder =
      'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADElEQVR42mP8z8AARQAFFQH+Af9UwwAAAABJRU5ErkJggg=='
    const snapshot = document.createElement('div')
    snapshot.className = 'innerdocbody cdc-legacy-snapshot'
    snapshot.dataset['cdcLegacySnapshot'] = 'true'
    snapshot.innerHTML = `<div class="ace-line lineguid-image"><img src="${placeholder}" data-src="${realSrc}" alt="real image" /></div>`

    const legacyDoc = new Doc()

    expect(legacyDoc.init({ container: snapshot })).toBe(true)
    const { images } = legacyDoc.intoMarkdownAST()
    const sources = await images[0]?.data?.fetchSources?.()

    expect(sources?.src).toBe(realSrc)
    expect(images[0]?.url).toBe(realSrc)
  })

  test('does not export Feishu loading spinner svg as a drawing image', () => {
    const snapshot = document.createElement('div')
    snapshot.className = 'innerdocbody cdc-legacy-snapshot'
    snapshot.dataset['cdcLegacySnapshot'] = 'true'
    snapshot.innerHTML = [
      '<div class="ace-line lineguid-spinner">',
      '选择对应的Model显示对应的plant信息',
      '<svg viewBox="25 25 50 50" style="animation: doc-image-loading-rotate-animation;"><circle cx="50" cy="50" r="20"/></svg>',
      '</div>',
    ].join('')

    const legacyDoc = new Doc()

    expect(legacyDoc.init({ container: snapshot })).toBe(true)
    const { root, images } = legacyDoc.intoMarkdownAST()

    expect(images).toHaveLength(0)
    expect(JSON.stringify(root)).toContain('选择对应的Model显示对应的plant信息')
  })

  test('exports meaningful embedded svg drawings', async () => {
    const snapshot = document.createElement('div')
    snapshot.className = 'innerdocbody cdc-legacy-snapshot'
    snapshot.dataset['cdcLegacySnapshot'] = 'true'
    snapshot.innerHTML = [
      '<div class="ace-line lineguid-drawing">',
      '<svg width="120" height="100" viewBox="0 0 120 100">',
      '<rect x="10" y="10" width="40" height="20"></rect>',
      '<rect x="70" y="10" width="40" height="20"></rect>',
      '<path d="M50 20 L70 20"></path>',
      '<text x="10" y="80">Model</text>',
      '</svg>',
      '</div>',
    ].join('')

    const legacyDoc = new Doc()

    expect(legacyDoc.init({ container: snapshot })).toBe(true)
    const { images } = legacyDoc.intoMarkdownAST()
    const sources = await images[0]?.data?.fetchSources?.()

    expect(images).toHaveLength(1)
    expect(sources?.src).toContain('data:image/svg+xml')
    expect(decodeURIComponent(sources?.src ?? '')).toContain('Model')
  })

  test('keeps snapshot lines that only contain legacy svg drawings', async () => {
    const snapshot = document.createElement('div')
    snapshot.className = 'innerdocbody cdc-legacy-snapshot'
    snapshot.dataset['cdcLegacySnapshot'] = 'true'
    snapshot.innerHTML = [
      '<div class="ace-line lineguid-pure-drawing">',
      '<svg width="160" height="120" viewBox="0 0 160 120">',
      '<rect x="10" y="10" width="50" height="30"></rect>',
      '<rect x="100" y="10" width="50" height="30"></rect>',
      '<path d="M60 25 L100 25"></path>',
      '<path d="M60 30 L100 35"></path>',
      '<path d="M60 35 L100 45"></path>',
      '<circle cx="80" cy="85" r="20"></circle>',
      '</svg>',
      '</div>',
    ].join('')

    const legacyDoc = new Doc()

    expect(legacyDoc.init({ container: snapshot })).toBe(true)
    const { images } = legacyDoc.intoMarkdownAST()
    const sources = await images[0]?.data?.fetchSources?.()

    expect(images).toHaveLength(1)
    expect(sources?.src).toContain('data:image/svg+xml')
  })

  test('exports compact text-only legacy svg drawings as images', async () => {
    const snapshot = document.createElement('div')
    snapshot.className = 'innerdocbody cdc-legacy-snapshot'
    snapshot.dataset['cdcLegacySnapshot'] = 'true'
    snapshot.innerHTML = [
      '<div class="ace-line lineguid-compact-drawing">',
      '<svg>',
      '<text>Model</text>',
      '<text>UPS</text>',
      '<text>Plant</text>',
      '<line x1="0" y1="0" x2="20" y2="20"></line>',
      '</svg>',
      '</div>',
    ].join('')

    const legacyDoc = new Doc()

    expect(legacyDoc.init({ container: snapshot })).toBe(true)
    const { root, images } = legacyDoc.intoMarkdownAST()
    const firstChild = root.children[0]
    const sources = await images[0]?.data?.fetchSources?.()

    expect(images).toHaveLength(1)
    expect(firstChild?.type).toBe('paragraph')
    expect(
      firstChild?.type === 'paragraph' && firstChild.children[0]?.type,
    ).toBe('image')
    expect(sources?.src).toContain('data:image/svg+xml')
  })

  test('parses collected non-ace legacy line nodes in snapshot mode', () => {
    const snapshot = document.createElement('div')
    snapshot.className = 'innerdocbody cdc-legacy-snapshot'
    snapshot.dataset['cdcLegacySnapshot'] = 'true'
    snapshot.innerHTML = [
      '<div class="ace-line lineguid-heading">4.20 更新</div>',
      '<div class="legacy-list-line">1. Model state：Active / InActive</div>',
      '<div class="legacy-list-line">Plant state Active / InActive</div>',
      '<div class="ace-line lineguid-next">更新内容</div>',
    ].join('')

    const legacyDoc = new Doc()

    expect(legacyDoc.init({ container: snapshot })).toBe(true)
    const markdownText = JSON.stringify(legacyDoc.intoMarkdownAST().root)

    expect(markdownText).toContain('Model state')
    expect(markdownText).toContain('Plant state')
    expect(markdownText).toContain('更新内容')
  })

  test('deduplicates repeated long legacy lines collected from virtual scroll overlap', () => {
    const snapshot = document.createElement('div')
    snapshot.className = 'innerdocbody cdc-legacy-snapshot'
    snapshot.dataset['cdcLegacySnapshot'] = 'true'
    snapshot.innerHTML = [
      '<div class="ace-line heading-h2 lineguid-update">4.20 更新</div>',
      '<div class="ace-line lineguid-model-1">Model state : Active/InActive 630 创建model默认为Active，不做Active到InActive逻辑设置</div>',
      '<div class="ace-line lineguid-active-1">Active 表示 Model可以被各功能模块调用、记录数据</div>',
      '<div class="ace-line lineguid-model-2">Model state : Active/InActive 630 创建model默认为Active，不做Active到InActive逻辑设置</div>',
      '<div class="ace-line lineguid-active-2">Active 表示 Model可以被各功能模块调用、记录数据</div>',
      '<div class="ace-line lineguid-next">Plant state Active/InActive</div>',
    ].join('')

    const legacyDoc = new Doc()

    expect(legacyDoc.init({ container: snapshot })).toBe(true)
    const markdown = Docx.stringify(legacyDoc.intoMarkdownAST().root)

    expect(
      markdown.match(/Model state : Active\/InActive/g) ?? [],
    ).toHaveLength(1)
    expect(
      markdown.match(/Active 表示 Model可以被各功能模块调用/g) ?? [],
    ).toHaveLength(1)
    expect(markdown).toContain('Plant state Active/InActive')
  })

  test('keeps short repeated labels while removing overlap noise', () => {
    const snapshot = document.createElement('div')
    snapshot.className = 'innerdocbody cdc-legacy-snapshot'
    snapshot.dataset['cdcLegacySnapshot'] = 'true'
    snapshot.innerHTML = [
      '<div class="ace-line lineguid-label-1">Active</div>',
      '<div class="ace-line lineguid-label-2">Active</div>',
      '<div class="ace-line lineguid-long-1">Active 表示该工厂可用</div>',
      '<div class="ace-line lineguid-long-2">Active 表示该工厂可用</div>',
    ].join('')

    const legacyDoc = new Doc()

    expect(legacyDoc.init({ container: snapshot })).toBe(true)
    const markdown = Docx.stringify(legacyDoc.intoMarkdownAST().root)

    expect(markdown.match(/^Active$/gm) ?? []).toHaveLength(2)
    expect(markdown.match(/Active 表示该工厂可用/g) ?? []).toHaveLength(1)
  })

  test('keeps repeated long lines when they belong to different legacy blocks', () => {
    const snapshot = document.createElement('div')
    snapshot.className = 'innerdocbody cdc-legacy-snapshot'
    snapshot.dataset['cdcLegacySnapshot'] = 'true'
    snapshot.innerHTML = [
      '<div class="ace-line lineguid-first">Active 表示该工厂可用</div>',
      '<div class="ace-line lineguid-image"><img src="https://example.com/image.png" alt="截图" /></div>',
      '<div class="ace-line lineguid-second">Active 表示该工厂可用</div>',
    ].join('')

    const legacyDoc = new Doc()

    expect(legacyDoc.init({ container: snapshot })).toBe(true)
    const markdown = Docx.stringify(legacyDoc.intoMarkdownAST().root)

    expect(markdown.match(/Active 表示该工厂可用/g) ?? []).toHaveLength(2)
  })
})
