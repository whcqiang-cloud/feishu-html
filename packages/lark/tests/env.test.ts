import { afterEach, describe, expect, test } from 'vitest'
import { doc } from '../src/doc'
import { docx } from '../src/docx'

const setWindowValue = (key: 'editor' | 'PageMain', value: unknown) => {
  Object.defineProperty(window, key, {
    configurable: true,
    value,
  })
}

afterEach(() => {
  delete (window as Partial<Window> & { editor?: unknown }).editor
  delete (window as Partial<Window> & { PageMain?: unknown }).PageMain
})

describe('Feishu document type detection', () => {
  test('keeps docx routing when PageMain and editor both exist', () => {
    setWindowValue('editor', {})
    setWindowValue('PageMain', {
      blockManager: {
        rootBlockModel: {},
      },
    })

    expect(docx.isDocx).toBe(true)
    expect(docx.isDoc).toBe(false)
    expect(doc.isDoc).toBe(false)
  })

  test('detects legacy docs only when editor exists without PageMain', () => {
    setWindowValue('editor', {})

    expect(docx.isDocx).toBe(false)
    expect(docx.isDoc).toBe(true)
    expect(doc.isDoc).toBe(true)
  })
})
