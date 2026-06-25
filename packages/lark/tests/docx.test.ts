import { test, describe, expect, it, afterEach } from 'vitest'
import type * as mdast from 'mdast'
import {
  BlockType,
  Transformer,
  mergeListItems,
  mergePhrasingContents,
  transformOperationsToPhrasingContents,
  type PageBlock,
} from '../src/docx'

const transformer = new Transformer()

afterEach(() => {
  document.body.innerHTML = ''
})

describe('mergeListItems()', () => {
  test('simple example', () => {
    const result = mergeListItems([
      {
        type: 'blockquote',
        children: [],
      },
      {
        type: 'listItem',
        children: [],
      },
      {
        type: 'listItem',
        data: {
          seq: 1,
        },
        children: [],
      },
      {
        type: 'listItem',
        checked: true,
        children: [],
      },
      {
        type: 'listItem',
        checked: false,
        children: [],
      },
      {
        type: 'blockquote',
        children: [],
      },
      {
        type: 'listItem',
        data: {
          seq: 2,
        },
        children: [],
      },
    ])
    const expectedResult: mdast.Nodes[] = [
      {
        type: 'blockquote',
        children: [],
      },
      {
        type: 'list',
        children: [
          {
            type: 'listItem',
            children: [],
          },
        ],
      },
      {
        type: 'list',
        ordered: true,
        start: 1,
        children: [
          {
            type: 'listItem',
            data: {
              seq: 1,
            },
            children: [],
          },
        ],
      },
      {
        type: 'list',
        children: [
          {
            type: 'listItem',
            checked: true,
            children: [],
          },
          {
            type: 'listItem',
            checked: false,
            children: [],
          },
        ],
      },
      {
        type: 'blockquote',
        children: [],
      },
      {
        type: 'list',
        ordered: true,
        start: 2,
        children: [
          {
            type: 'listItem',
            data: {
              seq: 2,
            },
            children: [],
          },
        ],
      },
    ]
    expect(result).toStrictEqual(expectedResult)
  })
})

describe('mergePhrasingContents()', () => {
  test('simple example', () => {
    const result = mergePhrasingContents([
      {
        type: 'strong',
        children: [],
      },
      {
        type: 'emphasis',
        children: [
          {
            type: 'text',
            value: 'a',
          },
        ],
      },
      {
        type: 'emphasis',
        children: [
          {
            type: 'text',
            value: 'b',
          },
        ],
      },
      {
        type: 'link',
        url: 'https://www.baidu.com',
        children: [
          {
            type: 'delete',
            children: [
              {
                type: 'strong',
                children: [
                  {
                    type: 'text',
                    value: 'a',
                  },
                ],
              },
            ],
          },
        ],
      },
      {
        type: 'link',
        url: 'https://www.baidu.com',
        children: [
          {
            type: 'strong',
            children: [
              {
                type: 'text',
                value: 'a',
              },
            ],
          },
        ],
      },
    ])
    const expectedResult: mdast.PhrasingContent[] = [
      {
        type: 'strong',
        children: [],
      },
      {
        type: 'emphasis',
        children: [
          {
            type: 'text',
            value: 'ab',
          },
        ],
      },
      {
        type: 'link',
        url: 'https://www.baidu.com',
        children: [
          {
            type: 'delete',
            children: [
              {
                type: 'strong',
                children: [
                  {
                    type: 'text',
                    value: 'a',
                  },
                ],
              },
            ],
          },
          {
            type: 'strong',
            children: [
              {
                type: 'text',
                value: 'a',
              },
            ],
          },
        ],
      },
    ]
    expect(result).toStrictEqual(expectedResult)
  })
})

describe('transformOperationsToPhrasingContents()', () => {
  describe('code span', () => {
    test('simple code span', () => {
      expect(
        transformOperationsToPhrasingContents([
          {
            insert: 'code',
            attributes: {
              inlineCode: 'true',
              author: '7096007617544896513',
            },
          },
          {
            insert: '\n',
            attributes: {
              fixEnter: 'true',
            },
          },
        ]).contents,
      ).toStrictEqual([{ type: 'inlineCode', value: 'code' }])
    })

    test('code span in strong emphasis', () => {
      expect(
        transformOperationsToPhrasingContents([
          {
            insert: 'a',
            attributes: {
              bold: 'true',
              author: '7096007617544896513',
            },
          },
          {
            insert: 'b',
            attributes: {
              bold: 'true',
              inlineCode: 'true',
              author: '7096007617544896513',
            },
          },
          {
            insert: 'c',
            attributes: {
              bold: 'true',
              author: '7096007617544896513',
            },
          },
          {
            insert: '\n',
            attributes: {
              fixEnter: 'true',
            },
          },
        ]).contents,
      ).toStrictEqual([
        {
          type: 'strong',
          children: [
            { type: 'text', value: 'a' },
            { type: 'inlineCode', value: 'b' },
            { type: 'text', value: 'c' },
          ],
        },
      ])
    })

    test('code span range intersect strong emphasis range', () => {
      expect(
        transformOperationsToPhrasingContents([
          {
            attributes: {
              bold: 'true',
              author: '7096007617544896513',
            },
            insert: 'a',
          },
          {
            insert: 'b',
            attributes: {
              inlineCode: 'true',
              bold: 'true',
              author: '7096007617544896513',
            },
          },
          {
            insert: 'c',
            attributes: {
              inlineCode: 'true',
              author: '7096007617544896513',
            },
          },
          {
            insert: '\n',
            attributes: {
              fixEnter: 'true',
            },
          },
        ]).contents,
      ).toStrictEqual([
        {
          type: 'strong',
          children: [
            {
              type: 'text',
              value: 'a',
            },
            {
              type: 'inlineCode',
              value: 'b',
            },
          ],
        },
        {
          type: 'inlineCode',
          value: 'c',
        },
      ])
    })
  })

  describe('emphasis and strong emphasis', () => {
    test('simple emphasis', () => {
      expect(
        transformOperationsToPhrasingContents([
          {
            insert: 'emphasis',
            attributes: {
              italic: 'true',
              author: '7096007617544896513',
            },
          },
          {
            insert: '\n',
            attributes: {
              fixEnter: 'true',
            },
          },
        ]).contents,
      ).toStrictEqual([
        {
          type: 'emphasis',
          children: [{ type: 'text', value: 'emphasis' }],
        },
      ])
    })

    test('simple strong emphasis', () => {
      expect(
        transformOperationsToPhrasingContents([
          {
            insert: 'strong emphasis',
            attributes: {
              bold: 'true',
              author: '7096007617544896513',
            },
          },
          {
            insert: '\n',
            attributes: {
              fixEnter: 'true',
            },
          },
        ]).contents,
      ).toStrictEqual([
        {
          type: 'strong',
          children: [{ type: 'text', value: 'strong emphasis' }],
        },
      ])
    })

    test('emphasis in strong emphasis', () => {
      expect(
        transformOperationsToPhrasingContents([
          {
            attributes: {
              bold: 'true',
              author: '7096007617544896513',
            },
            insert: 'a',
          },
          {
            insert: 'b',
            attributes: {
              italic: 'true',
              bold: 'true',
              author: '7096007617544896513',
            },
          },
          {
            attributes: {
              bold: 'true',
              author: '7096007617544896513',
            },
            insert: 'c',
          },
          {
            insert: '\n',
            attributes: {
              fixEnter: 'true',
            },
          },
        ]).contents,
      ).toStrictEqual([
        {
          type: 'strong',
          children: [
            {
              type: 'text',
              value: 'a',
            },
            {
              type: 'emphasis',
              children: [{ type: 'text', value: 'b' }],
            },
            {
              type: 'text',
              value: 'c',
            },
          ],
        },
      ])
    })

    test('emphasis range intersect strong emphasis range', () => {
      expect(
        transformOperationsToPhrasingContents([
          {
            attributes: {
              bold: 'true',
              author: '7096007617544896513',
            },
            insert: 'a',
          },
          {
            insert: 'b',
            attributes: {
              italic: 'true',
              bold: 'true',
              author: '7096007617544896513',
            },
          },
          {
            insert: 'c',
            attributes: {
              italic: 'true',
              author: '7096007617544896513',
            },
          },
          {
            insert: '\n',
            attributes: {
              fixEnter: 'true',
            },
          },
        ]).contents,
      ).toStrictEqual([
        {
          type: 'strong',
          children: [
            { type: 'text', value: 'a' },
            { type: 'emphasis', children: [{ type: 'text', value: 'b' }] },
          ],
        },
        {
          type: 'emphasis',
          children: [
            {
              type: 'text',
              value: 'c',
            },
          ],
        },
      ])
    })
  })

  describe('delete', () => {
    test('simple delete', () => {
      expect(
        transformOperationsToPhrasingContents([
          {
            insert: 'a',
            attributes: {
              strikethrough: 'true',
              author: '7096007617544896513',
            },
          },
          {
            insert: '\n',
            attributes: {
              fixEnter: 'true',
            },
          },
        ]).contents,
      ).toStrictEqual([
        {
          type: 'delete',
          children: [{ type: 'text', value: 'a' }],
        },
      ])
    })

    test('nesting are possible', () => {
      expect(
        transformOperationsToPhrasingContents([
          {
            attributes: {
              bold: 'true',
              author: '7096007617544896513',
            },
            insert: 'a',
          },
          {
            attributes: {
              italic: 'true',
              bold: 'true',
              author: '7096007617544896513',
            },
            insert: 'b',
          },
          {
            insert: 'c',
            attributes: {
              strikethrough: 'true',
              italic: 'true',
              bold: 'true',
              author: '7096007617544896513',
            },
          },
          {
            attributes: {
              bold: 'true',
              author: '7096007617544896513',
            },
            insert: 'd',
          },
          {
            insert: '\n',
            attributes: {
              fixEnter: 'true',
            },
          },
        ]).contents,
      ).toStrictEqual([
        {
          type: 'strong',
          children: [
            {
              type: 'text',
              value: 'a',
            },
            {
              type: 'emphasis',
              children: [
                {
                  type: 'text',
                  value: 'b',
                },
                { type: 'delete', children: [{ type: 'text', value: 'c' }] },
              ],
            },
            {
              type: 'text',
              value: 'd',
            },
          ],
        },
      ])
    })
  })

  describe('link', () => {
    test('simple link', () => {
      expect(
        transformOperationsToPhrasingContents([
          {
            insert: 'a',
            attributes: {
              'clientside-link-underline': 'true',
              author: '7096007617544896513',
              link: 'https%3A%2F%2Fwww.baidu.com',
            },
          },
          {
            insert: '\n',
            attributes: {
              fixEnter: 'true',
            },
          },
        ]).contents,
      ).toStrictEqual([
        {
          type: 'link',
          url: 'https://www.baidu.com',
          children: [{ type: 'text', value: 'a' }],
        },
      ])
    })
  })

  describe('underline', () => {
    test('wrap underlined text with <u></u>', () => {
      expect(
        transformOperationsToPhrasingContents([
          {
            insert: 'underlined',
            attributes: {
              underline: 'true',
            },
          },
          {
            insert: '\n',
            attributes: {
              fixEnter: 'true',
            },
          },
        ]).contents,
      ).toStrictEqual([
        {
          type: 'html',
          value: '<u>underlined</u>',
        },
      ])
    })
  })

  describe('mark priority', () => {
    test('strong > emphasis', () => {
      expect(
        transformOperationsToPhrasingContents([
          {
            attributes: {
              bold: 'true',
              author: '7096007617544896513',
            },
            insert: 'ab',
          },
          {
            insert: 'c',
            attributes: {
              bold: 'true',
              italic: 'true',
              author: '7096007617544896513',
            },
          },
          {
            attributes: {
              bold: 'true',
              author: '7096007617544896513',
            },
            insert: 'de',
          },
          {
            insert: '\n',
            attributes: {
              fixEnter: 'true',
            },
          },
        ]).contents,
      ).toStrictEqual([
        {
          type: 'strong',
          children: [
            { type: 'text', value: 'ab' },
            {
              type: 'emphasis',
              children: [{ type: 'text', value: 'c' }],
            },
            { type: 'text', value: 'de' },
          ],
        },
      ])
    })
  })

  describe('attributes not defined', () => {
    test('text', () => {
      expect(
        transformOperationsToPhrasingContents([
          {
            insert: 'emphasis',
          },
          {
            insert: '\n',
          },
        ]).contents,
      ).toStrictEqual([{ type: 'text', value: 'emphasis' }])
    })
  })

  describe('mention user', () => {
    test('simple case', () => {
      expect(
        transformOperationsToPhrasingContents([
          {
            attributes: {
              'inline-component': JSON.stringify({
                type: 'user',
                data: { uid: 'a' },
              }),
            },
            insert: '',
          },
          {
            attributes: {
              'inline-component': JSON.stringify({
                type: 'user',
                data: { uid: 'b' },
              }),
            },
            insert: '',
          },
          {
            insert: '\n',
          },
        ]).mentionUsers,
      ).toStrictEqual([
        { type: 'inlineCode', value: '', data: { mentionUserId: 'a' } },
        { type: 'inlineCode', value: '', data: { mentionUserId: 'b' } },
      ])
    })
  })
})

describe('transformer.transform()', () => {
  describe('divider', () => {
    test('one divider', () => {
      expect(
        transformer.transform({
          type: BlockType.PAGE,
          snapshot: {
            type: BlockType.PAGE,
          },
          children: [
            {
              type: BlockType.DIVIDER,
              children: [],
              snapshot: {
                type: BlockType.DIVIDER,
              },
            },
          ],
        }).root,
      ).toStrictEqual({
        type: 'root',
        children: [
          {
            type: 'thematicBreak',
          },
        ],
      })
    })

    test('two divider', () => {
      expect(
        transformer.transform({
          type: BlockType.PAGE,
          snapshot: {
            type: BlockType.PAGE,
          },
          children: [
            {
              type: BlockType.DIVIDER,
              children: [],
              snapshot: {
                type: BlockType.DIVIDER,
              },
            },
            {
              type: BlockType.DIVIDER,
              children: [],
              snapshot: {
                type: BlockType.DIVIDER,
              },
            },
          ],
        }).root,
      ).toStrictEqual({
        type: 'root',
        children: [
          {
            type: 'thematicBreak',
          },
          {
            type: 'thematicBreak',
          },
        ],
      })
    })
  })

  describe('heading', () => {
    test('heading one', () => {
      expect(
        transformer.transform({
          type: BlockType.PAGE,
          snapshot: {
            type: BlockType.PAGE,
          },
          children: [
            {
              type: BlockType.HEADING1,
              depth: 1,
              snapshot: {
                type: BlockType.HEADING1,
              },
              zoneState: {
                allText: '',
                content: {
                  ops: [
                    {
                      insert: 'heading one',
                      attributes: {},
                    },
                  ],
                },
              },
              children: [],
            },
          ],
        }).root,
      ).toStrictEqual({
        type: 'root',
        children: [
          {
            type: 'heading',
            depth: 1,
            children: [
              {
                type: 'text',
                value: 'heading one',
              },
            ],
          },
        ],
      })
    })
  })

  describe('synced reference', () => {
    test('reads content from inner block manager', () => {
      const { root } = transformer.transform({
        type: BlockType.PAGE,
        snapshot: {
          type: BlockType.PAGE,
        },
        children: [
          {
            type: BlockType.SYNCED_REFERENCE,
            snapshot: {
              type: BlockType.SYNCED_REFERENCE,
              src_page_id: 'source_page',
              src_block_id: 'source_block',
            },
            children: [],
            innerBlockManager: {
              rootBlockModel: {
                type: BlockType.PAGE,
                snapshot: {
                  type: BlockType.PAGE,
                },
                children: [
                  {
                    type: BlockType.TEXT,
                    snapshot: {
                      type: BlockType.TEXT,
                    },
                    zoneState: {
                      allText: '源内容\n',
                      content: {
                        ops: [
                          {
                            insert: '源内容',
                            attributes: {},
                          },
                        ],
                      },
                    },
                    children: [],
                  },
                ],
              },
            },
          },
        ],
      })

      expect(root).toStrictEqual({
        type: 'root',
        children: [
          {
            type: 'paragraph',
            children: [
              {
                type: 'text',
                value: '源内容',
              },
            ],
          },
        ],
      })
    })
  })

  describe('code', () => {
    test('simple example', () => {
      const { root } = transformer.transform({
        type: BlockType.PAGE,
        snapshot: {
          type: BlockType.PAGE,
        },
        children: [
          {
            type: BlockType.CODE,
            language: 'JavaScript',
            snapshot: {
              type: BlockType.CODE,
            },
            zoneState: {
              allText: 'const\n',
              content: {
                ops: [],
              },
            },
            children: [],
          },
        ],
      })
      const expectedRoot: mdast.Root = {
        type: 'root',
        children: [
          {
            type: 'code',
            lang: 'javascript',
            value: 'const',
          },
        ],
      }
      expect(root).toStrictEqual(expectedRoot)
    })
  })

  describe('blockquote', () => {
    test('simple example', () => {
      const { root } = transformer.transform({
        type: BlockType.PAGE,
        snapshot: {
          type: BlockType.PAGE,
        },
        children: [
          {
            type: BlockType.QUOTE_CONTAINER,
            snapshot: {
              type: BlockType.QUOTE_CONTAINER,
            },
            children: [
              {
                type: BlockType.ORDERED,
                snapshot: {
                  type: BlockType.ORDERED,
                  seq: '1',
                },
                zoneState: {
                  allText: '',
                  content: {
                    ops: [
                      {
                        insert: 'list item 1',
                        attributes: {},
                      },
                    ],
                  },
                },
                children: [],
              },
              {
                type: BlockType.ORDERED,
                snapshot: {
                  type: BlockType.ORDERED,
                  seq: '2',
                },
                zoneState: {
                  allText: '',
                  content: {
                    ops: [
                      {
                        insert: 'list item 2',
                        attributes: {},
                      },
                    ],
                  },
                },
                children: [],
              },
              {
                type: BlockType.DIAGRAM,
                snapshot: {
                  type: BlockType.DIAGRAM,
                },
                children: [],
              },
            ],
          },
        ],
      })
      const expectedRoot: mdast.Root = {
        type: 'root',
        children: [
          {
            type: 'blockquote',
            children: [
              {
                type: 'list',
                ordered: true,
                start: 1,
                children: [
                  {
                    type: 'listItem',
                    data: {
                      seq: 1,
                    },
                    children: [
                      {
                        type: 'paragraph',
                        children: [
                          {
                            type: 'text',
                            value: 'list item 1',
                          },
                        ],
                      },
                    ],
                  },
                  {
                    type: 'listItem',
                    data: {
                      seq: 2,
                    },
                    children: [
                      {
                        type: 'paragraph',
                        children: [
                          {
                            type: 'text',
                            value: 'list item 2',
                          },
                        ],
                      },
                    ],
                  },
                ],
              },
            ],
          },
        ],
      }
      expect(root).toStrictEqual(expectedRoot)
    })
  })

  describe('list', () => {
    test('simple example', () => {
      const { root } = transformer.transform({
        type: BlockType.PAGE,
        snapshot: {
          type: BlockType.PAGE,
        },
        children: [
          {
            type: BlockType.BULLET,
            snapshot: {
              type: BlockType.BULLET,
            },
            zoneState: {
              allText: '',
              content: {
                ops: [
                  {
                    insert: 'a',
                    attributes: {},
                  },
                ],
              },
            },
            children: [],
          },
          {
            type: BlockType.BULLET,
            snapshot: {
              type: BlockType.BULLET,
            },
            zoneState: {
              allText: '',
              content: {
                ops: [
                  {
                    insert: 'b',
                    attributes: {},
                  },
                ],
              },
            },
            children: [],
          },
          {
            type: BlockType.ORDERED,
            snapshot: {
              type: BlockType.ORDERED,
              seq: '2',
            },
            zoneState: {
              allText: '',
              content: {
                ops: [
                  {
                    insert: 'one',
                    attributes: {},
                  },
                ],
              },
            },
            children: [],
          },
          {
            type: BlockType.ORDERED,
            snapshot: {
              type: BlockType.ORDERED,
              seq: 'auto',
            },
            zoneState: {
              allText: '',
              content: {
                ops: [
                  {
                    insert: 'two',
                    attributes: {},
                  },
                ],
              },
            },
            children: [],
          },
          {
            type: BlockType.ORDERED,
            snapshot: {
              type: BlockType.ORDERED,
              seq: '1',
            },
            zoneState: {
              allText: '',
              content: {
                ops: [
                  {
                    insert: 'one',
                    attributes: {},
                  },
                ],
              },
            },
            children: [],
          },
          {
            type: BlockType.ORDERED,
            snapshot: {
              type: BlockType.ORDERED,
              seq: '2',
            },
            zoneState: {
              allText: '',
              content: {
                ops: [
                  {
                    insert: 'one',
                    attributes: {},
                  },
                ],
              },
            },
            children: [],
          },
          {
            type: BlockType.TODO,
            snapshot: {
              type: BlockType.TODO,
            },
            zoneState: {
              allText: '',
              content: {
                ops: [
                  {
                    insert: 'task one',
                    attributes: {},
                  },
                ],
              },
            },
            children: [],
          },
        ],
      })
      const expectedRoot: mdast.Root = {
        type: 'root',
        children: [
          {
            type: 'list',
            children: [
              {
                type: 'listItem',
                children: [
                  {
                    type: 'paragraph',
                    children: [
                      {
                        type: 'text',
                        value: 'a',
                      },
                    ],
                  },
                ],
              },
              {
                type: 'listItem',
                children: [
                  {
                    type: 'paragraph',
                    children: [
                      {
                        type: 'text',
                        value: 'b',
                      },
                    ],
                  },
                ],
              },
            ],
          },
          {
            type: 'list',
            ordered: true,
            start: 2,
            children: [
              {
                type: 'listItem',
                data: {
                  seq: 2,
                },
                children: [
                  {
                    type: 'paragraph',
                    children: [
                      {
                        type: 'text',
                        value: 'one',
                      },
                    ],
                  },
                ],
              },
              {
                type: 'listItem',
                data: {
                  seq: 'auto',
                },
                children: [
                  {
                    type: 'paragraph',
                    children: [
                      {
                        type: 'text',
                        value: 'two',
                      },
                    ],
                  },
                ],
              },
            ],
          },
          {
            type: 'list',
            ordered: true,
            start: 1,
            children: [
              {
                type: 'listItem',
                data: {
                  seq: 1,
                },
                children: [
                  {
                    type: 'paragraph',
                    children: [
                      {
                        type: 'text',
                        value: 'one',
                      },
                    ],
                  },
                ],
              },
              {
                type: 'listItem',
                data: {
                  seq: 2,
                },
                children: [
                  {
                    type: 'paragraph',
                    children: [
                      {
                        type: 'text',
                        value: 'one',
                      },
                    ],
                  },
                ],
              },
            ],
          },
          {
            type: 'list',
            children: [
              {
                type: 'listItem',
                checked: false,
                children: [
                  {
                    type: 'paragraph',
                    children: [
                      {
                        type: 'text',
                        value: 'task one',
                      },
                    ],
                  },
                ],
              },
            ],
          },
        ],
      }
      expect(root).toStrictEqual(expectedRoot)
    })
  })
})

describe('trim end enter', () => {
  describe('inline math', () => {
    it('with enter', () => {
      expect(
        transformer.transform({
          type: BlockType.PAGE,
          snapshot: {
            type: BlockType.PAGE,
          },
          children: [
            {
              type: BlockType.TEXT,
              snapshot: {
                type: BlockType.TEXT,
              },
              zoneState: {
                allText: '',
                content: {
                  ops: [
                    {
                      insert: '',
                      attributes: {
                        equation: 'math\n',
                      },
                    },
                  ],
                },
              },
              children: [],
            },
          ],
        }).root,
      ).toStrictEqual({
        type: 'root',
        children: [
          {
            type: 'paragraph',
            children: [
              {
                type: 'inlineMath',
                value: 'math',
              },
            ],
          },
        ],
      })
    })

    it('without enter', () => {
      expect(
        transformer.transform({
          type: BlockType.PAGE,
          snapshot: {
            type: BlockType.PAGE,
          },
          children: [
            {
              type: BlockType.TEXT,
              snapshot: {
                type: BlockType.TEXT,
              },
              zoneState: {
                allText: '',
                content: {
                  ops: [
                    {
                      insert: '',
                      attributes: {
                        equation: 'math',
                      },
                    },
                  ],
                },
              },
              children: [],
            },
          ],
        }).root,
      ).toStrictEqual({
        type: 'root',
        children: [
          {
            type: 'paragraph',
            children: [
              {
                type: 'inlineMath',
                value: 'math',
              },
            ],
          },
        ],
      })
    })
  })
})

describe('inline math', () => {
  test('inline equation with a single character', () => {
    expect(
      transformer.transform({
        type: BlockType.PAGE,
        snapshot: {
          type: BlockType.PAGE,
        },
        children: [
          {
            type: BlockType.TEXT,
            snapshot: {
              type: BlockType.TEXT,
            },
            zoneState: {
              allText: '',
              content: {
                ops: [
                  {
                    insert: '',
                    attributes: {
                      equation: 'a',
                    },
                  },
                ],
              },
            },
            children: [],
          },
        ],
      }).root,
    ).toStrictEqual({
      type: 'root',
      children: [
        {
          type: 'paragraph',
          children: [
            {
              type: 'inlineMath',
              value: 'a',
            },
          ],
        },
      ],
    })
  })
})

describe('embedded sheet', () => {
  test('exports the rendered sheet html when enabled', () => {
    document.body.innerHTML = `
      <section data-block-type="sheet">
        <div data-block-id="sheet-block">
          <table><thead><tr><th>Name</th></tr></thead><tbody><tr><td>Alice</td></tr></tbody></table>
        </div>
      </section>
    `

    const page: PageBlock = {
      id: 1,
      type: BlockType.PAGE,
      snapshot: {
        type: BlockType.PAGE,
      },
      children: [
        {
          id: 42,
          type: BlockType.SHEET,
          record: {
            id: 'sheet-block',
          },
          snapshot: {
            type: BlockType.SHEET,
          },
          children: [],
        },
      ],
    }

    const { root } = new Transformer({ bitable: true }).transform(page)

    const [sheet] = root.children

    expect(sheet.type).toBe('html')
    expect(sheet).toMatchObject({
      value:
        '<figure class="sheet" data-sheet-record-id="sheet-block" data-sheet-block-id="42"><div class="sheet-wrapper"><table><thead><tr><th>Name</th></tr></thead><tbody><tr><td>Alice</td></tr></tbody></table></div></figure>',
    })
  })

  test('snapshots a canvas-rendered sheet as an inline image', () => {
    document.body.innerHTML = `
      <section data-block-type="sheet">
        <div data-block-id="sheet-block">
          <canvas class="spreadsheet-canvas" width="10" height="12"></canvas>
        </div>
      </section>
    `

    const canvas = document.querySelector('canvas')
    if (!canvas) throw new Error('Expected canvas fixture')
    canvas.width = 10
    canvas.height = 12
    canvas.toDataURL = () => 'data:image/png;base64,sheet'

    const page: PageBlock = {
      id: 1,
      type: BlockType.PAGE,
      snapshot: {
        type: BlockType.PAGE,
      },
      children: [
        {
          id: 42,
          type: BlockType.SHEET,
          record: {
            id: 'sheet-block',
          },
          snapshot: {
            type: BlockType.SHEET,
          },
          children: [],
        },
      ],
    }

    const { root } = new Transformer({ bitable: true }).transform(page)

    const [sheet] = root.children

    expect(sheet.type).toBe('html')
    expect(sheet).toMatchObject({
      value:
        '<figure class="sheet" data-sheet-record-id="sheet-block" data-sheet-block-id="42"><div class="sheet-wrapper"><img class="sheet-snapshot" src="data:image/png;base64,sheet" alt="Embedded sheet snapshot" width="10" height="12"></div></figure>',
    })
  })

  test('exports a canvas snapshot without reading source canvas pixels', () => {
    document.body.innerHTML = `
      <section data-block-type="sheet">
        <div data-block-id="sheet-block">
          <canvas class="spreadsheet-canvas" width="10" height="12"></canvas>
        </div>
      </section>
    `

    const canvas = document.querySelector('canvas')
    if (!canvas) throw new Error('Expected canvas fixture')

    canvas.width = 10
    canvas.height = 12
    canvas.toDataURL = () => 'data:image/png;base64,sheet'
    canvas.getContext = () => {
      throw new Error('Source sheet canvas pixels should not be read')
    }

    const page: PageBlock = {
      id: 1,
      type: BlockType.PAGE,
      snapshot: {
        type: BlockType.PAGE,
      },
      children: [
        {
          id: 42,
          type: BlockType.SHEET,
          record: {
            id: 'sheet-block',
          },
          snapshot: {
            type: BlockType.SHEET,
          },
          children: [],
        },
      ],
    }

    const { root } = new Transformer({ bitable: true }).transform(page)

    const [sheet] = root.children

    expect(sheet.type).toBe('html')
    expect(sheet).toMatchObject({
      value:
        '<figure class="sheet" data-sheet-record-id="sheet-block" data-sheet-block-id="42"><div class="sheet-wrapper"><img class="sheet-snapshot" src="data:image/png;base64,sheet" alt="Embedded sheet snapshot" width="10" height="12"></div></figure>',
    })
  })

  test('does not export loading-only sheet dom as content', () => {
    document.body.innerHTML = `
      <section data-block-type="sheet">
        <div data-block-id="sheet-block">
          <div class="docx-block-loading-container"></div>
          <div class="sheet-block-container"><div></div></div>
        </div>
      </section>
    `

    const page: PageBlock = {
      id: 1,
      type: BlockType.PAGE,
      snapshot: {
        type: BlockType.PAGE,
      },
      children: [
        {
          id: 42,
          type: BlockType.SHEET,
          record: {
            id: 'sheet-block',
          },
          snapshot: {
            type: BlockType.SHEET,
          },
          children: [],
        },
      ],
    }

    const { root } = new Transformer({ bitable: true }).transform(page)

    const [sheet] = root.children

    expect(sheet.type).toBe('html')
    expect(sheet).toMatchObject({
      value:
        '<figure class="sheet" data-sheet-record-id="sheet-block" data-sheet-block-id="42"><div class="sheet-wrapper"><p class="sheet-missing">Sheet content is not loaded in the current page.</p></div></figure>',
    })
  })

  test('uses the visible sheet after locating a block without stable DOM ids', async () => {
    const page: PageBlock = {
      id: 1,
      type: BlockType.PAGE,
      snapshot: {
        type: BlockType.PAGE,
      },
      children: [
        {
          id: 42,
          type: BlockType.SHEET,
          record: {
            id: 'sheet-block',
          },
          snapshot: {
            type: BlockType.SHEET,
          },
          children: [],
        },
      ],
    }

    const { root } = new Transformer({
      bitable: true,
      locateBlockWithRecordId: () => {
        document.body.innerHTML = `
          <section data-sheet-element="embeddedSheetContainer">
            <canvas class="spreadsheet-canvas" width="10" height="12"></canvas>
          </section>
        `

        const section = document.querySelector('section')
        if (!section) throw new Error('Expected sheet fixture')
        section.getBoundingClientRect = () =>
          ({
            width: 10,
            height: 12,
            top: 0,
            left: 0,
            right: 10,
            bottom: 12,
          }) as DOMRect

        const canvas = document.querySelector('canvas')
        if (!canvas) throw new Error('Expected canvas fixture')
        canvas.width = 10
        canvas.height = 12
        canvas.toDataURL = () => 'data:image/png;base64,visible-sheet'

        return Promise.resolve(true)
      },
    }).transform(page)

    const [sheet] = root.children
    expect(sheet.type).toBe('html')

    const html = await (sheet as mdast.Html).data?.fetchHtml?.()

    expect(html).toBe(
      '<figure class="sheet" data-sheet-record-id="sheet-block" data-sheet-block-id="42"><div class="sheet-wrapper"><img class="sheet-snapshot" src="data:image/png;base64,visible-sheet" alt="Embedded sheet snapshot" width="10" height="12"></div></figure>',
    )
  })

  test('stitches a scrollable canvas-rendered sheet', async () => {
    const drawCalls: unknown[][] = []
    const originalGetContextDescriptor = Object.getOwnPropertyDescriptor(
      HTMLCanvasElement.prototype,
      'getContext',
    )
    const originalToDataURLDescriptor = Object.getOwnPropertyDescriptor(
      HTMLCanvasElement.prototype,
      'toDataURL',
    )

    Object.defineProperty(HTMLCanvasElement.prototype, 'getContext', {
      configurable: true,
      value: function getContext(this: HTMLCanvasElement) {
        if (this.classList.contains('spreadsheet-canvas')) {
          return {
            getImageData: () => ({
              data: new Uint8ClampedArray([0, 0, 0, 255]),
            }),
          }
        }

        return {
          drawImage: (...args: unknown[]) => {
            drawCalls.push(args)
          },
        }
      },
    })
    Object.defineProperty(HTMLCanvasElement.prototype, 'toDataURL', {
      configurable: true,
      value: function toDataURL(this: HTMLCanvasElement) {
        return this.classList.contains('spreadsheet-canvas')
          ? `data:image/png;base64,viewport-${String(scrollLeft)}`
          : 'data:image/png;base64,stitched'
      },
    })

    document.body.innerHTML = `
      <section data-block-type="sheet">
        <div data-block-id="sheet-block">
          <div class="scrollable-wrapper">
            <canvas class="spreadsheet-canvas" width="100" height="50"></canvas>
          </div>
        </div>
      </section>
    `

    const scrollable = document.querySelector<HTMLElement>(
      '.scrollable-wrapper',
    )
    const canvas = document.querySelector('canvas')
    if (!scrollable || !canvas) throw new Error('Expected sheet fixtures')

    let scrollLeft = 0
    Object.defineProperties(scrollable, {
      clientWidth: { value: 100, configurable: true },
      clientHeight: { value: 50, configurable: true },
      scrollWidth: { value: 200, configurable: true },
      scrollHeight: { value: 50, configurable: true },
      scrollLeft: {
        get: () => scrollLeft,
        set: (value: number) => {
          scrollLeft = value
        },
        configurable: true,
      },
    })
    scrollable.getBoundingClientRect = () =>
      ({
        width: 100,
        height: 50,
        top: 0,
        left: 0,
        right: 100,
        bottom: 50,
      }) as DOMRect

    canvas.width = 100
    canvas.height = 50
    canvas.getBoundingClientRect = () =>
      ({
        width: 100,
        height: 50,
        top: 0,
        left: 0,
        right: 100,
        bottom: 50,
      }) as DOMRect

    const page: PageBlock = {
      id: 1,
      type: BlockType.PAGE,
      snapshot: {
        type: BlockType.PAGE,
      },
      children: [
        {
          id: 42,
          type: BlockType.SHEET,
          record: {
            id: 'sheet-block',
          },
          snapshot: {
            type: BlockType.SHEET,
          },
          children: [],
        },
      ],
    }

    try {
      const { root } = new Transformer({ bitable: true }).transform(page)

      const [sheet] = root.children
      expect(sheet.type).toBe('html')

      const html = await (sheet as mdast.Html).data?.fetchHtml?.()

      expect(html).toBe(
        '<figure class="sheet" data-sheet-record-id="sheet-block" data-sheet-block-id="42"><div class="sheet-wrapper"><img class="sheet-snapshot" src="data:image/png;base64,stitched" alt="Embedded sheet snapshot" width="200" height="50"></div></figure>',
      )
      expect(drawCalls).toHaveLength(2)
      expect(drawCalls.map(call => call.slice(1, 3))).toStrictEqual([
        [0, 0],
        [100, 0],
      ])
    } finally {
      if (originalGetContextDescriptor) {
        Object.defineProperty(
          HTMLCanvasElement.prototype,
          'getContext',
          originalGetContextDescriptor,
        )
      }
      if (originalToDataURLDescriptor) {
        Object.defineProperty(
          HTMLCanvasElement.prototype,
          'toDataURL',
          originalToDataURLDescriptor,
        )
      }
    }
  })

  test('falls back to the visible sheet snapshot when scroll does not change canvas content', async () => {
    const originalGetContextDescriptor = Object.getOwnPropertyDescriptor(
      HTMLCanvasElement.prototype,
      'getContext',
    )
    const originalToDataURLDescriptor = Object.getOwnPropertyDescriptor(
      HTMLCanvasElement.prototype,
      'toDataURL',
    )

    Object.defineProperty(HTMLCanvasElement.prototype, 'getContext', {
      configurable: true,
      value: function getContext(this: HTMLCanvasElement) {
        if (this.classList.contains('spreadsheet-canvas')) {
          return {
            getImageData: () => ({
              data: new Uint8ClampedArray([0, 0, 0, 255]),
            }),
          }
        }

        return {
          drawImage: () => undefined,
        }
      },
    })
    Object.defineProperty(HTMLCanvasElement.prototype, 'toDataURL', {
      configurable: true,
      value: function toDataURL(this: HTMLCanvasElement) {
        return this.classList.contains('spreadsheet-canvas')
          ? 'data:image/png;base64,viewport'
          : 'data:image/png;base64,stitched'
      },
    })

    document.body.innerHTML = `
      <section data-block-type="sheet">
        <div data-block-id="sheet-block">
          <div class="scrollable-wrapper">
            <canvas class="spreadsheet-canvas" width="100" height="50"></canvas>
          </div>
        </div>
      </section>
    `

    const scrollable = document.querySelector<HTMLElement>(
      '.scrollable-wrapper',
    )
    const canvas = document.querySelector('canvas')
    if (!scrollable || !canvas) throw new Error('Expected sheet fixtures')

    let scrollLeft = 0
    Object.defineProperties(scrollable, {
      clientWidth: { value: 100, configurable: true },
      clientHeight: { value: 50, configurable: true },
      scrollWidth: { value: 200, configurable: true },
      scrollHeight: { value: 50, configurable: true },
      scrollLeft: {
        get: () => scrollLeft,
        set: (value: number) => {
          scrollLeft = value
        },
        configurable: true,
      },
    })
    scrollable.getBoundingClientRect = () =>
      ({
        width: 100,
        height: 50,
        top: 0,
        left: 0,
        right: 100,
        bottom: 50,
      }) as DOMRect

    canvas.width = 100
    canvas.height = 50
    canvas.getBoundingClientRect = () =>
      ({
        width: 100,
        height: 50,
        top: 0,
        left: 25,
        right: 125,
        bottom: 50,
      }) as DOMRect

    const page: PageBlock = {
      id: 1,
      type: BlockType.PAGE,
      snapshot: {
        type: BlockType.PAGE,
      },
      children: [
        {
          id: 42,
          type: BlockType.SHEET,
          record: {
            id: 'sheet-block',
          },
          snapshot: {
            type: BlockType.SHEET,
          },
          children: [],
        },
      ],
    }

    try {
      const { root } = new Transformer({ bitable: true }).transform(page)

      const [sheet] = root.children
      expect(sheet.type).toBe('html')

      const html = await (sheet as mdast.Html).data?.fetchHtml?.()

      expect(html).toBe(
        '<figure class="sheet" data-sheet-record-id="sheet-block" data-sheet-block-id="42"><div class="sheet-wrapper"><img class="sheet-snapshot" src="data:image/png;base64,viewport" alt="Embedded sheet snapshot" width="100" height="50"></div></figure>',
      )
    } finally {
      if (originalGetContextDescriptor) {
        Object.defineProperty(
          HTMLCanvasElement.prototype,
          'getContext',
          originalGetContextDescriptor,
        )
      }
      if (originalToDataURLDescriptor) {
        Object.defineProperty(
          HTMLCanvasElement.prototype,
          'toDataURL',
          originalToDataURLDescriptor,
        )
      }
    }
  })

  test('can locate a rendered sheet by block id when record id is absent', () => {
    document.body.innerHTML = `
      <div data-block-type="sheet">
        <div data-block-id="42" role="grid"><div role="row"><div role="columnheader">Name</div></div></div>
      </div>
    `

    const page: PageBlock = {
      id: 1,
      type: BlockType.PAGE,
      snapshot: {
        type: BlockType.PAGE,
      },
      children: [
        {
          id: 42,
          type: BlockType.SHEET,
          snapshot: {
            type: BlockType.SHEET,
          },
          children: [],
        },
      ],
    }

    const { root } = new Transformer({ bitable: true }).transform(page)

    const [sheet] = root.children

    expect(sheet.type).toBe('html')
    expect(sheet).toMatchObject({
      value:
        '<figure class="sheet" data-sheet-block-id="42"><div class="sheet-wrapper"><div data-block-id="42" role="grid"><div role="row"><div role="columnheader">Name</div></div></div></div></figure>',
    })
  })
})
