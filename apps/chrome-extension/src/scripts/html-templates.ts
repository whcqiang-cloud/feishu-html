export interface AttachmentInfo {
  name: string
  token: string
}

export interface HtmlTemplateOptions {
  pageTitle: string
  bodyHtml: string
  attachments: AttachmentInfo[]
  includeStyles: boolean
  printFriendly: boolean
}

const CSS_STYLES = /* css */ `
  :root {
    --color-bg: #ffffff;
    --color-text: #1f2328;
    --color-text-secondary: #656d76;
    --color-border: #d0d7de;
    --color-code-bg: #f6f8fa;
    --color-blockquote-border: #d0d7de;
    --color-table-stripe: #f6f8fa;
    --color-link: #0969da;
  }

  * {
    box-sizing: border-box;
  }

  body {
    font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', 'Noto Sans',
      Helvetica, Arial, sans-serif, 'Apple Color Emoji', 'Segoe UI Emoji';
    font-size: 16px;
    line-height: 1.6;
    color: var(--color-text);
    background-color: var(--color-bg);
    max-width: 900px;
    margin: 0 auto;
    padding: 32px 24px;
  }

  h1, h2, h3, h4, h5, h6 {
    margin-top: 24px;
    margin-bottom: 16px;
    font-weight: 600;
    line-height: 1.25;
  }

  h1 { font-size: 2em; border-bottom: 1px solid var(--color-border); padding-bottom: 0.3em; }
  h2 { font-size: 1.5em; border-bottom: 1px solid var(--color-border); padding-bottom: 0.3em; }
  h3 { font-size: 1.25em; }
  h4 { font-size: 1em; }

  p {
    margin-top: 0;
    margin-bottom: 16px;
  }

  a {
    color: var(--color-link);
    text-decoration: none;
  }

  a:hover {
    text-decoration: underline;
  }

  img {
    max-width: 100%;
    height: auto;
    display: block;
    margin: 16px auto;
  }

  pre {
    background-color: var(--color-code-bg);
    border-radius: 6px;
    padding: 16px;
    overflow-x: auto;
    font-size: 0.875em;
    line-height: 1.45;
  }

  code {
    font-family: 'SFMono-Regular', Consolas, 'Liberation Mono', Menlo, monospace;
    font-size: 0.875em;
  }

  :not(pre) > code {
    background-color: var(--color-code-bg);
    border-radius: 4px;
    padding: 0.2em 0.4em;
  }

  blockquote {
    margin: 0 0 16px;
    padding: 0 1em;
    color: var(--color-text-secondary);
    border-left: 0.25em solid var(--color-blockquote-border);
  }

  table {
    width: 100%;
    border-collapse: collapse;
    margin-bottom: 16px;
    overflow-x: auto;
    display: block;
  }

  table th, table td {
    border: 1px solid var(--color-border);
    padding: 8px 12px;
    text-align: left;
  }

  table th {
    background-color: var(--color-code-bg);
    font-weight: 600;
  }

  table tr:nth-child(even) {
    background-color: var(--color-table-stripe);
  }

  ul, ol {
    margin-top: 0;
    margin-bottom: 16px;
    padding-left: 2em;
  }

  li + li {
    margin-top: 0.25em;
  }

  hr {
    border: 0;
    border-top: 1px solid var(--color-border);
    margin: 24px 0;
  }

  .math {
    overflow-x: auto;
  }

  .math.display {
    display: block;
    margin: 16px 0;
    text-align: center;
  }

  figure.bitable {
    margin: 16px 0;
    overflow-x: auto;
  }

  figure.bitable figcaption {
    font-size: 0.9em;
    color: var(--color-text-secondary);
    margin-bottom: 8px;
    text-align: center;
  }

  .bitable-wrapper {
    overflow-x: auto;
    -webkit-overflow-scrolling: touch;
  }

  .bitable-wrapper table {
    display: table;
    border-collapse: collapse;
    width: auto;
    min-width: 100%;
  }

  .bitable-wrapper table th,
  .bitable-wrapper table td {
    border: 1px solid var(--color-border);
    padding: 6px 12px;
    text-align: left;
    white-space: nowrap;
  }

  .bitable-wrapper table th {
    background-color: var(--color-code-bg);
    font-weight: 600;
    position: sticky;
    top: 0;
  }

  footer {
    margin-top: 48px;
    padding-top: 24px;
    border-top: 1px solid var(--color-border);
    color: var(--color-text-secondary);
    font-size: 0.875em;
  }

  footer h3 {
    font-size: 1em;
    border-bottom: none;
    margin-bottom: 8px;
  }

  footer ul {
    margin-bottom: 8px;
  }

  @media print {
    body {
      max-width: none;
      padding: 0;
    }

    pre, blockquote {
      page-break-inside: avoid;
    }

    h1, h2, h3, h4, h5, h6 {
      page-break-after: avoid;
    }
  }
`

const PRINT_FRIENDLY_CSS = /* css */ `
  @media print {
    body {
      font-size: 12pt;
      line-height: 1.5;
    }

    a {
      color: inherit;
      text-decoration: underline;
    }

    pre, code {
      background-color: #f5f5f5 !important;
      -webkit-print-color-adjust: exact;
      print-color-adjust: exact;
    }

    table th {
      background-color: #f5f5f5 !important;
      -webkit-print-color-adjust: exact;
      print-color-adjust: exact;
    }

    img {
      max-width: 100% !important;
      page-break-inside: avoid;
    }
  }
`

export function wrapIntoFullHtml(options: HtmlTemplateOptions): string {
  const { pageTitle, bodyHtml, attachments, includeStyles, printFriendly } =
    options

  const title = escapeHtml(pageTitle || 'Document')

  let styles = ''
  if (includeStyles) {
    styles = CSS_STYLES
    if (printFriendly) {
      styles += '\n' + PRINT_FRIENDLY_CSS
    }
  }

  let footerHtml = ''
  if (attachments.length > 0) {
    const items = attachments
      .map(
        a =>
          `<li><strong>${escapeHtml(a.name)}</strong> <span style="color: var(--color-text-secondary);">(${escapeHtml(a.token)})</span></li>`,
      )
      .join('\n')
    footerHtml = `
    <h3>Attachments</h3>
    <p>The following files are attached to this document and need to be obtained separately from Feishu:</p>
    <ul>${items}</ul>`
  }

  const now = new Date().toISOString()
  footerHtml += `
    <p style="margin-top: 16px;">Converted at: ${now}</p>`

  return `<!DOCTYPE html>
<html lang="zh-CN">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>${title}</title>
  <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/katex@0.16.11/dist/katex.min.css" integrity="sha384-nB0miv6/jRmo5UMMR1wu3Gz6NLsoTkbqJghGIsx//Rlm+ZU03BU6SQNC66uf4l5+" crossorigin="anonymous">
  <script defer src="https://cdn.jsdelivr.net/npm/katex@0.16.11/dist/katex.min.js" integrity="sha384-7zkQWkzuo3B5mTepMUcHkMB5jZaolc2xDwL6VFqjFALcbeS9Ggm/Yr2r3Dy4lfFg" crossorigin="anonymous"></script>
  <script defer src="https://cdn.jsdelivr.net/npm/katex@0.16.11/dist/contrib/auto-render.min.js" integrity="sha384-43gviWU0YVjaDtb/GhzOouOXtZMP/7XUzwPTstBeZFe/+rCMvRwr4yROQP43s0Xk" crossorigin="anonymous"
    onload="renderMathInElement(document.body, { delimiters: [{left: '\\\\(', right: '\\\\)', display: false}, {left: '\\\\[', right: '\\\\]', display: true}] });"></script>
  <style>
${styles}
  </style>
</head>
<body>
  <article class="markdown-body">
${bodyHtml}
  </article>
  <footer>
${footerHtml}
  </footer>
</body>
</html>`
}

function escapeHtml(str: string): string {
  return str
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;')
}
