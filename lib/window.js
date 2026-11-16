import { collapseWhitespace, truncateText } from './text.js'

/**
 * Render a bounded read window.
 * @param {any} doc
 * @param {{ offset: number, limit: number, pages: any, focus: string, maxChars: number }} options
 */
export function renderRead(doc, options) {
  if (options.focus === 'outline') return renderOutline(doc, options)
  if (options.focus === 'tables') return renderTables(doc, options)
  return renderContent(doc, options)
}

/**
 * @param {any} doc
 * @param {{ offset: number, limit: number, pages: any, maxChars: number }} options
 */
function renderContent(doc, options) {
  const lines = doc.lines
  const selection = options.pages
  const startIndex = resolveStartIndex(lines, options.offset, selection)
  if (startIndex >= lines.length) {
    throw new Error(`offset ${options.offset} is out of range for "${doc.path}" (${doc.lineCount} lines)`)
  }
  const maxPage = selection === null ? undefined : Math.max(...selection.pages)
  /** @type {any[]} */
  const selected = []
  let chars = 0
  let truncatedByChars = false
  let index = startIndex
  for (; index < lines.length; index += 1) {
    const line = lines[index]
    if (selection !== null && (line.page === undefined || !selection.pages.has(line.page))) {
      if (maxPage !== undefined && line.page !== undefined && line.page > maxPage) break
      continue
    }
    if (selected.length >= options.limit) break
    const cost = line.text.length + 6
    if (chars + cost > options.maxChars && selected.length > 0) {
      truncatedByChars = true
      break
    }
    selected.push(line)
    chars += cost
  }

  const header = buildHeader(doc, {
    first: selected[0],
    last: selected.at(-1),
    selection,
    showOutline: options.offset === 1,
  })
  if (selected.length === 0) {
    return `${header}\n\n(No content in the selected range. ${selection === null ? '' : `Selected pages: ${selection.text}. `}Use offset=1 or a different page range.)`
  }
  const body = renderLines(selected)
  const endLine = selected.at(-1).number
  let footer
  if (truncatedByChars) {
    footer = `(Output capped at ${Math.round(options.maxChars / 1024)} KB; showing lines ${selected[0].number}-${endLine}. Continue with offset=${endLine + 1}.)`
  } else if (endLine < doc.lineCount) {
    const pages = pageRangeText(selected)
    footer = `(Showing lines ${selected[0].number}-${endLine} of ${doc.lineCount}${pages === '' ? '' : `; ${pages}`}. Continue with offset=${endLine + 1}${selection === null ? '' : ' or a different pages range'}.)`
  } else {
    footer = `(End of document - total ${doc.lineCount} lines${pageRangeText(selected) === '' ? '' : `; ${pageRangeText(selected)}`})`
  }
  return `${header}\n\n${body}\n\n${footer}`
}

/**
 * @param {any[]} lines
 * @param {number} offset
 * @param {{ pages: Set<number> } | null} selection
 */
function resolveStartIndex(lines, offset, selection) {
  if (selection === null || offset > 1) return offset - 1
  const first = lines.findIndex((line) => line.page !== undefined && selection.pages.has(line.page))
  return first === -1 ? 0 : first
}

/** @param {any[]} selected */
function renderLines(selected) {
  const out = []
  for (let index = 0; index < selected.length; index += 1) {
    const line = selected[index]
    out.push(`${line.number}: ${line.text}`)
    if (line.blankAfter && index < selected.length - 1) out.push('')
  }
  return out.join('\n')
}

/**
 * @param {any} doc
 * @param {{ first?: any, last?: any, selection: any, showOutline: boolean }} context
 */
function buildHeader(doc, context) {
  const meta = doc.meta
  const parts = [`[${doc.kind}] ${doc.path}`]
  if (doc.kind === 'pdf') parts.push(`${meta.pages} page${meta.pages === 1 ? '' : 's'}`)
  else parts.push(`pages unknown (DOCX)`)
  parts.push(`${doc.lineCount} lines`)
  parts.push(`${formatInt(doc.charCount)} chars`)
  if (meta.headings > 0) parts.push(`${meta.headings} headings${meta.headingsInferred ? ' (inferred)' : ''}`)
  if (doc.kind === 'docx' && meta.tables > 0) parts.push(`${meta.tables} tables`)
  if (doc.kind === 'pdf' && meta.tables > 0) parts.push(`${meta.tables} layout pages`)
  if (meta.images > 0) parts.push(`${meta.images} image${meta.images === 1 ? '' : 's'}`)
  if (meta.title) parts.push(`title: ${truncateText(meta.title, 80)}`)
  const header = [parts.join(' | ')]

  if (doc.outline.length > 0) {
    if (context.showOutline) {
      const entries = doc.outline.slice(0, 8).map((entry) => {
        const page = entry.page === undefined ? '' : `p${entry.page} `
        return `${entry.line}: ${page}${truncateText(entry.text, 60)}`
      })
      const more = doc.outline.length > 8 ? ` | +${doc.outline.length - 8} more` : ''
      header.push(`outline: ${entries.join(' | ')}${more}`)
    } else {
      header.push(`outline: ${doc.outline.length} headings (use focus="outline")`)
    }
  } else {
    header.push('outline: none detected')
  }

  const notes = []
  if (meta.tocOmitted) notes.push(`static TOC omitted (${meta.tocCount} entries)`)
  if (Array.isArray(meta.removedMargins) && meta.removedMargins.length > 0) notes.push('repeated headers/footers removed')
  if (meta.headingsInferred) notes.push('heading levels are inferred')
  if (doc.kind === 'docx') notes.push('DOCX has no fixed page numbers')
  if (notes.length > 0) header.push(`note: ${notes.join('; ')}`)
  if (context.selection !== null) header.push(`selection: pages ${context.selection.text}`)
  return header.join('\n')
}

/**
 * @param {any} doc
 * @param {{ pages: any, maxChars: number }} options
 */
function renderOutline(doc, options) {
  const header = buildHeader(doc, { selection: options.pages, showOutline: false })
  let entries = doc.outline
  if (options.pages !== null) {
    entries = entries.filter((entry) => entry.page !== undefined && options.pages.pages.has(entry.page))
  }
  if (entries.length === 0) {
    return `${header}\n\n(No headings detected${options.pages === null ? '' : ' in the selected pages'}. Read content with focus="content", or use focus="tables".)`
  }
  const shown = entries.slice(0, 100)
  const body = shown.map((entry) => {
    const indent = '  '.repeat(Math.max(0, Math.min(6, entry.level - 1)))
    const page = entry.page === undefined ? '' : ` [p${entry.page}]`
    return `${entry.line}: ${indent}${'#'.repeat(Math.max(1, Math.min(6, entry.level)))} ${truncateText(entry.text, 100)}${page}`
  }).join('\n')
  const more = entries.length > shown.length ? `\n(Showing ${shown.length} of ${entries.length} headings.)` : ''
  return `${header}\n\n${body}${more}\n\n(Use offset=<line> with focus="content" to read a section.)`
}

/**
 * @param {any} doc
 * @param {{ offset: number, limit: number, pages: any, maxChars: number }} options
 */
function renderTables(doc, options) {
  const header = buildHeader(doc, { selection: options.pages, showOutline: false })
  const ranges = doc.blockRanges.filter((range) => range.type === 'table' || range.type === 'layout')
  const selected = ranges.filter((range) => {
    if (options.pages === null) return true
    for (let lineIndex = range.lineStart; lineIndex <= range.lineEnd; lineIndex += 1) {
      const page = doc.lines[lineIndex - 1]?.page
      if (page !== undefined && options.pages.pages.has(page)) return true
    }
    return false
  })
  if (selected.length === 0) {
    return `${header}\n\n(No tables or layout blocks detected${options.pages === null ? '' : ' in the selected pages'}. Use focus="content".)`
  }
  const out = []
  let chars = 0
  let count = 0
  let shownTables = 0
  for (const range of selected) {
    const start = Math.max(range.lineStart, options.offset)
    if (start > range.lineEnd) continue
    const chunk = []
    for (let lineIndex = start; lineIndex <= range.lineEnd; lineIndex += 1) {
      const line = doc.lines[lineIndex - 1]
      if (line === undefined) continue
      if (options.pages !== null && (line.page === undefined || !options.pages.pages.has(line.page))) continue
      const cost = line.text.length + 6
      if (chars + cost > options.maxChars || count >= options.limit) break
      chunk.push(`${line.number}: ${line.text}`)
      chars += cost
      count += 1
    }
    if (chunk.length > 0) {
      out.push(chunk.join('\n'))
      shownTables += 1
    }
    if (count >= options.limit || chars >= options.maxChars) break
  }
  const kindNote = doc.kind === 'pdf' ? '\n(PDF layout blocks preserve column order, not exact cell boundaries.)' : ''
  return `${header}\n\n${out.join('\n\n')}${kindNote}\n\n(Showing ${shownTables} of ${selected.length} table/layout blocks. Use focus="content" and offset to read surrounding prose.)`
}

/** @param {any[]} lines */
function pageRangeText(lines) {
  const pages = lines.map((line) => line.page).filter((page) => typeof page === 'number')
  if (pages.length === 0) return ''
  const min = Math.min(...pages)
  const max = Math.max(...pages)
  return min === max ? `page ${min}` : `pages ${min}-${max}`
}

/** @param {number} value */
export function formatInt(value) {
  return String(value).replace(/\B(?=(\d{3})+(?!\d))/g, ',')
}
