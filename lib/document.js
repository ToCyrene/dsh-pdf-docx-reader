import { escapeTableCell } from './text.js'

/**
 * @typedef {{
 *   text: string,
 *   plain: string,
 *   number: number,
 *   page?: number,
 *   headingPath?: string,
 *   blankAfter?: boolean,
 *   blockIndex: number,
 * }} DocumentLine
 * @typedef {{
 *   kind: 'pdf' | 'docx',
 *   path: string,
 *   lines: DocumentLine[],
 *   outline: Array<{ level: number, text: string, line: number, page?: number }>,
 *   meta: Record<string, any>,
 *   lineCount: number,
 *   charCount: number,
 * }} ReaderDocument
 */

/**
 * Render parsed blocks into a virtual, line-addressed Markdown document.
 * @param {{
 *   kind: 'pdf' | 'docx',
 *   path: string,
 *   blocks: any[],
 *   meta: Record<string, any>,
 *   pageMarkers?: boolean,
 * }} input
 * @returns {ReaderDocument}
 */
export function buildDocument(input) {
  const { kind, path, blocks } = input
  const pageMarkers = input.pageMarkers !== false
  /** @type {DocumentLine[]} */
  const lines = []
  /** @type {ReaderDocument['outline']} */
  const outline = []
  /** @type {Array<{ level: number, text: string }>} */
  const headingStack = []
  let currentPage
  let currentBlockIndex = 0

  /**
   * @param {string} text
   * @param {string} plain
   * @param {{ page?: number, blankAfter?: boolean }} [options]
   */
  function push(text, plain, options = {}) {
    if (text.length === 0 && plain.length === 0 && !options.blankAfter) return
    const page = options.page ?? currentPage
    const headingPath = headingStack.length === 0
      ? undefined
      : headingStack.map((entry) => entry.text).join(' › ').slice(0, 160)
    lines.push({
      number: lines.length + 1,
      text,
      plain,
      ...(page === undefined ? {} : { page }),
      ...(headingPath === undefined ? {} : { headingPath }),
      ...(options.blankAfter ? { blankAfter: true } : {}),
      blockIndex: currentBlockIndex,
    })
  }

  for (let blockIndex = 0; blockIndex < blocks.length; blockIndex += 1) {
    const block = blocks[blockIndex]
    if (block === undefined || block === null) continue
    currentBlockIndex = blockIndex
    if (block.type === 'page') {
      currentPage = block.page
      if (pageMarkers) {
        const extra = block.images > 0 ? `, ${block.images} image${block.images === 1 ? '' : 's'}` : ''
        push(`[p${block.page}${extra}]`, '', { page: block.page })
      }
      continue
    }

    if (block.type === 'heading') {
      const level = Math.max(1, Math.min(6, block.level ?? 1))
      while (headingStack.length > 0 && headingStack[headingStack.length - 1].level >= level) headingStack.pop()
      headingStack.push({ level, text: block.plain || block.text })
      const lineNumber = lines.length + 1
      push(`${'#'.repeat(level)} ${block.text}`.trimEnd(), block.plain || block.text, { blankAfter: true })
      outline.push({
        level,
        text: block.plain || block.text,
        line: lineNumber,
        ...(block.page === undefined ? {} : { page: block.page }),
      })
      continue
    }

    if (block.type === 'paragraph') {
      const plain = block.plain ?? block.text
      const pieces = String(block.text).split('\n')
      const plainPieces = String(plain).split('\n')
      for (let index = 0; index < pieces.length; index += 1) {
        const isLast = index === pieces.length - 1
        push(pieces[index], plainPieces[index] ?? '', { blankAfter: isLast })
      }
      continue
    }

    if (block.type === 'list') {
      push(block.text, block.plain ?? block.text)
      continue
    }

    if (block.type === 'table') {
      const rows = Array.isArray(block.rows) ? block.rows : []
      if (rows.length === 0) continue
      push(block.marker ?? '[table]', '', { blankAfter: false })
      const columns = Math.max(...rows.map((row) => row.length))
      const normalized = rows.map((row) => {
        const copy = row.slice()
        while (copy.length < columns) copy.push(' ')
        return copy
      })
      const header = normalized[0]
      push(`| ${header.map(escapeTableCell).join(' | ')} |`, header.join(' | '))
      push(`| ${header.map(() => '---').join(' | ')} |`, '')
      for (let rowIndex = 1; rowIndex < normalized.length; rowIndex += 1) {
        const row = normalized[rowIndex]
        push(`| ${row.map(escapeTableCell).join(' | ')} |`, row.join(' | '), {
          blankAfter: rowIndex === normalized.length - 1,
        })
      }
      if (normalized.length === 1) {
        const last = lines[lines.length - 1]
        if (last !== undefined) last.blankAfter = true
      }
      continue
    }

    if (block.type === 'image') {
      const dimension = block.width > 0 && block.height > 0 ? `, ${block.width}×${block.height}` : ''
      const alt = block.alt ? `: ${block.alt}` : ''
      push(`[image ${block.index}${alt}${dimension}]`, block.alt || `image ${block.index}`, { blankAfter: true })
      continue
    }

    if (block.type === 'layout') {
      push(`[layout p${block.page}]`, '', { page: block.page, blankAfter: false })
      for (let index = 0; index < block.lines.length; index += 1) {
        const text = block.lines[index]
        push(text, text, {
          page: block.page,
          blankAfter: index === block.lines.length - 1,
        })
      }
      continue
    }

    if (block.type === 'note') {
      push(block.text, block.plain ?? block.text, { blankAfter: true })
    }
  }

  let charCount = 0
  for (const line of lines) charCount += line.text.length
  const blockRanges = []
  let currentRange
  for (let index = 0; index < lines.length; index += 1) {
    const line = lines[index]
    const block = blocks[line.blockIndex]
    if (currentRange === undefined || currentRange.blockIndex !== line.blockIndex) {
      currentRange = {
        blockIndex: line.blockIndex,
        type: block?.type ?? 'unknown',
        block,
        lineStart: index + 1,
        lineEnd: index + 1,
      }
      blockRanges.push(currentRange)
    } else {
      currentRange.lineEnd = index + 1
    }
  }
  const tables = blocks.filter((block) => block?.type === 'table').length
  const images = blocks.filter((block) => block?.type === 'image').length
  const pages = input.meta.pages ?? (kind === 'pdf' ? maxPage(lines) : undefined)
  const meta = {
    ...input.meta,
    ...(pages === undefined ? {} : { pages }),
    lines: lines.length,
    chars: charCount,
    headings: outline.length,
    tables: input.meta.tables ?? tables,
    images: input.meta.images ?? images,
  }

  return {
    kind,
    path,
    lines,
    outline,
    meta,
    blocks,
    blockRanges,
    lineCount: lines.length,
    charCount,
  }
}

/** @param {DocumentLine[]} lines */
function maxPage(lines) {
  let max = 0
  for (const line of lines) {
    if (line.page !== undefined) max = Math.max(max, line.page)
  }
  return max > 0 ? max : undefined
}
