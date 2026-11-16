import { createRequire } from 'node:module'
import { dirname, join } from 'node:path'
import { DocReaderError } from '../errors.js'
import { collapseWhitespace, isCJK, joinText, normalizePdfText } from '../text.js'

const require = createRequire(import.meta.url)

/** @type {Promise<any> | undefined} */
let pdfjsPromise

/** Lazily load pdf.js so plugin activation stays cheap. */
async function loadPdfjs() {
  if (pdfjsPromise === undefined) {
    pdfjsPromise = import('pdfjs-dist/legacy/build/pdf.mjs')
  }
  return pdfjsPromise
}

/** Resolve the pdfjs asset directories next to the installed package. */
function pdfjsAssetUrls() {
  let packageJson
  try {
    packageJson = require.resolve('pdfjs-dist/package.json')
  } catch {
    packageJson = require.resolve('pdfjs-dist')
  }
  const root = dirname(packageJson)
  return {
    cMapUrl: `${join(root, 'cmaps')}/`,
    standardFontDataUrl: `${join(root, 'standard_fonts')}/`,
  }
}

/**
 * Parse a PDF into the shared block model.
 * @param {Uint8Array} bytes
 * @param {{ pageMarkers?: boolean }} _config
 * @param {AbortSignal} signal
 * @returns {Promise<{ blocks: any[], meta: Record<string, any> }>}
 */
export async function parsePdf(bytes, _config, signal) {
  const pdfjs = await loadPdfjs()
  const assets = pdfjsAssetUrls()
  const loadingTask = pdfjs.getDocument({
    data: bytes,
    cMapUrl: assets.cMapUrl,
    cMapPacked: true,
    standardFontDataUrl: assets.standardFontDataUrl,
    isEvalSupported: false,
    useWorkerFetch: false,
    disableFontFace: true,
    verbosity: 0,
  })

  let doc
  try {
    doc = await loadingTask.promise
  } catch (error) {
    const name = error instanceof Error ? error.name : ''
    if (name === 'PasswordException') {
      throw new DocReaderError('This PDF is password-protected; decrypt it before reading', 'PDF_ENCRYPTED', { cause: error })
    }
    throw new DocReaderError(
      `Cannot open PDF (${error instanceof Error ? error.message : String(error)})`,
      'PDF_INVALID',
      { cause: error },
    )
  }

  try {
    signal.throwIfAborted()
    const numPages = doc.numPages
    if (numPages === 0) throw new DocReaderError('PDF has no pages', 'PDF_INVALID')

    /** @type {Array<{ number: number, width: number, height: number, items: any[], imageCount: number }>} */
    const pages = []
    for (let number = 1; number <= numPages; number += 1) {
      signal.throwIfAborted()
      const page = await doc.getPage(number)
      const viewport = page.getViewport({ scale: 1 })
      const textContent = await page.getTextContent({ includeMarkedContent: false, disableNormalization: false })
      const items = normalizeItems(textContent.items, viewport)
      let imageCount = 0
      try {
        const operatorList = await page.getOperatorList()
        imageCount = countImages(operatorList, pdfjs.OPS)
      } catch {
        // Image counting is diagnostic only; text extraction remains valid.
      }
      pages.push({ number, width: viewport.width, height: viewport.height, items, imageCount })
      page.cleanup()
    }

    const metadata = await doc.getMetadata().catch(() => undefined)
    const info = metadata?.info ?? {}
    const bookmarks = await extractBookmarks(doc)
    const bookmarkTitles = new Map(bookmarks.map((entry) => [normalizeTitle(entry.text), entry.text]).filter(([key]) => key.length > 0))
    const bodySize = estimateBodySize(pages)
    const marginLines = collectRepeatedMargins(pages)
    const blocks = []
    let totalChars = 0

    for (const page of pages) {
      signal.throwIfAborted()
      const allLines = buildLines(page.items)
      const withoutMargins = allLines.filter((line) => !isMarginLine(line, page.height) || !marginLines.repeated.has(normalizeMargin(line.text)))
      const lines = withoutMargins.filter((line) => !isWatermark(line.text))
      const removed = allLines.filter((line) => !lines.includes(line))
      if (removed.length > 0) marginLines.removed.push(...removed.map((line) => line.text))

      totalChars += lines.reduce((sum, line) => sum + line.text.length, 0)
      blocks.push({ type: 'page', page: page.number, images: page.imageCount })

      if (lines.length === 0) {
        blocks.push({
          type: 'paragraph',
          text: `[p${page.number}: no extractable text${page.imageCount > 0 ? `, ${page.imageCount} image(s)` : ''}]`,
          plain: '',
        })
        continue
      }

      if (isTableLikePage(page, lines)) {
        const layout = layoutLines(lines)
        const first = layout[0]
        if (first !== undefined && first.length <= 40 && !/[。；;]$/.test(first)
          && /^(?:[一二三四五六七八九十百]+、|第[一二三四五六七八九十百\d]+[章节篇部]|[（(][一二三四五六七八九十百\d]+[）)]|\d+(?:\.\d+){0,5}[.、]\s*\S|\d+(?:\.\d+){0,5}\s+\S)/.test(first)) {
          blocks.push({ type: 'heading', level: 1, text: first, plain: first, page: page.number })
          layout.shift()
        }
        blocks.push({ type: 'layout', page: page.number, lines: layout })
        continue
      }

      const columns = splitItemColumns(page.items, page.width)
      for (const columnItems of columns) {
        const columnLines = buildLines(columnItems).filter((line) => {
          if (isWatermark(line.text)) return false
          if (isMarginLine(line, page.height) && marginLines.repeated.has(normalizeMargin(line.text))) return false
          return true
        })
        if (columnLines.length > 0) appendPageColumn(blocks, columnLines, page, bodySize, bookmarkTitles)
      }
    }

    if (totalChars < 20) {
      throw new DocReaderError(
        `No extractable text in ${numPages} page(s); this PDF is probably scanned. Use an OCR-capable reader (for example read_pdf with MinerU) or render the pages.`,
        'PDF_NO_TEXT',
      )
    }

    const outline = inferOutline(blocks)
    const headings = blocks.filter((block) => block.type === 'heading').length
    return {
      blocks,
      meta: {
        title: typeof info.Title === 'string' ? info.Title.trim() : '',
        pages: numPages,
        pagesKnown: true,
        headings,
        tables: blocks.filter((block) => block.type === 'layout').length,
        images: pages.reduce((sum, page) => sum + page.imageCount, 0),
        removedMargins: [...new Set(marginLines.removed)].slice(0, 4),
        bookmarks,
        outline: outline.length > 0 ? outline : bookmarks,
      },
    }
  } finally {
    try {
      await loadingTask.destroy()
    } catch {
      // already destroyed
    }
  }
}

/**
 * Transform pdf.js text items into top-down page coordinates.
 * @param {any[]} items
 * @param {any} viewport
 */
function normalizeItems(items, viewport) {
  const m = viewport.transform
  const out = []
  for (const item of items) {
    if (typeof item.str !== 'string' || item.str.length === 0) continue
    if (!Array.isArray(item.transform)) continue
    const t = multiply(m, item.transform)
    const text = item.str.replace(/\u00a0/g, ' ')
    if (text.trim().length === 0 && !text.includes(' ')) continue
    out.push({
      text,
      x: t[4],
      y: t[5],
      width: typeof item.width === 'number' ? Math.abs(item.width) : 0,
      height: typeof item.height === 'number' ? Math.abs(item.height) : 10,
      fontName: item.fontName,
      dir: item.dir,
    })
  }
  return out
}

/** @param {number[]} m1 @param {number[]} m2 */
function multiply(m1, m2) {
  return [
    m1[0] * m2[0] + m1[2] * m2[1],
    m1[1] * m2[0] + m1[3] * m2[1],
    m1[0] * m2[2] + m1[2] * m2[3],
    m1[1] * m2[2] + m1[3] * m2[3],
    m1[0] * m2[4] + m1[2] * m2[5] + m1[4],
    m1[1] * m2[4] + m1[3] * m2[5] + m1[5],
  ]
}

/**
 * Count painted images on a page. Duplicate XObject names count once.
 * @param {any} operatorList
 * @param {Record<string, number>} OPS
 */
function countImages(operatorList, OPS) {
  const names = new Set()
  let inline = 0
  const imageOps = new Set([
    OPS.paintImageXObject,
    OPS.paintImageXObjectRepeat,
    OPS.paintInlineImageXObject,
    OPS.paintImageMaskXObject,
  ].filter((value) => typeof value === 'number'))
  for (let index = 0; index < operatorList.fnArray.length; index += 1) {
    const fn = operatorList.fnArray[index]
    if (fn === OPS.paintInlineImageXObject) {
      inline += 1
    } else if (fn === OPS.paintImageXObject || fn === OPS.paintImageXObjectRepeat) {
      const name = operatorList.argsArray[index]?.[0]
      if (typeof name === 'string') names.add(name)
    } else if (fn === OPS.paintImageMaskXObject) {
      inline += 1
    }
  }
  return names.size + inline
}

/**
 * Group items into visual lines by baseline, then order each line by x.
 * @param {any[]} items
 */
function buildLines(items) {
  if (items.length === 0) return []
  const sorted = [...items].sort((a, b) => a.y - b.y || a.x - b.x)
  const heights = sorted.map((item) => item.height).sort((a, b) => a - b)
  const medianHeight = heights[Math.floor(heights.length / 2)] || 10
  const tolerance = Math.max(1.5, medianHeight * 0.35)
  /** @type {Array<{ y: number, items: any[] }>} */
  const groups = []
  for (const item of sorted) {
    const current = groups[groups.length - 1]
    if (current !== undefined && Math.abs(item.y - current.y) <= tolerance) {
      current.items.push(item)
      current.y = (current.y * (current.items.length - 1) + item.y) / current.items.length
    } else {
      groups.push({ y: item.y, items: [item] })
    }
  }
  return groups.map((group) => lineFromItems(group.items, group.y))
}

/**
 * @param {any[]} items
 * @param {number} y
 */
function lineFromItems(items, y) {
  const sorted = [...items].sort((a, b) => a.x - b.x)
  let text = ''
  let previous
  for (const item of sorted) {
    const value = item.text
    if (previous === undefined) {
      text += value
    } else {
      const gap = item.x - (previous.x + previous.width)
      const fontSize = Math.max(previous.height, item.height, 1)
      const left = previous.text.at(-1) ?? ''
      const right = value[0] ?? ''
      const bothLatin = /[A-Za-z0-9]$/.test(left) && /^[A-Za-z0-9]/.test(right)
      const cjkLatin = (isCJK(left) && /[A-Za-z]/.test(right)) || (/[A-Za-z]/.test(left) && isCJK(right))
      const needsSpace = gap > Math.max(1.5, fontSize * 0.25)
        && !(isCJK(left) && isCJK(right))
        && !/\s$/.test(text)
        && !/^\s/.test(value)
        && (bothLatin || (cjkLatin && gap > fontSize * 0.35) || gap > fontSize * 0.8)
      if (needsSpace) text += ' '
      text += value
    }
    previous = item
  }
  const sizes = sorted.map((item) => item.height).sort((a, b) => a - b)
  return {
    y,
    xStart: sorted[0].x,
    xEnd: Math.max(...sorted.map((item) => item.x + item.width)),
    text: normalizePdfText(text),
    size: sizes[Math.floor(sizes.length / 2)] || 10,
    items: sorted,
  }
}

/** @param {Array<any>} pages */
function estimateBodySize(pages) {
  const counts = new Map()
  for (const page of pages) {
    for (const item of page.items) {
      const key = Math.round(item.height * 2) / 2
      counts.set(key, (counts.get(key) ?? 0) + 1)
    }
  }
  let best = 10
  let bestCount = 0
  for (const [size, count] of counts) {
    if (count > bestCount) {
      best = size
      bestCount = count
    }
  }
  return best
}

/**
 * Repeated top/bottom margin lines and page-number lines.
 * @param {Array<any>} pages
 */
function collectRepeatedMargins(pages) {
  const counts = new Map()
  const removed = []
  for (const page of pages) {
    const lines = buildLines(page.items)
    const margin = lines.filter((line) => isMarginLine(line, page.height))
    const seen = new Set()
    for (const line of margin) {
      const key = normalizeMargin(line.text)
      if (key.length === 0 || seen.has(key)) continue
      seen.add(key)
      counts.set(key, (counts.get(key) ?? 0) + 1)
    }
  }
  const repeated = new Set()
  const threshold = Math.max(2, Math.ceil(pages.length * 0.3))
  for (const [key, count] of counts) {
    if (count >= threshold) repeated.add(key)
  }
  return { repeated, removed }
}

/** @param {number} y @param {number} height */
function isMarginLine(line, height) {
  return line.y < height * 0.08 || line.y > height * 0.92
}

/** @param {string} text */
function normalizeMargin(text) {
  return text
    .replace(/第\s*\d+\s*页/g, '')
    .replace(/page\s*\d+/gi, '')
    .replace(/\d+/g, '#')
    .replace(/[\s\-—–_·.]+/g, '')
    .replace(/[（(]\s*[#]?\s*[）)]/g, '')
    .slice(0, 80)
}

/**
 * Table-like pages keep visual lines and column separators instead of a
 * fabricated Markdown table.
 * @param {any} page
 * @param {any[]} lines
 */
function isTableLikePage(page, lines) {
  if (lines.length < 4) return false
  const xClusters = countXClusters(page.items)
  const averageItems = page.items.length / lines.length
  const averageChars = lines.reduce((sum, line) => sum + line.text.length, 0) / lines.length
  return (averageItems >= 2.5 && averageChars < 22)
    || (xClusters >= 6 && averageChars < 45)
}

/** @param {any[]} items */
function countXClusters(items) {
  if (items.length === 0) return 0
  const intervals = items.map((item) => [item.x, item.x + Math.max(item.width, 1)]).sort((a, b) => a[0] - b[0])
  let clusters = 1
  let end = intervals[0][1]
  for (let index = 1; index < intervals.length; index += 1) {
    const [start, stop] = intervals[index]
    if (start > end + 4) {
      clusters += 1
      end = stop
    } else {
      end = Math.max(end, stop)
    }
  }
  return clusters
}

/**
 * @param {any[]} lines
 */
function layoutLines(lines) {
  const out = []
  for (const line of lines) {
    let text = ''
    let previous
    for (const item of line.items) {
      if (previous === undefined) {
        text += item.text
      } else {
        const gap = item.x - (previous.x + previous.width)
        const separator = gap > Math.max(6, line.size * 0.5) ? ' │ ' : ' '
        text += separator + item.text
      }
      previous = item
    }
    const normalized = collapseWhitespace(text)
    if (normalized.length > 0) out.push(normalized)
  }
  return out
}

/**
 * Split page items into up to two columns at a significant central gutter.
 * Doing this before line grouping keeps text from opposite columns out of the
 * same visual line.
 * @param {any[]} items
 * @param {number} pageWidth
 */
function splitItemColumns(items, pageWidth) {
  if (items.length < 12) return [items]
  const intervals = items
    .map((item) => [item.x, item.x + Math.max(item.width, 1)])
    .sort((a, b) => a[0] - b[0])
  const merged = []
  for (const [start, end] of intervals) {
    const last = merged[merged.length - 1]
    if (last !== undefined && start <= last[1] + 1) last[1] = Math.max(last[1], end)
    else merged.push([start, end])
  }
  let bestGap = 0
  let bestStart = 0
  for (let index = 1; index < merged.length; index += 1) {
    const gap = merged[index][0] - merged[index - 1][1]
    if (gap > bestGap && gap > Math.max(18, pageWidth * 0.04)) {
      bestGap = gap
      bestStart = merged[index - 1][1]
    }
  }
  if (bestGap === 0) return [items]
  const center = bestStart + bestGap / 2
  if (center < pageWidth * 0.25 || center > pageWidth * 0.75) return [items]
  const left = items.filter((item) => item.x < center)
  const right = items.filter((item) => item.x >= center)
  if (left.length < 4 || right.length < 4) return [items]
  return [left, right]
}

/**
 * @param {any[]} blocks
 * @param {any[]} lines
 * @param {{ number: number, width: number, height: number }} page
 * @param {number} bodySize
 * @param {Map<string, string>} bookmarkTitles
 */
function appendPageColumn(blocks, lines, page, bodySize, bookmarkTitles) {
  const sorted = [...lines].sort((a, b) => a.y - b.y)
  /** @type {{ lines: any[], text: string, heading: boolean, level: number, page: number } | undefined} */
  let paragraph
  const medianHeight = median(sorted.map((line) => line.size)) || bodySize
  const headingSizes = new Set()

  const flush = () => {
    if (paragraph === undefined) return
    const text = paragraph.text.trim()
    if (text.length > 0) {
      if (paragraph.heading) {
        blocks.push({ type: 'heading', level: paragraph.level, text, plain: text, page: paragraph.page })
      } else {
        blocks.push({ type: 'paragraph', text, plain: text, page: paragraph.page })
      }
    }
    paragraph = undefined
  }

  for (let index = 0; index < sorted.length; index += 1) {
    const line = sorted[index]
    const text = line.text.trim()
    if (text.length === 0) continue
    const heading = isListRun(sorted, index) ? false : isHeading(line, bodySize, bookmarkTitles)
    if (heading) {
      flush()
      const bookmark = bookmarkTitles.get(normalizeTitle(text))
      const headingText = bookmark ?? text
      headingSizes.add(Math.round(line.size * 2) / 2)
      const level = headingLevel(headingText, line.size, bodySize, headingSizes)
      blocks.push({ type: 'heading', level, text: headingText, plain: headingText, page: page.number })
      continue
    }
    if (paragraph !== undefined) {
      const previousLine = sorted[index - 1]
      const gap = previousLine === undefined ? 0 : line.y - previousLine.y
      const startsNew = gap > medianHeight * 1.8
        || /[。！？!?；;]$/.test(paragraph.lines.at(-1)?.text ?? '')
        || /^[•·▪◦\-*]\s|^\d+[.、)]\s|^[（(]\d+[）)]/.test(text)
      if (!startsNew) {
        paragraph.text = joinText(paragraph.text, text)
        paragraph.lines.push(line)
        continue
      }
      flush()
    }
    paragraph = {
      lines: [line],
      text,
      heading: false,
      level: 1,
      page: page.number,
    }
  }
  flush()
}

/** @param {any} line @param {number} bodySize @param {Map<string, string>} bookmarkTitles */
function isHeading(line, bodySize, bookmarkTitles) {
  const text = line.text.trim()
  if (text.length === 0) return false
  if (bookmarkTitles.size > 0) {
    if (bookmarkTitles.has(normalizeTitle(text))) return true
    return line.size >= bodySize * 1.6
  }
  if (line.size >= bodySize * 1.35) return true
  if (line.size >= bodySize * 1.15 && text.length <= 30) return true
  if (text.length > 0 && text.length <= 30 && !/[。；;]$/.test(text)) {
    if (/^第[一二三四五六七八九十百\d]+[章节篇部]/.test(text)) return true
    if (/^[一二三四五六七八九十百]+、/.test(text)) return true
    if (/^[（(][一二三四五六七八九十百]+[）)]/.test(text)) return true
    if (/^\d+(?:\.\d+){0,5}[.、]\s*\S/.test(text)) return true
    if (/^\d+(?:\.\d+){0,5}\s+\S/.test(text)) return true
  }
  return false
}

/**
 * A run of three or more consecutive list-marker lines is a list, not a set
 * of headings. This keeps `1. ... 2. ... 3. ...` out of the outline while
 * leaving a lone `3、section title` eligible as a heading.
 * @param {any[]} lines
 * @param {number} index
 */
function isListRun(lines, index) {
  const pattern = /^(?:\d+[.、)]|[（(]\d+[）)]|[-*•])\s*/
  if (!pattern.test(lines[index].text.trim())) return false
  let count = 1
  for (let cursor = index + 1; cursor < lines.length && cursor <= index + 2; cursor += 1) {
    if (pattern.test(lines[cursor].text.trim())) count += 1
    else break
  }
  return count >= 3
}

/** @param {string} text @param {number} size @param {number} bodySize @param {Set<number>} sizes */
function headingLevel(text, size, bodySize, sizes) {
  const numeric = /^(\d+(?:\.\d+){0,5})(?:[.、]\s*|\s+)\S/.exec(text)
  if (numeric !== null) return Math.min(6, numeric[1].split('.').length)
  if (/^第[一二三四五六七八九十百\d]+[章篇部]/.test(text)) return 1
  if (/^第[一二三四五六七八九十百\d]+节/.test(text)) return 2
  if (/^[一二三四五六七八九十百]+、/.test(text)) return 1
  if (/^[（(][一二三四五六七八九十百]+[）)]/.test(text)) return 2
  const ordered = [...sizes].sort((a, b) => b - a)
  const rank = ordered.findIndex((value) => Math.abs(value - size) < 0.6)
  return Math.max(1, Math.min(3, rank === -1 ? 1 : rank + 1))
}

/** @param {string} text */
function normalizeTitle(text) {
  return text.replace(/\s+/g, '').replace(/[：:·]/g, '').toLowerCase()
}

/** @param {string} text */
function isWatermark(text) {
  return /^(严禁复制|严禁外传|内部资料|仅供内部使用|机密|confidential)[。.!！]?$/i.test(text.trim())
}

/** @param {number[]} values */
function median(values) {
  if (values.length === 0) return 0
  const sorted = [...values].sort((a, b) => a - b)
  return sorted[Math.floor(sorted.length / 2)]
}

/**
 * Flatten pdf.js bookmarks into page/level entries.
 * @param {any} doc
 */
async function extractBookmarks(doc) {
  const outline = await doc.getOutline().catch(() => undefined)
  if (!Array.isArray(outline) || outline.length === 0) return []
  const entries = []
  const walk = async (items, level) => {
    for (const item of items) {
      let page
      try {
        const destination = typeof item.dest === 'string' ? await doc.getDestination(item.dest) : item.dest
        if (Array.isArray(destination) && destination[0] !== undefined) {
          page = (await doc.getPageIndex(destination[0])) + 1
        }
      } catch {
        page = undefined
      }
      entries.push({
        level,
        text: String(item.title ?? '').trim(),
        ...(page === undefined ? {} : { page }),
      })
      if (Array.isArray(item.items) && item.items.length > 0) await walk(item.items, level + 1)
    }
  }
  await walk(outline, 1)
  return entries
}

/**
 * Build an outline from inferred headings, carrying page numbers.
 * @param {any[]} blocks
 */
function inferOutline(blocks) {
  const outline = []
  let page
  for (const block of blocks) {
    if (block.type === 'page') {
      page = block.page
      continue
    }
    if (block.type === 'heading') {
      outline.push({
        level: block.level,
        text: block.plain || block.text,
        ...(page === undefined ? {} : { page }),
      })
    }
  }
  return outline
}
