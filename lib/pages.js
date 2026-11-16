import { DocReaderError } from './errors.js'

/**
 * Parse a forgiving PDF page selection: "3", "3-5", "1,3-5", "第3-5页",
 * "3~5", "3—5". Returns null when no selection is given.
 * @param {unknown} input
 * @param {number | undefined} totalPages
 * @returns {{ pages: Set<number>, ranges: Array<[number, number]>, text: string } | null}
 */
export function parsePageSelection(input, totalPages) {
  if (input === undefined || input === null) return null
  const raw = String(input).trim()
  if (raw.length === 0) return null
  const normalized = raw
    .replace(/[第页\s　]/g, '')
    .replace(/[~～—–－至到]/g, '-')
    .replace(/[，,]/g, ',')
  const pages = new Set()
  const ranges = []
  for (const part of normalized.split(',')) {
    if (part.length === 0) continue
    const match = /^(\d+)(?:-(\d+))?$/.exec(part)
    if (match === null) {
      throw new DocReaderError(
        `Invalid pages "${raw}"; use "3", "3-5", "1,3-5", or "第3-5页"`,
        'INVALID_PAGES',
      )
    }
    let start = Number.parseInt(match[1], 10)
    let end = match[2] === undefined ? start : Number.parseInt(match[2], 10)
    if (start > end) [start, end] = [end, start]
    if (totalPages !== undefined && (start < 1 || end > totalPages)) {
      throw new DocReaderError(
        `pages ${raw} is outside this PDF (1-${totalPages})`,
        'INVALID_PAGES',
      )
    }
    for (let page = start; page <= end; page += 1) pages.add(page)
    ranges.push([start, end])
  }
  if (pages.size === 0) {
    throw new DocReaderError(`Invalid pages "${raw}"; use "3", "3-5", or "第3-5页"`, 'INVALID_PAGES')
  }
  return { pages, ranges, text: raw }
}

/**
 * @param {number | undefined} page
 * @param {{ pages: Set<number> } | null} selection
 */
export function pageSelected(page, selection) {
  if (selection === null) return true
  if (page === undefined) return false
  return selection.pages.has(page)
}
