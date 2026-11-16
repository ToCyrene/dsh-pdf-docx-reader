/** Text helpers shared by the PDF and DOCX pipelines. */

const CJK_RANGES = [
  [0x2e80, 0x2fff],
  [0x3000, 0x303f],
  [0x3040, 0x30ff],
  [0x3400, 0x4dbf],
  [0x4e00, 0x9fff],
  [0xac00, 0xd7af],
  [0xf900, 0xfaff],
  [0xfe30, 0xfe4f],
  [0xff00, 0xffef],
  [0x20000, 0x2fa1f],
]

const PUNCTUATION = new Set([
  '，', '。', '；', '：', '？', '！', '、', '）', '》', '】', '」', '』',
  ',', '.', ';', ':', '?', '!', ')', ']', '}', '>', '”', '’',
])

/** @param {string} char */
export function isCJK(char) {
  const code = char.codePointAt(0)
  if (code === undefined) return false
  for (const [start, end] of CJK_RANGES) {
    if (code >= start && code <= end) return true
  }
  return false
}

/** @param {string} text */
export function hasCJK(text) {
  for (const char of text) {
    if (isCJK(char)) return true
  }
  return false
}

/** @param {string} char */
function isLatinLetter(char) {
  return /[A-Za-z0-9]/.test(char)
}

/**
 * Join two text fragments as PDF/DOCX prose would join them: no space
 * between CJK characters or before CJK punctuation, a space between Latin
 * words, and no space after a Latin hyphenated fragment.
 * @param {string} left
 * @param {string} right
 */
export function joinText(left, right) {
  if (left.length === 0) return right
  if (right.length === 0) return left
  const a = left.at(-1) ?? ''
  const b = right[0] ?? ''
  if (PUNCTUATION.has(b)) return left + right
  if (isCJK(a) || isCJK(b)) return left + right
  if (/\s$/.test(left) || /^\s/.test(right)) return left + right
  if (a === '-' && isLatinLetter(b)) return left + right
  return `${left} ${right}`
}

/**
 * Normalize PDF text artifacts without flattening CJK punctuation.
 * @param {string} text
 */
export function normalizePdfText(text) {
  return text
    .replace(/\u00ad/g, '')
    .replace(/\u00a0/g, ' ')
    .replace(/\ufb01/g, 'fi')
    .replace(/\ufb02/g, 'fl')
    .replace(/\ufb00/g, 'ff')
    .replace(/\ufb03/g, 'ffi')
    .replace(/\ufb04/g, 'ffl')
    .replace(/[\u200b\u200c\u200d\ufeff]/g, '')
    .replace(/[ \t]+/g, ' ')
    .trim()
}

/**
 * Collapse all whitespace to single spaces; used by search snippets.
 * @param {string} text
 */
export function collapseWhitespace(text) {
  return text.replace(/\s+/g, ' ').trim()
}

/**
 * Escape Markdown inline control characters in user text. Kept deliberately
 * small: the agent reads plain text, and excessive escaping costs tokens.
 * @param {string} text
 */
export function escapeInline(text) {
  return text
    .replace(/([\\`*_[\]])/g, '\\$1')
    .replace(/^(\s*)([#>+-])(\s)/, '$1\\$2$3')
    .replace(/^(\s*)(\d+)([.)])(\s)/, '$1$2\\$3$4')
}

/**
 * Escape a Markdown table cell.
 * @param {string} text
 */
export function escapeTableCell(text) {
  return text.replace(/\\/g, '\\\\').replace(/\|/g, '\\|').replace(/\r?\n/g, '<br>')
}

/**
 * Strip the most common Markdown markers for a plain-text search line.
 * Rendering code supplies the exact plain text where it can, so this is only
 * a fallback for blocks that do not carry one.
 * @param {string} text
 */
export function stripMarkdown(text) {
  return text
    .replace(/^#{1,6}\s+/, '')
    .replace(/^\s*[-*+]\s+/, '')
    .replace(/\[([^\]]*)\]\([^)]*\)/g, '$1')
    .replace(/[*_~`]/g, '')
}

/**
 * @param {string} text
 * @param {number} max
 */
export function truncateText(text, max) {
  if (text.length <= max) return text
  return `${text.slice(0, max)}…`
}
