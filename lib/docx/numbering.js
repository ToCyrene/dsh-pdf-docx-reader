import { attr, child, childrenOf, findAll, parseXml } from '../xml.js'

const ROMAN = [
  [1000, 'M'], [900, 'CM'], [500, 'D'], [400, 'CD'], [100, 'C'], [90, 'XC'],
  [50, 'L'], [40, 'XL'], [10, 'X'], [9, 'IX'], [5, 'V'], [4, 'IV'], [1, 'I'],
]

/** @param {number} value */
function roman(value) {
  let remaining = Math.max(1, Math.floor(value))
  let out = ''
  for (const [amount, glyph] of ROMAN) {
    while (remaining >= amount) {
      out += glyph
      remaining -= amount
    }
  }
  return out
}

/** @param {number} value */
function letters(value) {
  let remaining = Math.max(1, Math.floor(value))
  let out = ''
  while (remaining > 0) {
    remaining -= 1
    out = String.fromCharCode(97 + (remaining % 26)) + out
    remaining = Math.floor(remaining / 26)
  }
  return out
}

const CHINESE_DIGITS = ['零', '一', '二', '三', '四', '五', '六', '七', '八', '九']
const CHINESE_UNITS = ['', '十', '百', '千']

/** @param {number} value */
export function chineseNumber(value) {
  const n = Math.floor(value)
  if (!Number.isFinite(n) || n <= 0) return String(value)
  if (n < 10) return CHINESE_DIGITS[n]
  if (n < 20) return `十${n % 10 === 0 ? '' : CHINESE_DIGITS[n % 10]}`
  if (n < 10000) {
    const digits = String(n)
    let out = ''
    let pendingZero = false
    for (let index = 0; index < digits.length; index += 1) {
      const digit = Number(digits[index])
      const unit = CHINESE_UNITS[digits.length - 1 - index]
      if (digit === 0) {
        pendingZero = out.length > 0
        continue
      }
      if (pendingZero) {
        out += '零'
        pendingZero = false
      }
      out += CHINESE_DIGITS[digit] + unit
    }
    return out
  }
  return String(n)
}

/**
 * Format one numbering counter.
 * @param {number} value
 * @param {string} format
 * @param {boolean} legal
 */
export function formatNumber(value, format, legal = false) {
  if (legal) return String(value)
  switch (format) {
    case 'decimal':
      return String(value)
    case 'decimalZero':
      return value < 10 ? `0${value}` : String(value)
    case 'upperRoman':
      return roman(value)
    case 'lowerRoman':
      return roman(value).toLowerCase()
    case 'upperLetter':
      return letters(value).toUpperCase()
    case 'lowerLetter':
      return letters(value)
    case 'chineseCounting':
    case 'chineseCountingThousand':
    case 'chineseLegalSimplified':
    case 'ideographTraditional':
    case 'ideographZodiac':
      return chineseNumber(value)
    case 'ordinal':
    case 'cardinalText':
    case 'ordinalText':
      return String(value)
    case 'none':
      return ''
    default:
      return String(value)
  }
}

/** @param {string} text */
function bulletGlyph(text) {
  if (text.length === 0) return '-'
  const code = text.codePointAt(0) ?? 0
  if (code >= 0xf000 && code <= 0xf0ff) {
    if (code === 0xf06c || code === 0xf06e) return '-'
    if (code === 0xf0a7 || code === 0xf0d8) return '-'
    return '-'
  }
  if (text === '•' || text === '·' || text === '▪' || text === '◦') return '-'
  return text
}

/**
 * WordprocessingML numbering definitions plus per-numId counters.
 */
export class NumberingResolver {
  /** @param {string | undefined} numberingXml */
  constructor(numberingXml) {
    /** @type {Map<string, Map<number, { start: number, numFmt: string, lvlText: string, suff?: string, isLgl: boolean }>>} */
    this.abstracts = new Map()
    /** @type {Map<number, { abstractId: string, overrides: Map<number, { start?: number, lvlText?: string }> }>} */
    this.nums = new Map()
    /** @type {Map<number, number[]>} */
    this.counters = new Map()
    if (numberingXml === undefined) return
    const root = parseXml(numberingXml)
    if (root === undefined) return
    for (const abstract of findAll(root, 'w:abstractNum')) {
      const id = attr(abstract, 'abstractNumId')
      if (id === undefined) continue
      const levels = new Map()
      for (const level of childrenOf(abstract, 'w:lvl')) {
        const ilvl = Number.parseInt(attr(level, 'ilvl') ?? '0', 10)
        if (!Number.isFinite(ilvl)) continue
        const start = Number.parseInt(attr(child(level, 'w:start'), 'val') ?? '1', 10)
        const numFmt = attr(child(level, 'w:numFmt'), 'val') ?? 'decimal'
        const lvlText = attr(child(level, 'w:lvlText'), 'val') ?? ''
        const suff = attr(child(level, 'w:suff'), 'val')
        const isLgl = child(level, 'w:isLgl') !== undefined
        levels.set(ilvl, { start: Number.isFinite(start) ? start : 1, numFmt, lvlText, suff, isLgl })
      }
      this.abstracts.set(id, levels)
    }
    for (const num of findAll(root, 'w:num')) {
      const numId = Number.parseInt(attr(num, 'numId') ?? '', 10)
      if (!Number.isFinite(numId)) continue
      const abstractId = attr(child(num, 'w:abstractNumId'), 'val')
      if (abstractId === undefined) continue
      const overrides = new Map()
      for (const override of childrenOf(num, 'w:lvlOverride')) {
        const ilvl = Number.parseInt(attr(override, 'ilvl') ?? '', 10)
        if (!Number.isFinite(ilvl)) continue
        const startOverride = child(override, 'w:startOverride')
        const levelOverride = child(override, 'w:lvl')
        const entry = {}
        if (startOverride !== undefined) {
          const start = Number.parseInt(attr(startOverride, 'val') ?? '', 10)
          if (Number.isFinite(start)) entry.start = start
        }
        if (levelOverride !== undefined) {
          const lvlText = attr(child(levelOverride, 'w:lvlText'), 'val')
          if (lvlText !== undefined) entry.lvlText = lvlText
        }
        overrides.set(ilvl, entry)
      }
      this.nums.set(numId, { abstractId, overrides })
    }
  }

  /** @param {number | undefined} numId */
  has(numId) {
    return numId !== undefined && this.nums.has(numId)
  }

  /**
   * The effective level definition for one numId/ilvl.
   * @param {number} numId
   * @param {number} ilvl
   */
  level(numId, ilvl) {
    const num = this.nums.get(numId)
    if (num === undefined) return undefined
    const abstract = this.abstracts.get(num.abstractId)
    const level = abstract?.get(ilvl)
    if (level === undefined) return undefined
    const override = num.overrides.get(ilvl)
    return {
      ...level,
      ...(override?.start !== undefined ? { start: override.start } : {}),
      ...(override?.lvlText !== undefined ? { lvlText: override.lvlText } : {}),
    }
  }

  /**
   * Whether this paragraph's numbering is a bullet (and therefore a list,
   * never a heading).
   * @param {number} numId
   * @param {number} ilvl
   */
  isBullet(numId, ilvl) {
    const level = this.level(numId, ilvl)
    return level !== undefined && (level.numFmt === 'bullet' || level.numFmt === 'none')
  }

  /**
   * Advance the counter for this paragraph and return its visible label.
   * Must be called exactly once per numbered paragraph, in document order.
   * @param {number} numId
   * @param {number} ilvl
   * @returns {{ label: string, marker: string, bullet: boolean } | undefined}
   */
  consume(numId, ilvl) {
    const num = this.nums.get(numId)
    const level = this.level(numId, ilvl)
    if (num === undefined || level === undefined) return undefined
    let counters = this.counters.get(numId)
    if (counters === undefined) {
      counters = []
      this.counters.set(numId, counters)
    }
    for (let index = 0; index < ilvl; index += 1) {
      if (counters[index] === undefined) {
        const startLevel = this.level(numId, index)
        counters[index] = startLevel?.start ?? 1
      }
    }
    if (counters[ilvl] === undefined) counters[ilvl] = level.start - 1
    counters[ilvl] = (counters[ilvl] ?? 0) + 1
    for (let index = ilvl + 1; index < counters.length; index += 1) counters[index] = undefined

    if (level.numFmt === 'bullet') {
      return { label: '', marker: bulletGlyph(level.lvlText), bullet: true }
    }
    if (level.numFmt === 'none') return { label: '', marker: '', bullet: false }

    const label = level.lvlText.replace(/%([1-9])/g, (_match, digitText) => {
      const index = Number(digitText) - 1
      const reference = this.level(numId, index)
      const value = counters[index] ?? reference?.start ?? 1
      return formatNumber(value, reference?.numFmt ?? 'decimal', level.isLgl || reference?.isLgl === true)
    })
    const suffix = level.suff === 'nothing' ? '' : ' '
    return { label: `${label}${suffix}`, marker: '', bullet: false }
  }

  /** Reset all counters; used only by tests. */
  reset() {
    this.counters.clear()
  }
}
