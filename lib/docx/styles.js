import { attr, child, findAll, onOff, parseXml } from '../xml.js'

/**
 * @typedef {{
 *   bold?: boolean,
 *   italic?: boolean,
 *   strike?: boolean,
 *   size?: number,
 *   font?: string,
 *   outlineLvl?: number,
 *   numId?: number,
 *   ilvl?: number,
 *   jc?: string,
 *   before?: number,
 *   after?: number,
 *   line?: number,
 *   firstLine?: number,
 * }} ResolvedStyle
 */

/** @param {string | undefined} value */
function int(value) {
  if (value === undefined) return undefined
  const parsed = Number.parseInt(value, 10)
  return Number.isFinite(parsed) ? parsed : undefined
}

/** @param {import('../xml.js').XmlElement | undefined} rPr */
export function runProps(rPr) {
  if (rPr === undefined) return {}
  const b = child(rPr, 'w:b')
  const i = child(rPr, 'w:i')
  const strike = child(rPr, 'w:strike')
  const sz = child(rPr, 'w:sz')
  const fonts = child(rPr, 'w:rFonts')
  return {
    ...(b !== undefined ? { bold: onOff(b) } : {}),
    ...(i !== undefined ? { italic: onOff(i) } : {}),
    ...(strike !== undefined ? { strike: onOff(strike) } : {}),
    ...(sz !== undefined ? { size: int(attr(sz, 'val')) } : {}),
    ...(fonts !== undefined ? { font: attr(fonts, 'eastAsia') ?? attr(fonts, 'ascii') } : {}),
  }
}

/** @param {import('../xml.js').XmlElement | undefined} pPr */
export function paragraphProps(pPr) {
  if (pPr === undefined) return {}
  const outline = child(pPr, 'w:outlineLvl')
  const numPr = child(pPr, 'w:numPr')
  const jc = child(pPr, 'w:jc')
  const spacing = child(pPr, 'w:spacing')
  const ind = child(pPr, 'w:ind')
  let numId
  let ilvl
  if (numPr !== undefined) {
    const numIdEl = child(numPr, 'w:numId')
    const ilvlEl = child(numPr, 'w:ilvl')
    numId = int(attr(numIdEl, 'val'))
    ilvl = int(attr(ilvlEl, 'val'))
  }
  return {
    ...(outline !== undefined ? { outlineLvl: int(attr(outline, 'val')) } : {}),
    ...(numId !== undefined && numId !== 0 ? { numId } : {}),
    ...(ilvl !== undefined ? { ilvl } : {}),
    ...(jc !== undefined ? { jc: attr(jc, 'val') } : {}),
    ...(spacing !== undefined
      ? {
          before: int(attr(spacing, 'before')),
          after: int(attr(spacing, 'after')),
          line: int(attr(spacing, 'line')),
        }
      : {}),
    ...(ind !== undefined ? { firstLine: int(attr(ind, 'firstLine')) } : {}),
  }
}

/**
 * Resolve paragraph/character style inheritance for a DOCX package.
 */
export class StyleResolver {
  /**
   * @param {string | undefined} stylesXml
   * @param {string | undefined} documentXml
   */
  constructor(stylesXml, documentXml) {
    /** @type {Map<string, { id: string, name: string, type: string, basedOn?: string, pPr?: import('../xml.js').XmlElement, rPr?: import('../xml.js').XmlElement }>} */
    this.styles = new Map()
    /** @type {Map<string, ResolvedStyle>} */
    this.cache = new Map()
    this.defaultSize = 21
    this.defaultFont = undefined
    this.defaultParagraph = {}

    if (stylesXml !== undefined) {
      const root = parseXml(stylesXml)
      if (root !== undefined) {
        const docDefaults = findAll(root, 'w:docDefaults')[0]
        if (docDefaults !== undefined) {
          const rPrDefault = findAll(docDefaults, 'w:rPrDefault')[0]
          const rPr = rPrDefault === undefined ? undefined : findAll(rPrDefault, 'w:rPr')[0]
          const props = runProps(rPr)
          if (props.size !== undefined) this.defaultSize = props.size
          if (props.font !== undefined) this.defaultFont = props.font
          const pPrDefault = findAll(docDefaults, 'w:pPrDefault')[0]
          const pPr = pPrDefault === undefined ? undefined : findAll(pPrDefault, 'w:pPr')[0]
          this.defaultParagraph = paragraphProps(pPr)
        }
        for (const style of findAll(root, 'w:style')) {
          const id = attr(style, 'styleId')
          if (id === undefined) continue
          const nameEl = child(style, 'w:name')
          const basedOnEl = child(style, 'w:basedOn')
          this.styles.set(id, {
            id,
            name: attr(nameEl, 'val') ?? id,
            type: attr(style, 'type') ?? 'paragraph',
            basedOn: attr(basedOnEl, 'val'),
            pPr: child(style, 'w:pPr'),
            rPr: child(style, 'w:rPr'),
          })
        }
      }
    }

    // A paragraph can carry direct formatting that is not in styles.xml.
    if (documentXml !== undefined) {
      const root = parseXml(documentXml)
      if (root !== undefined) {
        const body = findAll(root, 'w:body')[0]
        const firstParagraph = body === undefined ? undefined : findAll(body, 'w:p')[0]
        const pPr = firstParagraph === undefined ? undefined : child(firstParagraph, 'w:pPr')
        const rPr = pPr === undefined ? undefined : child(pPr, 'w:rPr')
        const props = runProps(rPr)
        if (props.size !== undefined && this.defaultSize === 21) this.defaultSize = props.size
      }
    }
  }

  /** @param {string} id */
  get(id) {
    return this.styles.get(id)
  }

  /**
   * Resolve a style id, merging its basedOn chain. Direct paragraph/run
   * formatting is applied by the caller on top.
   * @param {string | undefined} id
   * @returns {ResolvedStyle}
   */
  resolve(id) {
    if (id === undefined) return {}
    const cached = this.cache.get(id)
    if (cached !== undefined) return cached
    const seen = new Set()
    let current = this.styles.get(id)
    const chain = []
    while (current !== undefined && !seen.has(current.id)) {
      seen.add(current.id)
      chain.push(current)
      current = current.basedOn === undefined ? undefined : this.styles.get(current.basedOn)
    }
    /** @type {ResolvedStyle} */
    const merged = {}
    for (let index = chain.length - 1; index >= 0; index -= 1) {
      const style = chain[index]
      const p = paragraphProps(style.pPr)
      const r = runProps(style.rPr)
      if (r.bold !== undefined) merged.bold = r.bold
      if (r.italic !== undefined) merged.italic = r.italic
      if (r.strike !== undefined) merged.strike = r.strike
      if (r.size !== undefined) merged.size = r.size
      if (r.font !== undefined) merged.font = r.font
      if (p.outlineLvl !== undefined) merged.outlineLvl = p.outlineLvl
      if (p.numId !== undefined) merged.numId = p.numId
      if (p.ilvl !== undefined) merged.ilvl = p.ilvl
      if (p.jc !== undefined) merged.jc = p.jc
      if (p.before !== undefined) merged.before = p.before
      if (p.after !== undefined) merged.after = p.after
      if (p.line !== undefined) merged.line = p.line
      if (p.firstLine !== undefined) merged.firstLine = p.firstLine
    }
    this.cache.set(id, merged)
    return merged
  }

  /** @param {string | undefined} id */
  styleName(id) {
    return id === undefined ? undefined : this.styles.get(id)?.name
  }

  /**
   * Heading level from a paragraph style: 1..9, 0 for a Title style, or
   * undefined for body styles.
   * @param {string | undefined} id
   */
  headingLevel(id) {
    if (id === undefined) return undefined
    const style = this.styles.get(id)
    if (style === undefined) return undefined
    const name = style.name.trim()
    const lower = name.toLowerCase()
    if (lower === 'title' || name === '标题' || lower === 'document title') return 0
    const byName = /^(?:heading|标题)\s*([1-9])$/i.exec(name)
    if (byName !== null) return Number(byName[1])
    const byId = /^Heading([1-9])$/i.exec(id)
    if (byId !== null) return Number(byId[1])
    if (lower.startsWith('heading ')) {
      const number = Number.parseInt(lower.slice('heading '.length), 10)
      if (Number.isFinite(number)) return number
    }
    return undefined
  }

  /** @param {string | undefined} id */
  isTocStyle(id) {
    if (id === undefined) return false
    const name = (this.styleName(id) ?? '').toLowerCase()
    return name.startsWith('toc') || name.includes('目录') || name.includes('目錄')
  }

  /** @param {import('../xml.js').XmlElement | undefined} pPr */
  paragraphFormatting(pPr) {
    return { ...this.defaultParagraph, ...paragraphProps(pPr) }
  }

  /** @param {import('../xml.js').XmlElement | undefined} rPr */
  runFormatting(rPr) {
    return { ...runProps(rPr) }
  }

  /** Body font size in half-points (for heading-ratio heuristics). */
  bodySize() {
    return this.defaultSize
  }
}
