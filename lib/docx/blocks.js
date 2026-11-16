import { DocReaderError } from '../errors.js'
import {
  attr,
  child,
  childrenOf,
  elementChildren,
  findAll,
  isElement,
  isTag,
  parseXml,
  textContent,
} from '../xml.js'
import { escapeInline, normalizePdfText } from '../text.js'
import { NumberingResolver } from './numbering.js'
import { StyleResolver, paragraphProps, runProps } from './styles.js'
import { readDocxPackage } from './unzip.js'

const MAX_TABLE_ROWS = 500
const MAX_TABLE_COLUMNS = 20
const MAX_FOOTNOTES = 20

/** @param {import('../xml.js').XmlElement | undefined} element */
function textOf(element) {
  if (element === undefined) return ''
  return normalizePdfText(textContent(element))
}

/**
 * Parse a DOCX package into the shared block model.
 * @param {Uint8Array} bytes
 * @param {{ inferHeadings: boolean }} config
 * @returns {{ blocks: any[], meta: Record<string, any> }}
 */
export function parseDocx(bytes, config) {
  const { parts, headerParts, footerParts } = readDocxPackage(bytes)
  const documentXml = parts.get('word/document.xml')
  if (documentXml === undefined) throw new DocReaderError('DOCX has no word/document.xml', 'DOCX_INVALID')

  const styles = new StyleResolver(parts.get('word/styles.xml'), documentXml)
  const numbering = new NumberingResolver(parts.get('word/numbering.xml'))
  const rels = parseRelationships(parts.get('word/_rels/document.xml.rels'))
  const footnotes = parseNotes(parts.get('word/footnotes.xml'))
  const endnotes = parseNotes(parts.get('word/endnotes.xml'))
  const coreTitle = parseCoreTitle(parts.get('docProps/core.xml'))

  const root = parseXml(documentXml)
  const body = root === undefined ? undefined : findAll(root, 'w:body')[0]
  if (body === undefined) throw new DocReaderError('DOCX document body is missing', 'DOCX_INVALID')

  const elements = elementChildren(body)
  const toc = detectToc(elements, styles)
  const tocSet = toc?.indices ?? new Set()
  addAuxiliaryTocIndices(elements, styles, tocSet)
  const ctx = {
    styles,
    numbering,
    rels,
    footnotes,
    endnotes,
    config,
    imageCounter: 0,
    footnoteRefs: new Set(),
    textboxCounter: 0,
    pageBreaks: 0,
    headingCount: 0,
    tableCount: 0,
    lastHeadingLevel: 0,
    lastWasHeading: false,
    parenCnLevel: undefined,
    currentLevels: [],
  }

  /** @type {any[]} */
  const blocks = []

  for (let index = 0; index < elements.length; index += 1) {
    const element = elements[index]
    if (isTag(element, 'w:p')) {
      if (tocSet.has(index)) continue
      const parsed = parseParagraph(element, ctx, { inHeading: false })
      if (parsed === undefined) continue
      emitParagraphBlocks(parsed, ctx, blocks)
    } else if (isTag(element, 'w:tbl')) {
      const table = parseTable(element, ctx)
      if (table !== undefined) blocks.push(table)
    } else if (isTag(element, 'w:sdt')) {
      const content = child(element, 'w:sdtContent')
      if (content !== undefined) {
        for (const nested of elementChildren(content)) {
          if (isTag(nested, 'w:p')) {
            const parsed = parseParagraph(nested, ctx, { inHeading: false })
            if (parsed !== undefined) emitParagraphBlocks(parsed, ctx, blocks)
          } else if (isTag(nested, 'w:tbl')) {
            const table = parseTable(nested, ctx)
            if (table !== undefined) blocks.push(table)
          }
        }
      }
    } else if (isTag(element, 'w:altChunk')) {
      blocks.push({ type: 'paragraph', text: '[altChunk not extracted]', plain: 'altChunk' })
    }
  }

  appendFootnotes(blocks, ctx, footnotes, endnotes)
  const headers = collectHeadersFooters(headerParts, parts)
  const footers = collectHeadersFooters(footerParts, parts)
  const sectionCount = findAll(body, 'w:sectPr').length

  return {
    blocks,
    meta: {
      title: coreTitle,
      pagesKnown: false,
      headings: ctx.headingCount,
      headingsInferred: config.inferHeadings && ctx.headingCount > 0,
      tables: ctx.tableCount,
      images: ctx.imageCounter,
      tocCount: toc?.entries.length ?? 0,
      tocOmitted: (toc?.entries.length ?? 0) > 0,
      tocEntries: toc?.entries ?? [],
      headers,
      footers,
      footnotes: [...ctx.footnoteRefs].length,
      sections: sectionCount,
      pageBreaks: ctx.pageBreaks,
    },
  }
}

/**
 * @param {import('../xml.js').XmlElement} paragraph
 * @param {ReturnType<typeof createContext> extends never ? never : any} ctx
 * @param {{ inHeading: boolean }} options
 */
function parseParagraph(paragraph, ctx, options) {
  const pPr = child(paragraph, 'w:pPr')
  const pStyleId = attr(child(pPr, 'w:pStyle'), 'val')
  const style = ctx.styles.resolve(pStyleId)
  const direct = ctx.styles.paragraphFormatting(pPr)
  const paragraphRPr = child(pPr, 'w:rPr')
  const paragraphRun = ctx.styles.runFormatting(paragraphRPr)

  const inline = collectInline(paragraph, ctx, {
    inHeading: options.inHeading,
    styleBold: paragraphRun.bold ?? style.bold ?? false,
    styleItalic: paragraphRun.italic ?? style.italic ?? false,
    styleStrike: paragraphRun.strike ?? style.strike ?? false,
  })
  const plain = normalizePdfText(inline.plain)
  const markdown = inline.markdown.trim()
  if (plain.length === 0 && inline.images.length === 0 && inline.textboxes.length === 0) return undefined

  const numId = direct.numId ?? style.numId
  const ilvl = direct.ilvl ?? style.ilvl ?? 0
  const outline = direct.outlineLvl ?? style.outlineLvl
  const bold = (paragraphRun.bold ?? style.bold ?? false) || inline.boldChars >= Math.max(1, inline.textChars * 0.6)
  const styleLevel = ctx.styles.headingLevel(pStyleId)
  const isTocStyle = ctx.styles.isTocStyle(pStyleId)

  return {
    paragraph,
    pStyleId,
    style,
    plain,
    markdown,
    images: inline.images,
    textboxes: inline.textboxes,
    pageBreak: inline.pageBreak,
    hasTocField: inline.hasTocField,
    bold,
    numId,
    ilvl,
    outline,
    styleLevel,
    isTocStyle,
    jc: direct.jc ?? style.jc,
    size: inline.maxSize ?? style.size ?? ctx.styles.bodySize(),
  }
}

/**
 * @param {ReturnType<typeof parseParagraph>} parsed
 * @param {any} ctx
 * @param {any[]} blocks
 */
function emitParagraphBlocks(parsed, ctx, blocks) {
  if (parsed === undefined) return
  if (parsed.pageBreak) ctx.pageBreaks += 1
  const numbering = resolveNumbering(parsed, ctx)

  const heading = detectHeading(parsed, numbering, ctx)
  if (heading !== undefined) {
    const label = heading.label
    const text = stripEmphasis(`${label}${parsed.markdown}`).trim()
    const plain = `${label}${parsed.plain}`.trim()
    blocks.push({
      type: 'heading',
      level: Math.min(6, Math.max(1, heading.level)),
      text,
      plain,
    })
    ctx.headingCount += 1
    ctx.lastHeadingLevel = heading.level
    ctx.lastWasHeading = true
    if (heading.level === 1) ctx.parenCnLevel = undefined
    if (ctx.currentLevels.length >= heading.level) ctx.currentLevels.length = heading.level - 1
    ctx.currentLevels.push(text)
    for (const image of parsed.images) blocks.push(image)
    for (const textbox of parsed.textboxes) {
      blocks.push({ type: 'paragraph', text: `[text box] ${escapeInline(textbox)}`, plain: textbox })
    }
    return
  }

  if (numbering !== undefined && (numbering.bullet || numbering.label.length > 0)) {
    const indent = '  '.repeat(Math.max(0, Math.min(5, parsed.ilvl)))
    const marker = numbering.bullet ? numbering.marker : numbering.label.trimEnd()
    const spacer = marker.length > 0 ? ' ' : ''
    const prefix = numbering.bullet ? `${indent}- ` : `${indent}${marker}${spacer}`
    blocks.push({
      type: 'list',
      level: parsed.ilvl,
      text: `${prefix}${parsed.markdown}`.replace(/\s+$/g, ''),
      plain: `${marker}${spacer}${parsed.plain}`.trim(),
      bullet: numbering.bullet,
    })
  } else if (parsed.markdown.length > 0) {
    blocks.push({ type: 'paragraph', text: parsed.markdown, plain: parsed.plain })
  }

  for (const image of parsed.images) blocks.push(image)
  for (const textbox of parsed.textboxes) {
    blocks.push({ type: 'paragraph', text: `[text box] ${escapeInline(textbox)}`, plain: textbox })
  }
  if (parsed.plain.length > 0) ctx.lastWasHeading = false
}

/**
 * @param {ReturnType<typeof parseParagraph>} parsed
 * @param {any} ctx
 */
function resolveNumbering(parsed, ctx) {
  if (parsed.numId === undefined || !ctx.numbering.has(parsed.numId)) return undefined
  return ctx.numbering.consume(parsed.numId, parsed.ilvl ?? 0)
}

/**
 * @param {ReturnType<typeof parseParagraph>} parsed
 * @param {{ label: string, marker: string, bullet: boolean } | undefined} numbering
 * @param {any} ctx
 */
function detectHeading(parsed, numbering, ctx) {
  if (parsed.styleLevel !== undefined) {
    if (parsed.styleLevel === 0) return { level: 1, label: '' }
    return { level: parsed.styleLevel, label: numbering?.label ?? '' }
  }
  if (parsed.outline !== undefined) {
    return { level: Math.max(1, parsed.outline + 1), label: numbering?.label ?? '' }
  }

  const text = parsed.plain
  const shortEnough = text.length > 0 && text.length <= 80
  const sentenceEnd = /[。；;！？!?]$/.test(text)

  if (numbering !== undefined && !numbering.bullet && numbering.label.length > 0) {
    const level = Math.min(6, Math.max(1, (parsed.ilvl ?? 0) + 1, numberingLabelLevel(numbering.label)))
    const autoHeading = parsed.bold && shortEnough
      || (!sentenceEnd && text.length <= 40 && (parsed.size >= ctx.styles.bodySize() * 1.2))
    if (autoHeading) return { level, label: numbering.label }
    return undefined
  }

  if (parsed.isTocStyle) return undefined

  if (/^[（(](由|请|各|如|注|说明|以下|备注)/.test(text)) return undefined

  const prefix = manualHeadingPrefix(text)
  if (prefix !== undefined && shortEnough && !sentenceEnd) {
    if (parsed.bold || text.length <= 60) {
      let level = prefix.level
      if (prefix.kind === 'parenCn') {
        if (ctx.parenCnLevel === undefined) {
          ctx.parenCnLevel = Math.max(2, Math.min(6, ctx.lastHeadingLevel > 0 ? ctx.lastHeadingLevel : 2))
        }
        level = ctx.parenCnLevel
      } else if (prefix.kind === 'cn' && ctx.lastHeadingLevel >= 3) {
        level = ctx.lastHeadingLevel
      } else if (prefix.kind === 'parenNum') {
        level = Math.min(6, Math.max(2, (ctx.lastHeadingLevel || 2) + 1))
      }
      return { level, label: '' }
    }
  }

  if (ctx.config.inferHeadings && parsed.bold && text.length > 0 && text.length <= 40 && !sentenceEnd) {
    if (/^[（(].{0,24}[）)]$/.test(text)) return undefined
    if (/[：:]$/.test(text)) return undefined
    const level = ctx.lastWasHeading
      ? Math.max(1, ctx.lastHeadingLevel)
      : (ctx.lastHeadingLevel > 0 ? Math.min(6, ctx.lastHeadingLevel + 1) : 1)
    return { level, label: '' }
  }
  return undefined
}

/**
 * @param {string} text
 * @returns {{ level: number, kind: string, label: string } | undefined}
 */
export function manualHeadingPrefix(text) {
  const trimmed = text.trim()
  let match = /^(\d+(?:\.\d+){0,5})(?:[.、．]\s*|\s+)/.exec(trimmed)
  if (match !== null) {
    return { level: Math.min(6, match[1].split('.').length), kind: 'num', label: match[0] }
  }
  match = /^([一二三四五六七八九十百]+)、\s*/.exec(trimmed)
  if (match !== null) return { level: 1, kind: 'cn', label: match[0] }
  match = /^[（(]([一二三四五六七八九十百]+)[）)]\s*/.exec(trimmed)
  if (match !== null) return { level: 2, kind: 'parenCn', label: match[0] }
  match = /^[（(](\d+)[）)]\s*/.exec(trimmed)
  if (match !== null) return { level: 2, kind: 'parenNum', label: match[0] }
  match = /^第([一二三四五六七八九十百\d]+)([章节篇部])/.exec(trimmed)
  if (match !== null) return { level: match[2] === '章' || match[2] === '篇' || match[2] === '部' ? 1 : 2, kind: 'chapter', label: match[0] }
  return undefined
}

/** @param {string} label */
function numberingLabelLevel(label) {
  const numeric = /^(\d+(?:\.\d+)+)/.exec(label.trim())
  return numeric === null ? 0 : numeric[1].split('.').length
}

/**
 * Walk paragraph/run children in order and collect Markdown, plain text,
 * images, footnote refs, and text boxes.
 * @param {import('../xml.js').XmlElement} element
 * @param {any} ctx
 * @param {{ inHeading: boolean, styleBold: boolean, styleItalic: boolean, styleStrike: boolean }} options
 */
function collectInline(element, ctx, options) {
  const out = {
    markdown: '',
    plain: '',
    images: [],
    textboxes: [],
    pageBreak: false,
    hasTocField: false,
    textChars: 0,
    boldChars: 0,
    maxSize: 0,
  }
  walkInline(element, ctx, options, out)
  return out
}

/**
 * @param {import('../xml.js').XmlElement} parent
 * @param {any} ctx
 * @param {{ inHeading: boolean, styleBold: boolean, styleItalic: boolean, styleStrike: boolean }} options
 * @param {any} out
 */
function walkInline(parent, ctx, options, out) {
  for (const node of parent.children) {
    if (!isElement(node)) continue
    const tag = node.tag
    if (tag === 'w:r') {
      addRun(node, ctx, options, out)
    } else if (tag === 'w:hyperlink') {
      const relId = attr(node, 'r:id')
      const target = relId === undefined ? undefined : ctx.rels.get(relId)?.target
      const before = out.markdown.length
      walkInline(node, ctx, options, out)
      if (target !== undefined && /^(https?:|mailto:)/i.test(target)) {
        const text = out.markdown.slice(before)
        if (text.length > 0) out.markdown = `${out.markdown.slice(0, before)}[${text}](${target})`
      }
    } else if (tag === 'w:ins' || tag === 'w:smartTag' || tag === 'w:customXml') {
      walkInline(node, ctx, options, out)
    } else if (tag === 'w:del') {
      // tracked deletion: omit
    } else if (tag === 'w:sdt') {
      const content = child(node, 'w:sdtContent')
      if (content !== undefined) walkInline(content, ctx, options, out)
    } else if (tag === 'w:fldSimple') {
      const instruction = attr(node, 'w:instr') ?? ''
      if (/TOC/i.test(instruction)) out.hasTocField = true
      walkInline(node, ctx, options, out)
    } else if (tag === 'w:drawing' || tag === 'w:pict') {
      addDrawing(node, ctx, out)
    } else if (tag === 'm:oMath' || tag === 'm:oMathPara') {
      const math = textOf(node)
      if (math.length > 0) {
        out.markdown += `[equation: ${escapeInline(math)}]`
        out.plain += math
        out.textChars += math.length
      }
    } else if (tag === 'w:tab') {
      out.markdown += ' '
      out.plain += ' '
    } else if (tag === 'w:br') {
      if (attr(node, 'type') === 'page') out.pageBreak = true
      out.markdown += ' '
      out.plain += ' '
    } else if (tag === 'w:bookmarkStart' || tag === 'w:bookmarkEnd'
      || tag === 'w:proofErr' || tag === 'w:commentRangeStart' || tag === 'w:commentRangeEnd') {
      // structural markers
    }
  }
}

/**
 * @param {import('../xml.js').XmlElement} run
 * @param {any} ctx
 * @param {{ inHeading: boolean, styleBold: boolean, styleItalic: boolean, styleStrike: boolean }} options
 * @param {any} out
 */
function addRun(run, ctx, options, out) {
  const rPr = child(run, 'w:rPr')
  const rStyleId = attr(child(rPr, 'w:rStyle'), 'val')
  const runStyle = ctx.styles.resolve(rStyleId)
  const direct = runProps(rPr)
  const bold = direct.bold ?? runStyle.bold ?? options.styleBold
  const italic = direct.italic ?? runStyle.italic ?? options.styleItalic
  const strike = direct.strike ?? runStyle.strike ?? options.styleStrike
  const size = direct.size ?? runStyle.size ?? 0
  out.maxSize = Math.max(out.maxSize, size)

  for (const node of run.children) {
    if (!isElement(node)) continue
    const tag = node.tag
    if (tag === 'w:t') {
      const text = textContent(node)
      if (text.length > 0) {
        appendFormatted(out, text, { bold, italic, strike, inHeading: options.inHeading })
      }
    } else if (tag === 'w:tab') {
      out.markdown += ' '
      out.plain += ' '
    } else if (tag === 'w:br') {
      if (attr(node, 'type') === 'page') out.pageBreak = true
      out.markdown += ' '
      out.plain += ' '
    } else if (tag === 'w:cr') {
      out.markdown += ' '
      out.plain += ' '
    } else if (tag === 'w:noBreakHyphen') {
      appendFormatted(out, '-', { bold, italic, strike, inHeading: options.inHeading })
    } else if (tag === 'w:softHyphen') {
      // invisible
    } else if (tag === 'w:sym') {
      const code = attr(node, 'char')
      if (code !== undefined) {
        const charCode = Number.parseInt(code, 16)
        if (Number.isFinite(charCode)) {
          appendFormatted(out, String.fromCharCode(charCode), { bold, italic, strike, inHeading: options.inHeading })
        }
      }
    } else if (tag === 'w:drawing' || tag === 'w:pict') {
      addDrawing(node, ctx, out)
    } else if (tag === 'w:footnoteReference') {
      const id = attr(node, 'w:id')
      if (id !== undefined) {
        ctx.footnoteRefs.add(`f:${id}`)
        out.markdown += `[^${id}]`
      }
    } else if (tag === 'w:endnoteReference') {
      const id = attr(node, 'w:id')
      if (id !== undefined) {
        ctx.footnoteRefs.add(`e:${id}`)
        out.markdown += `[^e${id}]`
      }
    } else if (tag === 'w:instrText') {
      if (/TOC/i.test(textContent(node))) out.hasTocField = true
    } else if (tag === 'w:object') {
      out.markdown += '[object]'
      out.plain += '[object]'
    } else if (tag === 'w:delText') {
      // deleted text
    }
  }
}

/**
 * @param {any} out
 * @param {string} text
 * @param {{ bold: boolean, italic: boolean, strike: boolean, inHeading: boolean }} format
 */
function appendFormatted(out, text, format) {
  const normalized = text.replace(/\r/g, '').replace(/\n/g, ' ')
  if (normalized.length === 0) return
  out.plain += normalized
  out.textChars += normalized.length
  if (format.bold) out.boldChars += normalized.length
  const isSpace = normalized.trim().length === 0
  if (isSpace) {
    out.markdown += normalized
    return
  }
  let rendered = escapeInline(normalized)
  if (format.bold && !format.inHeading) rendered = `**${rendered}**`
  if (format.italic) rendered = `*${rendered}*`
  if (format.strike) rendered = `~~${rendered}~~`
  out.markdown += rendered
}

/**
 * @param {import('../xml.js').XmlElement} node
 * @param {any} ctx
 * @param {any} out
 */
function addDrawing(node, ctx, out) {
  const textboxes = collectTextboxes(node)
  for (const textbox of textboxes) {
    if (textbox.length > 0) out.textboxes.push(textbox)
    ctx.textboxCounter += 1
  }
  const image = extractImage(node, ctx)
  if (image !== undefined) {
    ctx.imageCounter += 1
    const indexed = { ...image, index: ctx.imageCounter, type: 'image' }
    out.images.push(indexed)
  } else if (textboxes.length === 0) {
    const chart = findFirst(node, ['c:chart', 'dgm:relIds'])
    if (chart !== undefined) {
      out.markdown += ' [chart]'
      out.plain += ' [chart]'
    }
  }
}

/**
 * @param {import('../xml.js').XmlElement} node
 * @returns {string[]}
 */
function collectTextboxes(node) {
  const contents = findAll(node, 'w:txbxContent')
  const out = []
  for (const content of contents) {
    const paragraphs = findAll(content, 'w:p')
    const text = paragraphs.map((paragraph) => textOf(paragraph)).filter((value) => value.length > 0).join(' / ')
    if (text.length > 0) out.push(text)
  }
  return out
}

/**
 * @param {import('../xml.js').XmlElement} node
 * @param {any} ctx
 * @returns {{ alt: string, width: number, height: number, target?: string } | undefined}
 */
function extractImage(node, ctx) {
  const blip = findFirst(node, ['a:blip', 'v:imagedata'])
  if (blip === undefined) return undefined
  const relId = attr(blip, 'r:embed') ?? attr(blip, 'r:id')
  const target = relId === undefined ? undefined : ctx.rels.get(relId)?.target
  const extent = findFirst(node, ['wp:extent', 'a:ext'])
  const cx = Number.parseInt(attr(extent, 'cx') ?? '', 10)
  const cy = Number.parseInt(attr(extent, 'cy') ?? '', 10)
  const width = Number.isFinite(cx) ? Math.round((cx / 914400) * 96) : 0
  const height = Number.isFinite(cy) ? Math.round((cy / 914400) * 96) : 0
  const docPr = findFirst(node, ['wp:docPr', 'pic:cNvPr'])
  const alt = (attr(docPr, 'descr') ?? attr(docPr, 'title') ?? '').trim()
  return { alt, width, height, target }
}

/**
 * @param {import('../xml.js').XmlElement} node
 * @param {string[]} tags
 */
function findFirst(node, tags) {
  for (const tag of tags) {
    const found = findAll(node, tag)[0]
    if (found !== undefined) return found
  }
  return undefined
}

/**
 * @param {import('../xml.js').XmlElement} table
 * @param {any} ctx
 */
function parseTable(table, ctx) {
  const rows = childrenOf(table, 'w:tr')
  if (rows.length === 0) return undefined

  /** @type {string[][]} */
  const matrix = []
  /** @type {(string | undefined)[]} */
  const vMergeText = []
  let maxColumns = 0
  let merged = false

  for (const row of rows.slice(0, MAX_TABLE_ROWS)) {
    const cells = childrenOf(row, 'w:tc')
    /** @type {string[]} */
    const values = []
    let column = 0
    for (const cell of cells.slice(0, MAX_TABLE_COLUMNS * 2)) {
      const tcPr = child(cell, 'w:tcPr')
      const spanRaw = Number.parseInt(attr(child(tcPr, 'w:gridSpan'), 'val') ?? '1', 10)
      const span = Number.isFinite(spanRaw) && spanRaw > 0 ? Math.min(spanRaw, MAX_TABLE_COLUMNS) : 1
      if (span > 1) merged = true
      const vMerge = child(tcPr, 'w:vMerge')
      const vMergeValue = vMerge === undefined ? undefined : attr(vMerge, 'val')
      const isContinuation = vMerge !== undefined && vMergeValue !== 'restart'
      let text = cellText(cell, ctx)
      if (isContinuation) {
        merged = true
        text = vMergeText[column] ?? ''
      } else if (vMerge !== undefined) {
        merged = true
        vMergeText[column] = text
      } else {
        vMergeText[column] = undefined
      }
      for (let index = 0; index < span; index += 1) {
        values[column + index] = index === 0 ? text : ''
      }
      column += span
      if (column >= MAX_TABLE_COLUMNS) break
    }
    while (values.length < Math.min(column, MAX_TABLE_COLUMNS)) values.push('')
    maxColumns = Math.max(maxColumns, Math.min(values.length, MAX_TABLE_COLUMNS))
    matrix.push(values)
  }

  if (matrix.length === 0 || maxColumns === 0) return undefined
  const normalized = matrix.map((row) => {
    const copy = row.slice(0, maxColumns)
    while (copy.length < maxColumns) copy.push('')
    return copy
  })
  ctx.tableCount += 1
  const truncated = rows.length > MAX_TABLE_ROWS || maxColumns >= MAX_TABLE_COLUMNS
  return {
    type: 'table',
    rows: normalized,
    header: true,
    marker: `[table ${ctx.tableCount}: ${normalized.length}×${maxColumns}${merged ? ', merged cells' : ''}${truncated ? ', truncated' : ''}]`,
    merged,
    truncated,
  }
}

/**
 * @param {import('../xml.js').XmlElement} cell
 * @param {any} ctx
 */
function cellText(cell, ctx) {
  const parts = []
  for (const node of elementChildren(cell)) {
    if (isTag(node, 'w:p')) {
      const inline = collectInline(node, ctx, { inHeading: false, styleBold: false, styleItalic: false, styleStrike: false })
      const text = inline.markdown.trim()
      if (text.length > 0) parts.push(text)
      for (const image of inline.images) parts.push(`[image ${image.index}]`)
    } else if (isTag(node, 'w:tbl')) {
      const nested = parseTable(node, ctx)
      if (nested !== undefined) {
        parts.push(nested.rows.map((row) => row.join(' / ')).join(' / '))
      }
    }
  }
  const text = parts.join('<br>').replace(/\s+/g, ' ').trim()
  return text.length > 0 ? stripEmphasis(text) : ' '
}

/**
 * @param {import('../xml.js').XmlElement[]} elements
 * @param {StyleResolver} styles
 */
function detectToc(elements, styles) {
  let start = -1
  for (let index = 0; index < elements.length; index += 1) {
    const element = elements[index]
    if (!isTag(element, 'w:p')) continue
    const text = visibleText(element).trim()
    if (/^(目\s*录|contents|table of contents)$/i.test(text)) {
      start = index
      break
    }
  }
  if (start === -1) return undefined

  const indices = new Set()
  const entries = []
  let blankRun = 0
  for (let index = start; index < elements.length; index += 1) {
    const element = elements[index]
    if (!isTag(element, 'w:p')) {
      if (entries.length > 0) break
      continue
    }
    if (index === start) {
      indices.add(index)
      continue
    }
    const text = visibleText(element).trim()
    if (text.length === 0) {
      blankRun += 1
      if (blankRun > 2 && entries.length > 0) break
      indices.add(index)
      continue
    }
    blankRun = 0
    const pStyleId = attr(child(child(element, 'w:pPr'), 'w:pStyle'), 'val')
    const isEntry = styles.isTocStyle(pStyleId)
      || /[\t.·…]\s*\d+\s*$/.test(text)
      || /^(?:[一二三四五六七八九十百]+、|\d+(?:\.\d+)*[.、]?|[（(][一二三四五六七八九十百\d]+[）)])\s*\S.*\s+\d+\s*$/.test(text)
    if (!isEntry) {
      if (entries.length >= 3) break
      return undefined
    }
    indices.add(index)
    const title = text.replace(/[\t.·…]+\s*\d+\s*$/, '').replace(/\s+\d+\s*$/, '').replace(/\s+/g, ' ').trim()
    entries.push({ text: title, level: tocEntryLevel(text) })
  }
  if (entries.length < 3) return undefined
  return { indices, entries }
}

/**
 * Skip secondary tables of contents ("图 目 录", "表 目 录", "List of Figures")
 * after the main TOC, without adding them to the outline.
 * @param {import('../xml.js').XmlElement[]} elements
 * @param {StyleResolver} styles
 * @param {Set<number>} indices
 */
function addAuxiliaryTocIndices(elements, styles, indices) {
  for (let index = 0; index < elements.length; index += 1) {
    const element = elements[index]
    if (!isTag(element, 'w:p')) continue
    const text = visibleText(element).trim()
    if (!/^(图\s*目\s*录|表\s*目\s*录|插图目录|附表目录|list of (figures|tables))$/i.test(text)) continue
    indices.add(index)
    let entries = 0
    for (let cursor = index + 1; cursor < elements.length; cursor += 1) {
      const candidate = elements[cursor]
      if (!isTag(candidate, 'w:p')) break
      const candidateText = visibleText(candidate).trim()
      if (candidateText.length === 0) {
        indices.add(cursor)
        continue
      }
      const pStyleId = attr(child(child(candidate, 'w:pPr'), 'w:pStyle'), 'val')
      const isEntry = styles.isTocStyle(pStyleId)
        || /^(图|表|figure|table)\s*[\d一二三四五六七八九十]+.*\s+\d+\s*$/i.test(candidateText)
      if (!isEntry) break
      indices.add(cursor)
      entries += 1
    }
  }
}

/**
 * Visible paragraph text without field instructions or tracked deletions.
 * @param {import('../xml.js').XmlElement} element
 */
function visibleText(element) {
  let out = ''
  const walk = (node) => {
    if (!isElement(node)) return
    if (node.tag === 'w:instrText' || node.tag === 'w:delText' || node.tag === 'w:fldChar') return
    if (node.tag === 'w:t') {
      out += textContent(node)
      return
    }
    if (node.tag === 'w:tab' || node.tag === 'w:br' || node.tag === 'w:cr') {
      out += ' '
      return
    }
    for (const childNode of node.children) walk(childNode)
  }
  walk(element)
  return out
}

/** @param {string} markdown */
function stripEmphasis(markdown) {
  let text = markdown
  text = text.replace(/\*\*(.+?)\*\*/g, '$1')
  text = text.replace(/\*(.+?)\*/g, '$1')
  text = text.replace(/~~(.+?)~~/g, '$1')
  text = text.replace(/\\([\\`*_[\]])/g, '$1')
  return text
}

/** @param {string} text */
function tocEntryLevel(text) {
  const trimmed = text.trim()
  if (/^[一二三四五六七八九十百]+、/.test(trimmed)) return 1
  const numeric = /^(\d+(?:\.\d+)*)/.exec(trimmed)
  if (numeric !== null) return Math.min(6, numeric[1].split('.').length)
  if (/^[（(][一二三四五六七八九十百]+[）)]/.test(trimmed)) return 2
  return 1
}

/**
 * @param {string | undefined} xml
 * @returns {Map<string, { target: string, mode?: string, type?: string }>}
 */
function parseRelationships(xml) {
  const map = new Map()
  if (xml === undefined) return map
  const root = parseXml(xml)
  if (root === undefined) return map
  for (const relationship of findAll(root, 'Relationship')) {
    const id = attr(relationship, 'Id')
    const target = attr(relationship, 'Target')
    if (id === undefined || target === undefined) continue
    map.set(id, {
      target,
      mode: attr(relationship, 'TargetMode'),
      type: attr(relationship, 'Type'),
    })
  }
  return map
}

/**
 * @param {string | undefined} xml
 * @returns {Map<string, string>}
 */
function parseNotes(xml) {
  const map = new Map()
  if (xml === undefined) return map
  const root = parseXml(xml)
  if (root === undefined) return map
  for (const note of [...findAll(root, 'w:footnote'), ...findAll(root, 'w:endnote')]) {
    const id = attr(note, 'w:id')
    if (id === undefined || id === '-1' || id === '0') continue
    const paragraphs = findAll(note, 'w:p').map((paragraph) => textOf(paragraph)).filter((text) => text.length > 0)
    if (paragraphs.length > 0) map.set(id, paragraphs.join(' '))
  }
  return map
}

/**
 * @param {string | undefined} xml
 */
function parseCoreTitle(xml) {
  if (xml === undefined) return ''
  const root = parseXml(xml)
  if (root === undefined) return ''
  const title = findAll(root, 'dc:title')[0]
  return title === undefined ? '' : textOf(title)
}

/**
 * @param {string[]} names
 * @param {Map<string, string>} parts
 */
function collectHeadersFooters(names, parts) {
  const values = []
  for (const name of names) {
    const xml = parts.get(name)
    if (xml === undefined) continue
    const root = parseXml(xml)
    if (root === undefined) continue
    const text = findAll(root, 'w:p').map((paragraph) => textOf(paragraph)).filter((value) => value.length > 0).join(' ')
    if (text.length > 0 && !values.includes(text)) values.push(text)
  }
  return values.slice(0, 3)
}

/**
 * @param {any[]} blocks
 * @param {any} ctx
 * @param {Map<string, string>} footnotes
 * @param {Map<string, string>} endnotes
 */
function appendFootnotes(blocks, ctx, footnotes, endnotes) {
  const refs = [...ctx.footnoteRefs].slice(0, MAX_FOOTNOTES)
  if (refs.length === 0) return
  blocks.push({ type: 'paragraph', text: '**Footnotes**', plain: 'Footnotes' })
  for (const ref of refs) {
    const [kind, id] = ref.split(':')
    const text = kind === 'f' ? footnotes.get(id) : endnotes.get(id)
    if (text === undefined) continue
    const label = kind === 'f' ? id : `e${id}`
    blocks.push({ type: 'paragraph', text: `[^${label}]: ${escapeInline(text)}`, plain: text })
  }
}
