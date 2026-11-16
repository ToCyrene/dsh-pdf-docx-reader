import { DocReaderError } from './errors.js'
import { loadDocument } from './load.js'
import { parsePageSelection } from './pages.js'
import { normalizeToolError } from './read-tool.js'
import { collapseWhitespace, truncateText } from './text.js'

/**
 * Create the `pdf_docx_search` tool definition.
 * @param {any} ctx
 * @param {any} config
 * @param {any} cache
 */
export function createSearchTool(ctx, config, cache) {
  return {
    name: 'pdf_docx_search',
    description: 'Search one PDF or DOCX for literal text. Returns line, page/heading, and surrounding context; read more with pdf_docx_read using the returned line. Use a short literal phrase rather than a regex.',
    parameters: {
      type: 'object',
      additionalProperties: false,
      properties: {
        file_path: {
          type: 'string',
          description: 'PDF or DOCX path, relative to the session working directory.',
        },
        query: {
          type: 'string',
          description: 'Literal text to find, 1-200 characters. Case-insensitive by default.',
        },
        pages: {
          type: 'string',
          description: 'PDF only: restrict to pages such as "3" or "3-5".',
        },
        context: {
          type: 'integer',
          minimum: 0,
          maximum: 1000,
          description: `Characters of context before and after each match. Defaults to ${config.searchContext}.`,
        },
        max_hits: {
          type: 'integer',
          minimum: 1,
          maximum: 50,
          description: `Maximum matches to show. Defaults to ${config.searchMaxHits}.`,
        },
        case_sensitive: {
          type: 'boolean',
          description: 'Match ASCII case exactly. Defaults to false.',
        },
      },
      required: ['file_path', 'query'],
    },
    output: {
      schema: {
        type: 'object',
        additionalProperties: false,
        properties: { text: { type: 'string' } },
        required: ['text'],
      },
      render: (_args, value) => [{ type: 'text', text: value.text }],
    },
    timeoutMs: 120000,
    isConcurrencySafe: () => true,
    async execute(args, exec) {
      const input = parseSearchArguments(args, config)
      try {
        const { doc } = await loadDocument({
          ctx,
          exec,
          filePath: input.filePath,
          config,
          cache,
        })
        if (input.pages !== undefined && doc.kind !== 'pdf') {
          throw new DocReaderError('pages is PDF-only; DOCX has no fixed page numbers.', 'INVALID_PAGES')
        }
        const selection = parsePageSelection(input.pages, doc.meta.pages)
        const text = searchDocument(doc, {
          query: input.query,
          selection,
          context: input.context,
          maxHits: input.maxHits,
          caseSensitive: input.caseSensitive,
          maxChars: config.readMaxChars,
        })
        return { text }
      } catch (error) {
        throw normalizeToolError(error, input.filePath)
      }
    },
    presentCall(args) {
      return {
        card: 'generic',
        title: `Search ${args.file_path} for "${truncateText(String(args.query ?? ''), 40)}"`,
        kind: 'search',
      }
    },
  }
}

/**
 * @param {any} args
 * @param {any} config
 */
export function parseSearchArguments(args, config) {
  if (args === null || typeof args !== 'object' || Array.isArray(args)) {
    throw new DocReaderError('invalid arguments: expected an object', 'INVALID_ARGS')
  }
  if (typeof args.file_path !== 'string' || args.file_path.trim().length === 0) {
    throw new DocReaderError('file_path must be a non-empty string', 'INVALID_ARGS')
  }
  if (typeof args.query !== 'string' || args.query.trim().length === 0) {
    throw new DocReaderError('query must be a non-empty literal string', 'INVALID_ARGS')
  }
  const query = args.query.trim()
  if (query.length > 200) {
    throw new DocReaderError('query must be at most 200 characters', 'INVALID_ARGS')
  }
  const context = args.context === undefined ? config.searchContext : args.context
  if (!Number.isSafeInteger(context) || context < 0 || context > 1000) {
    throw new DocReaderError('context must be an integer between 0 and 1000', 'INVALID_ARGS')
  }
  const maxHits = args.max_hits === undefined ? config.searchMaxHits : args.max_hits
  if (!Number.isSafeInteger(maxHits) || maxHits < 1 || maxHits > 50) {
    throw new DocReaderError('max_hits must be an integer between 1 and 50', 'INVALID_ARGS')
  }
  if (args.pages !== undefined && typeof args.pages !== 'string') {
    throw new DocReaderError('pages must be a string such as "3" or "3-5"', 'INVALID_ARGS')
  }
  return {
    filePath: args.file_path,
    query,
    context,
    maxHits,
    caseSensitive: args.case_sensitive === true,
    pages: args.pages,
  }
}

/**
 * @param {any} doc
 * @param {{
 *   query: string,
 *   selection: any,
 *   context: number,
 *   maxHits: number,
 *   caseSensitive: boolean,
 *   maxChars: number,
 * }} options
 */
export function searchDocument(doc, options) {
  const index = searchIndex(doc)
  const needle = options.caseSensitive ? options.query : options.query.toLowerCase()
  const haystack = options.caseSensitive ? index.text : index.text.toLowerCase()
  const hits = []
  let total = 0
  let totalCountCapped = false
  let chars = 0
  let position = 0
  while (position <= haystack.length - needle.length) {
    const found = haystack.indexOf(needle, position)
    if (found === -1) break
    const lineIndex = lineIndexAt(index.starts, found)
    const line = doc.lines[lineIndex]
    const selected = options.selection === null
      || (line?.page !== undefined && options.selection.pages.has(line.page))
    if (selected) {
      total += 1
      if (hits.length < options.maxHits) {
        const lineStart = index.starts[lineIndex]
        const lineEnd = lineIndex + 1 < index.starts.length ? index.starts[lineIndex + 1] - 1 : index.text.length
        const snippet = buildSnippet(index.text, found, needle.length, options.context, lineStart, lineEnd)
        const cost = snippet.snippet.length + 40
        if (chars + cost <= options.maxChars || hits.length === 0) {
          chars += cost
          hits.push({
            line: lineIndex + 1,
            ...(line?.page === undefined ? {} : { page: line.page }),
            ...(line?.headingPath === undefined ? {} : { heading: line.headingPath }),
            ...snippet,
          })
        }
      }
    }
    if (total >= 1000) {
      totalCountCapped = true
      break
    }
    position = found + Math.max(1, needle.length)
  }

  const header = buildSearchHeader(doc, options, total, totalCountCapped, hits.length)
  if (hits.length === 0) {
    return `${header}\n\n(No literal match. Try a shorter phrase${options.caseSensitive ? ', or set case_sensitive=false' : ''}. A scanned PDF has no text layer and needs an OCR-capable reader.)`
  }
  const body = hits.map((hit) => {
    const page = hit.page === undefined ? '' : ` p${hit.page}`
    const heading = hit.heading === undefined ? '' : ` [${truncateText(hit.heading, 80)}]`
    return `L${hit.line}${page}${heading} ${hit.snippet}`
  }).join('\n')
  const shownNote = hits.length < total
    ? `\n(Showing ${hits.length} of ${totalCountCapped ? '1000+' : total} matches; narrow the query or add pages.)`
    : ''
  return `${header}\n\n${body}${shownNote}`
}

/**
 * @param {any} doc
 * @param {{ query: string, selection: any, caseSensitive: boolean }} options
 * @param {number} total
 * @param {boolean} capped
 * @param {number} shown
 */
function buildSearchHeader(doc, options, total, capped, shown) {
  const parts = [
    `[${doc.kind}] ${doc.path}`,
    `query="${options.query}"`,
    options.caseSensitive ? 'literal, case-sensitive' : 'literal, case-insensitive',
    `${capped ? '1000+' : total} matches${shown < total ? ` (showing ${shown})` : ''}`,
  ]
  if (options.selection !== null) parts.push(`pages ${options.selection.text}`)
  if (doc.kind === 'docx') parts.push('pages unknown (DOCX)')
  return parts.join(' | ')
}

/**
 * Lazily build the line-addressed plain-text index used by search.
 * @param {any} doc
 */
export function searchIndex(doc) {
  if (doc.searchIndex !== undefined) return doc.searchIndex
  const starts = new Array(doc.lines.length)
  let offset = 0
  const parts = []
  for (let index = 0; index < doc.lines.length; index += 1) {
    const line = doc.lines[index]
    starts[index] = offset
    parts.push(line.plain)
    offset += line.plain.length + 1
  }
  const index = { text: parts.join('\n'), starts }
  doc.searchIndex = index
  return index
}

/**
 * @param {number[]} starts
 * @param {number} offset
 */
function lineIndexAt(starts, offset) {
  let low = 0
  let high = starts.length - 1
  let best = 0
  while (low <= high) {
    const middle = (low + high) >> 1
    if (starts[middle] <= offset) {
      best = middle
      low = middle + 1
    } else {
      high = middle - 1
    }
  }
  return best
}

/**
 * @param {string} text
 * @param {number} start
 * @param {number} length
 * @param {number} context
 * @param {number} lineStart
 * @param {number} lineEnd
 */
function buildSnippet(text, start, length, context, lineStart, lineEnd) {
  const beforeRaw = text.slice(Math.max(lineStart, start - context), start)
  const matchRaw = text.slice(start, start + length)
  const afterRaw = text.slice(start + length, Math.min(lineEnd, start + length + context))
  const before = collapseWhitespace(beforeRaw)
  const match = collapseWhitespace(matchRaw) || matchRaw
  const after = collapseWhitespace(afterRaw)
  return {
    snippet: `${before.length > 0 ? `…${before}` : ''}⟦${match}⟧${after.length > 0 ? `${after}…` : ''}`,
    match,
  }
}
