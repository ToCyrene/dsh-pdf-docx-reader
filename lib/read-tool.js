import { DocReaderError } from './errors.js'
import { loadDocument } from './load.js'
import { parsePageSelection } from './pages.js'
import { renderRead } from './window.js'

/**
 * Create the `pdf_docx_read` tool definition.
 * @param {any} ctx
 * @param {any} config
 * @param {any} cache
 */
export function createReadTool(ctx, config, cache) {
  return {
    name: 'pdf_docx_read',
    description: 'Read a PDF or DOCX as line-numbered Markdown with page and heading anchors. Use offset/limit to page through, pages to select PDF pages, and focus="outline" or focus="tables" when that is what you need. Images are placeholders; DOCX has no fixed page numbers.',
    parameters: {
      type: 'object',
      additionalProperties: false,
      properties: {
        file_path: {
          type: 'string',
          description: 'PDF or DOCX path, relative to the session working directory.',
        },
        offset: {
          type: 'integer',
          minimum: 1,
          description: '1-based content line to start at. Defaults to 1.',
        },
        limit: {
          type: 'integer',
          minimum: 1,
          maximum: 1000,
          description: `Maximum content lines to return. Defaults to ${config.readLimit}.`,
        },
        pages: {
          type: 'string',
          description: 'PDF only: a page or range such as "3" or "3-5"; also accepts "1,3-5" and "第3-5页".',
        },
        focus: {
          type: 'string',
          enum: ['content', 'outline', 'tables'],
          description: 'content (default), outline, or tables.',
        },
      },
      required: ['file_path'],
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
      const input = parseReadArguments(args, config)
      try {
        const { doc } = await loadDocument({
          ctx,
          exec,
          filePath: input.filePath,
          config,
          cache,
        })
        if (input.pages !== undefined && doc.kind !== 'pdf') {
          throw new DocReaderError('pages is PDF-only; DOCX has no fixed page numbers. Use offset/limit instead.', 'INVALID_PAGES')
        }
        const selection = parsePageSelection(input.pages, doc.meta.pages)
        const text = renderRead(doc, {
          offset: input.offset,
          limit: input.limit,
          pages: selection,
          focus: input.focus,
          maxChars: config.readMaxChars,
        })
        return { text }
      } catch (error) {
        throw normalizeToolError(error, input.filePath)
      }
    },
    presentCall(args) {
      const parts = [`Read ${args.file_path}`]
      if (typeof args.pages === 'string' && args.pages.length > 0) parts.push(`(pages ${args.pages})`)
      else if (typeof args.offset === 'number' && args.offset > 1) parts.push(`(from line ${args.offset})`)
      if (typeof args.focus === 'string' && args.focus !== 'content') parts.push(`— ${args.focus}`)
      return { card: 'generic', title: parts.join(' '), kind: 'read' }
    },
  }
}

/**
 * @param {any} args
 * @param {any} config
 */
export function parseReadArguments(args, config) {
  if (args === null || typeof args !== 'object' || Array.isArray(args)) {
    throw new DocReaderError('invalid arguments: expected an object', 'INVALID_ARGS')
  }
  if (typeof args.file_path !== 'string' || args.file_path.trim().length === 0) {
    throw new DocReaderError('file_path must be a non-empty string', 'INVALID_ARGS')
  }
  const offset = args.offset === undefined ? 1 : args.offset
  if (!Number.isSafeInteger(offset) || offset < 1) {
    throw new DocReaderError('offset must be a positive integer', 'INVALID_ARGS')
  }
  const limit = args.limit === undefined ? config.readLimit : args.limit
  if (!Number.isSafeInteger(limit) || limit < 1 || limit > 1000) {
    throw new DocReaderError('limit must be an integer between 1 and 1000', 'INVALID_ARGS')
  }
  const focus = args.focus ?? 'content'
  if (!['content', 'outline', 'tables'].includes(focus)) {
    throw new DocReaderError('focus must be content, outline, or tables', 'INVALID_ARGS')
  }
  if (args.pages !== undefined && typeof args.pages !== 'string') {
    throw new DocReaderError('pages must be a string such as "3" or "3-5"', 'INVALID_ARGS')
  }
  return {
    filePath: args.file_path,
    offset,
    limit,
    focus,
    pages: args.pages,
  }
}

/**
 * @param {unknown} error
 * @param {string} filePath
 */
export function normalizeToolError(error, filePath) {
  if (error instanceof DocReaderError) return error
  if (error !== null && typeof error === 'object' && typeof error.code === 'string' && error.code.startsWith('FS_')) {
    return error
  }
  if (error !== null && typeof error === 'object' && error.name === 'AbortError') return error
  const message = error instanceof Error ? error.message : String(error)
  return new DocReaderError(`Failed to read "${filePath}": ${message}`, 'DOC_READ_ERROR', { cause: error })
}
