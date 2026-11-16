import { DocumentCache } from './lib/cache.js'
import { resolveConfig } from './lib/config.js'
import { createReadTool } from './lib/read-tool.js'
import { createSearchTool } from './lib/search-tool.js'

/** Cordis plugin name used by Loader diagnostics. */
export const name = 'pdf-docx-reader'

/** Services required by the reader tools. */
export const inject = ['tools', 'fs', 'workingDirectory', 'systemPrompt']

/**
 * Register the PDF/DOCX read and search tools plus a one-line routing hint.
 * The plugin is Host-only; parsing stays local and no network or model
 * request is added.
 * @param {any} ctx
 * @param {unknown} config
 */
export function apply(ctx, config) {
  const resolved = resolveConfig(config)
  const cache = new DocumentCache(resolved.cacheDocuments)

  ctx.effect(() => {
    const disposeRead = ctx.tools.register(createReadTool(ctx, resolved, cache))
    const disposeSearch = ctx.tools.register(createSearchTool(ctx, resolved, cache))
    const disposeGuidance = ctx.systemPrompt.section({
      name: 'tool:pdf-docx-reader',
      order: ctx.systemPrompt.getSectionOrder('TOOL_READ') + 1,
      text: ({ scope }) => ctx.tools.get('pdf_docx_read', scope) === undefined
        ? ''
        : 'For .pdf and .docx files, use pdf_docx_read or pdf_docx_search instead of read, shell, or Python extraction.',
    })
    return () => {
      disposeGuidance()
      disposeRead()
      disposeSearch()
      cache.clear()
    }
  })
}
