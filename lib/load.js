import { parseDocx } from './docx/blocks.js'
import { DocReaderError } from './errors.js'
import { buildDocument } from './document.js'
import { parsePdf } from './pdf/extract.js'
import { formatBytes, resolveTarget, verifyMagic } from './resolve.js'

const PARSER_VERSION = 1

/**
 * Resolve, parse, and cache one document. The cache key covers the file
 * version, parser version, and options that change the rendered lines.
 * @param {{
 *   ctx: any,
 *   exec: any,
 *   filePath: string,
 *   config: any,
 *   cache: any,
 * }} request
 */
export async function loadDocument({ ctx, exec, filePath, config, cache }) {
  const { target, info, kind } = await resolveTarget(ctx, exec, filePath)
  if (info.size !== undefined && info.size > config.maxFileBytes) {
    throw new DocReaderError(
      `"${target.displayPath}" is ${formatBytes(info.size)}; the limit is ${formatBytes(config.maxFileBytes)}`,
      'FILE_TOO_LARGE',
    )
  }
  const key = [
    kind,
    target.targetKey,
    String(info.version),
    String(info.size ?? 0),
    config.inferHeadings ? 'infer' : 'styles',
    config.pageMarkers ? 'markers' : 'plain',
    `v${PARSER_VERSION}`,
  ].join('|')

  const doc = await cache.load(key, async () => {
    const raw = await ctx.fs.readBytes(target, exec.signal, config.maxFileBytes)
    const bytes = raw instanceof Uint8Array ? Uint8Array.from(raw) : new Uint8Array(raw)
    verifyMagic(bytes, kind, target.displayPath)
    const parsed = kind === 'pdf'
      ? await parsePdf(bytes, config, exec.signal)
      : parseDocx(bytes, config)
    return buildDocument({
      kind,
      path: target.displayPath,
      blocks: parsed.blocks,
      meta: parsed.meta,
      pageMarkers: config.pageMarkers,
    })
  })
  return { doc, target, info, kind }
}
