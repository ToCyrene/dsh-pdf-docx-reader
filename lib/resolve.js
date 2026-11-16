import { DocReaderError } from './errors.js'

const PDF_MAGIC = [0x25, 0x50, 0x44, 0x46, 0x2d]
const ZIP_MAGIC = [0x50, 0x4b, 0x03, 0x04]
const OLE_MAGIC = [0xd0, 0xcf, 0x11, 0xe0]

/** @param {Uint8Array} bytes @param {number[]} magic */
function startsWith(bytes, magic) {
  if (bytes.length < magic.length) return false
  for (let index = 0; index < magic.length; index += 1) {
    if (bytes[index] !== magic[index]) return false
  }
  return true
}

/**
 * Resolve a model path to a regular file through the mounted filesystem.
 * @param {any} ctx
 * @param {any} exec
 * @param {string} filePath
 * @returns {Promise<{ target: any, info: any, kind: 'pdf' | 'docx' }>}
 */
export async function resolveTarget(ctx, exec, filePath) {
  if (typeof filePath !== 'string' || filePath.trim().length === 0) {
    throw new DocReaderError('file_path must be a non-empty string', 'INVALID_ARGS')
  }
  const extension = extensionOf(filePath)
  if (extension === 'doc' || extension === 'dot') {
    throw new DocReaderError(
      'Legacy .doc files are not supported; convert to .docx first (office_convert)',
      'DOC_UNSUPPORTED',
    )
  }
  if (!['pdf', 'docx', 'docm'].includes(extension)) {
    throw new DocReaderError(
      `Unsupported file type ".${extension || 'unknown'}"; only .pdf, .docx, and .docm are supported`,
      'UNSUPPORTED_TYPE',
    )
  }
  const cwd = exec.agent === undefined ? undefined : await ctx.workingDirectory.ensure(exec.agent, exec.signal)
  let target
  try {
    target = await ctx.fs.resolve(filePath, { ...(cwd === undefined ? {} : { cwd }), signal: exec.signal })
  } catch (error) {
    if (error instanceof DocReaderError) throw error
    throw new DocReaderError(
      `Cannot resolve "${filePath}": ${error instanceof Error ? error.message : String(error)}`,
      'PATH_INVALID',
      { cause: error },
    )
  }
  const info = await ctx.fs.stat(target, exec.signal)
  if (info === undefined) {
    throw new DocReaderError(`cannot read "${target.displayPath}": not found`, 'FS_NOT_FOUND')
  }
  if (info.type !== 'file') {
    throw new DocReaderError(`cannot read "${target.displayPath}": not a regular file`, 'FS_NOT_REGULAR_FILE')
  }
  return { target, info, kind: extension === 'pdf' ? 'pdf' : 'docx' }
}

/**
 * @param {Uint8Array} bytes
 * @param {'pdf' | 'docx'} kind
 * @param {string} displayPath
 */
export function verifyMagic(bytes, kind, displayPath) {
  if (startsWith(bytes, OLE_MAGIC)) {
    throw new DocReaderError(
      `"${displayPath}" is a legacy or encrypted Office file, not a .docx ZIP package; convert or decrypt it first`,
      'DOCX_INVALID',
    )
  }
  if (kind === 'pdf' && !startsWith(bytes, PDF_MAGIC)) {
    throw new DocReaderError(`"${displayPath}" does not look like a PDF (missing %PDF header)`, 'PDF_INVALID')
  }
  if (kind === 'docx' && !startsWith(bytes, ZIP_MAGIC)) {
    throw new DocReaderError(`"${displayPath}" does not look like a DOCX package`, 'DOCX_INVALID')
  }
}

/** @param {string} path */
export function extensionOf(path) {
  const match = /\.([A-Za-z0-9]+)$/.exec(path.trim())
  return match === null ? '' : match[1].toLowerCase()
}

/** @param {number} bytes */
export function formatBytes(bytes) {
  if (bytes < 1024) return `${bytes} B`
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`
  return `${(bytes / (1024 * 1024)).toFixed(1)} MB`
}
