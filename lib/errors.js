/**
 * Error type used by the PDF/DOCX reader. The tool pipeline surfaces the
 * message to the model; the code is kept for tests and future structured
 * failure reporting.
 */
export class DocReaderError extends Error {
  /**
   * @param {string} message
   * @param {string} [code]
   * @param {{ cause?: unknown }} [options]
   */
  constructor(message, code = 'DOC_READ_ERROR', options) {
    super(message, options)
    this.name = 'DocReaderError'
    this.code = code
  }
}

/**
 * Wrap an unknown failure with a reader-specific prefix while preserving the
 * original cause for logs.
 * @param {unknown} error
 * @param {string} context
 * @returns {DocReaderError}
 */
export function asReaderError(error, context) {
  if (error instanceof DocReaderError) return error
  const detail = error instanceof Error ? error.message : String(error)
  return new DocReaderError(`${context}: ${detail}`, 'DOC_READ_ERROR', { cause: error })
}
