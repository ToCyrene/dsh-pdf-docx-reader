import { DocReaderError } from './errors.js'

/** @typedef {{
 *   readLimit: number,
 *   readMaxChars: number,
 *   searchContext: number,
 *   searchMaxHits: number,
 *   cacheDocuments: number,
 *   inferHeadings: boolean,
 *   pageMarkers: boolean,
 *   maxFileBytes: number,
 * }} ReaderConfig
 */

export const DEFAULT_CONFIG = Object.freeze({
  readLimit: 300,
  readMaxChars: 48 * 1024,
  searchContext: 160,
  searchMaxHits: 10,
  cacheDocuments: 8,
  inferHeadings: true,
  pageMarkers: true,
  maxFileBytes: 100 * 1024 * 1024,
})

/**
 * Merge and validate a Loader row's `config`. Every field is optional; a
 * present field must have the documented type and range.
 * @param {unknown} input
 * @returns {ReaderConfig}
 */
export function resolveConfig(input) {
  if (input !== undefined && (input === null || typeof input !== 'object' || Array.isArray(input))) {
    throw new DocReaderError('pdf-docx-reader: config must be a mapping', 'INVALID_CONFIG')
  }
  const source = input ?? {}
  /** @type {ReaderConfig} */
  const config = { ...DEFAULT_CONFIG }
  for (const key of /** @type {const} */ ([
    'readLimit',
    'readMaxChars',
    'searchContext',
    'searchMaxHits',
    'cacheDocuments',
    'maxFileBytes',
  ])) {
    if (source[key] === undefined) continue
    const value = source[key]
    if (!Number.isSafeInteger(value) || value < 0) {
      throw new DocReaderError(`pdf-docx-reader: config.${key} must be a non-negative integer`, 'INVALID_CONFIG')
    }
    config[key] = value
  }
  for (const key of /** @type {const} */ (['inferHeadings', 'pageMarkers'])) {
    if (source[key] === undefined) continue
    if (typeof source[key] !== 'boolean') {
      throw new DocReaderError(`pdf-docx-reader: config.${key} must be a boolean`, 'INVALID_CONFIG')
    }
    config[key] = source[key]
  }
  if (config.readLimit < 1 || config.readLimit > 1000) {
    throw new DocReaderError('pdf-docx-reader: config.readLimit must be between 1 and 1000', 'INVALID_CONFIG')
  }
  if (config.readMaxChars < 1024) {
    throw new DocReaderError('pdf-docx-reader: config.readMaxChars must be at least 1024', 'INVALID_CONFIG')
  }
  if (config.searchContext > 1000) {
    throw new DocReaderError('pdf-docx-reader: config.searchContext must be at most 1000', 'INVALID_CONFIG')
  }
  if (config.searchMaxHits < 1 || config.searchMaxHits > 50) {
    throw new DocReaderError('pdf-docx-reader: config.searchMaxHits must be between 1 and 50', 'INVALID_CONFIG')
  }
  if (config.maxFileBytes < 1024) {
    throw new DocReaderError('pdf-docx-reader: config.maxFileBytes must be at least 1024', 'INVALID_CONFIG')
  }
  return config
}
