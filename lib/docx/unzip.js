import { unzipSync } from 'fflate'
import { DocReaderError } from '../errors.js'

/** Uncompressed DOCX package cap, independent of the downloaded file size. */
const MAX_UNCOMPRESSED_BYTES = 200 * 1024 * 1024
/** Single XML part cap; document.xml of a very large document can be tens of MB. */
const MAX_PART_BYTES = 64 * 1024 * 1024
const MAX_ENTRIES = 5000

const REQUIRED_PARTS = ['word/document.xml']

/**
 * Read the OOXML parts this reader needs, enforcing zip-bomb and entry caps.
 * @param {Uint8Array} bytes
 * @returns {{ parts: Map<string, string>, headerParts: string[], footerParts: string[], entries: number }}
 */
export function readDocxPackage(bytes) {
  let entries = 0
  let totalUncompressed = 0
  /** @type {string[]} */
  const headerParts = []
  /** @type {string[]} */
  const footerParts = []

  let raw
  try {
    raw = unzipSync(bytes, {
      filter: (file) => {
        entries += 1
        if (entries > MAX_ENTRIES) {
          throw new DocReaderError(`DOCX has more than ${MAX_ENTRIES} package entries`, 'DOCX_TOO_COMPLEX')
        }
        totalUncompressed += file.originalSize
        if (totalUncompressed > MAX_UNCOMPRESSED_BYTES) {
          throw new DocReaderError('DOCX expands beyond the 200 MB package limit', 'DOCX_TOO_LARGE')
        }
        if (file.originalSize > MAX_PART_BYTES) {
          throw new DocReaderError(`DOCX part "${file.name}" exceeds the 64 MB part limit`, 'DOCX_TOO_LARGE')
        }
        if (/^word\/header\d*\.xml$/.test(file.name)) {
          headerParts.push(file.name)
          return true
        }
        if (/^word\/footer\d*\.xml$/.test(file.name)) {
          footerParts.push(file.name)
          return true
        }
        return (
          file.name === 'word/document.xml'
          || file.name === 'word/styles.xml'
          || file.name === 'word/numbering.xml'
          || file.name === 'word/_rels/document.xml.rels'
          || file.name === 'word/footnotes.xml'
          || file.name === 'word/endnotes.xml'
          || file.name === 'word/comments.xml'
          || file.name === 'docProps/core.xml'
        )
      },
    })
  } catch (error) {
    if (error instanceof DocReaderError) throw error
    throw new DocReaderError(
      `Not a readable DOCX package (${error instanceof Error ? error.message : String(error)})`,
      'DOCX_INVALID',
      { cause: error },
    )
  }

  for (const required of REQUIRED_PARTS) {
    if (raw[required] === undefined) {
      throw new DocReaderError('Not a Word document: word/document.xml is missing', 'DOCX_INVALID')
    }
  }

  const decoder = new TextDecoder('utf-8')
  /** @type {Map<string, string>} */
  const parts = new Map()
  for (const [name, data] of Object.entries(raw)) {
    parts.set(name, decoder.decode(data).replace(/^\uFEFF/, ''))
  }
  return { parts, headerParts, footerParts, entries }
}
