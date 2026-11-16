import fs from 'node:fs'
import path from 'node:path'
import { fileURLToPath } from 'node:url'
import { apply } from '../index.js'

export const PLUGIN_DIR = path.dirname(path.dirname(fileURLToPath(import.meta.url)))
export const FIXTURES = path.join(PLUGIN_DIR, 'test', 'fixtures')

/**
 * Minimal Host harness for exercising the tools without a running DSH.
 * @param {Record<string, unknown>} [config]
 */
export function createHarness(config = {}) {
  const registry = new Map()
  const ctx = {
    fs: {
      async resolve(requestedPath, options = {}) {
        const cwd = options.cwd ?? PLUGIN_DIR
        const displayPath = path.isAbsolute(requestedPath) ? requestedPath : path.resolve(cwd, requestedPath)
        return { targetKey: displayPath, displayPath }
      },
      async stat(target) {
        try {
          const stat = fs.statSync(target.targetKey)
          return {
            version: String(stat.mtimeMs),
            type: stat.isFile() ? 'file' : stat.isDirectory() ? 'directory' : 'other',
            size: stat.size,
          }
        } catch {
          return undefined
        }
      },
      async readBytes(target, _signal, maxBytes) {
        const bytes = fs.readFileSync(target.targetKey)
        if (bytes.length > maxBytes) {
          const error = new Error('FS_TOO_LARGE')
          error.code = 'FS_TOO_LARGE'
          throw error
        }
        return bytes
      },
    },
    workingDirectory: {
      async ensure() {
        return PLUGIN_DIR
      },
    },
    tools: {
      register(definition) {
        registry.set(definition.name, definition)
        return () => registry.delete(definition.name)
      },
      get(name) {
        return registry.get(name)
      },
    },
    systemPrompt: {
      section() {
        return () => {}
      },
      getSectionOrder() {
        return 1
      },
    },
    effect(callback) {
      const dispose = callback()
      return dispose
    },
  }
  apply(ctx, { readLimit: 300, ...config })
  const exec = {
    signal: new AbortController().signal,
    agent: { id: 'test-agent' },
    callId: 'test-call',
    name: 'test',
    arguments: {},
  }
  return {
    ctx,
    registry,
    exec,
    read: registry.get('pdf_docx_read'),
    search: registry.get('pdf_docx_search'),
  }
}

/** @param {string} name */
export function sample(name) {
  return path.join(PLUGIN_DIR, name)
}

/** @param {string} name */
export function fixture(name) {
  return path.join(FIXTURES, name)
}

/**
 * Resolve an optional real-world test document. Publishing and CI must not
 * depend on one machine's directory layout; point this at a local folder via
 * `DSH_PDF_DOCX_TEST_DATA` when the larger documents are available.
 * @param {string} name
 * @returns {string | undefined}
 */
export function external(name) {
  const root = process.env.DSH_PDF_DOCX_TEST_DATA
  if (typeof root !== 'string' || root.trim().length === 0) return undefined
  return path.join(root, name)
}

/** @param {string | undefined} file */
export function has(file) {
  return typeof file === 'string' && fs.existsSync(file)
}
