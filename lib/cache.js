const MAX_TOTAL_CHARS = 20 * 1024 * 1024

/**
 * In-memory LRU of parsed documents plus in-flight deduplication. A parsed
 * document is immutable except for its lazily built search index.
 */
export class DocumentCache {
  /** @param {number} maxDocuments */
  constructor(maxDocuments) {
    this.maxDocuments = Math.max(0, maxDocuments)
    /** @type {Map<string, { doc: any, chars: number }>} */
    this.entries = new Map()
    /** @type {Map<string, Promise<any>>} */
    this.inflight = new Map()
    this.totalChars = 0
  }

  /** @param {string} key */
  get(key) {
    const entry = this.entries.get(key)
    if (entry === undefined) return undefined
    this.entries.delete(key)
    this.entries.set(key, entry)
    return entry.doc
  }

  /** @param {string} key @param {any} doc */
  set(key, doc) {
    if (this.maxDocuments === 0) return
    const chars = typeof doc.charCount === 'number' ? doc.charCount : 0
    const existing = this.entries.get(key)
    if (existing !== undefined) {
      this.totalChars -= existing.chars
      this.entries.delete(key)
    }
    this.entries.set(key, { doc, chars })
    this.totalChars += chars
    while (this.entries.size > this.maxDocuments || this.totalChars > MAX_TOTAL_CHARS) {
      const oldest = this.entries.keys().next()
      if (oldest.done === true) break
      const entry = this.entries.get(oldest.value)
      this.entries.delete(oldest.value)
      if (entry !== undefined) this.totalChars -= entry.chars
    }
  }

  /**
   * Load once per key even when several tool calls race on the same file.
   * @param {string} key
   * @param {() => Promise<any>} loader
   */
  async load(key, loader) {
    const cached = this.get(key)
    if (cached !== undefined) return cached
    const running = this.inflight.get(key)
    if (running !== undefined) return running
    const promise = (async () => {
      const doc = await loader()
      this.set(key, doc)
      return doc
    })()
    this.inflight.set(key, promise)
    try {
      return await promise
    } finally {
      this.inflight.delete(key)
    }
  }

  clear() {
    this.entries.clear()
    this.inflight.clear()
    this.totalChars = 0
  }
}
