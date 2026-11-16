import test from 'node:test'
import assert from 'node:assert/strict'
import { DocumentCache } from '../lib/cache.js'

test('LRU evicts the oldest document', () => {
  const cache = new DocumentCache(1)
  cache.set('a', { charCount: 10 })
  cache.set('b', { charCount: 20 })
  assert.equal(cache.get('a'), undefined)
  assert.notEqual(cache.get('b'), undefined)
})

test('load deduplicates concurrent work for one key', async () => {
  const cache = new DocumentCache(2)
  let calls = 0
  const loader = async () => {
    calls += 1
    await new Promise((resolve) => setTimeout(resolve, 10))
    return { charCount: 1 }
  }
  const [first, second] = await Promise.all([cache.load('same', loader), cache.load('same', loader)])
  assert.equal(calls, 1)
  assert.equal(first, second)
})

test('disabled cache does not retain documents', async () => {
  const cache = new DocumentCache(0)
  const doc = await cache.load('a', async () => ({ charCount: 1 }))
  assert.notEqual(doc, undefined)
  assert.equal(cache.get('a'), undefined)
})
