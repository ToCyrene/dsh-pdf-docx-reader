import test from 'node:test'
import assert from 'node:assert/strict'
import fs from 'node:fs'
import { createHarness, fixture, sample } from './helpers.js'

const TEST_PDF = sample('test.pdf')
const TEST_DOCX = sample('test.docx')

test('registers exactly the two reader tools', () => {
  const harness = createHarness()
  assert.deepEqual([...harness.registry.keys()].sort(), ['pdf_docx_read', 'pdf_docx_search'])
  assert.ok(harness.read)
  assert.ok(harness.search)
})

test('read returns numbered Markdown and a continuation footer', async () => {
  const harness = createHarness()
  const result = await harness.read.execute({ file_path: TEST_PDF, limit: 5 }, harness.exec)
  assert.match(result.text, /^\[pdf\] /)
  assert.match(result.text, /^1: \[p1/m)
  assert.match(result.text, /Continue with offset=6/)
  assert.doesNotMatch(result.text, /undefined:/)
})

test('read outline and pages work for PDF', async () => {
  const harness = createHarness()
  const outline = await harness.read.execute({ file_path: TEST_PDF, focus: 'outline' }, harness.exec)
  assert.match(outline.text, /# 这是标题1/)
  const pages = await harness.read.execute({ file_path: TEST_PDF, pages: '第1页', limit: 4 }, harness.exec)
  assert.match(pages.text, /selection: pages 第1页/)
})

test('search returns line, match markers, and context', async () => {
  const harness = createHarness()
  const result = await harness.search.execute({ file_path: TEST_PDF, query: '测试文件', context: 20 }, harness.exec)
  assert.match(result.text, /1 matches/)
  assert.match(result.text, /⟦测试文件⟧/)
  assert.match(result.text, /L2/)
})

test('DOCX read works and pages is rejected', async () => {
  const harness = createHarness()
  const result = await harness.read.execute({ file_path: TEST_DOCX, limit: 10 }, harness.exec)
  assert.match(result.text, /\[docx\]/)
  assert.match(result.text, /# 这是标题1/)
  await assert.rejects(
    harness.read.execute({ file_path: TEST_DOCX, pages: '1' }, harness.exec),
    /pages is PDF-only/,
  )
})

test('focus tables returns generated tables', async () => {
  const harness = createHarness()
  const result = await harness.read.execute({ file_path: fixture('structured.docx'), focus: 'tables' }, harness.exec)
  assert.match(result.text, /\[table 1: 3×3\]/)
  assert.match(result.text, /\[table 2: 2×3, merged cells\]/)
})

test('read rejects an out-of-range offset', async () => {
  const harness = createHarness()
  await assert.rejects(
    harness.read.execute({ file_path: TEST_PDF, offset: 9999 }, harness.exec),
    /out of range/,
  )
})

test('unsupported and missing files fail with actionable messages', async () => {
  const harness = createHarness()
  await assert.rejects(
    harness.read.execute({ file_path: `${TEST_PDF}.txt` }, harness.exec),
    /Unsupported file type/,
  )
  await assert.rejects(
    harness.read.execute({ file_path: sample('missing.pdf') }, harness.exec),
    /not found/,
  )
})

test('search validates query and reports zero matches cleanly', async () => {
  const harness = createHarness()
  await assert.rejects(harness.search.execute({ file_path: TEST_PDF, query: '   ' }, harness.exec), /query/)
  const result = await harness.search.execute({ file_path: TEST_PDF, query: '不存在的字符串' }, harness.exec)
  assert.match(result.text, /0 matches/)
  assert.match(result.text, /No literal match/)
})

test('cached document survives a second read', async () => {
  const first = createHarness()
  const read1 = await first.read.execute({ file_path: TEST_DOCX, limit: 5 }, first.exec)
  const read2 = await first.read.execute({ file_path: TEST_DOCX, limit: 5 }, first.exec)
  assert.equal(read1.text, read2.text)
  assert.equal(fs.existsSync(TEST_DOCX), true)
})
