import test from 'node:test'
import assert from 'node:assert/strict'
import fs from 'node:fs'
import { parsePdf } from '../lib/pdf/extract.js'
import { buildDocument } from '../lib/document.js'
import { external, has, sample } from './helpers.js'

const TASK_BOOK = external('（已压缩）新一代人工智能国家科技重大专项课题任务书2026ZD0127703.pdf')
const APPLICATION = external('（已压缩）新一代人工智能国家科技重大专项申报书SQ2026AAA012210_副本.pdf')

async function load(file) {
  const { blocks, meta } = await parsePdf(new Uint8Array(fs.readFileSync(file)), {}, new AbortController().signal)
  const doc = buildDocument({ kind: 'pdf', path: file, blocks, meta, pageMarkers: true })
  return { blocks, meta, doc }
}

test('test.pdf keeps bookmarked headings and image count', async () => {
  const { meta, doc } = await load(sample('test.pdf'))
  assert.equal(meta.pages, 1)
  assert.equal(meta.images, 1)
  assert.equal(doc.outline.length, 2)
  assert.equal(doc.outline[0].text, '这是标题1')
  assert.equal(doc.outline[1].text, '这是标题2')
  assert.ok(doc.lines.some((line) => line.text === '# 这是标题1'))
  assert.ok(doc.lines.some((line) => line.text === '1. 这是 1'))
})

test('task book PDF detects layout pages and section headings', { skip: !has(TASK_BOOK) }, async () => {
  const { meta, blocks, doc } = await load(TASK_BOOK)
  assert.equal(meta.pages, 18)
  assert.ok(blocks.filter((block) => block.type === 'layout').length >= 3)
  assert.ok(doc.outline.some((entry) => entry.text.includes('二、课题研究内容')))
  const firstLayout = blocks.find((block) => block.type === 'layout')
  assert.ok(firstLayout.lines.some((line) => line.includes('│')))
})

test('application PDF extracts readable body text', { skip: !has(APPLICATION) }, async () => {
  const { doc } = await load(APPLICATION)
  assert.ok(doc.outline.some((entry) => entry.text.startsWith('3.1')))
  assert.ok(doc.lines.some((line) => line.text.includes('超智融合')))
})

test('invalid PDF bytes are rejected', async () => {
  await assert.rejects(parsePdf(new Uint8Array([1, 2, 3, 4]), {}, new AbortController().signal), /PDF/i)
})

test('an aborted parse stops before reading pages', async () => {
  const controller = new AbortController()
  controller.abort()
  await assert.rejects(
    parsePdf(new Uint8Array(fs.readFileSync(sample('test.pdf'))), {}, controller.signal),
    (error) => error.name === 'AbortError',
  )
})
