import test from 'node:test'
import assert from 'node:assert/strict'
import fs from 'node:fs'
import { parseDocx } from '../lib/docx/blocks.js'
import { buildDocument } from '../lib/document.js'
import { external, fixture, has, sample } from './helpers.js'

const CHAPTER_5 = external('第5章.docx')
const IMPLEMENTATION = external('课题3实施方案.docx')

function load(file) {
  const { blocks, meta } = parseDocx(new Uint8Array(fs.readFileSync(file)), { inferHeadings: true })
  return { blocks, meta, doc: buildDocument({ kind: 'docx', path: file, blocks, meta, pageMarkers: true }) }
}

test('test.docx keeps headings, list, and image placeholder', () => {
  const { meta, doc } = load(sample('test.docx'))
  assert.equal(meta.images, 1)
  assert.equal(doc.outline.length, 2)
  assert.equal(doc.outline[0].text, '这是标题1')
  assert.equal(doc.outline[1].level, 2)
  assert.ok(doc.lines.some((line) => line.text === '1. 这是1'))
  assert.ok(doc.lines.some((line) => line.text.startsWith('[image 1')))
})

test('structured fixture renders tables, merged cells, and numbered lists', () => {
  const { meta, doc } = load(fixture('structured.docx'))
  assert.equal(meta.tables, 2)
  assert.equal(meta.images, 1)
  assert.equal(doc.outline.length, 4)
  assert.ok(doc.lines.some((line) => line.text === '1. 第一条'))
  assert.ok(doc.lines.some((line) => line.text.includes('[table 1: 3×3]')))
  assert.ok(doc.lines.some((line) => line.text.includes('[table 2: 2×3, merged cells]')))
  assert.ok(doc.lines.some((line) => line.text.includes('| 指标3.1 | 第三方测试 |')))
})

test('第5章.docx infers the expected heading hierarchy', { skip: !has(CHAPTER_5) }, () => {
  const { doc } = load(CHAPTER_5)
  assert.equal(doc.outline.length, 4)
  assert.deepEqual(doc.outline.map((entry) => entry.level), [1, 2, 2, 2])
  assert.equal(doc.outline[0].text, '一、项目成果呈现形式及测试方法')
})

test('课题3实施方案.docx preserves auto-numbered headings, tables, and TOC omission', { skip: !has(IMPLEMENTATION) }, () => {
  const { meta, doc } = load(IMPLEMENTATION)
  assert.equal(meta.tocCount, 27)
  assert.equal(meta.tables, 12)
  assert.equal(meta.images, 8)
  const outlineText = doc.outline.map((entry) => entry.text)
  assert.ok(outlineText.includes('一、课题概要'))
  assert.ok(outlineText.includes('1.1 课题简介'))
  assert.ok(outlineText.includes('2.2.1 主要研究工作'))
  assert.ok(outlineText.includes('2.3.1 任务一：面向超智融合的高性能科学I/O中间件'))
  assert.ok(doc.lines.some((line) => line.text.includes('[table 12:')))
  assert.ok(!doc.lines.some((line) => line.text.includes('课题目标与项目任务要求的对应说明\t')))
})

test('invalid DOCX bytes are rejected', () => {
  assert.throws(() => parseDocx(new Uint8Array([1, 2, 3]), { inferHeadings: true }), /DOCX|ZIP|package/i)
})
