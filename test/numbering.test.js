import test from 'node:test'
import assert from 'node:assert/strict'
import { NumberingResolver } from '../lib/docx/numbering.js'

const numberingXml = `<?xml version="1.0" encoding="UTF-8"?>
<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:abstractNum w:abstractNumId="0">
    <w:lvl w:ilvl="0">
      <w:start w:val="1"/>
      <w:numFmt w:val="chineseCounting"/>
      <w:lvlText w:val="%1、"/>
      <w:suff w:val="nothing"/>
    </w:lvl>
    <w:lvl w:ilvl="1">
      <w:start w:val="1"/>
      <w:numFmt w:val="decimal"/>
      <w:lvlText w:val="%1.%2"/>
      <w:suff w:val="space"/>
      <w:isLgl/>
    </w:lvl>
  </w:abstractNum>
  <w:abstractNum w:abstractNumId="1">
    <w:lvl w:ilvl="0">
      <w:start w:val="1"/>
      <w:numFmt w:val="bullet"/>
      <w:lvlText w:val="\uF06C"/>
    </w:lvl>
  </w:abstractNum>
  <w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>
  <w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num>
</w:numbering>`

test('numbering renders hierarchical labels and advances counters', () => {
  const resolver = new NumberingResolver(numberingXml)
  assert.equal(resolver.consume(1, 0).label, '一、')
  assert.equal(resolver.consume(1, 1).label, '1.1 ')
  assert.equal(resolver.consume(1, 1).label, '1.2 ')
  assert.equal(resolver.consume(1, 0).label, '二、')
  assert.equal(resolver.consume(1, 1).label, '2.1 ')
})

test('bullet numbering is recognized and mapped to a list marker', () => {
  const resolver = new NumberingResolver(numberingXml)
  assert.equal(resolver.isBullet(2, 0), true)
  assert.deepEqual(resolver.consume(2, 0), { label: '', marker: '-', bullet: true })
})

test('unknown numbering ids are harmless', () => {
  const resolver = new NumberingResolver(undefined)
  assert.equal(resolver.has(99), false)
  assert.equal(resolver.consume(99, 0), undefined)
})
