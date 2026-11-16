import test from 'node:test'
import assert from 'node:assert/strict'
import { parsePageSelection } from '../lib/pages.js'
import { chineseNumber, formatNumber } from '../lib/docx/numbering.js'

test('page selection accepts common forms', () => {
  assert.deepEqual([...parsePageSelection('3', 10).pages], [3])
  assert.deepEqual([...parsePageSelection('3-5', 10).pages], [3, 4, 5])
  assert.deepEqual([...parsePageSelection('1,3-5', 10).pages], [1, 3, 4, 5])
  assert.deepEqual([...parsePageSelection('第3-5页', 10).pages], [3, 4, 5])
  assert.deepEqual([...parsePageSelection('3~5', 10).pages], [3, 4, 5])
  assert.deepEqual([...parsePageSelection('5-3', 10).pages], [3, 4, 5])
  assert.equal(parsePageSelection(undefined, 10), null)
})

test('page selection rejects malformed or out-of-range values', () => {
  assert.throws(() => parsePageSelection('abc', 10), /Invalid pages/)
  assert.throws(() => parsePageSelection('0', 10), /outside/)
  assert.throws(() => parsePageSelection('11', 10), /outside/)
  assert.throws(() => parsePageSelection('3-', 10), /Invalid pages/)
})

test('Chinese numbering covers common counts', () => {
  assert.equal(chineseNumber(1), '一')
  assert.equal(chineseNumber(10), '十')
  assert.equal(chineseNumber(11), '十一')
  assert.equal(chineseNumber(20), '二十')
  assert.equal(chineseNumber(21), '二十一')
  assert.equal(chineseNumber(101), '一百零一')
})

test('number formats cover roman, letters, and decimal zero', () => {
  assert.equal(formatNumber(4, 'lowerRoman'), 'iv')
  assert.equal(formatNumber(4, 'upperRoman'), 'IV')
  assert.equal(formatNumber(2, 'lowerLetter'), 'b')
  assert.equal(formatNumber(7, 'decimalZero'), '07')
  assert.equal(formatNumber(7, 'decimal', true), '7')
})
