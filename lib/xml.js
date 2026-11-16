import { XMLParser } from 'fast-xml-parser'

const parser = new XMLParser({
  preserveOrder: true,
  ignoreAttributes: false,
  attributeNamePrefix: '@_',
  textNodeName: '#text',
  trimValues: false,
  parseTagValue: false,
  parseAttributeValue: false,
  processEntities: true,
  allowBooleanAttributes: true,
})

/**
 * @typedef {{ tag: string, attrs: Record<string, string>, children: XmlNode[] }} XmlElement
 * @typedef {{ tag: '#text', text: string }} XmlText
 * @typedef {XmlElement | XmlText} XmlNode
 */

/**
 * Parse an OOXML document into a small ordered tree. Namespace prefixes are
 * preserved (`w:p`, `m:t`, ...) so callers can distinguish WordprocessingML
 * from DrawingML or math text.
 * @param {string} xml
 * @returns {XmlElement | undefined} the document root, excluding the XML declaration
 */
export function parseXml(xml) {
  const raw = parser.parse(xml)
  for (const node of raw) {
    if (node['#text'] !== undefined) continue
    const tag = tagOf(node)
    if (tag === '?xml') continue
    return toTree(node)
  }
  return undefined
}

/** @param {Record<string, unknown>} node */
function tagOf(node) {
  for (const key of Object.keys(node)) {
    if (key !== ':@') return key
  }
  return ''
}

/**
 * @param {Record<string, any>} node
 * @returns {XmlNode}
 */
function toTree(node) {
  if (node['#text'] !== undefined) return { tag: '#text', text: String(node['#text']) }
  const tag = tagOf(node)
  const attrs = {}
  const rawAttrs = node[':@']
  if (rawAttrs !== undefined && rawAttrs !== null && typeof rawAttrs === 'object') {
    for (const [key, value] of Object.entries(rawAttrs)) {
      if (key.startsWith('@_')) attrs[key.slice(2)] = String(value)
    }
  }
  const rawChildren = node[tag]
  const children = Array.isArray(rawChildren) ? rawChildren.map(toTree) : []
  return { tag, attrs, children }
}

/** @param {XmlNode} node @returns {node is XmlElement} */
export function isElement(node) {
  return node.tag !== '#text'
}

/** @param {XmlNode} node @returns {node is XmlText} */
export function isText(node) {
  return node.tag === '#text'
}

/** @param {string} tag */
function localName(tag) {
  const index = tag.indexOf(':')
  return index === -1 ? tag : tag.slice(index + 1)
}

/** @param {string} tag @param {string} name */
export function isTag(node, name) {
  return node.tag === name
}

/**
 * Get an attribute by exact name, falling back to a local-name match so a
 * caller can ask for `val` without spelling `w:val`.
 * @param {XmlElement} element
 * @param {string} name
 */
export function attr(element, name) {
  if (element === undefined) return undefined
  if (Object.hasOwn(element.attrs, name)) return element.attrs[name]
  const wanted = localName(name)
  for (const [key, value] of Object.entries(element.attrs)) {
    if (localName(key) === wanted) return value
  }
  return undefined
}

/** @param {XmlElement} element @param {string} tag */
export function child(element, tag) {
  if (element === undefined) return undefined
  for (const node of element.children) {
    if (isElement(node) && isTag(node, tag)) return node
  }
  return undefined
}

/** @param {XmlElement} element @param {string} tag @returns {XmlElement[]} */
export function childrenOf(element, tag) {
  if (element === undefined) return []
  return element.children.filter((node) => isElement(node) && isTag(node, tag))
}

/** @param {XmlElement} element @returns {XmlElement[]} */
export function elementChildren(element) {
  if (element === undefined) return []
  return element.children.filter(isElement)
}

/**
 * Recursively collect text nodes under an element.
 * @param {XmlNode} node
 * @returns {string}
 */
export function textContent(node) {
  if (node === undefined) return ''
  if (isText(node)) return node.text
  let out = ''
  for (const childNode of node.children) out += textContent(childNode)
  return out
}

/**
 * Recursively find elements by tag, preserving order.
 * @param {XmlNode} node
 * @param {string} tag
 * @param {XmlElement[]} [out]
 * @returns {XmlElement[]}
 */
export function findAll(node, tag, out = []) {
  if (node === undefined || !isElement(node)) return out
  if (isTag(node, tag)) out.push(node)
  for (const childNode of node.children) findAll(childNode, tag, out)
  return out
}

/**
 * WordprocessingML on/off property: present without a value means true;
 * `0`, `false`, `off` mean false.
 * @param {XmlElement | undefined} element
 */
export function onOff(element) {
  if (element === undefined) return false
  const value = attr(element, 'val')
  if (value === undefined) return true
  return !['0', 'false', 'off'].includes(value.toLowerCase())
}
