/**
 * Shared test harness: load the browser bundle the way the module table does, and
 * expose the pieces the docx-layer tests need.
 *
 * The bundle is not an ES module -- it calls window.__ModuleLoader__.load({ id,
 * factory }) and requires 'react' -- so it is evaluated in a vm context with a
 * fake loader and a stub react, exactly like the browser boot graph.
 */
import { readFileSync } from 'node:fs'
import { fileURLToPath } from 'node:url'
import { dirname, join } from 'node:path'
import vm from 'node:vm'
import { inflateRawSync } from 'node:zlib'

const here = dirname(fileURLToPath(import.meta.url))

/** Load lib/client.js and return its __internals test seam. */
export function loadInternals() {
  let definition
  const sandbox = {
    window: { __ModuleLoader__: { load: (value) => { definition = value } } },
    require: (specifier) => {
      if (specifier === 'react')
        return {
          createElement: () => null,
          useState: (initial) => [initial, () => {}],
          useEffect: () => {},
          useRef: (initial) => ({ current: initial }),
        }
      throw new Error('unexpected require: ' + specifier)
    },
    console,
    TextEncoder,
    TextDecoder,
    Blob,
    Response,
    DecompressionStream,
    CompressionStream,
    btoa: (v) => Buffer.from(v, 'binary').toString('base64'),
    atob: (v) => Buffer.from(v, 'base64').toString('binary'),
    setTimeout,
    clearInterval,
    setInterval,
    URL,
    fetch: () => Promise.resolve({ json: () => Promise.resolve({ ok: false }) }),
  }
  vm.runInContext(readFileSync(join(here, '..', 'lib', 'client.js'), 'utf8'), vm.createContext(sandbox), {
    filename: 'client.js',
  })
  return definition.factory(sandbox.require).__internals
}

/** Read word/document.xml out of a docx byte buffer. */
export function documentXmlOf(bytes) {
  const dv = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength)
  let eocd = -1
  for (let i = bytes.length - 22; i >= 0; i--) {
    if (dv.getUint32(i, true) === 0x06054b50) { eocd = i; break }
  }
  if (eocd < 0) throw new Error('no central directory')
  const count = dv.getUint16(eocd + 10, true)
  let p = dv.getUint32(eocd + 16, true)
  const entries = []
  for (let i = 0; i < count; i++) {
    const method = dv.getUint16(p + 10, true)
    const compSize = dv.getUint32(p + 20, true)
    const nameLen = dv.getUint16(p + 28, true)
    const extraLen = dv.getUint16(p + 30, true)
    const commentLen = dv.getUint16(p + 32, true)
    const localOffset = dv.getUint32(p + 42, true)
    const name = new TextDecoder().decode(bytes.subarray(p + 46, p + 46 + nameLen)).replace(/\\/g, '/')
    entries.push({ name, method, compSize, localOffset })
    p += 46 + nameLen + extraLen + commentLen
  }
  const entry = entries.find((candidate) => candidate.name === 'word/document.xml')
  if (entry === undefined) throw new Error('word/document.xml missing')
  const lo = entry.localOffset
  const lnameLen = dv.getUint16(lo + 26, true)
  const lextraLen = dv.getUint16(lo + 28, true)
  const start = lo + 30 + lnameLen + lextraLen
  const raw = bytes.subarray(start, start + entry.compSize)
  const data = entry.method === 0 ? raw : new Uint8Array(inflateRawSync(raw))
  return new TextDecoder().decode(data)
}

/** Assertion helper with a PASS/FAIL line and a failure count. */
export function createChecker() {
  const state = { failures: 0 }
  function check(label, condition, detail) {
    if (condition) {
      console.log('  PASS  ' + label)
    } else {
      state.failures += 1
      console.log('  FAIL  ' + label + (detail === undefined ? '' : '\n        ' + detail))
    }
  }
  return { check, state }
}

// ── a DOM-shaped surface, without a browser ──────────────────────────────────

export function fakeText(value) {
  return { nodeType: 3, nodeValue: value, textContent: value, childNodes: [] }
}

export function fakeElement(tag, attrs, children) {
  const list = children || []
  return {
    nodeType: 1,
    tagName: tag.toUpperCase(),
    attributes: attrs || {},
    childNodes: list,
    textContent: list.map((child) => child.textContent).join(''),
    getAttribute: (name) => (Object.prototype.hasOwnProperty.call(attrs || {}, name) ? attrs[name] : null),
    querySelectorAll: () => [],
  }
}

/**
 * Build the DOM the editor's surface renders for a parsed document: one node per
 * editable paragraph carrying its block id, the same wrapper order blocksToHtml
 * emits (inline style span, then u/strong/em), read-only markers on protected
 * paragraphs, and a read-only table for a table.
 *
 * Driving the real collector and the real save path with this is what makes the
 * "surface ⇄ document" round trip testable headlessly -- and it is deliberately
 * built from the same run formatting, so a size or colour the surface cannot show
 * again shows up as a byte difference instead of passing silently.
 */
export function fakeDomFor(internals, parsed) {
  const runNodes = (run) => {
    const css = internals.runStyleCss(run.fmt)
    const nodes = []
    String(run.text).split('\n').forEach((line, index) => {
      if (index > 0) nodes.push(fakeElement('br', {}, []))
      if (line === '') return
      let node = fakeText(line)
      if (css !== '') node = fakeElement('span', { style: css }, [node])
      if (run.fmt.u) node = fakeElement('u', {}, [node])
      if (run.fmt.b) node = fakeElement('strong', {}, [node])
      if (run.fmt.i) node = fakeElement('em', {}, [node])
      nodes.push(node)
    })
    return nodes
  }

  const children = []
  for (const block of parsed.blocks) {
    if (block.kind !== 'p') {
      if (block.name === 'tbl')
        children.push(fakeElement('table', { 'data-docx-table': '1', contenteditable: 'false' }, []))
      continue
    }
    const attrs = { 'data-block-id': String(block.id), 'data-docx-style': block.style }
    if (block.numbered === true) attrs['data-docx-list'] = '1'
    if (block.protected === true) {
      attrs['data-docx-protected'] = '1'
      attrs.contenteditable = 'false'
    }
    const nodes = block.runs.flatMap(runNodes)
    children.push(fakeElement('p', attrs, nodes.length === 0 ? [fakeElement('br', {}, [])] : nodes))
  }
  return fakeElement('div', {}, children)
}
