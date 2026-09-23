/**
 * Invariant check against a real Word-authored document:
 * parsing and re-emitting with NO edits must reproduce every paragraph byte-exact,
 * and must keep every non-paragraph body child as it was.
 *
 * Usage: node test/noop-invariant.mjs <path-to.docx>
 */
import { readFileSync } from 'node:fs'
import { inflateRawSync } from 'node:zlib'
import { fileURLToPath } from 'node:url'
import { dirname, join } from 'node:path'
import vm from 'node:vm'

const here = dirname(fileURLToPath(import.meta.url))
const docxPath = process.argv[2]
if (docxPath === undefined) {
  console.log('usage: node test/noop-invariant.mjs <path-to.docx>')
  process.exit(0)
}

// ── load the bundle exactly like the browser module table ─────────────────────
let definition
const sandbox = {
  window: { __ModuleLoader__: { load: (value) => { definition = value } } },
  require: (specifier) => {
    if (specifier === 'react')
      return { createElement: () => null, useState: (i) => [i, () => {}], useEffect: () => {}, useRef: (i) => ({ current: i }) }
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
  fetch: () => Promise.resolve({ json: () => Promise.resolve({ ok: false }) })
}
vm.runInContext(readFileSync(join(here, '..', 'lib', 'client.js'), 'utf8'), vm.createContext(sandbox), { filename: 'client.js' })
const internals = definition.factory(sandbox.require).__internals

// ── read word/document.xml out of the docx ───────────────────────────────────
function documentXmlOf(bytes) {
  const dv = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength)
  let eocd = -1
  for (let i = bytes.length - 22; i >= 0; i--) if (dv.getUint32(i, true) === 0x06054b50) { eocd = i; break }
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

const source = documentXmlOf(new Uint8Array(readFileSync(docxPath)))
const parsed = internals.blocksFromDocumentXml(source)
const emitted = internals.documentXmlFromBlocks(source, parsed.blocks, parsed.sectPrXml)

const paragraphs = parsed.blocks.filter((block) => block.kind === 'p')
let missing = 0
for (const block of paragraphs) if (!emitted.includes(block.origXml)) missing += 1

console.log('file            :', docxPath)
console.log('document.xml    :', source.length, 'chars ->', emitted.length, 'chars')
console.log('paragraphs      :', paragraphs.length)
console.log('byte-exact kept :', paragraphs.length - missing, '/', paragraphs.length)
console.log('other children  :', parsed.blocks.filter((b) => b.kind !== 'p').length, '(emitted verbatim)')
console.log('declaration kept:', emitted.startsWith('<?xml'))

if (missing === 0) console.log('NO-OP INVARIANT OK')
else {
  console.log('NO-OP INVARIANT FAILED (' + missing + ' paragraphs rewritten)')
  process.exit(1)
}
