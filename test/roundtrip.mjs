/**
 * Docx round-trip test.
 *
 * The bundle is loaded the way the browser module table loads it, then the pure
 * parse ⇄ emit path is exercised directly through the `__internals` seam. This is
 * the regression net for the two bugs that made edits vanish or formatting bleed:
 * a baseline taken from the current content, and paragraphs rebuilt from scratch.
 */
import { readFileSync } from 'node:fs'
import { fileURLToPath } from 'node:url'
import { dirname, join } from 'node:path'
import vm from 'node:vm'

const here = dirname(fileURLToPath(import.meta.url))
const clientPath = join(here, '..', 'lib', 'client.js')

const fakeReact = {
  createElement: (...args) => ({ type: args[0], props: args[1] }),
  useState: (initial) => [initial, () => {}],
  useEffect: () => {},
  useRef: (initial) => ({ current: initial })
}

let definition
const sandbox = {
  window: { __ModuleLoader__: { load: (value) => { definition = value } } },
  require: (specifier) => {
    if (specifier === 'react') return fakeReact
    throw new Error('unexpected require: ' + specifier)
  },
  fetch: () => Promise.resolve({ json: () => Promise.resolve({ ok: false }) }),
  console,
  TextEncoder,
  TextDecoder,
  Blob,
  Response,
  DecompressionStream,
  CompressionStream,
  btoa: (value) => Buffer.from(value, 'binary').toString('base64'),
  atob: (value) => Buffer.from(value, 'base64').toString('binary'),
  setTimeout,
  clearInterval,
  setInterval,
  URL
}

vm.runInContext(readFileSync(clientPath, 'utf8'), vm.createContext(sandbox), { filename: 'client.js' })
const internals = definition.factory(sandbox.require).__internals
if (internals === undefined) throw new Error('bundle exposes no __internals seam')

const DOC = [
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
  '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>',
  '<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Title</w:t></w:r></w:p>',
  '<w:p><w:pPr><w:ind w:firstLine="420"/><w:jc w:val="center"/></w:pPr>',
  '<w:r><w:rPr><w:sz w:val="28"/><w:color w:val="FF0000"/></w:rPr><w:t>Red</w:t></w:r>',
  '<w:r><w:t xml:space="preserve"> and plain</w:t></w:r></w:p>',
  '<w:p><w:pPr><w:ind w:firstLine="420"/></w:pPr><w:r><w:t>Indented</w:t></w:r></w:p>',
  '<w:p/>',
  '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr>',
  '</w:body></w:document>'
].join('')

const P2 = 1
const P3 = 2

let failures = 0
function check(label, condition, detail) {
  if (condition) {
    console.log('  PASS  ' + label)
    return
  }
  failures += 1
  console.log('  FAIL  ' + label)
  if (detail !== undefined) console.log('        ' + detail)
}
function contains(label, haystack, needle) {
  check(label, haystack.includes(needle), 'missing: ' + needle)
}

const parsed = internals.blocksFromDocumentXml(DOC)

// ── 1. an untouched document is emitted byte-exact ────────────────────────────
const untouched = internals.documentXmlFromBlocks(DOC, parsed.blocks, parsed.sectPrXml)
check('untouched: every paragraph byte-exact', parsed.blocks.every((block) => block.kind !== 'p' || untouched.includes(block.origXml)))
contains('untouched: empty paragraph survives as <w:p/>', untouched, '<w:p/>')

/** Collect-path baseline: current runs, original formatting inherited. */
function collect(index, runs, overrides = {}) {
  const orig = parsed.blocks[index]
  return internals.withBaseline(
    {
      kind: 'p',
      id: orig.id,
      style: overrides.style !== undefined ? overrides.style : orig.style,
      numbered: overrides.numbered !== undefined ? overrides.numbered : orig.numbered,
      runs
    },
    orig
  )
}
function textRuns(block) {
  return block.runs.map((run) => ({ text: run.text, fmt: { b: run.fmt.b, i: run.fmt.i, u: run.fmt.u } }))
}

// ── 2. edit text in the second run: pPr and the untouched run survive ─────────
{
  const runs = textRuns(parsed.blocks[P2])
  runs[1].text = ' and plain EDITED'
  const out = internals.documentXmlFromBlocks(DOC, [collect(P2, runs)], parsed.sectPrXml)
  contains('text edit: pPr reused byte-exact', out, parsed.blocks[P2].origPPrXml)
  contains('text edit: indentation kept', out, '<w:ind w:firstLine="420"/>')
  contains('text edit: alignment kept', out, '<w:jc w:val="center"/>')
  contains('text edit: untouched run byte-exact', out, '<w:r><w:rPr><w:sz w:val="28"/><w:color w:val="FF0000"/></w:rPr><w:t>Red</w:t></w:r>')
  contains('text edit: new text written', out, ' and plain EDITED')
}

// ── 3. edit text inside the formatted run: its rPr survives ───────────────────
{
  const runs = textRuns(parsed.blocks[P2])
  runs[0].text = 'Redder'
  const out = internals.documentXmlFromBlocks(DOC, [collect(P2, runs)], parsed.sectPrXml)
  contains('formatted-run text edit: size kept', out, '<w:sz w:val="28"/>')
  contains('formatted-run text edit: colour kept', out, '<w:color w:val="FF0000"/>')
  contains('formatted-run text edit: new text written', out, '<w:t xml:space="preserve">Redder</w:t>')
}

// ── 4. toggle bold on the formatted run: other rPr children kept, order valid ──
{
  const runs = textRuns(parsed.blocks[P2])
  runs[0].fmt.b = true
  const out = internals.documentXmlFromBlocks(DOC, [collect(P2, runs)], parsed.sectPrXml)
  const actual = out.match(/<w:rPr>.*?<\/w:rPr>/)?.[0]
  check(
    'bold toggle: b inserted before sz/color, size and colour kept',
    actual === '<w:rPr><w:b/><w:sz w:val="28"/><w:color w:val="FF0000"/></w:rPr>',
    'actual: ' + String(actual)
  )
  // Regression guard: an edit that changes no text at all must still be detected.
  // Change detection compares text AND formatting, so this paragraph must not be
  // emitted as the original bytes.
  check(
    'bold toggle: format-only edit detected',
    out !== parsed.blocks[P2].origXml,
    'paragraph was emitted unchanged'
  )
}

// ── 5. change the style of an indented paragraph: pPr is patched, not replaced ─
{
  const out = internals.documentXmlFromBlocks(DOC, [collect(P3, textRuns(parsed.blocks[P3]), { style: 'Heading2' })], parsed.sectPrXml)
  contains('style change: new pStyle', out, '<w:pStyle w:val="Heading2"/>')
  contains('style change: indentation kept', out, '<w:ind w:firstLine="420"/>')
  check(
    'style change: pStyle comes first',
    out.includes('<w:pPr><w:pStyle w:val="Heading2"/><w:ind w:firstLine="420"/></w:pPr>'),
    out.match(/<w:pPr>.*?<\/w:pPr>/)?.[0]
  )
}

// ── 6. turn an indented paragraph into a list item: numPr patched in place ────
{
  const out = internals.documentXmlFromBlocks(DOC, [collect(P3, textRuns(parsed.blocks[P3]), { numbered: true })], parsed.sectPrXml)
  contains('list toggle: numPr added', out, '<w:numId w:val="1"/>')
  contains('list toggle: indentation kept', out, '<w:ind w:firstLine="420"/>')
  check(
    'list toggle: numPr before ind (schema order)',
    out.indexOf('<w:numPr>') < out.indexOf('<w:ind '),
    out.match(/<w:pPr>.*?<\/w:pPr>/)?.[0]
  )
}

// ── 7. a paragraph typed in the editor gets its style written ─────────────────
{
  const fresh = internals.withBaseline(
    { kind: 'p', id: -1, style: 'Heading1', numbered: false, runs: [{ text: 'Brand new', fmt: { b: false, i: false, u: false } }] },
    undefined
  )
  const out = internals.documentXmlFromBlocks(DOC, [fresh], parsed.sectPrXml)
  contains('new paragraph: pStyle written', out, '<w:pPr><w:pStyle w:val="Heading1"/></w:pPr>')
  contains('new paragraph: text written', out, 'Brand new')
}

// ── 8. an empty paragraph collected from the surface stays empty ──────────────
{
  const empty = collect(3, [])
  const out = internals.documentXmlFromBlocks(DOC, [empty], parsed.sectPrXml)
  contains('empty paragraph: stays <w:p/>', out, '<w:p/>')
}

// ── 9. HTML for the surface: no <br> between runs ─────────────────────────────
{
  const html = internals.blocksToHtml([parsed.blocks[0], parsed.blocks[P2]])
  check('surface HTML: exactly one block per paragraph', (html.match(/<p /g) || []).length + (html.match(/<h1 /g) || []).length === 2, html)
  check(
    'surface HTML: no <br> inserted between runs of one paragraph',
    !html.includes('Red<br>'),
    html
  )
}

console.log('')
if (failures === 0) {
  console.log('ROUND TRIP OK')
} else {
  console.log('ROUND TRIP FAILED (' + failures + ')')
  process.exit(1)
}
