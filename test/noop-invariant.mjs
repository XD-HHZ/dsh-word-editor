/**
 * Invariant check against a real Word-authored document: opening and saving with
 * NO edits must reproduce document.xml byte-for-byte, and must keep every
 * non-paragraph body child (tables, content controls, markers).
 *
 * It drives `saveDocumentXml` -- the same function the save button drives -- with
 * the block list an unedited editor would produce. Testing the lower-level emitter
 * with the full parsed block list instead is what previously hid the fact that the
 * save path dropped every non-paragraph body child.
 *
 * Usage: node test/noop-invariant.mjs <path-to.docx>
 */
import { readFileSync } from 'node:fs'
import { loadInternals, documentXmlOf } from './harness.mjs'

const docxPath = process.argv[2]
if (docxPath === undefined) {
  console.log('usage: node test/noop-invariant.mjs <path-to.docx>')
  process.exit(0)
}

const internals = loadInternals()
const source = documentXmlOf(new Uint8Array(readFileSync(docxPath)))
const parsed = internals.blocksFromDocumentXml(source)

// An unedited editor collects every editable paragraph and nothing else.
const collected = parsed.blocks
  .filter((block) => block.kind === 'p' && block.protected !== true)
  .map((block) =>
    internals.withBaseline(
      {
        kind: 'p',
        id: block.id,
        style: block.style,
        numbered: block.numbered,
        runs: block.runs.map((run) => ({ text: run.text, fmt: { ...run.fmt } })),
      },
      block
    )
  )

const emitted = internals.saveDocumentXml(source, parsed.blocks, collected, parsed.sectPrXml)

const paragraphs = parsed.blocks.filter((block) => block.kind === 'p')
const protectedCount = paragraphs.filter((block) => block.protected === true).length
const others = parsed.blocks.filter((block) => block.kind !== 'p')
const lost = others.filter((block) => !emitted.includes(block.xml))

console.log('file            :', docxPath)
console.log('document.xml    :', source.length, 'chars ->', emitted.length, 'chars')
console.log('paragraphs      :', paragraphs.length, '(' + protectedCount + ' read-only)')
console.log('other children  :', others.length, '(emitted verbatim)')
console.log('declaration kept:', emitted.startsWith('<?xml'))

if (emitted === source) {
  console.log('NO-OP INVARIANT OK (byte-for-byte)')
} else if (lost.length > 0) {
  console.log('NO-OP INVARIANT FAILED: dropped ' + lost.map((block) => block.name).join(', '))
  process.exit(1)
} else {
  console.log('NO-OP INVARIANT FAILED: document.xml was rewritten')
  process.exit(1)
}
