/**
 * Structure tests: what happens to the parts of a document the editor does not
 * edit -- tables, hyperlinks, images, fields, line breaks, bookmarks, content
 * controls -- when the user opens a file and saves it.
 *
 * This is the data-loss test suite. It deliberately drives the SAME function the
 * save button drives (`saveDocumentXml`) rather than the lower-level emitter, so
 * a regression in the save path cannot hide behind a green invariant test.
 *
 * Usage: node test/structure.mjs
 */
import { loadInternals, createChecker } from './harness.mjs'
import { TABLE, MIXED, BOOKMARKED } from './fixtures.mjs'

const internals = loadInternals()
const { check, state } = createChecker()

/**
 * What the editor's DOM yields for a document: one block per EDITABLE paragraph,
 * carrying its baseline. Protected paragraphs are not collected (they render
 * contenteditable=false), and non-paragraph body children never reach the DOM at
 * all -- that is exactly the shape of list the save path has to cope with.
 */
function collectLikeDom(parsed) {
  return parsed.blocks
    .filter((block) => block.kind === 'p' && block.protected !== true)
    .map((block) =>
      internals.withBaseline(
        {
          kind: 'p',
          id: block.id,
          style: block.style,
          numbered: block.numbered,
          runs: block.runs.map((run) => ({ text: run.text, fmt: { b: run.fmt.b, i: run.fmt.i, u: run.fmt.u } })),
        },
        block
      )
    )
}

function paragraphWith(parsed, text) {
  return parsed.blocks.find((block) => block.kind === 'p' && internals.runsText(block.runs).includes(text))
}

// ── 1. open + save with no edits must reproduce the file byte-for-byte ───────
for (const [label, source] of [['table', TABLE], ['mixed', MIXED]]) {
  const parsed = internals.blocksFromDocumentXml(source)
  const saved = internals.saveDocumentXml(source, parsed.blocks, collectLikeDom(parsed), parsed.sectPrXml)
  check(
    label + ': no-op save is byte-exact',
    saved === source,
    'length ' + source.length + ' -> ' + saved.length
  )
  const dropped = parsed.blocks.filter((block) => block.kind === 'other' && !saved.includes(block.xml))
  check(
    label + ': every non-paragraph body child survives',
    dropped.length === 0,
    'dropped: ' + dropped.map((block) => block.name).join(', ')
  )
}

// ── 2. paragraphs the flat rebuild cannot reproduce are marked read-only ─────
{
  const parsed = internals.blocksFromDocumentXml(MIXED)
  const protectedTexts = parsed.blocks
    .filter((block) => block.kind === 'p' && block.protected === true)
    .map((block) => internals.runsText(block.runs))
  const editableTexts = parsed.blocks
    .filter((block) => block.kind === 'p' && block.protected !== true)
    .map((block) => internals.runsText(block.runs))

  check('hyperlink paragraph is protected', protectedTexts.some((text) => text.includes('the docs')), 'protected: ' + JSON.stringify(protectedTexts))
  check('drawing paragraph is protected', protectedTexts.some((text) => text.includes('Logo')), 'protected: ' + JSON.stringify(protectedTexts))
  check('line-break paragraph is protected', protectedTexts.some((text) => text.includes('line one')), 'protected: ' + JSON.stringify(protectedTexts))
  check('field paragraph is protected', protectedTexts.some((text) => text.includes('1')), 'protected: ' + JSON.stringify(protectedTexts))
  check('plain paragraphs stay editable', editableTexts.some((text) => text === 'Hello') && editableTexts.some((text) => text.includes('tail')), 'editable: ' + JSON.stringify(editableTexts))
}

// ── 3. a bookmarked heading must stay editable AND keep its bookmark ────────
{
  const parsed = internals.blocksFromDocumentXml(BOOKMARKED)
  const heading = paragraphWith(parsed, 'Chapter One')
  check('bookmarked heading is editable', heading !== undefined && heading.protected !== true)

  const collected = collectLikeDom(parsed).map((block) =>
    block.id === heading.id ? { ...block, runs: [{ text: 'Chapter One edited', fmt: { b: false, i: false, u: false } }] } : block
  )
  const saved = internals.saveDocumentXml(BOOKMARKED, parsed.blocks, collected, parsed.sectPrXml)
  check('bookmark kept after editing the heading', saved.includes('<w:bookmarkStart w:id="1" w:name="_Toc123"/>') && saved.includes('<w:bookmarkEnd w:id="1"/>'), saved)
  check('heading text edit landed', saved.includes('Chapter One edited'), saved)
  check('heading pStyle kept', saved.includes('<w:pStyle w:val="Heading1"/>'), saved)
}

// ── 4. editing a normal paragraph must not disturb anything around it ───────
{
  const parsed = internals.blocksFromDocumentXml(MIXED)
  const hello = paragraphWith(parsed, 'Hello')
  const collected = collectLikeDom(parsed).map((block) =>
    block.id === hello.id ? { ...block, runs: [{ text: 'Hello edited', fmt: { b: false, i: false, u: false } }] } : block
  )
  const saved = internals.saveDocumentXml(MIXED, parsed.blocks, collected, parsed.sectPrXml)
  check('edit: new text written', saved.includes('Hello edited'), '')
  check('edit: table still byte-exact', saved.includes('<w:tbl>') && saved.includes('<w:tcPr><w:tcW w:w="4675" w:type="dxa"/></w:tcPr>'), '')
  check('edit: hyperlink element untouched', saved.includes('<w:hyperlink r:id="rId7" w:history="1">'), '')
  check('edit: drawing untouched', saved.includes('<a:blip r:embed="rId9"/>'), '')
  check('edit: block-level content control untouched', saved.includes('<w:sdtContent>'), '')
  check('edit: sectPr untouched', saved.includes('<w:pgMar w:top="1440"'), '')
}

// ── 5. a protected paragraph can never be rewritten, whatever the DOM says ──
{
  const parsed = internals.blocksFromDocumentXml(MIXED)
  const link = paragraphWith(parsed, 'the docs')
  // Simulate a collector that DID include it and reported mangled text.
  const tampered = internals.withBaseline(
    { kind: 'p', id: link.id, style: link.style, numbered: link.numbered, runs: [{ text: 'MANGLED', fmt: { b: false, i: false, u: false } }] },
    link
  )
  const collected = collectLikeDom(parsed).concat([tampered])
  const saved = internals.saveDocumentXml(MIXED, parsed.blocks, collected, parsed.sectPrXml)
  check('protected paragraph keeps its original bytes', saved.includes('<w:hyperlink r:id="rId7"'), saved.slice(0, 200))
  check('protected paragraph ignores mangled text', !saved.includes('MANGLED'))
  check('protected paragraph still byte-exact', saved === MIXED)
}

// ── 6. the surface tells the user what it will not edit ─────────────────────
{
  const parsed = internals.blocksFromDocumentXml(MIXED)
  const html = internals.blocksToHtml(parsed.blocks)
  check('surface marks protected paragraphs', /data-docx-protected="1"[^>]*contenteditable="false"|contenteditable="false"[^>]*data-docx-protected="1"/.test(html), '')
  check('surface renders the table', html.includes('<table') && html.includes('A1') && html.includes('B2'), '')
  check('surface renders the content control text', html.includes('Inside a content control'), '')
  check('surface does not render bare bookmark markers', !html.includes('bookmarkStart'), '')

  const table = internals.blocksFromDocumentXml(TABLE)
  const tableHtml = internals.blocksToHtml(table.blocks)
  check('table-only document is not blank', tableHtml.includes('<table') && tableHtml.includes('After the table'), tableHtml)
}

// ── 7. the whole loop: a DOM-shaped collection feeds the real save path ─────
// Everything above hands `saveDocumentXml` a hand-built list. Here the list comes
// out of the real collector, driven by a fake DOM shaped the way the surface
// builds it, so "the editor cannot lose what it does not render" is tested end to
// end rather than assumed.
{
  const fakeText = (value) => ({ nodeType: 3, nodeValue: value, textContent: value, childNodes: [] })
  const fakeElement = (tag, attrs, children) => {
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

  /** Build the DOM the surface would render for a parsed document. */
  const surfaceFor = (parsed) => {
    const children = []
    for (const block of parsed.blocks) {
      if (block.kind !== 'p') {
        if (block.name === 'tbl') children.push(fakeElement('table', { 'data-docx-table': '1', contenteditable: 'false' }, []))
        continue
      }
      const attrs = { 'data-block-id': String(block.id), 'data-docx-style': block.style }
      if (block.protected === true) {
        attrs['data-docx-protected'] = '1'
        attrs.contenteditable = 'false'
      }
      const runs = block.runs.map((run) => fakeText(run.text))
      children.push(fakeElement('p', attrs, runs.length === 0 ? [fakeElement('br', {}, [])] : runs))
    }
    return fakeElement('div', {}, children)
  }

  for (const [label, source] of [['table', TABLE], ['mixed', MIXED]]) {
    const parsed = internals.blocksFromDocumentXml(source)
    const collected = internals.collectEditorBlocks(surfaceFor(parsed), { blocks: parsed.blocks })
    const byId = collected.map((block) => block.id)
    const protectedIds = parsed.blocks.filter((block) => block.kind === 'p' && block.protected === true).map((block) => block.id)
    check(
      label + ': collector skips every protected paragraph',
      protectedIds.every((id) => !byId.includes(id)),
      'collected ids: ' + JSON.stringify(byId)
    )
    const saved = internals.saveDocumentXml(source, parsed.blocks, collected, parsed.sectPrXml)
    check(label + ': collected DOM round-trips byte-exact', saved === source, 'length ' + source.length + ' -> ' + saved.length)
  }

  // A paragraph typed in the editor (no id) must be inserted, not dropped.
  {
    const parsed = internals.blocksFromDocumentXml(TABLE)
    const root = surfaceFor(parsed)
    root.childNodes.splice(1, 0, fakeElement('p', { 'data-docx-style': 'Normal' }, [fakeText('brand new')]))
    const collected = internals.collectEditorBlocks(root, { blocks: parsed.blocks })
    const saved = internals.saveDocumentXml(TABLE, parsed.blocks, collected, parsed.sectPrXml)
    check('new paragraph is written', /<w:t[^>]*>brand new<\/w:t>/.test(saved), saved.match(/<w:body>[\s\S]{0,160}/)?.[0])
    check('new paragraph did not replace a neighbour', saved.includes('<w:t>Hello</w:t>'), '')
    check('new paragraph did not drop the table', saved.includes('<w:tbl>'), '')
  }
}

console.log('')
if (state.failures === 0) {
  console.log('STRUCTURE OK')
} else {
  console.error(state.failures + ' check(s) failed')
  process.exit(1)
}
