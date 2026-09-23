/**
 * Docx round-trip test.
 *
 * The bundle is loaded the way the browser module table loads it, then the pure
 * parse ⇄ emit path is exercised directly through the `__internals` seam. This is
 * the regression net for the bugs that made edits vanish or formatting bleed:
 * a baseline taken from the current content, paragraphs rebuilt from scratch,
 * and formatting-only edits that were never detected as changes.
 */
import { loadInternals, fakeDomFor } from './harness.mjs'

const internals = loadInternals()
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
/** A copy of a block's runs as the editor's DOM reports them, formatting included. */
function textRuns(block) {
  return block.runs.map((run) => ({ text: run.text, fmt: { ...run.fmt } }))
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

// ── 10. font size and colour: shown, settable, clearable, never lost ─────────
{
  // The surface has to show the document's own size/colour, or it lies about the text.
  const html = internals.blocksToHtml([parsed.blocks[P2]])
  contains('surface HTML: run size rendered', html, 'font-size:14pt')
  contains('surface HTML: run colour rendered', html, 'color:#FF0000')

  // Changing only the size must count as a change, exactly like bold-only.
  const sized = textRuns(parsed.blocks[P2])
  sized[1].fmt.sz = 24
  const sizedOut = internals.documentXmlFromBlocks(DOC, [collect(P2, sized)], parsed.sectPrXml)
  contains('size-only edit: sz written', sizedOut, '<w:sz w:val="24"/>')
  check('size-only edit: detected', sizedOut !== parsed.blocks[P2].origXml, 'emitted unchanged')

  // Changing only the colour must too.
  const coloured = textRuns(parsed.blocks[P2])
  coloured[1].fmt.color = '00FF00'
  const colouredOut = internals.documentXmlFromBlocks(DOC, [collect(P2, coloured)], parsed.sectPrXml)
  contains('colour-only edit: color written', colouredOut, '<w:color w:val="00FF00"/>')

  // Setting size and colour on a run that had no rPr: schema order is color, then sz.
  const fresh = internals.withBaseline(
    { kind: 'p', id: -1, style: 'Normal', numbered: false, runs: [{ text: 'plain', fmt: { b: false, i: false, u: false, sz: 21, color: '3366CC' } }] },
    undefined
  )
  const freshOut = internals.documentXmlFromBlocks(DOC, [fresh], parsed.sectPrXml)
  contains('new run rPr: color before sz (schema order)', freshOut, '<w:rPr><w:color w:val="3366CC"/><w:sz w:val="21"/></w:rPr>')

  // Clearing asks for removal explicitly; the colour around it stays.
  const cleared = textRuns(parsed.blocks[P2])
  cleared[0].fmt = { b: false, i: false, u: false, sz: undefined, color: 'FF0000', szClear: true }
  const clearedOut = internals.documentXmlFromBlocks(DOC, [collect(P2, cleared)], parsed.sectPrXml)
  check('clear size: sz removed', !clearedOut.includes('<w:sz'), clearedOut.match(/<w:rPr>.*?<\/w:rPr>/)?.[0])
  contains('clear size: colour kept', clearedOut, '<w:color w:val="FF0000"/>')

  // A run whose size the collector could not state must keep the document's own size:
  // "not stated" is not "delete it".
  const unstated = textRuns(parsed.blocks[P2])
  unstated[0].text = 'Redder'
  unstated[0].fmt = { b: false, i: false, u: false }
  const unstatedOut = internals.documentXmlFromBlocks(DOC, [collect(P2, unstated)], parsed.sectPrXml)
  contains('unstated size: document size survives a text edit', unstatedOut, '<w:sz w:val="28"/>')
  contains('unstated colour: document colour survives a text edit', unstatedOut, '<w:color w:val="FF0000"/>')

  // The collector reads the surface's own markup back into run formatting.
  const fakeText = (value) => ({ nodeType: 3, nodeValue: value, childNodes: [] })
  const fakeElement = (tag, attrs, children) => ({
    nodeType: 1,
    tagName: tag.toUpperCase(),
    childNodes: children,
    getAttribute: (name) => (Object.prototype.hasOwnProperty.call(attrs, name) ? attrs[name] : null),
  })
  const wrapped = fakeElement('p', {}, [
    fakeElement('span', { style: 'font-size:10.5pt;color:#3366cc' }, [fakeText('styled')]),
    fakeText(' plain'),
    fakeElement('span', { style: 'color: rgb(255, 0, 0)' }, [fakeText(' rgb')]),
    fakeElement('span', { 'data-docx-reset': 'sz,color' }, [fakeText(' reset')]),
  ])
  const collected = internals.collectRuns(wrapped)
  check('collector: one run per formatted stretch', collected.length === 4, 'runs: ' + JSON.stringify(collected.map((run) => run.text)))
  check('collector: half-point size from pt', collected[0].fmt.sz === 21, JSON.stringify(collected[0].fmt))
  check('collector: colour upper-cased', collected[0].fmt.color === '3366CC', JSON.stringify(collected[0].fmt))
  check('collector: rgb() colour', collected[2].fmt.color === 'FF0000', JSON.stringify(collected[2]))
  check('collector: reset clears size and colour', collected[3].fmt.sz === undefined && collected[3].fmt.color === undefined, JSON.stringify(collected[3]))
  // What Word actually writes: rFonts, colour, w:sz AND w:szCs for complex script text.
  const WORD_LIKE = [
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n',
    '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>',
    '<w:p><w:r><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:color w:val="FF0000"/><w:sz w:val="28"/><w:szCs w:val="28"/></w:rPr><w:t>Red</w:t></w:r></w:p>',
    '<w:p><w:r><w:rPr><w:color w:val="auto"/><w:sz w:val="21"/></w:rPr><w:t>Auto</w:t></w:r></w:p>',
    '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr>',
    '</w:body></w:document>',
  ].join('')
  const wordParsed = internals.blocksFromDocumentXml(WORD_LIKE)
  check('Word rPr: w:sz read as half-points', wordParsed.blocks[0].runs[0].fmt.sz === 28, JSON.stringify(wordParsed.blocks[0].runs[0].fmt))
  check('Word rPr: colour read', wordParsed.blocks[0].runs[0].fmt.color === 'FF0000', JSON.stringify(wordParsed.blocks[0].runs[0].fmt))
  check('Word rPr: w:color val="auto" stays inherited', wordParsed.blocks[1].runs[0].fmt.color === undefined, JSON.stringify(wordParsed.blocks[1].runs[0].fmt))
  check('Word rPr: half-point size read', wordParsed.blocks[1].runs[0].fmt.sz === 21, JSON.stringify(wordParsed.blocks[1].runs[0].fmt))

  // The no-op save has to be byte-exact for that shape too.
  const wordDom = internals.collectEditorBlocks(fakeDomFor(internals, wordParsed), { blocks: wordParsed.blocks })
  check(
    'Word rPr: surface round trip is byte-exact',
    internals.saveDocumentXml(WORD_LIKE, wordParsed.blocks, wordDom, wordParsed.sectPrXml) === WORD_LIKE
  )

  // Resizing keeps the font and colour, and moves w:szCs together with w:sz.
  const resized = { ...wordParsed.blocks[0], runs: [{ text: 'Red', fmt: { b: false, i: false, u: false, sz: 24, color: 'FF0000' } }] }
  const resizedOut = internals.documentXmlFromBlocks(WORD_LIKE, [resized], wordParsed.sectPrXml)
  contains('resize: w:sz updated', resizedOut, '<w:sz w:val="24"/>')
  contains('resize: w:szCs follows', resizedOut, '<w:szCs w:val="24"/>')
  contains('resize: font kept', resizedOut, '<w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/>')
  contains('resize: colour kept', resizedOut, '<w:color w:val="FF0000"/>')
}

console.log('')
if (failures === 0) {
  console.log('ROUND TRIP OK')
} else {
  console.log('ROUND TRIP FAILED (' + failures + ')')
  process.exit(1)
}
