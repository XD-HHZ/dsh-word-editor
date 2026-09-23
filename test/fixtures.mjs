/**
 * document.xml fixtures for structures a real Word file contains but the editor
 * does not edit: tables, hyperlinks, images, fields, line breaks, bookmarks and
 * block-level content controls.
 *
 * They are XML strings rather than .docx binaries on purpose: the repo ships no
 * fixtures, CI needs no network, and the thing under test is the document.xml
 * layer, not the zip container.
 */

const HEAD =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' +
  '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"' +
  ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"' +
  ' xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"' +
  ' xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"' +
  ' xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">'

const SECT_PR =
  '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>' +
  '<w:pgMar w:top="1440" w:right="1800" w:bottom="1440" w:left="1800" w:header="851" w:footer="992" w:gutter="0"/>' +
  '</w:sectPr>'

/** Build a document.xml from body children plus the tail sectPr. */
function doc(body) {
  return HEAD + '<w:body>' + body + SECT_PR + '</w:body></w:document>'
}

const P_HELLO = '<w:p><w:r><w:t>Hello</w:t></w:r></w:p>'

const P_STYLED =
  '<w:p><w:pPr><w:pStyle w:val="Heading1"/><w:ind w:firstLine="420"/><w:jc w:val="center"/></w:pPr>' +
  '<w:r><w:rPr><w:sz w:val="28"/><w:color w:val="FF0000"/></w:rPr><w:t>Red</w:t></w:r>' +
  '<w:r><w:t xml:space="preserve"> and plain</w:t></w:r></w:p>'

const P_EMPTY = '<w:p/>'

/** A table whose cells each hold a paragraph -- the paragraphs are NOT editable. */
const TBL =
  '<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/><w:tblW w:w="0" w:type="auto"/></w:tblPr>' +
  '<w:tblGrid><w:gridCol w:w="4675"/><w:gridCol w:w="4675"/></w:tblGrid>' +
  '<w:tr><w:tc><w:tcPr><w:tcW w:w="4675" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>A1</w:t></w:r></w:p></w:tc>' +
  '<w:tc><w:tcPr><w:tcW w:w="4675" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>B1</w:t></w:r></w:p></w:tc></w:tr>' +
  '<w:tr><w:tc><w:p><w:r><w:t>A2</w:t></w:r></w:p></w:tc>' +
  '<w:tc><w:p><w:r><w:t>B2</w:t></w:r></w:p></w:tc></w:tr></w:tbl>'

export const TABLE = doc(P_HELLO + TBL + '<w:p><w:r><w:t>After the table</w:t></w:r></w:p>')

const HYPERLINK =
  '<w:p><w:r><w:t xml:space="preserve">See </w:t></w:r>' +
  '<w:hyperlink r:id="rId7" w:history="1">' +
  '<w:r><w:rPr><w:rStyle w:val="Hyperlink"/></w:rPr><w:t>the docs</w:t></w:r>' +
  '</w:hyperlink>' +
  '<w:r><w:t xml:space="preserve"> for more.</w:t></w:r></w:p>'

const DRAWING =
  '<w:p><w:r><w:t xml:space="preserve">Logo: </w:t></w:r>' +
  '<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0">' +
  '<wp:extent cx="190500" cy="190500"/><wp:docPr id="1" name="Picture 1"/>' +
  '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
  '<pic:pic><pic:nvPicPr><pic:cNvPr id="1" name="logo.png"/><pic:cNvPicPr/></pic:nvPicPr>' +
  '<pic:blipFill><a:blip r:embed="rId9"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>' +
  '<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="190500" cy="190500"/></a:xfrm>' +
  '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic>' +
  '</a:graphicData></a:graphic></wp:inline></w:drawing></w:r>' +
  '<w:r><w:t xml:space="preserve"> attached</w:t></w:r></w:p>'

/** A soft line break and a tab inside runs: text survives, but the marks do not. */
const BREAK_AND_TAB =
  '<w:p><w:r><w:t>line one</w:t><w:br/><w:t>line two</w:t></w:r>' +
  '<w:r><w:tab/><w:t>after tab</w:t></w:r></w:p>'

const BOOKMARK_AND_FIELD =
  '<w:p><w:bookmarkStart w:id="0" w:name="mark1"/>' +
  '<w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
  '<w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r>' +
  '<w:r><w:fldChar w:fldCharType="end"/></w:r>' +
  '<w:r><w:t>1</w:t></w:r>' +
  '<w:bookmarkEnd w:id="0"/></w:p>'

/**
 * A heading Word wrapped in a bookmark for its table of contents. The paragraph is
 * ordinary text, so it MUST stay editable -- a bookmark is zero-width and can be
 * re-emitted in place instead of forcing the whole paragraph read-only.
 */
export const BOOKMARKED = doc(
  '<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr>' +
  '<w:bookmarkStart w:id="1" w:name="_Toc123"/>' +
  '<w:r><w:t>Chapter One</w:t></w:r>' +
  '<w:bookmarkEnd w:id="1"/></w:p>' +
  '<w:p><w:r><w:t>body</w:t></w:r></w:p>'
)

/** Block-level content control wrapping a paragraph. */
const SDT =
  '<w:sdt><w:sdtPr><w:alias w:val="Title"/></w:sdtPr><w:sdtContent>' +
  '<w:p><w:r><w:t>Inside a content control</w:t></w:r></w:p>' +
  '</w:sdtContent></w:sdt>'

/** Everything at once, interleaved with editable paragraphs. */
export const MIXED = doc(
  P_HELLO +
  HYPERLINK +
  TBL +
  P_STYLED +
  DRAWING +
  BREAK_AND_TAB +
  BOOKMARK_AND_FIELD +
  SDT +
  P_EMPTY +
  '<w:p><w:r><w:t>tail</w:t></w:r></w:p>'
)
