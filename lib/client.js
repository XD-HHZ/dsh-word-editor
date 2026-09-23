window.__ModuleLoader__.load({
	id: "@xd-hhz/dsh-word-editor",
	factory: (require) => {
		var module = { exports: {} };
		var exports = module.exports;
		Object.defineProperty(exports, Symbol.toStringTag, { value: "Module" });
		const React = require("react");

		/** Document-preview implementation id; also the key this plugin's body registers under. */
		const IMPL_ID = "dsh-word-editor";
		/** Host bridge prefix. */
		const ROUTE = "/word-editor";
		/** Shown on content the editor keeps verbatim but does not edit. */
		const PROTECTED_HINT = "此处含表格、图片、超链接、域或换行等内容，本编辑器原样保留但不修改；需要编辑请用 Word 打开。";
		/** Point sizes the toolbar offers; `w:sz` stores them doubled, as half-points. */
		const FONT_SIZES = [9, 10.5, 12, 14, 16, 18, 24, 36];

		// ───────────────────────────── host bridge ─────────────────────────────

		async function bridgeRead(address) {
			const response = await fetch(ROUTE + "/read", {
				method: "POST",
				headers: { "content-type": "application/json" },
				body: JSON.stringify({ address })
			});
			return await response.json();
		}

		async function bridgeSave(payload) {
			const response = await fetch(ROUTE + "/save", {
				method: "POST",
				headers: { "content-type": "application/json" },
				body: JSON.stringify(payload)
			});
			return await response.json();
		}

		// ───────────────────────────── byte helpers ────────────────────────────

		function toBase64(bytes) {
			let binary = "";
			const chunk = 0x8000;
			for (let i = 0; i < bytes.length; i += chunk) {
				binary += String.fromCharCode.apply(null, Array.prototype.slice.call(bytes.subarray(i, i + chunk)));
			}
			return btoa(binary);
		}

		function fromBase64(value) {
			const binary = atob(value);
			const out = new Uint8Array(binary.length);
			for (let i = 0; i < binary.length; i++) out[i] = binary.charCodeAt(i);
			return out;
		}

		function utf8Encode(text) {
			return new TextEncoder().encode(text);
		}

		function utf8Decode(bytes) {
			return new TextDecoder("utf-8").decode(bytes);
		}

		// ─────────────────────────────── zip layer ─────────────────────────────

		function findEocd(dv, bytes) {
			for (let i = bytes.length - 22; i >= Math.max(0, bytes.length - 65557); i--) {
				if (dv.getUint32(i, true) === 0x06054b50) return i;
			}
			return -1;
		}

		async function inflateRaw(bytes) {
			const stream = new Blob([bytes]).stream().pipeThrough(new DecompressionStream("deflate-raw"));
			return new Uint8Array(await new Response(stream).arrayBuffer());
		}

		async function deflateRaw(bytes) {
			const stream = new Blob([bytes]).stream().pipeThrough(new CompressionStream("deflate-raw"));
			return new Uint8Array(await new Response(stream).arrayBuffer());
		}

		const CRC_TABLE = (() => {
			const table = new Int32Array(256);
			for (let n = 0; n < 256; n++) {
				let c = n;
				for (let k = 0; k < 8; k++) c = (c & 1) ? (0xEDB88320 ^ (c >>> 1)) : (c >>> 1);
				table[n] = c;
			}
			return table;
		})();

		function crc32(bytes) {
			let c = -1;
			for (let i = 0; i < bytes.length; i++) c = CRC_TABLE[(c ^ bytes[i]) & 0xFF] ^ (c >>> 8);
			return (c ^ -1) >>> 0;
		}

		function dosDateTime() {
			const now = new Date();
			const time = ((now.getHours() << 11) | (now.getMinutes() << 5) | (now.getSeconds() >> 1)) & 0xFFFF;
			const date = (((now.getFullYear() - 1980) << 9) | ((now.getMonth() + 1) << 5) | now.getDate()) & 0xFFFF;
			return { time, date };
		}

		/**
		 * Read a zip/docx: central directory first, then each local entry.
		 * Entry names are normalised to `/` so a zip written with backslashes
		 * (Windows `ZipFile.CreateFromDirectory`) still finds `word/document.xml`.
		 */
		async function unzip(bytes) {
			const dv = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
			const eocd = findEocd(dv, bytes);
			if (eocd < 0) throw new Error("不是有效的 zip/docx：找不到中央目录");
			const count = dv.getUint16(eocd + 10, true);
			const cdOffset = dv.getUint32(eocd + 16, true);
			const entries = [];
			let p = cdOffset;
			for (let i = 0; i < count; i++) {
				if (dv.getUint32(p, true) !== 0x02014b50) break;
				const method = dv.getUint16(p + 10, true);
				const time = dv.getUint16(p + 12, true);
				const date = dv.getUint16(p + 14, true);
				const compSize = dv.getUint32(p + 20, true);
				const nameLen = dv.getUint16(p + 28, true);
				const extraLen = dv.getUint16(p + 30, true);
				const commentLen = dv.getUint16(p + 32, true);
				const localOffset = dv.getUint32(p + 42, true);
				const name = utf8Decode(bytes.subarray(p + 46, p + 46 + nameLen)).replace(/\\/g, "/");
				entries.push({ name, method, time, date, compSize, localOffset });
				p += 46 + nameLen + extraLen + commentLen;
			}
			for (const entry of entries) {
				const lo = entry.localOffset;
				const localNameLen = dv.getUint16(lo + 26, true);
				const localExtraLen = dv.getUint16(lo + 28, true);
				const start = lo + 30 + localNameLen + localExtraLen;
				const raw = bytes.subarray(start, start + entry.compSize);
				if (entry.method === 0) entry.data = new Uint8Array(raw);
				else if (entry.method === 8) entry.data = await inflateRaw(raw);
				else throw new Error("不支持的 zip 压缩方式: " + entry.method);
			}
			return entries;
		}

		/** Rebuild a zip from entries; each part is deflated unless storing is smaller. */
		async function zip(entries) {
			const parts = [];
			const central = [];
			let offset = 0;
			const stamp = dosDateTime();
			for (const entry of entries) {
				const nameBytes = utf8Encode(String(entry.name).replace(/\\/g, "/"));
				const data = entry.data;
				let method = 8;
				let stored = await deflateRaw(data);
				if (stored.length >= data.length) {
					method = 0;
					stored = data;
				}
				const crc = crc32(data);
				const local = new Uint8Array(30 + nameBytes.length);
				const ldv = new DataView(local.buffer);
				ldv.setUint32(0, 0x04034b50, true);
				ldv.setUint16(4, 20, true);
				ldv.setUint16(6, 0, true);
				ldv.setUint16(8, method, true);
				ldv.setUint16(10, entry.time !== undefined ? entry.time : stamp.time, true);
				ldv.setUint16(12, entry.date !== undefined ? entry.date : stamp.date, true);
				ldv.setUint32(14, crc, true);
				ldv.setUint32(18, stored.length, true);
				ldv.setUint32(22, data.length, true);
				ldv.setUint16(26, nameBytes.length, true);
				ldv.setUint16(28, 0, true);
				local.set(nameBytes, 30);
				parts.push(local, stored);
				const cd = new Uint8Array(46 + nameBytes.length);
				const cdv = new DataView(cd.buffer);
				cdv.setUint32(0, 0x02014b50, true);
				cdv.setUint16(4, 20, true);
				cdv.setUint16(6, 20, true);
				cdv.setUint16(8, 0, true);
				cdv.setUint16(10, method, true);
				cdv.setUint16(12, entry.time !== undefined ? entry.time : stamp.time, true);
				cdv.setUint16(14, entry.date !== undefined ? entry.date : stamp.date, true);
				cdv.setUint32(16, crc, true);
				cdv.setUint32(20, stored.length, true);
				cdv.setUint32(24, data.length, true);
				cdv.setUint16(28, nameBytes.length, true);
				cdv.setUint16(30, 0, true);
				cdv.setUint16(32, 0, true);
				cdv.setUint16(34, 0, true);
				cdv.setUint16(36, 0, true);
				cdv.setUint32(38, 0, true);
				cdv.setUint32(42, offset, true);
				cd.set(nameBytes, 46);
				central.push(cd);
				offset += local.length + stored.length;
			}
			let cdSize = 0;
			for (const part of central) cdSize += part.length;
			const end = new Uint8Array(22);
			const edv = new DataView(end.buffer);
			edv.setUint32(0, 0x06054b50, true);
			edv.setUint16(8, entries.length, true);
			edv.setUint16(10, entries.length, true);
			edv.setUint32(12, cdSize, true);
			edv.setUint32(16, offset, true);
			const all = parts.concat(central, [end]);
			let total = 0;
			for (const part of all) total += part.length;
			const out = new Uint8Array(total);
			let o = 0;
			for (const part of all) {
				out.set(part, o);
				o += part.length;
			}
			return out;
		}

		// ─────────────────────────── minimal XML layer ─────────────────────────

		function decodeEntities(value) {
			return value.replace(/&(?:#x([0-9a-fA-F]+)|#([0-9]+)|(amp|lt|gt|quot|apos));/g, (match, hex, dec, named) => {
				if (hex !== undefined) return String.fromCodePoint(parseInt(hex, 16));
				if (dec !== undefined) return String.fromCodePoint(parseInt(dec, 10));
				if (named === "amp") return "&";
				if (named === "lt") return "<";
				if (named === "gt") return ">";
				if (named === "quot") return '"';
				if (named === "apos") return "'";
				return match;
			});
		}

		function escapeXml(value) {
			return String(value).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
		}

		function escapeAttr(value) {
			return escapeXml(value).replace(/"/g, "&quot;");
		}

		function escapeHtml(value) {
			return String(value).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
		}

		/**
		 * Parse the one root element this plugin needs. Node shape:
		 * `{ type:'element', name (local), rawName, attrs:[{name,value}], children }`
		 * or `{ type:'text', text }`. Unmodified nodes are not re-serialised from
		 * this tree: the original raw text of each paragraph is kept verbatim.
		 */
		function parseXml(src) {
			let i = 0;
			const localNameOf = (raw) => {
				const colon = raw.indexOf(":");
				return colon < 0 ? raw : raw.slice(colon + 1);
			};
			function parseAttrs(text) {
				const attrs = [];
				const re = /([^\s=\/]+)\s*=\s*(?:"([^"]*)"|'([^']*)')/g;
				let m;
				while ((m = re.exec(text)) !== null) attrs.push({ name: m[1], value: decodeEntities(m[2] !== undefined ? m[2] : m[3]) });
				return attrs;
			}
			function parseNodes() {
				const list = [];
				while (i < src.length) {
					const lt = src.indexOf("<", i);
					if (lt < 0) {
						const tail = src.slice(i);
						if (tail.trim() !== "") list.push({ type: "text", text: decodeEntities(tail) });
						i = src.length;
						break;
					}
					if (lt > i) {
						const text = src.slice(i, lt);
						if (text.trim() !== "") list.push({ type: "text", text: decodeEntities(text) });
					}
					if (src.startsWith("<!--", lt)) {
						const end = src.indexOf("-->", lt);
						i = end < 0 ? src.length : end + 3;
						continue;
					}
					if (src.startsWith("<?", lt)) {
						const end = src.indexOf("?>", lt);
						i = end < 0 ? src.length : end + 2;
						continue;
					}
					if (src.startsWith("<![CDATA[", lt)) {
						const end = src.indexOf("]]>", lt);
						list.push({ type: "text", text: src.slice(lt + 9, end) });
						i = end < 0 ? src.length : end + 3;
						continue;
					}
					if (src.startsWith("<!", lt)) {
						const end = src.indexOf(">", lt);
						i = end < 0 ? src.length : end + 1;
						continue;
					}
					if (src.startsWith("</", lt)) {
						i = lt;
						return list;
					}
					const gt = src.indexOf(">", lt);
					if (gt < 0) {
						i = src.length;
						break;
					}
					let tag = src.slice(lt + 1, gt);
					const selfClosing = tag.endsWith("/");
					if (selfClosing) tag = tag.slice(0, -1);
					const space = tag.search(/\s/);
					const rawName = space < 0 ? tag : tag.slice(0, space);
					const attrText = space < 0 ? "" : tag.slice(space + 1);
					const node = { type: "element", name: localNameOf(rawName), rawName, attrs: parseAttrs(attrText), children: [] };
					list.push(node);
					i = gt + 1;
					if (!selfClosing) {
						node.children = parseNodes();
						if (src.startsWith("</", i)) {
							const close = src.indexOf(">", i);
							i = close < 0 ? src.length : close + 1;
						}
					}
				}
				return list;
			}
			return parseNodes().find((node) => node.type === "element");
		}

		/** Attribute lookup by local name, so `w:val` answers to `val`. */
		function attr(node, name) {
			if (node === undefined || node === null || !node.attrs) return undefined;
			for (const a of node.attrs) {
				if (a.name === name) return a.value;
				const colon = a.name.indexOf(":");
				if (colon >= 0 && a.name.slice(colon + 1) === name) return a.value;
			}
			return undefined;
		}

		function childrenNamed(node, name) {
			return (node.children || []).filter((child) => child.type === "element" && child.name === name);
		}

		function firstChild(node, name) {
			return childrenNamed(node, name)[0];
		}

		function textOf(node) {
			let out = "";
			const walk = (n) => {
				if (n.type === "text") {
					out += n.text;
					return;
				}
				if (n.name === "t") {
					for (const child of n.children) if (child.type === "text") out += child.text;
					return;
				}
				if (n.name === "tab") {
					out += "\t";
					return;
				}
				if (n.name === "br" || n.name === "cr") {
					out += "\n";
					return;
				}
				for (const child of n.children || []) walk(child);
			};
			for (const child of node.children || []) walk(child);
			return out;
		}

		function serializeNode(node) {
			if (node.type === "text") return escapeXml(node.text);
			const attrs = node.attrs.map((a) => " " + a.name + '="' + escapeAttr(a.value) + '"').join("");
			if (!node.children || node.children.length === 0) return "<" + node.rawName + attrs + "/>";
			return "<" + node.rawName + attrs + ">" + node.children.map(serializeNode).join("") + "</" + node.rawName + ">";
		}

		// ─────────────────────────── docx ⇄ blocks ─────────────────────────────

		function runsText(runs) {
			return (runs || []).map((run) => run.text).join("");
		}

		/**
		 * Formatting signature of a run list: one entry per formatting stretch, carrying
		 * the b/i/u flags and how many characters they cover.
		 *
		 * Adjacent runs with identical flags collapse into one stretch, so a run the
		 * editor's DOM merged and a run the document kept separate produce the SAME
		 * signature. That keeps an untouched paragraph recognised as untouched while a
		 * format-only edit (bolding text without touching its characters) still counts
		 * as a change instead of being silently reverted on save.
		 */
		function formatSignature(runs) {
			const out = [];
			for (const run of runs || []) {
				const flag =
					(run.fmt.b ? "b" : "") + (run.fmt.i ? "i" : "") + (run.fmt.u ? "u" : "") +
					"@sz" + String(run.fmt.sz === undefined ? "" : run.fmt.sz) +
					"@c" + String(run.fmt.color === undefined ? "" : run.fmt.color);
				const last = out[out.length - 1];
				if (last !== undefined && last.flag === flag) {
					last.length += run.text.length;
					continue;
				}
				out.push({ flag, length: run.text.length });
			}
			return out;
		}

		function signaturesEqual(left, right) {
			if (left.length !== right.length) return false;
			for (let i = 0; i < left.length; i++) {
				if (left[i].flag !== right[i].flag || left[i].length !== right[i].length) return false;
			}
			return true;
		}

		/**
		 * Attach a change-detection baseline.
		 *
		 * `orig*` always describes the LAST SAVED state: for a freshly parsed
		 * document that is the paragraph's own content (and `origXml` is the raw
		 * XML the caller passes in), for a paragraph collected from the editor it
		 * is copied from the block it came from. Deriving the baseline from the
		 * current runs instead would make every paragraph compare equal and
		 * silently drop every edit.
		 */
		function withBaseline(block, orig) {
			if (orig !== undefined && orig !== null) {
				block.origXml = orig.origXml;
				block.origText = orig.origText;
				block.origStyle = orig.origStyle;
				block.origNumbered = orig.origNumbered;
				block.origPPrXml = orig.origPPrXml;
				block.origRuns = orig.origRuns;
				block.origMarkers = orig.origMarkers;
				block.protected = orig.protected === true;
			} else {
				block.origText = runsText(block.runs);
				block.origStyle = block.style;
				block.origNumbered = block.numbered;
				block.origMarkers = block.origMarkers === undefined ? [] : block.origMarkers;
				// Only default this: a paragraph parsed from the file already knows whether
				// it is protected, and this branch must not overwrite that verdict.
				if (block.protected === undefined) block.protected = false;
				// A paragraph typed in the editor has no original formatting at all;
				// the emitter builds its pPr from the current style instead of reusing one.
				block.origNew = true;
			}
			block.text = runsText(block.runs);
			return block;
		}

		function styleFromPPr(pPr) {
			if (pPr === undefined) return "Normal";
			const style = firstChild(pPr, "pStyle");
			if (style === undefined) return "Normal";
			const value = attr(style, "val");
			return value !== undefined ? String(value) : "Normal";
		}

		/**
		 * Paragraph-level children that carry no text: their position relative to the
		 * runs around them is all that matters. They can be re-emitted in place, so a
		 * bookmark or a comment range must NOT force the paragraph to be read-only.
		 */
		const MARKER_NAMES = new Set([
			"bookmarkStart", "bookmarkEnd", "commentRangeStart", "commentRangeEnd",
			"permStart", "permEnd", "proofErr",
			"moveFromRangeStart", "moveFromRangeEnd", "moveToRangeStart", "moveToRangeEnd",
		]);

		/**
		 * Run children a flat rebuild reproduces exactly: the properties and the text.
		 * Anything else in a run (a drawing, a field char, a tab, a hard line break) is
		 * content a run of plain text cannot stand in for.
		 */
		const RUN_SAFE_NAMES = new Set(["rPr", "t", "lastRenderedPageBreak"]);

		/** `w:sz` text ("28", "21") as a number of half-points, or undefined when absent/invalid. */
		function halfPointsOf(value) {
			if (value === undefined) return undefined;
			const parsed = Number(value);
			return Number.isInteger(parsed) && parsed > 0 ? parsed : undefined;
		}

		/** A 6-digit RRGGBB run colour, or undefined for "auto" and theme colours we must not pin down. */
		function hexColorOf(value) {
			if (value === undefined) return undefined;
			const match = /^#?([0-9a-fA-F]{6})$/.exec(String(value).trim());
			return match === null ? undefined : match[1].toUpperCase();
		}

		/**
		 * Split one paragraph into what the editor can edit and what it must preserve.
		 * @param node - the parsed `w:p`.
		 * @returns `{ runs, markers, protect }`, where markers are `{ at, xml }` to be
		 * re-emitted before the run at that index, and `protect` means the paragraph
		 * contains something a flat rebuild would destroy.
		 */
		function readParagraph(node) {
			const runs = [];
			const markers = [];
			let protect = false;
			const walk = (element) => {
				if (element.type !== "element") return;
				if (element.name === "r") {
					for (const child of element.children || []) {
						if (child.type === "element" && !RUN_SAFE_NAMES.has(child.name)) protect = true;
					}
					const rPr = firstChild(element, "rPr");
					const on = (tag) => {
						if (rPr === undefined) return false;
						const el = firstChild(rPr, tag);
						if (el === undefined) return false;
						const value = attr(el, "val");
						return value !== "0" && value !== "false" && value !== "none";
					};
					const val = (tag) => {
						if (rPr === undefined) return undefined;
						const el = firstChild(rPr, tag);
						if (el === undefined) return undefined;
						const value = attr(el, "val");
						return value === undefined ? undefined : String(value);
					};
					const text = textOf(element);
					if (text !== "")
						runs.push({
							text,
							fmt: {
								b: on("b"),
								i: on("i"),
								u: on("u"),
								// w:sz is in half-points; "auto" and theme colours are inherited,
								// so they stay undefined rather than becoming a fake explicit value.
								sz: halfPointsOf(val("sz")),
								color: hexColorOf(val("color")),
							},
							// Kept so an unchanged run can be re-emitted byte-exact and a
							// text-only edit can keep its exotic rPr (size, colour, fonts).
							xml: serializeNode(element),
							rPrXml: rPr !== undefined ? serializeNode(rPr) : "",
						});
					return;
				}
				if (MARKER_NAMES.has(element.name)) {
					markers.push({ at: runs.length, xml: serializeNode(element) });
					return;
				}
				if (element.name === "pPr") return;
				// A container we do not rebuild -- w:hyperlink, w:ins, w:del, w:sdt,
				// w:smartTag, w:fldSimple, w:oMath. Its text is still shown and editable-
				// looking, so keep walking to collect it, but the paragraph is read-only.
				protect = true;
				for (const child of element.children || []) walk(child);
			};
			for (const child of node.children || []) walk(child);
			return { runs, markers, protect };
		}

		function blocksFromDocumentXml(xml) {
			const root = parseXml(xml);
			if (root === undefined) throw new Error("document.xml 解析失败");
			const body = firstChild(root, "body");
			if (body === undefined) throw new Error("document.xml 缺少 w:body");
			const blocks = [];
			let sectPrXml = "";
			for (const child of body.children) {
				if (child.type !== "element") continue;
				if (child.name === "sectPr") {
					sectPrXml = serializeNode(child);
					continue;
				}
				const raw = serializeNode(child);
				if (child.name === "p") {
					const pPr = firstChild(child, "pPr");
					const style = styleFromPPr(pPr);
					const numPr = pPr !== undefined ? firstChild(pPr, "numPr") : undefined;
					const numbered = numPr !== undefined && firstChild(numPr, "numId") !== undefined;
					const read = readParagraph(child);
					const runs = read.runs;
					blocks.push(
						withBaseline(
							{
								kind: "p",
								id: blocks.length,
								style,
								numbered,
								runs,
								// True when the paragraph holds content a flat rebuild would
								// destroy: it renders read-only and is always re-emitted
								// byte-exact, so opening and saving can never damage it.
								protected: read.protect === true,
								markers: read.markers,
								origXml: raw,
								origPPrXml: pPr !== undefined ? serializeNode(pPr) : "",
								origMarkers: read.markers.map((marker) => ({ at: marker.at, xml: marker.xml })),
								origRuns: runs.map((run) => ({
									text: run.text,
									fmt: { b: run.fmt.b, i: run.fmt.i, u: run.fmt.u, sz: run.fmt.sz, color: run.fmt.color },
									xml: run.xml,
									rPrXml: run.rPrXml,
								})),
							},
							undefined
						)
					);
					continue;
				}
				blocks.push({ kind: "other", id: blocks.length, name: child.name, xml: raw, text: textOf(child), runs: [], style: "Normal", numbered: false, origXml: raw });
			}
			return { blocks, sectPrXml };
		}

		/**
		 * rPr children in the order OOXML's schema requires, so a toggled b/i/u lands
		 * where Word expects it instead of at the end of the element.
		 */
		const RPR_ORDER = [
			"rStyle", "rFonts", "b", "bCs", "i", "iCs", "caps", "smallCaps", "strike", "dstrike",
			"outline", "shadow", "emboss", "imprint", "noProof", "snapToGrid", "vanish", "webHidden",
			"color", "spacing", "w", "kern", "position", "sz", "szCs", "highlight", "u", "effect",
			"bdr", "shd", "fitText", "vertAlign", "rtl", "cs", "em", "lang", "eastAsianLayout",
			"specVanish", "oMath",
		];

		function rprRank(name) {
			const index = RPR_ORDER.indexOf(name);
			return index < 0 ? RPR_ORDER.length : index;
		}

		/** Turn one boolean run property on or off inside an existing rPr, keeping every other child. */
		function toggleRunProp(rPr, tag, on) {
			const existing = rPr.children.findIndex((child) => child.type === "element" && child.name === tag);
			if (!on) {
				if (existing >= 0) rPr.children.splice(existing, 1);
				return;
			}
			if (existing >= 0) rPr.children.splice(existing, 1);
			const node = { type: "element", name: tag, rawName: "w:" + tag, attrs: [], children: [] };
			const rank = rprRank(tag);
			let at = rPr.children.length;
			for (let i = 0; i < rPr.children.length; i++) {
				const child = rPr.children[i];
				if (child.type === "element" && rprRank(child.name) > rank) {
					at = i;
					break;
				}
			}
			rPr.children.splice(at, 0, node);
		}

		/**
		 * Set, replace or clear one run property that carries a `w:val` (w:sz, w:color).
		 * Clearing removes the element so the value falls back to the style, and an
		 * insertion lands in schema order rather than at the end of the rPr.
		 */
		function setRunValProp(rPr, tag, value) {
			const existing = rPr.children.findIndex((child) => child.type === "element" && child.name === tag);
			if (existing >= 0) {
				// Setting the value an element already carries must not move it: an rPr
				// written by Word keeps its own child order when nothing really changed.
				const current = rPr.children[existing].attrs.find((attribute) => attribute.name === "w:val");
				if (value !== undefined && current !== undefined && current.value === String(value)) return;
				rPr.children.splice(existing, 1);
			}
			if (value === undefined || value === null || value === "") return;
			const node = {
				type: "element",
				name: tag,
				rawName: "w:" + tag,
				attrs: [{ name: "w:val", value: String(value) }],
				children: [],
			};
			const rank = rprRank(tag);
			let at = rPr.children.length;
			for (let i = 0; i < rPr.children.length; i++) {
				const child = rPr.children[i];
				if (child.type === "element" && rprRank(child.name) > rank) {
					at = i;
					break;
				}
			}
			rPr.children.splice(at, 0, node);
		}

		/**
		 * Build a run's rPr. When the original rPr is available the boolean properties are
		 * toggled inside it, so size, colour and fonts survive an edit that only changed the
		 * text or one of bold/italic/underline.
		 */
		function runRPrXml(origRPrXml, fmt) {
			if (typeof origRPrXml !== "string" || origRPrXml === "") {
				const props = [];
				if (fmt.b) props.push("<w:b/>");
				if (fmt.i) props.push("<w:i/>");
				if (fmt.color !== undefined) props.push('<w:color w:val="' + escapeAttr(fmt.color) + '"/>');
				if (fmt.sz !== undefined) props.push('<w:sz w:val="' + String(fmt.sz) + '"/>');
				if (fmt.u) props.push('<w:u w:val="single"/>');
				return props.length > 0 ? "<w:rPr>" + props.join("") + "</w:rPr>" : "";
			}
			const rPr = parseXml(origRPrXml);
			if (rPr === undefined) return origRPrXml;
			toggleRunProp(rPr, "b", fmt.b);
			toggleRunProp(rPr, "i", fmt.i);
			toggleRunProp(rPr, "u", fmt.u);
			// A size or colour the collector did not state means "leave the document's own
			// value alone": only an explicit value or the toolbar's reset may change it.
			if (fmt.colorClear === true) setRunValProp(rPr, "color", undefined);
			else if (fmt.color !== undefined) setRunValProp(rPr, "color", fmt.color);
			// Word writes w:sz for Latin and w:szCs for complex script text. When a size is
			// actually changed they have to move together, or the two scripts disagree.
			const sized = rPr.children.some((child) => child.type === "element" && child.name === "szCs");
			if (fmt.szClear === true) {
				setRunValProp(rPr, "sz", undefined);
				if (sized) setRunValProp(rPr, "szCs", undefined);
			} else if (fmt.sz !== undefined) {
				setRunValProp(rPr, "sz", String(fmt.sz));
				if (sized) setRunValProp(rPr, "szCs", String(fmt.sz));
			}
			if (!rPr.children || rPr.children.length === 0) return "";
			return "<w:rPr>" + rPr.children.map(serializeNode).join("") + "</w:rPr>";
		}

		function runToXml(run, origRun) {
			const rPr = runRPrXml(origRun !== undefined ? origRun.rPrXml : undefined, run.fmt);
			const parts = [];
			String(run.text).split("\n").forEach((line, index) => {
				if (index > 0) parts.push("<w:br/>");
				line.split("\t").forEach((segment, segIndex) => {
					if (segIndex > 0) parts.push("<w:tab/>");
					if (segment !== "") parts.push('<w:t xml:space="preserve">' + escapeXml(segment) + "</w:t>");
				});
			});
			if (parts.length === 0) parts.push('<w:t xml:space="preserve"></w:t>');
			return "<w:r>" + rPr + parts.join("") + "</w:r>";
		}

		/**
		 * Whether a collected run left one property alone.
		 *
		 * `collectRuns` reports every field it can see, but a size or colour the surface could
		 * not represent (no span in the DOM) must not silently erase the value the document
		 * already had. So an unstated value counts as unchanged, and only the toolbar's
		 * explicit reset asks for removal.
		 */
		function runPropUnchanged(collected, original, clearFlag) {
			if (clearFlag === true) return false;
			if (collected === undefined) return true;
			return collected === original;
		}

		/**
		 * One run of a paragraph that changed: reuse the original run byte-exact when text
		 * and formatting are identical, otherwise rebuild it while carrying the original
		 * rPr's other properties over.
		 */
		function runXml(run, origRuns, index) {
			const orig = origRuns !== undefined && index < origRuns.length ? origRuns[index] : undefined;
			if (
				orig !== undefined &&
				orig.text === run.text &&
				orig.fmt.b === run.fmt.b &&
				orig.fmt.i === run.fmt.i &&
				orig.fmt.u === run.fmt.u &&
				runPropUnchanged(run.fmt.sz, orig.fmt.sz, run.fmt.szClear) &&
				runPropUnchanged(run.fmt.color, orig.fmt.color, run.fmt.colorClear)
			) {
				return orig.xml;
			}
			return runToXml(run, orig);
		}

		/** pPr children that must precede w:numPr, in schema order. */
		const BEFORE_NUM_PR = new Set(["pStyle", "keepNext", "keepLines", "pageBreakBefore", "framePr", "widowControl"]);

		/**
		 * pPr for a paragraph whose style or numbering changed: the original pPr is kept and
		 * only w:pStyle / w:numPr are added, replaced or removed, so indentation, alignment
		 * and spacing survive a style change.
		 */
		function pPrWithStyle(origPPrXml, style, numbered) {
			const pStyle =
				style !== "Normal" && style !== "ListParagraph"
					? { type: "element", name: "pStyle", rawName: "w:pStyle", attrs: [{ name: "w:val", value: style }], children: [] }
					: undefined;
			const numPr = numbered
				? {
						type: "element",
						name: "numPr",
						rawName: "w:numPr",
						attrs: [],
						children: [
							{ type: "element", name: "ilvl", rawName: "w:ilvl", attrs: [{ name: "w:val", value: "0" }], children: [] },
							{ type: "element", name: "numId", rawName: "w:numId", attrs: [{ name: "w:val", value: "1" }], children: [] },
						],
					}
				: undefined;
			const original = typeof origPPrXml === "string" && origPPrXml !== "" ? parseXml(origPPrXml) : undefined;
			const children =
				original !== undefined && original.children !== undefined
					? original.children.filter((child) => child.type !== "element" || (child.name !== "pStyle" && child.name !== "numPr"))
					: [];
			const merged = children.slice();
			if (pStyle !== undefined) merged.unshift(pStyle);
			if (numPr !== undefined) {
				let at = 0;
				while (at < merged.length && merged[at].type === "element" && BEFORE_NUM_PR.has(merged[at].name)) at += 1;
				merged.splice(at, 0, numPr);
			}
			if (merged.length === 0) return "";
			return "<w:pPr>" + merged.map(serializeNode).join("") + "</w:pPr>";
		}

		/**
		 * A paragraph that changed on screen. Its unchanged pPr is reused byte-exact; a style
		 * or list change patches only w:pStyle / w:numPr; each run keeps its original XML when
		 * it did not change. Only a paragraph the user actually edited is rebuilt at all.
		 */
		function paragraphXml(block) {
			let pPrXml;
			if (block.origNew === true) {
				pPrXml = pPrWithStyle("", block.style, block.numbered);
			} else if (block.origStyle === block.style && block.origNumbered === block.numbered) {
				pPrXml = typeof block.origPPrXml === "string" ? block.origPPrXml : "";
			} else {
				pPrXml = pPrWithStyle(block.origPPrXml, block.style, block.numbered);
			}
			const runs = block.runs || [];
			const markers = Array.isArray(block.markers) ? block.markers : block.origMarkers;
			const buckets = new Map();
			for (const marker of markers || []) {
				const at = Number.isInteger(marker.at) ? marker.at : 0;
				if (!buckets.has(at)) buckets.set(at, []);
				buckets.get(at).push(marker.xml);
			}
			let inner = "";
			for (let index = 0; index <= runs.length; index += 1) {
				const bucket = buckets.get(index);
				if (bucket !== undefined) inner += bucket.join("");
				if (index < runs.length) inner += runXml(runs[index], block.origRuns, index);
			}
			return "<w:p>" + pPrXml + inner + "</w:p>";
		}

		/** Emit document.xml: unchanged paragraphs keep their exact original bytes. */
		function documentXmlFromBlocks(originalXml, blocks, sectPrXml) {
			const root = parseXml(originalXml);
			if (root === undefined) throw new Error("document.xml 解析失败");
			// Everything before the root element -- the XML declaration and any whitespace
			// after it -- is kept byte-exact, so a no-op save is truly a no-op.
			const openAt = originalXml.indexOf("<" + root.rawName);
			const prefix = openAt > 0 ? originalXml.slice(0, openAt) : "";
			const head = "<" + root.rawName + root.attrs.map((a) => " " + a.name + '="' + escapeAttr(a.value) + '"').join("") + ">";
			const parts = [];
			for (const block of blocks) {
				if (block.kind === "other") {
					parts.push(block.xml);
					continue;
				}
				const current = runsText(block.runs);
				const unchanged =
					block.origXml !== undefined &&
					block.origText === current &&
					block.origStyle === block.style &&
					block.origNumbered === block.numbered &&
					// Text alone is not the whole change: bolding existing characters keeps
					// the text identical, so the formatting signature has to agree as well.
					signaturesEqual(formatSignature(block.runs), formatSignature(block.origRuns));
				parts.push(unchanged ? block.origXml : paragraphXml(block));
			}
			parts.push(sectPrXml);
			return prefix + head + "<w:body>" + parts.join("") + "</w:body></" + root.rawName + ">";
		}

		/**
		 * Fold the paragraphs the editor's DOM produced back into the block list parsed
		 * from the file.
		 *
		 * The DOM only ever contains editable paragraphs: non-paragraph body children
		 * (tables, block content controls, bookmark markers) never reach it, and
		 * protected paragraphs render read-only and are not collected. Emitting the DOM
		 * list alone would therefore DELETE those from the saved file -- which is why
		 * the skeleton parsed from the file, not the DOM, decides what the file contains.
		 *
		 * A protected paragraph keeps its original bytes no matter what the DOM claims,
		 * so it is impossible for the editor to rewrite one.
		 *
		 * @param skeleton - the block list parsed from the file, in document order.
		 * @param collected - paragraph blocks collected from the editor's DOM.
		 * @returns the block list to emit.
		 */
		function mergeCollected(skeleton, collected) {
			const byId = new Map();
			const insertedAfter = new Map();
			const leading = [];
			let anchor = null;
			for (const block of collected || []) {
				const id = block !== null && block !== undefined ? block.id : -1;
				if (Number.isInteger(id) && id >= 0 && !byId.has(id)) {
					byId.set(id, block);
					anchor = id;
					continue;
				}
				// No usable id, or two DOM paragraphs claiming one: a paragraph typed in
				// the editor. It goes right after the last paragraph we could identify.
				const at = anchor === null ? leading : insertedAfter.get(anchor) || (insertedAfter.set(anchor, []), insertedAfter.get(anchor));
				at.push(block);
			}
			const merged = leading.slice();
			for (const block of skeleton || []) {
				if (block.kind === "p") {
					const replacement = block.protected === true ? undefined : byId.get(block.id);
					merged.push(replacement !== undefined ? replacement : block);
				} else {
					merged.push(block);
				}
				const extra = insertedAfter.get(block.id);
				if (extra !== undefined) merged.push.apply(merged, extra);
			}
			return merged;
		}

		/**
		 * The save path. Everything the save button writes goes through here, so the
		 * data-loss tests exercise the real thing instead of a lower-level emitter.
		 * @param originalXml - the document.xml the file was loaded from.
		 * @param skeleton - the block list parsed from it.
		 * @param collected - paragraph blocks collected from the editor's DOM.
		 * @param sectPrXml - the body-level section properties.
		 * @returns the new document.xml.
		 */
		function saveDocumentXml(originalXml, skeleton, collected, sectPrXml) {
			return documentXmlFromBlocks(originalXml, mergeCollected(skeleton, collected), sectPrXml);
		}

		/**
		 * Editable HTML. A `<br>` is emitted only BETWEEN newline-separated
		 * segments of one run -- never between runs, which would split a
		 * paragraph into separate lines on the next save.
		 */
		/**
		 * The inline CSS that shows a run's own size and colour. `w:sz` is half-points, so it
		 * becomes pt; an absent property inherits from the paragraph style and emits nothing.
		 */
		function runStyleCss(fmt) {
			const parts = [];
			if (fmt.sz !== undefined) parts.push("font-size:" + String(fmt.sz / 2) + "pt");
			if (fmt.color !== undefined) parts.push("color:#" + fmt.color);
			return parts.join(";");
		}

		function blocksToHtml(blocks) {
			const out = [];
			for (const block of blocks) {
				if (block.kind === "other") {
					const rendered = renderNonParagraph(block);
					if (rendered !== "") out.push(rendered);
					continue;
				}
				const tag = block.style === "Heading1" ? "h1" : block.style === "Heading2" ? "h2" : block.style === "Heading3" ? "h3" : "p";
				const locked = block.protected === true;
				const attrs =
					' data-block-id="' + String(block.id) + '" data-docx-style="' + escapeHtml(block.style) + '"' + (block.numbered === true ? ' data-docx-list="1"' : "") +
					(locked ? ' data-docx-protected="1" contenteditable="false" title="' + escapeAttr(PROTECTED_HINT) + '" style="' + escapeAttr(styles.locked) + '"' : "");
				let inner = "";
				for (const run of block.runs || []) {
					const styleText = runStyleCss(run.fmt);
					const segments = String(run.text).split("\n");
					for (let index = 0; index < segments.length; index++) {
						if (index > 0) inner += "<br>";
						const segment = segments[index];
						if (segment === "") continue;
						let piece = escapeHtml(segment);
						// The document's own size/colour has to be visible, otherwise the
						// surface silently lies about what the paragraph looks like.
						if (styleText !== "") piece = '<span style="' + escapeAttr(styleText) + '">' + piece + "</span>";
						if (run.fmt.u) piece = "<u>" + piece + "</u>";
						if (run.fmt.b) piece = "<strong>" + piece + "</strong>";
						if (run.fmt.i) piece = "<em>" + piece + "</em>";
						inner += piece;
					}
				}
				if (inner === "") inner = "<br>";
				out.push("<" + tag + attrs + ">" + inner + "</" + tag + ">");
			}
			return out.join("");
		}

		/**
		 * One body child that is not a plain paragraph. Tables and block content controls
		 * are shown so the surface never hides part of the document, but read-only: the
		 * editor keeps them byte-exact and cannot edit inside them.
		 * @param block - a `kind: "other"` block.
		 * @returns HTML, or "" for a zero-width marker that has nothing to show.
		 */
		function renderNonParagraph(block) {
			if (MARKER_NAMES.has(block.name)) return "";
			if (block.name === "tbl") {
				const root = parseXml(block.xml);
				const rows = root === undefined ? [] : childrenNamed(root, "tr");
				if (rows.length === 0) return "";
				const body = rows
					.map((row) => {
						const cells = childrenNamed(row, "tc").map((cell) => {
							const text = childrenNamed(cell, "p").map((item) => textOf(item)).join(" ").trim();
							return '<td style="' + escapeAttr(styles.tableCell) + '">' + (text === "" ? "&nbsp;" : escapeHtml(text)) + "</td>";
						});
						return "<tr>" + cells.join("") + "</tr>";
					})
					.join("");
				return (
					'<table data-docx-table="1" contenteditable="false" title="' + escapeAttr(PROTECTED_HINT) + '" style="' +
					escapeAttr(styles.table) + '"><tbody>' + body + "</tbody></table>"
				);
			}
			const text = String(block.text || "").trim();
			if (text === "") return "";
			return (
				'<div data-docx-protected="1" contenteditable="false" title="' + escapeAttr(PROTECTED_HINT) + '" style="' +
				escapeAttr(styles.locked) + '">' + escapeHtml(text) + "</div>"
			);
		}

		// ─────────────────────────── editor → blocks ───────────────────────────

		/**
		 * Walk up from a selection endpoint to the block that owns it.
		 *
		 * Tells the toolbar whether a selection may be restyled at all: read-only content
		 * must not be touched, and an inline span cannot span two block elements.
		 * @param node - a DOM node a selection starts or ends in.
		 * @returns `{ editable, block }`.
		 */
		function blockOfNode(node) {
			let current = node;
			while (current !== null && current !== undefined) {
				if (current.nodeType === 1 && typeof current.getAttribute === "function") {
					if (current.getAttribute("data-docx-protected") === "1") return { editable: false, block: current };
					if (current.getAttribute("data-block-id") !== null) return { editable: true, block: current };
				}
				current = current.parentNode;
			}
			return { editable: false, block: null };
		}

		/** A CSS length in pt or px as half-points, which is the unit `w:sz` uses. */
		function cssFontSizeHalfPoints(value) {
			const match = /([\d.]+)\s*(pt|px)/i.exec(String(value));
			if (match === null) return undefined;
			const size = Number(match[1]);
			if (!Number.isFinite(size) || size <= 0) return undefined;
			const pt = match[2].toLowerCase() === "pt" ? size : (size * 72) / 96;
			const half = Math.round(pt * 2);
			return half > 0 ? half : undefined;
		}

		/** A CSS colour (#rrggbb or rgb()/rgba()) as RRGGBB, or undefined when it is not one. */
		function cssColorHex(value) {
			const text = String(value).trim();
			const hex = /^#([0-9a-fA-F]{6})$/.exec(text);
			if (hex !== null) return hex[1].toUpperCase();
			const rgb = /^rgba?\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)/i.exec(text);
			if (rgb === null) return undefined;
			const part = (value) => Math.max(0, Math.min(255, Number(value))).toString(16).padStart(2, "0").toUpperCase();
			return part(rgb[1]) + part(rgb[2]) + part(rgb[3]);
		}

		/** Whether two runs carry the same formatting, so adjacent DOM pieces can merge. */
		function sameRunFormat(left, right) {
			return (
				left.b === right.b && left.i === right.i && left.u === right.u && left.sz === right.sz && left.color === right.color
			);
		}

		function collectRuns(el) {
			const runs = [];
			const walk = (node, fmt) => {
				if (node.nodeType === 3) {
					if (node.nodeValue !== "") runs.push({ text: node.nodeValue, fmt: { b: fmt.b, i: fmt.i, u: fmt.u, sz: fmt.sz, color: fmt.color } });
					return;
				}
				if (node.nodeType !== 1) return;
				const tag = String(node.tagName).toLowerCase();
				const next = { b: fmt.b, i: fmt.i, u: fmt.u, sz: fmt.sz, color: fmt.color };
				if (tag === "b" || tag === "strong") next.b = true;
				if (tag === "i" || tag === "em") next.i = true;
				if (tag === "u") next.u = true;
				if (tag === "br") {
					runs.push({ text: "\n", fmt: next });
					return;
				}
				// The toolbar's "back to default" spans carry an explicit reset: a span with
				// no declaration cannot undo a size or colour inherited from an outer one.
				const reset = node.getAttribute("data-docx-reset");
				if (reset !== null && reset !== undefined) {
					for (const key of String(reset).split(",")) {
						if (key === "sz") next.sz = undefined;
						if (key === "color") next.color = undefined;
					}
				}
				const style = node.getAttribute("style");
				if (style !== null && style !== undefined && String(style) !== "") {
					const text = String(style);
					if (/font-weight\s*:\s*(bold|[6-9]00)/i.test(text)) next.b = true;
					if (/font-style\s*:\s*italic/i.test(text)) next.i = true;
					if (/text-decoration[^;]*underline/i.test(text)) next.u = true;
					const size = /font-size\s*:\s*([^;]+)/i.exec(text);
					if (size !== null) {
						const half = cssFontSizeHalfPoints(size[1]);
						if (half !== undefined) next.sz = half;
					}
					const color = /(?:^|;)\s*color\s*:\s*([^;]+)/i.exec(text);
					if (color !== null) {
						const hex = cssColorHex(color[1]);
						if (hex !== undefined) next.color = hex;
					}
				}
				const fontColor = node.getAttribute("color");
				if (tag === "font" && fontColor !== null && fontColor !== undefined) {
					const hex = cssColorHex(fontColor);
					if (hex !== undefined) next.color = hex;
				}
				for (let i = 0; i < node.childNodes.length; i++) walk(node.childNodes[i], next);
			};
			for (let i = 0; i < el.childNodes.length; i++) walk(el.childNodes[i], { b: false, i: false, u: false, sz: undefined, color: undefined });
			const merged = [];
			for (const run of runs) {
				const last = merged[merged.length - 1];
				if (last !== undefined && sameRunFormat(last.fmt, run.fmt)) last.text += run.text;
				else merged.push({ text: run.text, fmt: { b: run.fmt.b, i: run.fmt.i, u: run.fmt.u, sz: run.fmt.sz, color: run.fmt.color } });
			}
			return merged;
		}

		/**
		 * One paragraph from the editor. A lone `<br>` is the placeholder the surface renders
		 * for an empty paragraph, so it collects as NO runs: that keeps an untouched empty
		 * paragraph byte-exact on save (and preserves any non-text content such as a drawing).
		 */
		function blockFromElement(el, state) {
			const idAttr = el.getAttribute("data-block-id");
			const id = idAttr === null ? -1 : Number(idAttr);
			let orig = state.blocks.find((block) => block.id === id);
			if (orig === undefined) orig = state.blocks.find((block) => block.kind === "p" && block.text === el.textContent);
			let runs = collectRuns(el);
			if (runs.length === 1 && runs[0].text === "\n") runs = [];
			let style = el.getAttribute("data-docx-style");
			if (style === null || style === "") {
				const tag = String(el.tagName).toLowerCase();
				style = tag === "h1" ? "Heading1" : tag === "h2" ? "Heading2" : tag === "h3" ? "Heading3" : orig !== undefined ? orig.style : "Normal";
			}
			const numbered = (orig !== undefined && orig.numbered) || el.getAttribute("data-docx-list") === "1";
			return withBaseline({ kind: "p", id: orig !== undefined ? orig.id : -1, style, numbered, runs }, orig);
		}

		function collectEditorBlocks(rootEl, state) {
			const blocks = [];
			if (rootEl === null) return blocks;
			const nodes = rootEl.childNodes;
			for (let i = 0; i < nodes.length; i++) {
				const node = nodes[i];
				if (node.nodeType !== 1) continue;
				// Content the editor preserves verbatim is skipped here, so it is never
				// rebuilt from the DOM and never reaches the emitted document.
				if (typeof node.getAttribute === "function" && node.getAttribute("data-docx-protected") === "1") continue;
				const tag = String(node.tagName).toLowerCase();
				if (tag === "table" || tag === "tbody" || tag === "tr" || tag === "td" || tag === "th") continue;
				if (tag === "ul" || tag === "ol") {
					const items = node.querySelectorAll("li");
					for (let k = 0; k < items.length; k++) {
						const li = items[k];
						const runs = collectRuns(li);
						const orig = state.blocks.find((block) => block.id === Number(li.getAttribute("data-block-id")));
						blocks.push(withBaseline({ kind: "p", id: orig !== undefined ? orig.id : -1, style: "ListParagraph", numbered: true, runs }, orig));
					}
					continue;
				}
				if (tag === "p" || tag === "div" || tag === "h1" || tag === "h2" || tag === "h3") blocks.push(blockFromElement(node, state));
			}
			return blocks;
		}

		// ───────────────── one editable copy per document ──────────────────────

		const holderByPath = {};
		const listeners = new Set();
		let sequence = 0;
		const notify = () => {
			for (const listener of Array.from(listeners)) {
				try {
					listener();
				} catch {
					/* a broken listener must not stop the others */
				}
			}
		};

		function useHolder(path, myId) {
			const [, force] = React.useState(0);
			React.useEffect(() => {
				const listener = () => force((n) => n + 1);
				listeners.add(listener);
				return () => {
					listeners.delete(listener);
				};
			}, []);
			const claim = () => {
				if (path !== "" && holderByPath[path] !== myId) {
					holderByPath[path] = myId;
					notify();
				}
			};
			const release = () => {
				if (path !== "" && holderByPath[path] === myId) {
					delete holderByPath[path];
					notify();
				}
			};
			const holder = path === "" ? myId : holderByPath[path];
			return { isHolder: holder === undefined || holder === myId, claim, release, holder: holder === undefined ? myId : holder };
		}

		// ─────────────────────────────── styles ────────────────────────────────

		const styles = {
			root: { display: "flex", flexDirection: "column", height: "100%", minHeight: 0, gap: 6, padding: 8, boxSizing: "border-box" },
			toolbar: { display: "flex", flexWrap: "wrap", alignItems: "center", gap: 4 },
			button: { minWidth: 28, height: 26, padding: "0 8px", borderRadius: 6, border: "1px solid rgba(127,127,127,.35)", background: "transparent", color: "inherit", cursor: "pointer", fontSize: 12 },
			status: { fontSize: 11, opacity: 0.85, marginLeft: "auto", marginRight: 6 },
			save: { height: 26, padding: "0 12px", borderRadius: 6, border: "1px solid rgba(59,130,246,.5)", background: "rgba(59,130,246,.15)", color: "inherit", cursor: "pointer", fontSize: 12 },
			saveDisabled: { opacity: 0.45, cursor: "not-allowed" },
			hint: { fontSize: 11, opacity: 0.6, wordBreak: "break-all" },
			lock: { margin: "4px 0", padding: "8px 10px", borderRadius: 6, border: "1px solid rgba(245,158,11,.6)", background: "rgba(245,158,11,.12)", fontSize: 12, cursor: "pointer" },
			page: { flex: 1, minHeight: 0, overflow: "auto", padding: "24px 28px", background: "rgba(127,127,127,.06)", border: "1px solid rgba(127,127,127,.25)", borderRadius: 8, outline: "none", lineHeight: 1.75, fontSize: 14 },
			// Content the editor preserves verbatim and refuses to edit: shown, but
			// marked so the user is never surprised that it does not respond.
			locked: { borderLeft: "3px solid rgba(245,158,11,.7)", background: "rgba(245,158,11,.08)", padding: "2px 8px", margin: "4px 0", opacity: 0.92 },
			table: { borderCollapse: "collapse", width: "100%", margin: "8px 0", fontSize: 13 },
			tableCell: { border: "1px solid rgba(127,127,127,.4)", padding: "4px 6px", verticalAlign: "top" },
			select: { height: 26, borderRadius: 6, border: "1px solid rgba(127,127,127,.35)", background: "transparent", color: "inherit", fontSize: 12 },
			color: { width: 28, height: 26, padding: 0, border: "1px solid rgba(127,127,127,.35)", borderRadius: 6, background: "transparent", cursor: "pointer" }
		};

		// ─────────────────────────────── the editor ────────────────────────────

		function WordEditor(props) {
			const address = String(props.resourceAddress || "");
			const myIdRef = React.useRef("");
			if (myIdRef.current === "") {
				sequence += 1;
				myIdRef.current = "w" + String(sequence);
			}
			const myId = myIdRef.current;
			const [state, setState] = React.useState(null);
			const [status, setStatus] = React.useState("打开中…");
			const [error, setError] = React.useState(null);
			const [dirty, setDirty] = React.useState(false);
			const [html, setHtml] = React.useState("");
			const [probe, setProbe] = React.useState("");
			const rootRef = React.useRef(null);
			const stateRef = React.useRef(null);
			const loadedForRef = React.useRef("");
			const claimPath = state !== null ? String(state.path) : "";
			const claim = useHolder(claimPath, myId);

			React.useEffect(() => {
				if (address === "") return undefined;
				if (loadedForRef.current === address) return undefined;
				loadedForRef.current = address;
				let cancelled = false;
				setStatus("解析文档…");
				setError(null);
				setDirty(false);
				setHtml("");
				stateRef.current = null;
				setState(null);
				loadDocument(address)
					.then((loaded) => {
						if (cancelled) return;
						stateRef.current = loaded;
						setState(loaded);
						setHtml(blocksToHtml(loaded.blocks));
						setStatus("就绪 · " + String(loaded.blocks.length) + " 段");
					})
					.catch((cause) => {
						if (cancelled) return;
						setError(String(cause && cause.message ? cause.message : cause));
						setStatus("失败");
					});
				return () => {
					cancelled = true;
				};
			}, [address]);

			React.useEffect(() => () => claim.release(), []);

			const collect = () => {
				const current = stateRef.current;
				if (current === null) return { blocks: null, why: "no-state" };
				if (rootRef.current === null) return { blocks: null, why: "root-ref-null" };
				const blocks = collectEditorBlocks(rootRef.current, current);
				if (blocks.length === 0) return { blocks: null, why: "empty-dom" };
				// The collected list holds only editable paragraphs. `current.blocks` stays
				// the skeleton parsed from the file, so tables and other non-paragraph
				// children cannot be lost by saving.
				current.collected = blocks;
				const last = blocks[blocks.length - 1];
				setProbe(String(blocks.length) + " 段 / 末段尾『" + runsText(last.runs).slice(-14) + "』");
				return { blocks, why: "ok" };
			};

			React.useEffect(() => {
				if (!dirty) return undefined;
				const handle = setInterval(() => {
					if (claim.isHolder) collect();
				}, 1500);
				return () => clearInterval(handle);
			}, [dirty, claim.isHolder]);

			const exec = (command, value) => {
				try {
					document.execCommand(command, false, value === undefined ? null : value);
				} catch {
					/* an unsupported command leaves the selection alone */
				}
				setDirty(true);
			};

			/**
			 * Wrap the current selection in a span, which is how a size or colour reaches the
			 * document: the collector reads the span's inline style back into the run's rPr.
			 *
			 * Inline HTML cannot cross a block boundary, so a selection spanning two
			 * paragraphs (or touching read-only content) is refused instead of letting the
			 * browser restructure -- or merge -- the paragraphs to make the span fit.
			 */
			const wrapSelection = (attribute, value) => {
				try {
					const selection = document.getSelection();
					if (selection === null || selection.rangeCount === 0 || selection.isCollapsed) {
						setStatus("请先选中要改格式的文字");
						return;
					}
					const range = selection.getRangeAt(0);
					const start = blockOfNode(range.startContainer);
					const end = blockOfNode(range.endContainer);
					if (!start.editable || !end.editable) {
						setStatus("所选内容属于只读部分，不能改格式");
						return;
					}
					if (start.block !== end.block) {
						setStatus("请在同一段内选择要改格式的文字");
						return;
					}
					const text = selection.toString();
					if (text === "") {
						setStatus("请先选中要改格式的文字");
						return;
					}
					const open = attribute === "style" ? '<span style="' + escapeAttr(value) + '">' : '<span data-docx-reset="' + escapeAttr(value) + '">';
					document.execCommand("insertHTML", false, open + escapeHtml(text) + "</span>");
				} catch (cause) {
					setStatus("设置格式失败：" + String(cause && cause.message ? cause.message : cause));
				}
				setDirty(true);
			};

			const setFontSize = (pt) => {
				if (pt === "") {
					wrapSelection("data-docx-reset", "sz");
					return;
				}
				wrapSelection("style", "font-size:" + String(pt) + "pt");
			};

			const onSave = async () => {
				const current = stateRef.current;
				if (current === null) return;
				if (!claim.isHolder) {
					setStatus("已拒绝保存：此窗口不是当前编辑窗口");
					return;
				}
				setStatus("保存中…");
				try {
					const got = collect();
					if (got.blocks === null) {
						setStatus("已拒绝保存：无法读取编辑区（" + got.why + "）");
						return;
					}
					const collected = got.blocks;
					const newXml = saveDocumentXml(current.originalXml, current.blocks, collected, current.sectPrXml);
					const entries = current.entries.map((entry) =>
						entry.name === "word/document.xml"
							? { name: entry.name, data: utf8Encode(newXml), time: entry.time, date: entry.date }
							: { name: entry.name, data: entry.data, time: entry.time, date: entry.date }
					);
					const outBytes = await zip(entries);
					const result = await bridgeSave({ address, path: current.path, base64: toBase64(outBytes) });
					if (result !== null && result !== undefined && result.ok === true) {
						const reparsed = blocksFromDocumentXml(newXml);
						current.entries = entries;
						current.originalXml = newXml;
						current.blocks = reparsed.blocks;
						current.sectPrXml = reparsed.sectPrXml;
						setDirty(false);
						setStatus("已保存 · " + String(collected.length) + "段 · xml" + String(newXml.length) + " · " + String(result.size) + "B");
					} else {
						setStatus("保存失败：" + String(result && result.error));
					}
				} catch (cause) {
					setStatus("保存异常：" + String(cause && cause.message ? cause.message : cause));
				}
			};

			const button = (label, title, onClick) =>
				React.createElement("button", { type: "button", title, onClick, style: styles.button }, label);

			const toolbar = React.createElement(
				"div",
				{ style: styles.toolbar },
				button("B", "加粗", () => exec("bold")),
				button("I", "斜体", () => exec("italic")),
				button("U", "下划线", () => exec("underline")),
				React.createElement(
					"select",
					{
						key: "size",
						title: "字号（先选中文字）",
						defaultValue: "",
						style: styles.select,
						onChange: (event) => {
							setFontSize(event.target.value);
							event.target.value = "";
						},
					},
					React.createElement("option", { value: "" }, "字号"),
					...FONT_SIZES.map((pt) => React.createElement("option", { key: String(pt), value: String(pt) }, String(pt) + " pt"))
				),
				React.createElement("input", {
					key: "color",
					type: "color",
					title: "文字颜色（先选中文字）",
					defaultValue: "#ff0000",
					style: styles.color,
					onChange: (event) => wrapSelection("style", "color:" + String(event.target.value)),
				}),
				button("A", "清除所选文字的字号与颜色", () => wrapSelection("data-docx-reset", "sz,color")),
				button("H1", "一级标题", () => exec("formatBlock", "H1")),
				button("H2", "二级标题", () => exec("formatBlock", "H2")),
				button("H3", "三级标题", () => exec("formatBlock", "H3")),
				button("P", "正文", () => exec("formatBlock", "P")),
				button("• 列表", "无序列表", () => exec("insertUnorderedList")),
				button("1. 列表", "有序列表", () => exec("insertOrderedList")),
				button("◉ 检查", "按当前编辑区重新读取", () => collect()),
				React.createElement(
					"span",
					{ style: styles.status, key: "status" },
					error !== null ? error : "#" + myId + (claim.isHolder ? "(当前)" : "(非当前)") + " · " + status + (dirty ? " · 未保存" : "")
				),
				React.createElement(
					"button",
					{
						type: "button",
						key: "save",
						style: state === null || !claim.isHolder ? Object.assign({}, styles.save, styles.saveDisabled) : styles.save,
						onClick: onSave,
						disabled: state === null || !claim.isHolder
					},
					"保存"
				)
			);

			const overlay = claim.isHolder
				? null
				: React.createElement(
						"div",
						{ key: "lock", style: styles.lock, onClick: () => claim.claim() },
						"此文档已在编辑器 #" + String(claim.holder) + " 中打开。点这里接管编辑（接管后此处才可编辑）。"
					);

			const pageProps = {
				style: claim.isHolder ? styles.page : Object.assign({}, styles.page, { opacity: 0.75 }),
				ref: rootRef,
				contentEditable: claim.isHolder === true,
				suppressContentEditableWarning: true,
				spellCheck: false,
				onInput: () => {
					if (claim.isHolder) setDirty(true);
				}
			};
			if (html !== "") pageProps.dangerouslySetInnerHTML = { __html: html };

			return React.createElement(
				"div",
				{ style: styles.root, "data-word-editor": "1" },
				toolbar,
				React.createElement("div", { style: styles.hint, key: "probe" }, probe),
				React.createElement("div", { style: styles.hint, key: "path" }, state !== null ? String(state.path) : address),
				overlay,
				React.createElement("div", pageProps)
			);
		}

		/** Read one .docx through the host bridge and parse it into blocks. */
		async function loadDocument(address) {
			const rel = await bridgeRead(address);
			if (rel === null || rel === undefined || rel.ok !== true) throw new Error(String(rel && rel.error) || "读取失败");
			const bytes = fromBase64(String(rel.base64));
			const entries = await unzip(bytes);
			const docEntry = entries.find((entry) => entry.name === "word/document.xml");
			if (docEntry === undefined) throw new Error("不是 Word 文档：缺少 word/document.xml");
			const originalXml = utf8Decode(docEntry.data);
			const parsed = blocksFromDocumentXml(originalXml);
			return { path: String(rel.path), entries, originalXml, blocks: parsed.blocks, sectPrXml: parsed.sectPrXml };
		}

		// ───────────────────────────── registration ────────────────────────────

		const inject = ["slots", "documentPreviews"];

		function apply(ctx) {
			const registry = typeof ctx.get === "function" ? ctx.get("documentPreviews") : undefined;
			if (registry === undefined || registry === null) {
				console.error("[word-editor] documentPreviews service is unavailable; the .docx editor cannot claim documents");
				return;
			}
			ctx.effect(() =>
				registry.register({
					id: IMPL_ID,
					extensions: ["docx"],
					priority: "extension",
					title: () => "Word 编辑器",
					loading: "bytes-complete",
					wrap: false
				})
			);
			ctx.effect(() => ctx.slots.inject("sidebar.right.tab.document", () => ctx.slots.register({ name: "sidebar.right.tab.document", key: IMPL_ID }, WordEditor)));
		}

		exports.apply = apply;
		exports.inject = inject;
		// Test seam: the docx ⇄ blocks round trip is pure, so the checked-in test can
		// exercise it without a DOM. Nothing in the host runtime reads these.
		exports.__internals = {
			parseXml,
			runsText,
			withBaseline,
			blocksFromDocumentXml,
			documentXmlFromBlocks,
			mergeCollected,
			saveDocumentXml,
			collectEditorBlocks,
			collectRuns,
			formatSignature,
			runStyleCss,
			blockOfNode,
			blocksToHtml,
			paragraphXml,
		};
		return module.exports;
	}
});
