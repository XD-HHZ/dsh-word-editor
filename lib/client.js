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

		function runsOf(paragraph) {
			const runs = [];
			const walk = (node) => {
				if (node.type !== "element") return;
				if (node.name === "r") {
					const rPr = firstChild(node, "rPr");
					const on = (tag) => {
						if (rPr === undefined) return false;
						const el = firstChild(rPr, tag);
						if (el === undefined) return false;
						const value = attr(el, "val");
						return value !== "0" && value !== "false" && value !== "none";
					};
					const text = textOf(node);
					if (text !== "") runs.push({ text, fmt: { b: on("b"), i: on("i"), u: on("u") } });
					return;
				}
				for (const child of node.children || []) walk(child);
			};
			for (const child of paragraph.children || []) walk(child);
			return runs;
		}

		function runsText(runs) {
			return (runs || []).map((run) => run.text).join("");
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
			} else {
				block.origText = runsText(block.runs);
				block.origStyle = block.style;
				block.origNumbered = block.numbered;
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
					blocks.push(withBaseline({ kind: "p", id: blocks.length, style, numbered, runs: runsOf(child), origXml: raw }, undefined));
					continue;
				}
				blocks.push({ kind: "other", id: blocks.length, name: child.name, xml: raw, text: textOf(child), runs: [], style: "Normal", numbered: false, origXml: raw });
			}
			return { blocks, sectPrXml };
		}

		function runToXml(run) {
			const props = [];
			if (run.fmt.b) props.push("<w:b/>");
			if (run.fmt.i) props.push("<w:i/>");
			if (run.fmt.u) props.push('<w:u w:val="single"/>');
			const rPr = props.length > 0 ? "<w:rPr>" + props.join("") + "</w:rPr>" : "";
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

		function paragraphXml(block) {
			const bits = [];
			if (block.style !== "Normal" && block.style !== "ListParagraph") bits.push('<w:pStyle w:val="' + escapeAttr(block.style) + '"/>');
			if (block.numbered) bits.push('<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>');
			const pPr = bits.length > 0 ? "<w:pPr>" + bits.join("") + "</w:pPr>" : "";
			return "<w:p>" + pPr + (block.runs || []).map(runToXml).join("") + "</w:p>";
		}

		/** Emit document.xml: unchanged paragraphs keep their exact original bytes. */
		function documentXmlFromBlocks(originalXml, blocks, sectPrXml) {
			const root = parseXml(originalXml);
			if (root === undefined) throw new Error("document.xml 解析失败");
			let declaration = "";
			if (originalXml.startsWith("<?xml")) {
				const end = originalXml.indexOf("?>");
				if (end >= 0) declaration = originalXml.slice(0, end + 2);
			}
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
					block.origNumbered === block.numbered;
				parts.push(unchanged ? block.origXml : paragraphXml(block));
			}
			parts.push(sectPrXml);
			return declaration + head + "<w:body>" + parts.join("") + "</w:body></" + root.rawName + ">";
		}

		/**
		 * Editable HTML. A `<br>` is emitted only BETWEEN newline-separated
		 * segments of one run -- never between runs, which would split a
		 * paragraph into separate lines on the next save.
		 */
		function blocksToHtml(blocks) {
			const out = [];
			for (const block of blocks) {
				if (block.kind === "other") continue;
				const tag = block.style === "Heading1" ? "h1" : block.style === "Heading2" ? "h2" : block.style === "Heading3" ? "h3" : "p";
				const attrs =
					' data-block-id="' + String(block.id) + '" data-docx-style="' + escapeHtml(block.style) + '"' + (block.numbered === true ? ' data-docx-list="1"' : "");
				let inner = "";
				for (const run of block.runs || []) {
					const segments = String(run.text).split("\n");
					for (let index = 0; index < segments.length; index++) {
						if (index > 0) inner += "<br>";
						const segment = segments[index];
						if (segment === "") continue;
						let piece = escapeHtml(segment);
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

		// ─────────────────────────── editor → blocks ───────────────────────────

		function collectRuns(el) {
			const runs = [];
			const walk = (node, fmt) => {
				if (node.nodeType === 3) {
					if (node.nodeValue !== "") runs.push({ text: node.nodeValue, fmt: { b: fmt.b, i: fmt.i, u: fmt.u } });
					return;
				}
				if (node.nodeType !== 1) return;
				const tag = String(node.tagName).toLowerCase();
				const next = { b: fmt.b, i: fmt.i, u: fmt.u };
				if (tag === "b" || tag === "strong") next.b = true;
				if (tag === "i" || tag === "em") next.i = true;
				if (tag === "u") next.u = true;
				if (tag === "br") {
					runs.push({ text: "\n", fmt: next });
					return;
				}
				const style = node.getAttribute("style");
				if (style !== null && style !== undefined) {
					if (/font-weight\s*:\s*(bold|[6-9]00)/i.test(style)) next.b = true;
					if (/font-style\s*:\s*italic/i.test(style)) next.i = true;
					if (/text-decoration[^;]*underline/i.test(style)) next.u = true;
				}
				for (let i = 0; i < node.childNodes.length; i++) walk(node.childNodes[i], next);
			};
			for (let i = 0; i < el.childNodes.length; i++) walk(el.childNodes[i], { b: false, i: false, u: false });
			const merged = [];
			for (const run of runs) {
				const last = merged[merged.length - 1];
				if (last !== undefined && last.fmt.b === run.fmt.b && last.fmt.i === run.fmt.i && last.fmt.u === run.fmt.u) last.text += run.text;
				else merged.push({ text: run.text, fmt: { b: run.fmt.b, i: run.fmt.i, u: run.fmt.u } });
			}
			return merged;
		}

		function blockFromElement(el, state) {
			const idAttr = el.getAttribute("data-block-id");
			const id = idAttr === null ? -1 : Number(idAttr);
			let orig = state.blocks.find((block) => block.id === id);
			if (orig === undefined) orig = state.blocks.find((block) => block.kind === "p" && block.text === el.textContent);
			const runs = collectRuns(el);
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
				const tag = String(node.tagName).toLowerCase();
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
			page: { flex: 1, minHeight: 0, overflow: "auto", padding: "24px 28px", background: "rgba(127,127,127,.06)", border: "1px solid rgba(127,127,127,.25)", borderRadius: 8, outline: "none", lineHeight: 1.75, fontSize: 14 }
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
				current.blocks = blocks;
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
					const newXml = documentXmlFromBlocks(current.originalXml, collected, current.sectPrXml);
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
		return module.exports;
	}
});
