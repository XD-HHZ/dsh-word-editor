/**
 * Host half of dsh-word-editor.
 *
 * Provides the byte bridge the browser half needs, because the composed
 * filesystem service is text-shaped and cannot carry a .docx back to disk:
 *
 *   POST /word-editor/read  { address }            -> { ok, path, size, base64 }
 *   POST /word-editor/save  { address|path, base64 } -> { ok, path, size, verified }
 *   GET  /word-editor/ping                         -> { ok, version }
 *
 * Writes go through node:fs directly, so they are byte-exact; every save is
 * read back and compared before it is reported as successful.
 */
import { readFile, writeFile } from 'node:fs/promises'
import { isAbsolute, resolve as resolvePath } from 'node:path'

/** Host plugin name. */
export const name = 'word-editor'

/** The composed web server carries the browser bridge. */
export const inject = ['webServer']

const ROUTE_PREFIX = '/word-editor'
const VERSION = '1.0.0'
const MAX_BODY_BYTES = 256 * 1024 * 1024

/**
 * Decode one `/`-joined path whose segments may be percent-encoded.
 * @param value - raw path text.
 * @returns the decoded path.
 */
function decodeSegments(value) {
  return String(value)
    .split('/')
    .map((segment) => {
      try {
        return decodeURIComponent(segment)
      } catch {
        return segment
      }
    })
    .join('/')
}

/**
 * Split a `dsh-resource://file/...` address into its scope and path.
 * @param address - the resource address.
 * @returns its parts, or undefined when the address is not a file address.
 */
function parseAddress(address) {
  const match = /^dsh-resource:\/\/file\/(session|absolute)\/([\s\S]*)$/.exec(String(address ?? ''))
  if (match === null) return undefined
  const scope = match[1]
  const rest = match[2]
  if (scope === 'absolute') return { scope, path: decodeSegments(rest) }
  const slash = rest.indexOf('/')
  if (slash < 0) return { scope, sessionId: rest, path: '' }
  return { scope, sessionId: rest.slice(0, slash), path: decodeSegments(rest.slice(slash + 1)) }
}

/**
 * Resolve an address (or a plain path) to the absolute path this process writes.
 * @param ctx - the owning plugin context, used only to read a session workspace root.
 * @param input - `{ address }` or `{ path }` from the browser half.
 * @returns the absolute path, or an error string.
 */
function resolveTarget(ctx, input) {
  const explicit = input?.path
  if (typeof explicit === 'string' && explicit !== '' && isAbsolute(explicit)) return { path: explicit }
  const parsed = parseAddress(input?.address)
  if (parsed === undefined) return { error: 'unsupported address' }
  if (parsed.scope === 'absolute') {
    if (parsed.path === '') return { error: 'empty path' }
    return { path: parsed.path }
  }
  if (parsed.path === '') return { error: 'empty path' }
  if (isAbsolute(parsed.path)) return { path: parsed.path }
  const sessions = ctx.get('sessions')
  const session = sessions !== undefined && typeof sessions.get === 'function' ? sessions.get(parsed.sessionId) : undefined
  const cwd = session?.header?.cwd
  if (typeof cwd !== 'string' || cwd === '') return { error: 'unknown workspace for session ' + String(parsed.sessionId) }
  return { path: resolvePath(cwd, parsed.path) }
}

/**
 * Read a request body as JSON with a hard size cap.
 * @param req - the incoming request.
 * @returns the parsed body, or undefined when it is absent or malformed.
 */
async function readJsonBody(req) {
  const chunks = []
  let total = 0
  for await (const chunk of req) {
    total += chunk.length
    if (total > MAX_BODY_BYTES) throw new Error('request body too large')
    chunks.push(chunk)
  }
  if (total === 0) return undefined
  return JSON.parse(Buffer.concat(chunks).toString('utf8'))
}

/**
 * Send one JSON response.
 * @param res - the response to own.
 * @param status - HTTP status code.
 * @param value - JSON-serializable body.
 */
function sendJson(res, status, value) {
  const body = JSON.stringify(value)
  res.writeHead(status, {
    'content-type': 'application/json; charset=utf-8',
    'cache-control': 'no-store',
    'content-length': Buffer.byteLength(body),
  })
  res.end(body)
}

/**
 * Whether a browser-originated request comes from this same origin.
 *
 * A request with no Origin header is a local, non-browser caller (the harness
 * itself, curl); a browser always sends one for a POST, so a mismatched host is
 * rejected. This keeps a random page in the user's browser from writing files.
 * @param req - the incoming request.
 * @returns true when the request may proceed.
 */
function sameOrigin(req) {
  const origin = req.headers.origin
  if (typeof origin !== 'string' || origin === '') return true
  try {
    return new URL(origin).host === req.headers.host
  } catch {
    return false
  }
}

/**
 * Host plugin body: register the three bridge routes on the composed web server.
 * @param ctx - host root context carrying `webServer` (and optionally `sessions`).
 */
export function apply(ctx) {
  const register = (path, kind, handler) => {
    ctx.effect(() => ctx.webServer.register({ kind, path, handler }), 'word-editor: ' + path)
  }

  register(ROUTE_PREFIX + '/ping', 'exact', (req, res) => {
    sendJson(res, 200, { ok: true, version: VERSION })
  })

  register(ROUTE_PREFIX + '/read', 'exact', async (req, res) => {
    if (req.method !== 'POST') return sendJson(res, 405, { ok: false, error: 'POST required' })
    if (!sameOrigin(req)) return sendJson(res, 403, { ok: false, error: 'cross-origin request refused' })
    try {
      const body = await readJsonBody(req)
      const target = resolveTarget(ctx, body)
      if (target.error !== undefined) return sendJson(res, 400, { ok: false, error: target.error })
      let bytes
      try {
        bytes = await readFile(target.path)
      } catch (error) {
        return sendJson(res, 404, { ok: false, error: 'read failed: ' + String(error?.message ?? error), path: target.path })
      }
      return sendJson(res, 200, { ok: true, path: target.path, size: bytes.length, base64: bytes.toString('base64') })
    } catch (error) {
      return sendJson(res, 500, { ok: false, error: String(error?.message ?? error) })
    }
  })

  register(ROUTE_PREFIX + '/save', 'exact', async (req, res) => {
    if (req.method !== 'POST') return sendJson(res, 405, { ok: false, error: 'POST required' })
    if (!sameOrigin(req)) return sendJson(res, 403, { ok: false, error: 'cross-origin request refused' })
    try {
      const body = await readJsonBody(req)
      const target = resolveTarget(ctx, body)
      if (target.error !== undefined) return sendJson(res, 400, { ok: false, error: target.error })
      const bytes = Buffer.from(String(body?.base64 ?? ''), 'base64')
      if (bytes.length === 0) return sendJson(res, 400, { ok: false, error: 'empty payload' })
      await writeFile(target.path, bytes)
      const back = await readFile(target.path)
      const verified = back.length === bytes.length && back.equals(bytes)
      if (!verified) return sendJson(res, 500, { ok: false, error: 'written bytes differ from the payload', path: target.path })
      return sendJson(res, 200, { ok: true, path: target.path, size: bytes.length, verified: true })
    } catch (error) {
      return sendJson(res, 500, { ok: false, error: String(error?.message ?? error) })
    }
  })

  console.log('[word-editor] host half ready: ' + ROUTE_PREFIX + ' routes registered')
}
