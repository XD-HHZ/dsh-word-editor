/**
 * Host-half test: import lib/index.js exactly as the harness loader would, drive
 * `apply` with a fake context, and call the registered routes directly.
 *
 * The point is /word-editor/ping, which reports the package version read from
 * package.json at request time. This test fails if that read breaks or drifts.
 *
 * Usage: node test/host-ping.mjs
 */
import { readFileSync } from 'node:fs'
import { fileURLToPath } from 'node:url'
import { dirname, join } from 'node:path'

const here = dirname(fileURLToPath(import.meta.url))
const pkg = JSON.parse(readFileSync(join(here, '..', 'package.json'), 'utf8'))

let failures = 0
function check(label, condition, detail) {
  if (condition) {
    console.log('  PASS  ' + label)
  } else {
    failures += 1
    console.log('  FAIL  ' + label + (detail === undefined ? '' : '\n        ' + detail))
  }
}

const host = await import(new URL('../lib/index.js', import.meta.url).href)

// ── plugin shape ─────────────────────────────────────────────────────────────
check('exports name', host.name === 'word-editor', 'name: ' + String(host.name))
check(
  'declares webServer',
  Array.isArray(host.inject) && host.inject.includes('webServer'),
  'inject: ' + JSON.stringify(host.inject)
)
check('exports apply', typeof host.apply === 'function')

// ── drive apply with a fake context ──────────────────────────────────────────
const routes = []
const disposers = []
const ctx = {
  webServer: {
    register: (entry) => {
      routes.push(entry)
      return () => {}
    },
  },
  effect: (fn) => {
    const disposer = fn()
    disposers.push(disposer)
    return () => {}
  },
  get: () => undefined,
}

host.apply(ctx)

const paths = routes.map((route) => route.kind + ' ' + route.path).sort()
check(
  'registers exactly the three routes',
  paths.join(', ') === 'exact /word-editor/ping, exact /word-editor/read, exact /word-editor/save',
  'actually: ' + paths.join(', ')
)
check('every route is owned by an effect', disposers.length === routes.length, 'effects: ' + disposers.length)

const ping = routes.find((route) => route.path === '/word-editor/ping')

/** Call one handler with a minimal req/res pair and capture the response. */
async function call(handler, req) {
  let status
  let body = ''
  const res = {
    writeHead: (code) => { status = code },
    end: (chunk) => { body = chunk ?? '' },
  }
  await handler(req, res)
  return { status, body }
}

const response = await call(ping.handler, { method: 'GET', headers: {} })
check('ping answers 200', response.status === 200, 'status: ' + String(response.status))

let parsed
try {
  parsed = JSON.parse(response.body)
} catch (error) {
  parsed = undefined
  check('ping body is JSON', false, String(error))
}
if (parsed !== undefined) {
  check('ping body is JSON', true)
  check('ping reports ok', parsed.ok === true, 'body: ' + response.body)
  check(
    'ping version matches package.json',
    parsed.version === pkg.version,
    'ping: ' + String(parsed.version) + ', package.json: ' + pkg.version
  )
}

// The same handler must stay answerable for an unmatched method too (the harness
// probes with GET); the guard is only that it never throws.
const again = await call(ping.handler, { method: 'HEAD', headers: {} })
check('ping is idempotent and stateless', again.status === 200 && again.body === response.body)

console.log('')
if (failures === 0) {
  console.log('HOST PING OK')
} else {
  console.error(failures + ' check(s) failed')
  process.exit(1)
}
