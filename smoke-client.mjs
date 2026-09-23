// Smoke test: load the client bundle exactly the way the browser module table does.
import { readFileSync } from 'node:fs'
import { fileURLToPath } from 'node:url'
import { dirname, join } from 'node:path'
import vm from 'node:vm'

const here = dirname(fileURLToPath(import.meta.url))
const clientPath = join(here, 'lib', 'client.js')
const source = readFileSync(clientPath, 'utf8')

const loaded = []
const fakeReact = {
  createElement: (...args) => ({ type: args[0], props: args[1], children: args.slice(2) }),
  useState: (initial) => [initial, () => {}],
  useEffect: () => {},
  useRef: (initial) => ({ current: initial })
}

const sandbox = {
  window: {
    __ModuleLoader__: {
      load: (definition) => {
        loaded.push(definition)
      }
    }
  },
  require: (specifier) => {
    if (specifier === 'react') return fakeReact
    throw new Error('unexpected require: ' + specifier)
  },
  fetch: () => Promise.resolve({ json: () => Promise.resolve({ ok: false, error: 'stub' }) }),
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

const context = vm.createContext(sandbox)
vm.runInContext(source, context, { filename: 'client.js' })

console.log('module definitions loaded:', loaded.length)
if (loaded.length !== 1) throw new Error('expected exactly one module definition')
const definition = loaded[0]
console.log('id:', definition.id)
if (definition.id !== 'dsh-word-editor') throw new Error('unexpected id')

const exported = definition.factory(sandbox.require)
console.log('exports keys:', Object.keys(exported).join(','))
if (typeof exported.apply !== 'function') throw new Error('apply is not a function')
if (!Array.isArray(exported.inject)) throw new Error('inject is not an array')
console.log('inject:', exported.inject.join(','))

// Exercise apply() against a stub context and count registrations.
const registrations = { previews: 0, bodies: 0, effects: 0 }
const stubCtx = {
  get: (name) => {
    if (name !== 'documentPreviews') return undefined
    return {
      register: (definition2) => {
        registrations.previews += 1
        console.log('  preview registered:', definition2.id, JSON.stringify(definition2.extensions), definition2.priority, definition2.loading)
        return () => {}
      }
    }
  },
  slots: {
    inject: (key, callback) => {
      registrations.effects += 1
      callback()
      return () => {}
    },
    register: (options) => {
      registrations.bodies += 1
      console.log('  body registered:', options.name, options.key)
      return () => {}
    }
  },
  effect: (callback) => {
    registrations.effects += 1
    const disposer = callback()
    return typeof disposer === 'function' ? disposer : () => {}
  }
}

exported.apply(stubCtx)
console.log('registrations:', JSON.stringify(registrations))
if (registrations.previews !== 1 || registrations.bodies !== 1) throw new Error('registration incomplete')
console.log('CLIENT BUNDLE OK')
