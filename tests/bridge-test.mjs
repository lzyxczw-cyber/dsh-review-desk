/**
 * WPS bridge test: starts the loopback bridge with a fake Cordis context and
 * drives GET /health, the CORS preflight and POST /analyze (NDJSON frames).
 */
import assert from 'node:assert/strict'
import path from 'node:path'

process.env.REVIEW_DESK_WPS_BRIDGE = '1'
process.env.REVIEW_DESK_WPS_BRIDGE_PORT = '47839'
process.env.REVIEW_DESK_WPS_TOKEN = 'test-token'

const here = path.dirname(new URL(import.meta.url).pathname)
const mod = await import(path.join(here, '..', 'index.js'))

const BASE = 'http://127.0.0.1:47839'
const AUTH = { 'x-review-desk-token': 'test-token' }

const ctx = {
  effect(fn) {
    fn()
    return () => {}
  },
  inject(deps, callback) {
    callback(ctx)
    return () => {}
  },
  get() {
    return undefined
  },
  agentDefaultModel: {
    currentSelection: () => ({ provider: 'deepseek', model: 'deepseek-v4-flash', reasoningEffort: 'max' }),
  },
  llm: {
    listProviders: () => [{ id: 'deepseek', name: 'DeepSeek' }],
    async listModels(provider) {
      return [{
        provider,
        id: 'deepseek-v4-flash',
        name: 'DeepSeek V4 Flash',
        reasoning: { efforts: [{ id: 'low', name: '低' }, { id: 'high', name: '高' }], defaultEffort: 'low' },
      }]
    },
    async resolveModelInfo(provider, model) {
      return { provider, id: model, name: 'DeepSeek V4 Flash' }
    },
    async prepareCall(config) {
      llmCalls.push(config)
      return {
        config,
        stream: () => (async function* () {
          yield { type: 'text-delta', text: '### 结构判断\n本段是论证落点。' }
          yield { type: 'text-delta', text: '\n### 逻辑问题\n- **以偏概全** —— 样本不足。' }
          yield { type: 'finish', reason: { kind: 'stop' } }
        })(),
      }
    },
  },
  connection: { fetch: { register() { return () => {} } } },
  tools: { register() { return () => {} } },
}

const llmCalls = []

mod.apply(ctx)

/** Wait for the bridge listener; the plugin starts it asynchronously. */
async function health() {
  const started = Date.now()
  for (;;) {
    try {
      const response = await fetch(`${BASE}/health`)
      if (response.ok) return response.json()
    } catch {
      /* not listening yet */
    }
    if (Date.now() - started > 8000) throw new Error('bridge did not start')
    await new Promise(resolve => setTimeout(resolve, 100))
  }
}

const info = await health()
assert.equal(info.ok, true)
assert.equal(info.service, 'review-desk-wps-bridge')
assert.equal(info.auth, 'token')
console.log('ok  health')

// A foreign browser origin cannot drive the bridge, even on loopback.
const forbidden = await fetch(`${BASE}/health`, { headers: { origin: 'http://evil.example' } })
assert.equal(forbidden.status, 403)
console.log('ok  origin rejected')

const allowedOrigin = await fetch(`${BASE}/health`, { headers: { origin: 'http://127.0.0.1:3889' } })
assert.equal(allowedOrigin.status, 200)
assert.equal(allowedOrigin.headers.get('access-control-allow-origin'), 'http://127.0.0.1:3889')
console.log('ok  loopback origin reflected')

const preflight = await fetch(`${BASE}/analyze`, {
  method: 'OPTIONS',
  headers: { origin: 'http://127.0.0.1:3889', 'access-control-request-headers': 'content-type,x-review-desk-token' },
})
assert.equal(preflight.status, 204)
assert.match(preflight.headers.get('access-control-allow-headers') ?? '', /x-review-desk-token/u)
console.log('ok  cors preflight')

const withoutToken = await fetch(`${BASE}/models`)
assert.equal(withoutToken.status, 401)
console.log('ok  token required')

const catalogResponse = await fetch(`${BASE}/models`, { headers: AUTH })
assert.equal(catalogResponse.status, 200)
const catalog = await catalogResponse.json()
assert.equal(catalog.default.provider, 'deepseek')
assert.equal(catalog.default.model, 'deepseek-v4-flash')
assert.equal(catalog.default.reasoningEffort, 'max')
assert.equal(catalog.providers[0].models[0].reasoning.efforts.length, 2)
assert.equal(catalog.caps.default.expand, 8_000)
assert.equal(catalog.caps.max, 16_000)
console.log('ok  models catalog')

const response = await fetch(`${BASE}/analyze`, {
  method: 'POST',
  headers: { 'content-type': 'application/json', ...AUTH },
  body: JSON.stringify({ action: 'logic', text: '这是一段需要审核的文字。' }),
})
assert.equal(response.status, 200)
assert.match(response.headers.get('content-type') ?? '', /ndjson/u)
const frames = (await response.text())
  .trim()
  .split('\n')
  .filter(line => line.trim() !== '')
  .map(line => JSON.parse(line))
assert.equal(frames[0].type, 'meta')
assert.equal(frames[0].action, 'logic')
assert.equal(frames[0].provider, 'deepseek')
assert.ok(frames.some(frame => frame.type === 'delta' && frame.text.includes('结构判断')), 'streamed text')
assert.equal(frames.at(-1).type, 'done')
console.log('ok  analyze stream')

const capped = await fetch(`${BASE}/analyze`, {
  method: 'POST',
  headers: { 'content-type': 'application/json', ...AUTH },
  body: JSON.stringify({ action: 'expand', text: '一段需要扩写的文字。', maxTokens: 20_000 }),
})
assert.equal(capped.status, 200)
assert.equal(llmCalls.at(-1).maxTokens, 16_000, 'output cap is clamped')
console.log('ok  output cap override')

const invalid = await fetch(`${BASE}/analyze`, {
  method: 'POST',
  headers: { 'content-type': 'application/json', ...AUTH },
  body: JSON.stringify({ action: 'unknown', text: 'x' }),
})
assert.equal(invalid.status, 400)
const problem = await invalid.json()
assert.ok(typeof problem.error === 'string' && problem.error !== '')
console.log('ok  invalid payload')

assert.equal((await fetch(`${BASE}/nope`, { headers: AUTH })).status, 404)
console.log('ok  unknown route')

console.log('\nall bridge checks passed')
process.exit(0)
