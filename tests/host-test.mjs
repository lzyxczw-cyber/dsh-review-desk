/**
 * Host-half test: drives the registered route and tool through a fake Cordis
 * context whose `llm` and `web` behave like the real services.
 */
import fs from 'node:fs'
import path from 'node:path'
import assert from 'node:assert/strict'

const here = path.dirname(new URL(import.meta.url).pathname)
const mod = await import(path.join(here, '..', 'index.js'))

const calls = []
const callConfigs = []
const registered = { routes: new Map(), tool: null }

function makeStream({ system, messages }) {
  const user = messages[0].content[0].text
  const out = []
  if (system.includes('检索规划员')) {
    out.push('{"queries":["某银行 科技型企业贷款 2023 增速"],"claims":["科技型企业贷款余额 42.6 亿元"]}')
  } else if (system.includes('事实核查员')) {
    out.push('### 核查结论\n- **增速高于各项贷款**：有据可查 —— 依据 [S1]。\n- **42.6 亿元**：部分一致 —— 依据 [S1]。')
    out.push('\n### 需要留意的表述\n- “显著低于” → 建议补充具体差距。\n### 检索覆盖\n仅覆盖公开报道，缺监管口径。')
  } else if (system.includes('统稿')) {
    out.push('### 在全文中的位置\n承接上一段的实践描述。\n### 关联点\n- 呼应：与“技术流评价体系”形成支撑。\n- 冲突：未发现。\n- 缺口：缺收益闭环的数据。\n### 统稿建议\n1. 前移到第二段之后。')
  } else if (system.includes('撰稿人')) {
    out.push('<<<TEXT\n从实践效果看，采用“技术流”评价的客户不良率更低。\n>>>\n### 扩写说明\n- 补充了因果链条。\n- 需要作者补充真实不良率数据。')
  } else {
    out.push('### 结构判断\n本段是论证的落点。\n### 逻辑问题\n- **以偏概全** —— 摘录 → 样本不足 → 补统计口径。\n### 收紧建议\n1. 收敛结论。')
  }
  calls.push({ system, user })
  const chunks = [{ type: 'block-start', index: 0, blockType: 'text' }]
  for (const piece of out) chunks.push({ type: 'text-delta', index: 0, text: piece })
  chunks.push({ type: 'finish', reason: { kind: 'stop' } })
  return chunks
}

const ctx = {
  effect(fn, label) {
    fn()
    return () => {}
  },
  inject(deps, cb) {
    cb(ctx)
    return () => {}
  },
  get(name) {
    return name === 'web' ? web : undefined
  },
  agentDefaultModel: {
    currentSelection: () => ({ provider: 'deepseek', model: 'deepseek-v4-flash' }),
  },
  llm: {
    listProviders: () => [{ id: 'deepseek', name: 'DeepSeek' }],
    async listModels(provider) {
      return [{ provider, id: 'deepseek-v4-flash', name: 'DeepSeek V4 Flash' }]
    },
    async resolveModelInfo(provider, model) {
      return {
        provider,
        id: model,
        name: 'DeepSeek V4 Flash',
        reasoning: { efforts: [{ id: 'low', name: '低' }, { id: 'high', name: '高' }], defaultEffort: 'low' },
      }
    },
    async prepareCall(config) {
      callConfigs.push(config)
      return { config, stream: options => (async function* () { for (const chunk of makeStream(options)) yield chunk })() }
    },
  },
  connection: {
    fetch: {
      register(route) {
        registered.routes.set(route.path, route)
        return async () => {}
      },
    },
  },
  tools: {
    register(definition) {
      registered.tool = definition
      return () => {}
    },
  },
}

const searches = []
const web = {
  async search({ query, maxResults }) {
    searches.push({ query, maxResults })
    return {
      sources: [
        { url: 'https://example.com/a', title: '科技金融统计报告', snippet: '科技型中小企业贷款余额同比增长 21.9%', publishedAt: '2024-02-01' },
        { url: 'https://example.com/b', title: '某银行年报', summary: '科技型企业贷款余额 42.6 亿元' },
        { url: 'https://example.com/a', title: '重复来源', snippet: 'dup' },
      ],
      truncated: false,
    }
  },
}

mod.apply(ctx)
const analyzeRoute = registered.routes.get('/api/review-desk/analyze')
assert.ok(analyzeRoute, 'analyze route registered')
assert.deepEqual(analyzeRoute.methods, ['POST'])
assert.deepEqual(
  [...registered.routes.keys()].sort(),
  ['/api/review-desk/analyze', '/api/review-desk/engine', '/api/review-desk/original-pages'],
  'this plugin owns the analyze, engine-asset, and original-page routes',
)
assert.ok(registered.tool, 'tool registered')
assert.equal(registered.tool.name, 'review_manuscript')

/** POST one payload and collect its NDJSON frames. */
async function analyze(payload) {
  const request = new Request('http://127.0.0.1:9/api/review-desk/analyze', {
    method: 'POST',
    headers: { 'content-type': 'application/json' },
    body: JSON.stringify(payload),
  })
  const response = await analyzeRoute.fetch(request)
  const text = await response.text()
  const frames = text.split('\n').filter(Boolean).map(line => JSON.parse(line))
  return { response, frames, text }
}

// 1. source: planner → search → synthesis, with sources in a frame
{
  const { response, frames } = await analyze({
    action: 'source',
    text: '2023 年科技型企业贷款余额 42.6 亿元，不良率 0.31%。',
    manuscript: '全文……',
    title: '测试文稿',
  })
  assert.equal(response.status, 200)
  assert.match(response.headers.get('content-type'), /ndjson/)
  const types = frames.map(frame => frame.type)
  assert.deepEqual(types.slice(0, 2), ['meta', 'stage'])
  assert.ok(types.indexOf('sources') > 1 && types.indexOf('sources') < types.indexOf('delta'))
  const sources = frames.find(frame => frame.type === 'sources')
  assert.equal(sources.sources.length, 2, 'deduplicated by url')
  assert.equal(sources.claims.length, 1)
  assert.equal(sources.queries.length, 1)
  assert.equal(sources.sources[0].title, '科技金融统计报告')
  assert.equal(sources.sources[1].summary, '科技型企业贷款余额 42.6 亿元')
  assert.ok(types.includes('delta'))
  assert.equal(frames.at(-1).type, 'done')
  const system = calls.at(-1).system
  assert.match(system, /事实核查员/)
  assert.match(calls.at(-1).user, /\[S1\] 科技金融统计报告/)
  assert.match(calls.at(-1).user, /\[S2\] 某银行年报/)
  assert.doesNotMatch(calls.at(-1).user, /重复来源/, 'deduplicated evidence is not repeated')
  assert.equal(searches.length, 1)
  assert.equal(searches[0].query, '某银行 科技型企业贷款 2023 增速')
  console.log('ok  source:', frames.length, 'frames,', sources.sources.length, 'sources')
}

// 2. logic: single call, no search
{
  const before = searches.length
  const { frames } = await analyze({ action: 'logic', text: '因此可以证明全部客户都适用。' })
  assert.equal(searches.length, before, 'logic performs no search')
  const meta = frames.find(frame => frame.type === 'meta')
  assert.equal(meta.model, 'deepseek-v4-flash')
  assert.equal(meta.provider, 'deepseek')
  assert.match(calls.at(-1).system, /逻辑/)
  assert.equal(frames.at(-1).type, 'done')
  console.log('ok  logic:', frames.filter(f => f.type === 'delta').length, 'deltas')
}

// 3. expand: markers survive the frame stream
{
  const { frames } = await analyze({ action: 'expand', text: '技术流评价体系更有效。' })
  const raw = frames.filter(frame => frame.type === 'delta').map(frame => frame.text).join('')
  assert.match(raw, /<<<TEXT/)
  assert.match(raw, />>>/)
  assert.match(raw, /扩写说明/)
  console.log('ok  expand:', raw.length, 'chars')
}

// 4. relate: the manuscript reaches the prompt
{
  const { frames } = await analyze({ action: 'relate', text: '同时也要看到收益闭环。', manuscript: '第一段……\n\n第四段……' })
  assert.match(calls.at(-1).user, /【全文】/)
  assert.match(calls.at(-1).user, /第四段/)
  assert.equal(frames.at(-1).type, 'done')
  console.log('ok  relate')
}

// 5. validation: bad payloads are 4xx JSON, not streams
{
  for (const [payload, status] of [
    [{ action: 'nope', text: 'x' }, 400],
    [{ action: 'logic', text: '   ' }, 400],
    [{ action: 'logic' }, 400],
  ]) {
    const { response, frames, text } = await analyze(payload)
    assert.equal(response.status, status)
    const body = JSON.parse(text)
    assert.equal(typeof body.error, 'string')
    assert.equal(frames.length, 1, 'a JSON error body is not a frame stream')
  }
  const getRequest = new Request('http://127.0.0.1:9/api/review-desk/analyze', { method: 'GET' })
  const getResponse = await analyzeRoute.fetch(getRequest)
  assert.equal(getResponse.status, 405)
  console.log('ok  validation')
}

// 6. the tool returns the same text plus sources
{
  const exec = { signal: new AbortController().signal }
  const value = await registered.tool.execute({ action: 'source', text: '科技型企业贷款余额 42.6 亿元。' }, exec)
  assert.equal(value.action, 'source')
  assert.match(value.text, /核查结论/)
  assert.equal(value.sources.length, 2)
  const content = registered.tool.output.render({}, value)
  assert.deepEqual(content, [{ type: 'text', text: value.text }])
  console.log('ok  tool')
}

// 7. a failing model route becomes an `error` frame, not a rejection
{
  const failing = {
    ...ctx,
    llm: {
      async prepareCall() {
        throw new Error('模型不可用')
      },
    },
  }
  const local = { routes: new Map(), tool: null }
  const localCtx = {
    ...failing,
    effect(fn) { fn(); return () => {} },
    inject(deps, cb) { cb(localCtx); return () => {} },
    connection: { fetch: { register(route) { local.routes.set(route.path, route); return async () => {} } } },
    tools: { register(def) { local.tool = def; return () => {} } },
  }
  mod.apply(localCtx)
  const response = await local.routes.get('/api/review-desk/analyze').fetch(new Request('http://x/api/review-desk/analyze', {
    method: 'POST',
    headers: { 'content-type': 'application/json' },
    body: JSON.stringify({ action: 'logic', text: '一段文字' }),
  }))
  const frames = (await response.text()).split('\n').filter(Boolean).map(line => JSON.parse(line))
  const error = frames.find(frame => frame.type === 'error')
  assert.ok(error, 'error frame emitted')
  assert.match(error.message, /模型不可用/)
  console.log('ok  error frame')
}

// 8. malformed JSON body is 400
{
  const response = await analyzeRoute.fetch(new Request('http://x/api/review-desk/analyze', {
    method: 'POST',
    headers: { 'content-type': 'application/json' },
    body: '{not json',
  }))
  assert.equal(response.status, 400)
  console.log('ok  bad json')
}

// 9. a missing default model is reported to the user, not thrown at the route
{
  const local = { routes: new Map() }
  const localCtx = {
    ...ctx,
    agentDefaultModel: { currentSelection: () => ({ provider: '', model: '' }) },
    effect(fn) { fn(); return () => {} },
    inject(deps, cb) { cb(localCtx); return () => {} },
    connection: { fetch: { register(route) { local.routes.set(route.path, route); return async () => {} } } },
    tools: { register() { return () => {} } },
  }
  mod.apply(localCtx)
  const response = await local.routes.get('/api/review-desk/analyze').fetch(new Request('http://x/api/review-desk/analyze', {
    method: 'POST',
    headers: { 'content-type': 'application/json' },
    body: JSON.stringify({ action: 'logic', text: '一段文字' }),
  }))
  const frames = (await response.text()).split('\n').filter(Boolean).map(line => JSON.parse(line))
  assert.equal(frames.at(-1).type, 'error')
  assert.match(frames.at(-1).message, /默认模型/)
  console.log('ok  no model')
}


// 10. an explicit route and reasoning level travel with the request
{
  const { frames } = await analyze({
    action: 'logic',
    text: '一段用于路由测试的文字。',
    provider: 'deepseek',
    model: 'deepseek-v4-flash',
    reasoningEffort: 'high',
  })
  const meta = frames.find(frame => frame.type === 'meta')
  assert.equal(meta.provider, 'deepseek')
  assert.equal(meta.model, 'deepseek-v4-flash')
  assert.equal(meta.reasoningEffort, 'high')
  const config = callConfigs.at(-1)
  assert.equal(config.provider, 'deepseek')
  assert.equal(config.reasoningEffort, 'high')
  console.log('ok  explicit route + level')
}

// 11. an unknown route is refused before dispatch
{
  const { frames } = await analyze({ action: 'logic', text: '一段文字。', provider: 'gone', model: 'x' })
  assert.equal(frames.at(-1).type, 'error')
  assert.match(frames.at(-1).message, /已不可用/)
  console.log('ok  unknown route')
}

// 12. the model catalog lives on the shipped Remote, not on extra routes here
{
  assert.ok(registered.routes.has('/api/review-desk/analyze'))
  assert.ok(registered.routes.has('/api/review-desk/original-pages'))
  assert.ok(![...registered.routes.keys()].some(path => path.includes('model')), 'no model routes')
  console.log('ok  routes')
}

// 13. custom runs carry the requirement as the task
{
  const { frames } = await analyze({
    action: 'custom',
    text: '这段需要改成领导讲话口径。',
    instruction: '改写成领导讲话口径，控制在 120 字内',
    manuscript: '全文上下文。',
  })
  assert.equal(frames.at(-1).type, 'done')
  const last = calls.at(-1)
  assert.match(last.system, /任务要求/)
  assert.match(last.user, /【任务要求】\n改写成领导讲话口径，控制在 120 字内/)
  assert.match(last.user, /【全文（仅作上下文参考）】/)
  assert.equal(frames.find(frame => frame.type === 'meta').action, 'custom')

  const missing = await analyze({ action: 'custom', text: '一段文字。' })
  assert.equal(missing.response.status, 400)
  assert.match(JSON.parse(missing.text).error, /处理要求/)
  console.log('ok  custom action')
}


// 14. a rejected reasoning level degrades to the default with a warning
{
  const local = { routes: new Map() }
  const seenConfigs = []
  const ctxWithStrictModel = {
    ...ctx,
    llm: {
      listProviders: () => [{ id: 'deepseek', name: 'DeepSeek' }],
      async prepareCall(config) {
        seenConfigs.push(config)
        if (config.reasoningEffort !== undefined) throw new Error('unsupported reasoning effort')
        return { config, stream: options => (async function* () { for (const chunk of makeStream(options)) yield chunk })() }
      },
    },
    connection: { fetch: { register(route) { local.routes.set(route.path, route); return async () => {} } } },
    tools: { register() { return () => {} } },
  }
  mod.apply({ ...ctxWithStrictModel, effect(fn) { fn(); return () => {} }, inject(deps, cb) { cb(ctxWithStrictModel); return () => {} } })
  const response = await local.routes.get('/api/review-desk/analyze').fetch(new Request('http://x/api/review-desk/analyze', {
    method: 'POST',
    headers: { 'content-type': 'application/json' },
    body: JSON.stringify({ action: 'logic', text: '一段文字。', reasoningEffort: 'max' }),
  }))
  const frames = (await response.text()).split('\n').filter(Boolean).map(line => JSON.parse(line))
  const warning = frames.find(frame => frame.type === 'warning')
  assert.ok(warning, 'warning frame emitted')
  assert.equal(warning.code, 'reasoning-fallback')
  assert.equal(warning.effort, 'max')
  assert.equal(seenConfigs.length, 2, 'retried without the level')
  assert.equal(seenConfigs[1].reasoningEffort, undefined)
  assert.equal(frames.at(-1).type, 'done')
  console.log('ok  level fallback')
}


// 15. the original-page route renders through the Host office engine
{
  const local = { routes: new Map() }
  const rendered = []
  const officeKit = {
    async renderImages(request, signal) {
      rendered.push({ extension: request.extension, maxImages: request.maxImages, bytes: (await request.source.read(signal, 30_000_000)).bytes.length })
      return {
        pageCount: 3,
        dpi: 96,
        missingFonts: [],
        images: [
          { png: new Uint8Array([137, 80, 78, 71]), width: 100, height: 140, page: 1 },
          { png: new Uint8Array([137, 80, 78, 71]), width: 100, height: 140, page: 2 },
        ],
      }
    },
  }
  const localCtx = {
    ...ctx,
    get(name) { return name === 'officeKit' ? officeKit : ctx.get(name) },
    effect(fn) { fn(); return () => {} },
    inject(deps, cb) { cb(localCtx); return () => {} },
    connection: { fetch: { register(route) { local.routes.set(route.path, route); return async () => {} } } },
    tools: { register() { return () => {} } },
  }
  mod.apply(localCtx)
  const route = local.routes.get('/api/review-desk/original-pages')
  const response = await route.fetch(new Request('http://x/api/review-desk/original-pages?ext=docx', {
    method: 'POST',
    headers: { 'content-type': 'application/octet-stream' },
    body: new Uint8Array([1, 2, 3, 4, 5]),
  }))
  assert.equal(response.status, 200)
  const body = await response.json()
  assert.equal(body.pageCount, 3)
  assert.equal(body.pages.length, 2)
  assert.equal(body.pages[0].data.length > 0, true, 'base64 payload')
  assert.equal(body.truncated, true)
  assert.equal(rendered[0].extension, 'docx')
  assert.equal(rendered[0].bytes, 5)
  assert.equal(rendered[0].maxImages, 8)

  const badExt = await route.fetch(new Request('http://x/api/review-desk/original-pages?ext=txt', { method: 'POST', body: new Uint8Array([1]) }))
  assert.equal(badExt.status, 400)

  // A host without the engine answers 501 instead of failing the panel.
  const withoutEngine = { routes: new Map() }
  const bareCtx = {
    ...ctx,
    get(name) { return name === 'officeKit' ? undefined : ctx.get(name) },
    effect(fn) { fn(); return () => {} },
    inject(deps, cb) { cb(bareCtx); return () => {} },
    connection: { fetch: { register(route) { withoutEngine.routes.set(route.path, route); return async () => {} } } },
    tools: { register() { return () => {} } },
  }
  mod.apply(bareCtx)
  const bare = await withoutEngine.routes.get('/api/review-desk/original-pages').fetch(
    new Request('http://x/api/review-desk/original-pages?ext=pdf', { method: 'POST', body: new Uint8Array([1]) }),
  )
  assert.equal(bare.status, 501)
  console.log('ok  original pages')
}


// 16. a revision runs under the revision prompt and keeps the original for facts
{
  const { frames } = await analyze({
    action: 'custom',
    revise: true,
    text: '第一版改写。',
    original: '初始段落原文。',
    instruction: '再短一点',
    manuscript: '全文。',
  })
  assert.equal(frames.at(-1).type, 'done')
  const last = calls.at(-1)
  assert.match(last.system, /正在修订一段已经写好的文字/)
  assert.match(last.system, /<<<TEXT/)
  assert.match(last.user, /【任务要求】\n再短一点/)
  assert.match(last.user, /【当前文字】\n第一版改写。/)
  assert.match(last.user, /【原始段落（事实参照）】\n初始段落原文。/)
  assert.doesNotMatch(last.user, /【待审文字】/)

  const wrong = await analyze({ action: 'logic', revise: true, text: '一段文字。' })
  assert.equal(wrong.response.status, 400)
  assert.match(JSON.parse(wrong.text).error, /custom/)
  console.log('ok  revision')
}


{
  const engineRoute = registered.routes.get('/api/review-desk/engine')
  assert.ok(engineRoute, 'engine asset route registered')
  const ok = await engineRoute.fetch(new Request('http://x/api/review-desk/engine?name=superdoc.js'))
  assert.equal(ok.status, 200)
  assert.ok((ok.headers.get('content-type') ?? '').includes('javascript'), ok.headers.get('content-type'))
  assert.ok((await ok.text()).length > 1000, 'the bundle is served')
  const css = await engineRoute.fetch(new Request('http://x/api/review-desk/engine?name=superdoc.css'))
  assert.equal(css.status, 200)
  assert.ok((css.headers.get('content-type') ?? '').includes('text/css'))
  const missing = await engineRoute.fetch(new Request('http://x/api/review-desk/engine?name=..%2Findex.js'))
  assert.equal(missing.status, 404, 'only declared names are served')
  const wrongMethod = await engineRoute.fetch(new Request('http://x/api/review-desk/engine?name=superdoc.js', { method: 'POST' }))
  assert.equal(wrongMethod.status, 405)
}

console.log('\nall host checks passed')
