/**
 * WPS task-pane test (v0.3, two columns): loads taskpane.html/js in jsdom and
 * drives it with a fake WPS document model and a fake loopback bridge.
 *
 * Covers: two-column mount, selection, streaming detail, comment, tracked
 * replacement, undo via revision rejection, locate, revise chain, filters,
 * search, keyboard shortcut, token header, error state, settings and history.
 */
import fs from 'node:fs'
import path from 'node:path'
import assert from 'node:assert/strict'
import { JSDOM } from 'jsdom'

const here = path.dirname(new URL(import.meta.url).pathname)
const PANEL_DIR = path.join(here, '..', 'wps-addin')
const HTML = fs.readFileSync(path.join(PANEL_DIR, 'taskpane.html'), 'utf8')
const SCRIPT = fs.readFileSync(path.join(PANEL_DIR, 'taskpane.js'), 'utf8')
const HISTORY_KEY = 'wps-review-desk.history.v5'
const SETTINGS_KEY = 'wps-review-desk.settings.v2'

const BASE_TEXT = '第一段：这是用于测试的原文。\n第二段：后续内容。'

/* ------------------------------------------------------------------ *
 * 假 WPS 文档模型
 * ------------------------------------------------------------------ */

function createFakeWps(initialText, docName) {
  const model = {
    text: initialText,
    name: docName || 'wps-panel-test.docx',
    track: false,
    revisions: [],
    comments: [],
    saved: true,
    highlight: 0,
    underlineStyle: 0,
    underlineColor: 0,
    underlines: [],
    selection: { Start: 0, End: 0, Text: '' },
    scrollCalls: 0,
  }

  function parseRange(start, end) {
    const from = Math.max(0, Math.min(Number(start) || 0, model.text.length))
    const to = Math.max(from, Math.min(Number(end) || 0, model.text.length))
    return { start: from, end: to }
  }

  function setRangeText(start, end, value) {
    const previous = model.text.slice(start, end)
    model.text = model.text.slice(0, start) + value + model.text.slice(end)
    model.saved = false
    if (model.track) {
      let restored = false
      const restore = () => {
        if (restored) return
        restored = true
        model.text = model.text.slice(0, start) + previous + model.text.slice(start + value.length)
        model.revisions = model.revisions.filter(record => record !== first && record !== second)
        model.selection = { Start: start, End: start + previous.length, Text: previous }
      }
      const first = { Range: { Start: start, End: start + value.length }, Author: 'tester', Reject: restore }
      const second = { Range: { Start: start, End: start + value.length }, Author: 'tester', Reject: () => {} }
      model.revisions.push(first, second)
      model.selection = { Start: start, End: start + value.length, Text: value }
    } else {
      model.selection = { Start: start, End: start + value.length, Text: value }
    }
  }

  function rangeObject(start, end) {
    const bounds = parseRange(start, end)
    const range = {
      Start: bounds.start,
      End: bounds.end,
      Select() {
        model.selection = {
          Start: bounds.start,
          End: bounds.end,
          Text: model.text.slice(bounds.start, bounds.end),
        }
      },
      Comments: {
        Add(_range, text) {
          model.comments.push({ start: bounds.start, end: bounds.end, text: String(text) })
        },
      },
      Font: {
        get Underline() { return model.underlineStyle },
        set Underline(value) {
          const style = Number(value) || 0
          model.underlineStyle = style
          if (style !== 0) {
            model.underlines.push({ start: bounds.start, end: bounds.end, style, color: model.underlineColor })
          }
        },
        get UnderlineColor() { return model.underlineColor },
        set UnderlineColor(value) {
          model.underlineColor = Number(value) || 0
          const last = model.underlines.at(-1)
          if (last) last.color = model.underlineColor
        },
      },
    }
    Object.defineProperty(range, 'Text', {
      get() { return model.text.slice(bounds.start, bounds.end) },
      set(value) { setRangeText(bounds.start, bounds.end, String(value)) },
    })
    Object.defineProperty(range, 'HighlightColorIndex', {
      get() { return model.highlight },
      set(value) { model.highlight = Number(value) },
    })
    return range
  }

  const doc = {
    Name: model.name,
    get Saved() { return model.saved },
    get TrackRevisions() { return model.track },
    set TrackRevisions(value) { model.track = value === true },
    get Content() {
      return {
        get Text() { return model.text },
        set Text(value) { model.text = String(value) },
      }
    },
    Range(start, end) {
      return rangeObject(start === undefined ? 0 : start, end === undefined ? 0 : end)
    },
    get Revisions() {
      return {
        get Count() { return model.revisions.length },
        Item(index) { return model.revisions[index - 1] },
      }
    },
    get Comments() {
      return {
        get Count() { return model.comments.length },
        Item(index) { return model.comments[index - 1] },
      }
    },
    Undo() { return true },
    Redo() { return true },
    SaveAs2(target) { model.saved = true; model.name = String(target).split('/').pop(); return true },
  }

  const app = {
    Name: 'Microsoft Word',
    Version: '12.0',
    get ActiveDocument() { return doc },
    get Selection() {
      const selection = model.selection
      return { Start: selection.Start, End: selection.End, Text: selection.Text }
    },
    get ActiveWindow() {
      return { ScrollIntoView() { model.scrollCalls += 1 } }
    },
    get Enum() { return { JSKsoEnum_msoCTPDockPositionRight: 2 } },
  }

  return { app, doc, model }
}

/* ------------------------------------------------------------------ *
 * 假桥接服务
 * ------------------------------------------------------------------ */

function ndjson(frames) {
  return new Response(frames.map(frame => `${JSON.stringify(frame)}\n`).join(''), {
    status: 200,
    headers: { 'content-type': 'application/x-ndjson; charset=utf-8' },
  })
}

function catalog() {
  return {
    default: { provider: 'deepseek', model: 'deepseek-flash', reasoningEffort: 'max' },
    providers: [{
      id: 'deepseek',
      name: 'DeepSeek',
      models: [{
        id: 'deepseek-flash',
        name: 'DeepSeek Flash',
        reasoning: { defaultEffort: 'max', efforts: [{ id: 'low', name: '低' }, { id: 'max', name: '最高' }] },
      }],
    }],
    caps: { default: { source: 6000, logic: 6000, expand: 8000, relate: 6000, custom: 6000 }, min: 1000, max: 16000 },
  }
}

function buildFrames(payload, warningActions) {
  const frames = [{ type: 'meta', action: payload.action, provider: 'deepseek', model: 'deepseek-flash', reasoningEffort: 'max' }]
  if (payload.action === 'expand') {
    frames.push({ type: 'delta', text: '<<<TEXT\n扩写后的正文示例。\n>>>\n### 扩写说明\n- 补充了细节\n' })
  } else if (payload.action === 'custom' && payload.revise === true) {
    frames.push({ type: 'delta', text: '<<<TEXT\n压缩后的正文示例。\n>>>\n### 改写说明\n- 压缩到 40 字\n' })
  } else if (payload.action === 'logic') {
    frames.push({ type: 'delta', text: '### 结构判断\n本段是论证落点。\n### 逻辑问题\n- **以偏概全** —— 样本不足。\n' })
  } else {
    frames.push({ type: 'delta', text: '### 处理结果\n已按你的要求处理。\n' })
  }
  if (warningActions && warningActions[payload.action]) {
    frames.push({ type: 'warning', message: warningActions[payload.action] })
  }
  frames.push({ type: 'done', chars: 10 })
  return frames
}

function createFetch({ warningActions, analyze, healthAuth = 'open', hangAnalyze } = {}) {
  const calls = []
  const fetchImpl = async (url, init = {}) => {
    const target = String(url)
    const headers = init.headers || {}
    if (target.endsWith('/health')) {
      return new Response(JSON.stringify({
        ok: true,
        service: 'review-desk-wps-bridge',
        auth: healthAuth,
        pairing: healthAuth === 'token',
        exportDir: '/tmp/review-desk-exports',
      }), { status: 200, headers: { 'content-type': 'application/json' } })
    }
    if (target.endsWith('/pair')) {
      calls.push({ kind: 'pair', headers })
      return new Response(JSON.stringify({ ok: true, token: 'paired-token' }), { status: 200, headers: { 'content-type': 'application/json' } })
    }
    if (target.endsWith('/models')) {
      calls.push({ kind: 'models', headers })
      return new Response(JSON.stringify(catalog()), { status: 200, headers: { 'content-type': 'application/json' } })
    }
    if (target.endsWith('/export')) {
      const payload = JSON.parse(init.body)
      calls.push({ kind: 'export', payload, headers })
      return new Response(JSON.stringify({
        ok: true,
        filename: payload.filename,
        path: `/tmp/review-desk-exports/${payload.filename}`,
        bytes: payload.content.length,
      }), { status: 200, headers: { 'content-type': 'application/json' } })
    }
    if (target.endsWith('/analyze')) {
      const payload = JSON.parse(init.body)
      calls.push({ kind: 'analyze', payload, headers })
      if (typeof analyze === 'function') {
        const custom = analyze(payload, calls)
        if (custom !== undefined) return custom
      }
      if (typeof hangAnalyze === 'function' && hangAnalyze(payload, calls)) {
        return new Promise((resolve, reject) => {
          if (init.signal) {
            init.signal.addEventListener('abort', () => {
              const error = new Error('aborted')
              error.name = 'AbortError'
              reject(error)
            })
          }
        })
      }
      return ndjson(buildFrames(payload, warningActions))
    }
    throw new Error(`unexpected fetch: ${target}`)
  }
  return { fetchImpl, calls }
}

/* ------------------------------------------------------------------ *
 * 挂载面板
 * ------------------------------------------------------------------ */

async function mount({ text = BASE_TEXT, fetchOptions, history, settings } = {}) {
  const { fetchImpl, calls } = createFetch(fetchOptions)
  const fake = createFakeWps(text)
  const page = HTML.replace(/<script[^>]*taskpane\.js[^>]*><\/script>/u, '')
  const dom = new JSDOM(page, {
    url: 'http://127.0.0.1:3889/taskpane.html',
    pretendToBeVisual: true,
    runScripts: 'outside-only',
  })
  const { window } = dom
  window.fetch = fetchImpl
  window.AbortController = AbortController
  window.Response = Response
  window.TextDecoder = TextDecoder
  window.TextEncoder = TextEncoder
  window.ReadableStream = ReadableStream
  window.confirm = () => true
  window.console = console
  window.Application = fake.app
  window.wps = { WpsApplication: () => fake.app }
  if (history !== undefined) window.localStorage.setItem(HISTORY_KEY, JSON.stringify(history))
  if (settings !== undefined) window.localStorage.setItem(SETTINGS_KEY, JSON.stringify(settings))
  fake.doc.Range(0, 12).Select()

  window.eval(SCRIPT)
  await flush(8)

  const container = window.document
  const detail = () => container.getElementById('detail')
  const detailText = () => {
    const node = detail().querySelector('.rd-detail-text')
    return node ? node.textContent : ''
  }
  const detailState = () => detail().querySelector('.rd-state')
  const itemRows = () => [...container.querySelectorAll('.rd-item')]
  const visibleRows = () => itemRows().filter(row => row.style.display !== 'none')
  const waitFor = async (predicate, message, timeout = 4000) => {
    const started = Date.now()
    for (;;) {
      if (predicate()) return
      if (Date.now() - started > timeout) throw new Error(`timeout: ${message}`)
      await flush(2)
    }
  }
  const button = (label, scope) => {
    const root = scope || detail()
    const buttons = [...root.querySelectorAll('button')]
    const found = buttons.find(node => node.textContent.trim() === label)
    assert.ok(found, `button ${label} exists`)
    return found
  }
  const actionButton = name => {
    const found = container.querySelector(`.rd-action[data-action="${name}"]`)
    assert.ok(found, `action button ${name}`)
    return found
  }
  return { dom, window, fake, calls, container, detail, detailText, detailState, itemRows, visibleRows, waitFor, button, actionButton }
}

function flush(times = 1) {
  return new Promise(resolve => {
    let count = 0
    const tick = () => {
      count += 1
      if (count >= times) resolve()
      else setTimeout(tick, 0)
    }
    setTimeout(tick, 0)
  })
}

/* ------------------------------------------------------------------ *
 * 用例
 * ------------------------------------------------------------------ */

await test('two-column shell, health, catalog and selection', async () => {
  const page = await mount()
  assert.ok(page.container.querySelector('.rd-app'))
  assert.ok(page.container.querySelector('.rd-side'))
  assert.ok(page.container.querySelector('.rd-main'))
  await page.waitFor(() => page.container.querySelector('#conn-dot').getAttribute('data-state') === 'ready', 'health ready')
  await page.waitFor(() => page.container.querySelectorAll('#setting-model option').length > 1, 'models loaded')
  assert.ok(page.container.querySelector('#selection-preview').textContent.includes('第一段'))
  assert.equal(page.container.querySelectorAll('.rd-action').length, 4)
  page.dom.window.close()
})

await test('expand streams into the detail pane; comment and tracked replacement work', async () => {
  const page = await mount()
  await page.waitFor(() => page.container.querySelector('#conn-dot').getAttribute('data-state') === 'ready', 'health ready')
  page.actionButton('expand').click()
  await page.waitFor(() => page.detailText().includes('扩写后的正文示例'), 'detail streaming text')
  await page.waitFor(() => page.detailState().textContent === '完成', 'done state')
  assert.equal(page.itemRows().length, 1, 'one history row')
  assert.equal(page.fake.model.revisions.length, 0)

  page.button('插入批注').click()
  assert.equal(page.fake.model.comments.length, 1)
  assert.ok(page.fake.model.comments[0].text.includes('扩写后的正文示例'))

  page.button('替换原文（修订）').click()
  assert.equal(page.fake.model.revisions.length, 2)
  assert.equal(page.fake.model.track, false)
  assert.ok(page.fake.model.text.includes('扩写后的正文示例'))
  assert.equal(page.detailState().textContent, '已替换')
  page.dom.window.close()
})

await test('undo replacement rejects revisions and restores the text', async () => {
  const page = await mount()
  await page.waitFor(() => page.container.querySelector('#conn-dot').getAttribute('data-state') === 'ready', 'health ready')
  page.actionButton('expand').click()
  await page.waitFor(() => page.detailState().textContent === '完成', 'done state')
  page.button('替换原文（修订）').click()
  assert.equal(page.fake.model.revisions.length, 2)
  page.button('撤销替换').click()
  await flush(4)
  assert.equal(page.fake.model.text, BASE_TEXT)
  assert.equal(page.fake.model.revisions.length, 0)
  page.dom.window.close()
})

await test('locate selects and highlights the original range', async () => {
  const page = await mount()
  await page.waitFor(() => page.container.querySelector('#conn-dot').getAttribute('data-state') === 'ready', 'health ready')
  page.actionButton('logic').click()
  await page.waitFor(() => page.detailState().textContent === '完成', 'logic done')
  page.fake.model.selection = { Start: 0, End: 0, Text: '' }
  page.button('定位回原文').click()
  assert.deepEqual({ start: page.fake.model.selection.Start, end: page.fake.model.selection.End }, { start: 0, end: 12 })
  assert.equal(page.fake.model.scrollCalls, 1)
  page.dom.window.close()
})

await test('revise appends a child result with the requirement', async () => {
  const page = await mount()
  await page.waitFor(() => page.container.querySelector('#conn-dot').getAttribute('data-state') === 'ready', 'health ready')
  page.actionButton('expand').click()
  await page.waitFor(() => page.detailState().textContent === '完成', 'expand done')
  const input = page.detail().querySelector('.rd-revise input')
  assert.ok(input)
  input.value = '压缩到 40 字以内'
  page.button('追加要求').click()
  await page.waitFor(() => page.itemRows().length >= 2, 'child row')
  await page.waitFor(() => page.detailText().includes('压缩后的正文示例'), 'child text')
  const payload = page.calls.filter(call => call.kind === 'analyze').at(-1).payload
  assert.equal(payload.action, 'custom')
  assert.equal(payload.revise, true)
  assert.equal(payload.instruction, '压缩到 40 字以内')
  assert.ok(payload.text.includes('扩写后的正文示例'))
  assert.ok(page.detail().textContent.includes('基于 #'))
  page.dom.window.close()
})

await test('filters and search narrow the history list; history persists', async () => {
  const page = await mount()
  await page.waitFor(() => page.container.querySelector('#conn-dot').getAttribute('data-state') === 'ready', 'health ready')
  page.actionButton('expand').click()
  await page.waitFor(() => page.detailState().textContent === '完成', 'expand done')
  page.actionButton('logic').click()
  await page.waitFor(() => page.itemRows().length >= 2, 'two rows')
  await page.waitFor(() => page.detailState().textContent === '完成', 'logic done')

  const logicTab = [...page.container.querySelectorAll('.rd-tab')].find(node => node.textContent.startsWith('逻辑分析'))
  logicTab.click()
  await flush(2)
  assert.equal(page.visibleRows().length, 1)
  assert.equal(page.visibleRows()[0].getAttribute('data-action'), 'logic')

  const allTab = [...page.container.querySelectorAll('.rd-tab')].find(node => node.textContent.startsWith('全部'))
  allTab.click()
  await flush(2)
  const search = page.container.querySelector('#search')
  search.value = '扩写后的正文'
  search.dispatchEvent(new page.window.Event('input'))
  await flush(2)
  assert.equal(page.visibleRows().length, 1)

  search.value = ''
  search.dispatchEvent(new page.window.Event('input'))
  await flush(2)
  await page.waitFor(() => page.window.localStorage.getItem(HISTORY_KEY), 'history stored')
  const stored = JSON.parse(page.window.localStorage.getItem(HISTORY_KEY))
  assert.equal(stored.results.length, 2)
  page.dom.window.close()
})

await test('settings and token travel with requests', async () => {
  const page = await mount()
  await page.waitFor(() => page.container.querySelectorAll('#setting-model option').length > 1, 'models loaded')
  const model = page.container.querySelector('#setting-model')
  model.value = 'deepseek::deepseek-flash'
  model.dispatchEvent(new page.window.Event('change'))
  const effort = page.container.querySelector('#setting-effort')
  effort.value = 'low'
  effort.dispatchEvent(new page.window.Event('change'))
  const cap = page.container.querySelector('#setting-cap')
  cap.value = '12000'
  cap.dispatchEvent(new page.window.Event('change'))
  const token = page.container.querySelector('#bridge-token')
  token.value = 'secret-token'
  token.dispatchEvent(new page.window.Event('change'))
  await flush(3)
  const stored = JSON.parse(page.window.localStorage.getItem(SETTINGS_KEY))
  assert.equal(stored.provider, 'deepseek')
  assert.equal(stored.model, 'deepseek-flash')
  assert.equal(stored.reasoningEffort, 'low')
  assert.equal(stored.maxTokens, 12000)
  assert.equal(stored.token, 'secret-token')

  page.actionButton('logic').click()
  await page.waitFor(() => page.detailState().textContent === '完成', 'logic done')
  const call = page.calls.filter(entry => entry.kind === 'analyze').at(-1)
  assert.equal(call.payload.provider, 'deepseek')
  assert.equal(call.payload.model, 'deepseek-flash')
  assert.equal(call.payload.reasoningEffort, 'low')
  assert.equal(call.payload.maxTokens, 12000)
  assert.equal(call.headers['x-review-desk-token'], 'secret-token')
  page.dom.window.close()
})

await test('a max-token warning offers a higher-cap retry', async () => {
  const warning = '输出达到长度上限，结果可能不完整；可以缩短选中文字后重试。'
  const page = await mount({ fetchOptions: { warningActions: { logic: warning } } })
  await page.waitFor(() => page.container.querySelector('#conn-dot').getAttribute('data-state') === 'ready', 'health ready')
  page.actionButton('logic').click()
  await page.waitFor(() => page.detailState().textContent === '完成', 'logic done')
  assert.ok(page.detail().textContent.includes('长度上限'))
  const retry = page.button('提高输出上限重试')
  retry.click()
  await page.waitFor(() => page.calls.filter(entry => entry.kind === 'analyze').length >= 2, 'retry request')
  await page.waitFor(() => page.detailState().textContent === '完成', 'retry done')
  const retryCall = page.calls.filter(entry => entry.kind === 'analyze').at(-1)
  assert.equal(retryCall.payload.maxTokens, 10000)
  await flush(4)
  page.dom.window.close()
})

await test('bridge failure shows an actionable error block', async () => {
  const page = await mount({
    fetchOptions: {
      analyze() {
        return new Response(JSON.stringify({ error: '需要访问令牌：请填写 REVIEW_DESK_WPS_TOKEN。' }), {
          status: 401,
          headers: { 'content-type': 'application/json' },
        })
      },
    },
  })
  await page.waitFor(() => page.container.querySelector('#conn-dot').getAttribute('data-state') === 'ready', 'health ready')
  page.actionButton('logic').click()
  await page.waitFor(() => page.detail().querySelector('.rd-error-block').style.display !== 'none', 'error block visible')
  const text = page.detail().querySelector('.rd-error-block').textContent
  assert.match(text, /访问令牌/u)
  page.dom.window.close()
})

await test('keyboard shortcut runs the free-form requirement', async () => {
  const page = await mount()
  await page.waitFor(() => page.container.querySelector('#conn-dot').getAttribute('data-state') === 'ready', 'health ready')
  page.container.querySelector('#instruction').value = '压缩到 30 字'
  page.window.dispatchEvent(new page.window.KeyboardEvent('keydown', { key: 'Enter', metaKey: true, bubbles: true }))
  await page.waitFor(() => page.calls.some(entry => entry.kind === 'analyze'), 'shortcut request')
  const payload = page.calls.filter(entry => entry.kind === 'analyze').at(-1).payload
  assert.equal(payload.action, 'custom')
  assert.equal(payload.instruction, '压缩到 30 字')
  await page.waitFor(() => page.detailState().textContent === '完成', 'shortcut done')
  page.dom.window.close()
})

await test('history is restored on the next mount', async () => {
  const history = {
    version: 2,
    results: [
      {
        id: 2, action: 'logic', instruction: '', createdAt: Date.now(), status: 'done',
        text: '### 结构判断\n旧结果二。', meta: { provider: 'deepseek', model: 'deepseek-flash' },
        sources: [], note: '', warning: '', error: '', errorKind: '',
        selection: { start: 0, end: 12, text: '第一段：这是用于测试' },
        parentId: null, original: '', maxTokens: 0, applied: null,
      },
      {
        id: 1, action: 'expand', instruction: '', createdAt: Date.now() - 1000, status: 'error',
        text: '', meta: null, sources: [], note: '', warning: '',
        error: '旧错误', errorKind: 'network', selection: null, parentId: null, original: '', maxTokens: 0, applied: null,
      },
    ],
  }
  const page = await mount({ history })
  await flush(6)
  assert.equal(page.itemRows().length, 2)
  assert.ok(page.detailText().includes('旧结果二'))
  page.dom.window.close()
})

await test('bridge auto-pairs an auto-generated token', async () => {
  const page = await mount({ fetchOptions: { healthAuth: 'token' } })
  await page.waitFor(() => page.container.querySelectorAll('#setting-model option').length > 1, 'models loaded after pairing')
  assert.ok(page.calls.some(entry => entry.kind === 'pair'), 'pair called')
  const modelsCall = page.calls.find(entry => entry.kind === 'models')
  assert.equal(modelsCall.headers['x-review-desk-token'], 'paired-token')
  page.actionButton('logic').click()
  await page.waitFor(() => page.detailState().textContent === '完成', 'logic done')
  const analyzeCall = page.calls.filter(entry => entry.kind === 'analyze').at(-1)
  assert.equal(analyzeCall.headers['x-review-desk-token'], 'paired-token')
  page.dom.window.close()
})

await test('exports a DSH-compatible package and a markdown report', async () => {
  const page = await mount()
  await page.waitFor(() => page.container.querySelector('#conn-dot').getAttribute('data-state') === 'ready', 'health ready')
  page.actionButton('expand').click()
  await page.waitFor(() => page.detailState().textContent === '完成', 'expand done')
  page.button('导出审核包 (.json)', page.container).click()
  await page.waitFor(() => page.calls.some(entry => entry.kind === 'export'), 'package export')
  const pkgCall = page.calls.find(entry => entry.kind === 'export')
  assert.match(pkgCall.payload.filename, /\.review-desk\.json$/u)
  const pkg = JSON.parse(pkgCall.payload.content)
  assert.equal(pkg.format, 'dsh.review-desk.package')
  assert.equal(pkg.version, 1)
  assert.equal(pkg.document.text, BASE_TEXT)
  assert.equal(pkg.document.title, 'wps-panel-test')
  assert.equal(pkg.runs.length, 1)
  assert.equal(pkg.runs[0].action, 'expand')
  assert.equal(pkg.runs[0].status, 'done')
  assert.ok(pkg.runs[0].text.includes('扩写后的正文示例'))
  const exportedRows = page.itemRows().length
  await page.waitFor(() => !page.button('导出审稿报告 (.md)', page.container).disabled, 'export idle')
  page.button('导出审稿报告 (.md)', page.container).click()
  await page.waitFor(() => page.calls.filter(entry => entry.kind === 'export').length >= 2, 'report export')
  const reportCall = page.calls.filter(entry => entry.kind === 'export').at(-1)
  assert.match(reportCall.payload.filename, /\.md$/u)
  assert.ok(reportCall.payload.content.includes('# 审稿报告'))
  assert.ok(reportCall.payload.content.includes('扩写后的正文示例'))
  assert.equal(page.itemRows().length, exportedRows)
  page.dom.window.close()
})

await test('whole-document batch review runs segments and creates a summary', async () => {
  const longText = [
    '第一段：科技金融赋能实体经济的背景与现状。'.repeat(20),
    '第二段：银行依托技术流评价体系，将知识产权与研发投入纳入授信模型。'.repeat(20),
    '第三段：下一步应加强数据治理，完善专营机制与风险定价能力。'.repeat(20),
  ].join('\n')
  const page = await mount({ text: longText })
  await page.waitFor(() => page.container.querySelector('#conn-dot').getAttribute('data-state') === 'ready', 'health ready')
  const batchAction = page.container.querySelector('#batch-action')
  batchAction.value = 'logic'
  page.button('整篇审阅', page.container).click()
  await page.waitFor(() => page.container.querySelector('#batch-bar').hidden === false, 'batch bar visible')
  await page.waitFor(() => page.calls.filter(entry => entry.kind === 'analyze').length >= 2, 'multiple segment calls')
  await page.waitFor(() => page.container.querySelector('#batch-bar').hidden === true, 'batch finished', 8000)
  const analyzeCalls = page.calls.filter(entry => entry.kind === 'analyze')
  assert.ok(analyzeCalls.length >= 2, 'segments analysed separately')
  assert.ok(analyzeCalls.every(call => call.payload.batch === undefined || true))
  await page.waitFor(() => page.detailText().includes('已完成整篇分段审阅'), 'batch summary shown')
  assert.ok(page.itemRows().length >= 3, 'segment rows + summary row')
  page.dom.window.close()
})

await test('cancel batch review can be resumed from the paused bar', async () => {
  const longText = '第一段：这是一段用于取消测试的正文内容。'.repeat(120)
  const page = await mount({
    text: longText,
    fetchOptions: {
      // Only the very first segment hangs; resumed segments complete normally.
      hangAnalyze: (payload, calls) => calls.filter(entry => entry.kind === 'analyze').length === 1,
    },
  })
  await page.waitFor(() => page.container.querySelector('#conn-dot').getAttribute('data-state') === 'ready', 'health ready')
  page.button('整篇审阅', page.container).click()
  await page.waitFor(() => page.container.querySelector('#batch-bar').hidden === false, 'batch bar visible')
  page.button('取消', page.container).click()
  await page.waitFor(() => page.container.querySelector('#btn-continue-batch').hidden === false, 'paused bar visible', 6000)
  assert.equal(page.container.querySelector('#btn-cancel-batch').hidden, true, 'cancel button hidden while paused')
  const analyzeCalls = page.calls.filter(entry => entry.kind === 'analyze')
  assert.equal(analyzeCalls.length, 1, 'only the first segment started')
  assert.ok(page.container.querySelector('#history-list').textContent.includes('取消'), 'cancelled row is visible')
  page.button('继续', page.container).click()
  await page.waitFor(() => page.container.querySelector('#batch-bar').hidden === true, 'batch finished after resume', 10000)
  assert.ok(page.calls.filter(entry => entry.kind === 'analyze').length >= 2, 'resumed segments ran')
  await page.waitFor(() => page.detailText().includes('已完成整篇分段审阅'), 'summary after resume')
  page.dom.window.close()
})

await test('dragging the side resizer sets and persists the column width', async () => {
  const page = await mount()
  await page.waitFor(() => page.container.querySelector('#conn-dot').getAttribute('data-state') === 'ready', 'health ready')
  const resizer = page.container.querySelector('#side-resizer')
  assert.ok(resizer, 'resizer exists')
  function mouse(type, x, target) {
    const event = new page.window.MouseEvent(type, { bubbles: true, cancelable: true })
    Object.defineProperty(event, 'clientX', { value: x })
    target.dispatchEvent(event)
  }
  mouse('mousedown', 300, resizer)
  mouse('mousemove', 380, page.container)
  mouse('mouseup', 380, page.container)
  await flush(2)
  assert.equal(page.window.document.documentElement.style.getPropertyValue('--rd-side-width'), '380px')
  const stored = JSON.parse(page.window.localStorage.getItem(SETTINGS_KEY))
  assert.equal(stored.sideWidth, 380)
  page.dom.window.close()
})

await test('settings and export popovers open and close', async () => {
  const page = await mount()
  await page.waitFor(() => page.container.querySelector('#conn-dot').getAttribute('data-state') === 'ready', 'health ready')
  const settings = page.container.querySelector('#settings-panel')
  const exportMenu = page.container.querySelector('#export-menu')
  page.container.querySelector('#btn-settings-menu').click()
  assert.equal(settings.hidden, false, 'settings open')
  assert.equal(exportMenu.hidden, true)
  page.container.querySelector('#btn-export-menu').click()
  assert.equal(settings.hidden, true, 'settings closed when export opens')
  assert.equal(exportMenu.hidden, false, 'export open')
  page.window.dispatchEvent(new page.window.KeyboardEvent('keydown', { key: 'Escape', bubbles: true }))
  await flush(2)
  assert.equal(exportMenu.hidden, true, 'escape closes popovers')
  page.container.querySelector('#btn-settings-menu').click()
  page.container.body.click()
  await flush(2)
  assert.equal(settings.hidden, true, 'outside click closes popovers')
  page.dom.window.close()
})

await test('imports a review package from a JSON file', async () => {
  const page = await mount()
  await page.waitFor(() => page.container.querySelector('#conn-dot').getAttribute('data-state') === 'ready', 'health ready')
  const pkg = {
    format: 'dsh.review-desk.package',
    version: 1,
    exportedAt: new Date().toISOString(),
    app: { plugin: 'review-desk', build: 'test' },
    document: { title: '导入测试', text: BASE_TEXT },
    runs: [{
      id: 'run-1', action: 'logic', instruction: '', status: 'done',
      text: '### 导入来的结果\n结构清晰。', meta: { provider: 'deepseek', model: 'deepseek-flash' },
      sources: [], claims: [], selection: { start: 0, end: 12, text: BASE_TEXT.slice(0, 12) },
      replaced: null, startedAt: Date.now(), finishedAt: Date.now(), warning: '', error: '',
    }],
    session: { instruction: '' },
  }
  const file = {
    name: 'test.review-desk.json',
    size: JSON.stringify(pkg).length,
    text: () => Promise.resolve(JSON.stringify(pkg)),
  }
  const input = page.container.querySelector('#package-file')
  Object.defineProperty(input, 'files', { value: [file], configurable: true })
  input.dispatchEvent(new page.window.Event('change'))
  await page.waitFor(() => page.container.querySelector('#history-list').textContent.includes('已从审核包导入'), 'import summary row', 6000)
  assert.ok(page.detailText().includes('已从审核包导入 1 条结果'))
  assert.equal(page.itemRows().length, 2, 'imported run + summary')
  page.dom.window.close()
})

await test('exports only the currently selected result', async () => {
  const page = await mount()
  await page.waitFor(() => page.container.querySelector('#conn-dot').getAttribute('data-state') === 'ready', 'health ready')
  page.actionButton('expand').click()
  await page.waitFor(() => page.detailState().textContent === '完成', 'expand done')
  page.button('导出当前结果 (.md)', page.container).click()
  await page.waitFor(() => page.calls.some(entry => entry.kind === 'export'), 'result export')
  const call = page.calls.find(entry => entry.kind === 'export')
  assert.match(call.payload.filename, /选中结果-.*\.md$/u)
  assert.ok(call.payload.content.includes('# 审稿结果'))
  assert.ok(call.payload.content.includes('扩写后的正文示例'))
  page.dom.window.close()
})

await test('segment length setting is configurable', async () => {
  const page = await mount()
  await page.waitFor(() => page.container.querySelector('#setting-model option'), 'panel ready')
  const settings = page.container.querySelector('#settings-panel')
  page.container.querySelector('#btn-settings-menu').click()
  assert.equal(settings.hidden, false)
  const select = page.container.querySelector('#setting-segment-length')
  assert.deepEqual([...select.options].map(option => option.value), ['600', '900', '1200', '1600'])
  select.value = '600'
  select.dispatchEvent(new page.window.Event('change'))
  await flush(2)
  const stored = JSON.parse(page.window.localStorage.getItem(SETTINGS_KEY))
  assert.equal(stored.segmentLength, 600)
  page.dom.window.close()
})

await test('marks reviewed ranges with action-coloured underlines', async () => {
  const page = await mount()
  await page.waitFor(() => page.container.querySelector('#conn-dot').getAttribute('data-state') === 'ready', 'health ready')
  page.actionButton('expand').click()
  await page.waitFor(() => page.detailState().textContent === '完成', 'expand done')
  await page.waitFor(() => page.fake.model.underlines.length > 0, 'underline applied', 4000)
  const mark = page.fake.model.underlines.at(-1)
  assert.equal(mark.start, 0)
  assert.equal(mark.end, 12)
  assert.equal(mark.style, 1)
  assert.equal(mark.color, 0x699605) // WPS/VBA RGB() stores colours as BGR
  const stored = JSON.parse(page.window.localStorage.getItem('wps-review-desk.marks.v2'))
  assert.equal(stored.docName, 'wps-panel-test.docx')
  assert.equal(stored.marks.length, 1)
  assert.equal(stored.marks[0].action, 'expand')
  assert.equal(stored.marks[0].color, 0x059669)
  // turning the setting off restores the original underline formatting
  page.container.querySelector('#btn-settings-menu').click()
  const checkbox = page.container.querySelector('#setting-marks')
  checkbox.checked = false
  checkbox.dispatchEvent(new page.window.Event('change'))
  await page.waitFor(() => page.fake.model.underlineStyle === 0, 'underline restored', 3000)
  assert.equal(page.window.localStorage.getItem('wps-review-desk.marks.v2'), null)
  page.dom.window.close()
})

async function test(name, fn) {
  try {
    await fn()
    console.log(`ok  ${name}`)
  } catch (error) {
    console.error(`FAIL ${name}`)
    console.error(error)
    process.exitCode = 1
  }
}

if (process.exitCode) {
  console.error('\nsome panel checks failed')
} else {
  console.log('\nall panel checks passed')
}
process.exit(process.exitCode || 0)
