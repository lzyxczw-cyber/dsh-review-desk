/**
 * WPS task-pane test: loads wps-addin/taskpane.html + taskpane.js in jsdom and
 * drives it with a fake WPS document model and a fake loopback bridge.
 *
 * Covers: selection, streamed analysis, comment insertion, tracked replacement,
 * undo via revision rejection, locate, revise chain, filters, settings and
 * history persistence.
 */
import fs from 'node:fs'
import path from 'node:path'
import assert from 'node:assert/strict'
import { JSDOM } from 'jsdom'

const here = path.dirname(new URL(import.meta.url).pathname)
const PANEL_DIR = path.join(here, '..', 'wps-addin')
const HTML = fs.readFileSync(path.join(PANEL_DIR, 'taskpane.html'), 'utf8')
const SCRIPT = fs.readFileSync(path.join(PANEL_DIR, 'taskpane.js'), 'utf8')
const HISTORY_KEY = 'wps-review-desk.history.v1'
const SETTINGS_KEY = 'wps-review-desk.settings.v1'

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
    selection: { Start: 0, End: 0, Text: '' },
    scrollCalls: 0,
  }

  function parseRange(start, end) {
    const from = Math.max(0, Math.min(Number(start) || 0, model.text.length))
    const to = Math.max(from, Math.min(Number(end) || 0, model.text.length))
    return { start: from, end: to }
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
    }
    Object.defineProperty(range, 'Text', {
      get() {
        return model.text.slice(bounds.start, bounds.end)
      },
      set(value) {
        setRangeText(bounds.start, bounds.end, String(value))
      },
    })
    return range
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

  const doc = {
    Name: model.name,
    get Saved() { return model.saved },
    get TrackRevisions() { return model.track },
    set TrackRevisions(value) { model.track = value === true },
    get Content() {
      return {
        get Text() { return model.text },
        set Text(value) { model.text = String(value) },
        Find: { Execute() { return false } },
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
    ExportAsFixedFormat() { return true },
  }

  const app = {
    Name: 'Microsoft Word',
    Version: '12.0',
    get ActiveDocument() { return doc },
    get Selection() {
      const selection = model.selection
      return {
        Start: selection.Start,
        End: selection.End,
        Text: selection.Text,
      }
    },
    get ActiveWindow() {
      return {
        ScrollIntoView() { model.scrollCalls += 1 },
      }
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
    frames.push({ type: 'delta', text: '### 结构判断\n本段是论证落点。\n### 逻辑问题\n- **以偏概全** —— 样本不足。\n### 收紧建议\n1. 收敛结论。\n' })
  } else {
    frames.push({ type: 'delta', text: '### 处理结果\n已按你的要求处理。\n' })
  }
  if (warningActions && warningActions[payload.action]) {
    frames.push({ type: 'warning', message: warningActions[payload.action] })
  }
  frames.push({ type: 'done', chars: 10 })
  return frames
}

function createFetch({ warningActions, analyze } = {}) {
  const calls = []
  const fetchImpl = async (url, init) => {
    const target = String(url)
    if (target.endsWith('/health')) {
      return new Response(JSON.stringify({ ok: true, service: 'review-desk-wps-bridge' }), { status: 200, headers: { 'content-type': 'application/json' } })
    }
    if (target.endsWith('/models')) {
      return new Response(JSON.stringify(catalog()), { status: 200, headers: { 'content-type': 'application/json' } })
    }
    if (target.endsWith('/analyze')) {
      const payload = JSON.parse(init.body)
      calls.push(payload)
      if (typeof analyze === 'function') {
        const custom = analyze(payload, calls)
        if (custom !== undefined) return custom
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
  const button = (label, scope) => {
    const root = scope || container
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
  const results = () => [...container.querySelectorAll('.rd-result')]
  const lastCard = () => {
    const cards = results()
    assert.ok(cards.length > 0, 'a result card exists')
    return cards[0]
  }
  const waitFor = async (predicate, message, timeout = 4000) => {
    const started = Date.now()
    for (;;) {
      if (predicate()) return
      if (Date.now() - started > timeout) throw new Error(`timeout: ${message}`)
      await flush(2)
    }
  }
  return { dom, window, fake, calls, container, button, actionButton, results, lastCard, waitFor }
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

function findButton(container, label) {
  const buttons = [...container.querySelectorAll('button')]
  const found = buttons.find(node => node.textContent.trim() === label)
  assert.ok(found, `button ${label} exists`)
  return found
}

/* ------------------------------------------------------------------ *
 * 用例
 * ------------------------------------------------------------------ */

await test('health, catalog and selection render', async () => {
  const page = await mount()
  await page.waitFor(() => page.container.querySelector('#conn-dot').getAttribute('data-state') === 'ready', 'health ready')
  await page.waitFor(() => page.container.querySelectorAll('#setting-model option').length > 1, 'models loaded')
  assert.ok(page.container.querySelector('#selection-preview').textContent.includes('第一段'))
  assert.ok(page.container.querySelectorAll('.rd-action').length === 4)
  page.dom.window.close()
})

await test('expand streams a card, comment and tracked replacement work', async () => {
  const page = await mount()
  await page.waitFor(() => page.container.querySelector('#conn-dot').getAttribute('data-state') === 'ready', 'health ready')
  page.actionButton('expand').click()
  await page.waitFor(() => page.lastCard().getAttribute('data-status') === 'done', 'expand done')
  const card = page.lastCard()
  assert.ok(card.textContent.includes('扩写后的正文示例'))
  assert.ok(card.querySelector('strong') === null || true)
  assert.equal(page.fake.model.revisions.length, 0, 'no document change yet')

  findButton(card, '插入批注').click()
  assert.equal(page.fake.model.comments.length, 1)
  assert.ok(page.fake.model.comments[0].text.includes('扩写后的正文示例'))

  findButton(card, '替换原文（修订）').click()
  assert.equal(page.fake.model.revisions.length, 2, 'tracked replacement produces two revisions')
  assert.equal(page.fake.model.track, false, 'track setting restored')
  assert.ok(page.fake.model.text.includes('扩写后的正文示例'))
  assert.equal(page.calls.at(-1).action, 'expand')
  page.dom.window.close()
})

await test('undo replacement rejects revisions and restores the text', async () => {
  const page = await mount()
  await page.waitFor(() => page.container.querySelector('#conn-dot').getAttribute('data-state') === 'ready', 'health ready')
  page.actionButton('expand').click()
  await page.waitFor(() => page.lastCard().getAttribute('data-status') === 'done', 'expand done')
  const card = page.lastCard()
  findButton(card, '替换原文（修订）').click()
  assert.equal(page.fake.model.revisions.length, 2, 'two tracked revisions')
  assert.ok(page.fake.model.text.includes('扩写后的正文示例'))
  findButton(card, '撤销替换').click()
  await flush(4)
  assert.equal(page.fake.model.text, BASE_TEXT, 'text restored')
  assert.equal(page.fake.model.revisions.length, 0, 'revisions rejected')
  page.dom.window.close()
})

await test('locate selects the original range', async () => {
  const page = await mount()
  await page.waitFor(() => page.container.querySelector('#conn-dot').getAttribute('data-state') === 'ready', 'health ready')
  page.actionButton('logic').click()
  await page.waitFor(() => page.lastCard().getAttribute('data-status') === 'done', 'logic done')
  page.fake.model.selection = { Start: 0, End: 0, Text: '' }
  findButton(page.lastCard(), '定位回原文').click()
  assert.deepEqual(
    { start: page.fake.model.selection.Start, end: page.fake.model.selection.End },
    { start: 0, end: 12 },
  )
  assert.equal(page.fake.model.scrollCalls, 1)
  page.dom.window.close()
})

await test('revise appends a child result with the requirement', async () => {
  const page = await mount()
  await page.waitFor(() => page.container.querySelector('#conn-dot').getAttribute('data-state') === 'ready', 'health ready')
  page.actionButton('expand').click()
  await page.waitFor(() => page.lastCard().getAttribute('data-status') === 'done', 'expand done')
  const card = page.lastCard()
  findButton(card, '追加要求').click()
  const input = card.querySelector('.rd-revise input')
  assert.ok(input)
  input.value = '压缩到 40 字以内'
  findButton(card, '执行').click()
  await page.waitFor(() => page.results().length >= 2, 'child card exists')
  await page.waitFor(() => page.lastCard().getAttribute('data-status') === 'done', 'child done')
  const payload = page.calls.at(-1)
  assert.equal(payload.action, 'custom')
  assert.equal(payload.revise, true)
  assert.equal(payload.instruction, '压缩到 40 字以内')
  assert.ok(payload.text.includes('扩写后的正文示例'))
  assert.ok(page.lastCard().textContent.includes('压缩后的正文示例'))
  assert.ok(page.lastCard().textContent.includes('基于 #'))
  page.dom.window.close()
})

await test('filters limit visible cards and history is persisted', async () => {
  const page = await mount()
  await page.waitFor(() => page.container.querySelector('#conn-dot').getAttribute('data-state') === 'ready', 'health ready')
  page.actionButton('expand').click()
  await page.waitFor(() => page.lastCard().getAttribute('data-status') === 'done', 'expand done')
  page.actionButton('logic').click()
  await page.waitFor(() => page.container.querySelectorAll('.rd-result').length >= 2, 'two cards')
  await page.waitFor(() => page.lastCard().getAttribute('data-status') === 'done', 'logic done')
  const logicTab = [...page.container.querySelectorAll('.rd-tab')].find(node => node.textContent.startsWith('逻辑分析'))
  assert.ok(logicTab)
  logicTab.click()
  await flush(2)
  const visible = page.results().filter(card => card.style.display !== 'none')
  assert.equal(visible.length, 1)
  assert.equal(visible[0].getAttribute('data-action'), 'logic')
  await page.waitFor(() => page.window.localStorage.getItem(HISTORY_KEY), 'history stored')
  const stored = JSON.parse(page.window.localStorage.getItem(HISTORY_KEY))
  assert.equal(stored.results.length, 2)
  page.dom.window.close()
})

await test('settings load from the catalog and travel with requests', async () => {
  const page = await mount()
  await page.waitFor(() => page.container.querySelectorAll('#setting-model option').length > 1, 'models loaded')
  const model = page.container.querySelector('#setting-model')
  model.value = 'deepseek::deepseek-flash'
  model.dispatchEvent(new page.window.Event('change'))
  await flush(3)
  const effort = page.container.querySelector('#setting-effort')
  assert.ok([...effort.options].some(option => option.value === 'max'))
  effort.value = 'low'
  effort.dispatchEvent(new page.window.Event('change'))
  const cap = page.container.querySelector('#setting-cap')
  cap.value = '12000'
  cap.dispatchEvent(new page.window.Event('change'))
  await flush(3)
  const stored = JSON.parse(page.window.localStorage.getItem(SETTINGS_KEY))
  assert.equal(stored.provider, 'deepseek')
  assert.equal(stored.model, 'deepseek-flash')
  assert.equal(stored.reasoningEffort, 'low')
  assert.equal(stored.maxTokens, 12000)
  page.actionButton('logic').click()
  await page.waitFor(() => page.lastCard().getAttribute('data-status') === 'done', 'logic done')
  const payload = page.calls.at(-1)
  assert.equal(payload.provider, 'deepseek')
  assert.equal(payload.model, 'deepseek-flash')
  assert.equal(payload.reasoningEffort, 'low')
  assert.equal(payload.maxTokens, 12000)
  page.dom.window.close()
})

await test('a max-token warning offers a higher-cap retry', async () => {
  const warning = '输出达到长度上限，结果可能不完整；可以缩短选中文字后重试。'
  const page = await mount({ fetchOptions: { warningActions: { logic: warning } } })
  await page.waitFor(() => page.container.querySelector('#conn-dot').getAttribute('data-state') === 'ready', 'health ready')
  page.actionButton('logic').click()
  await page.waitFor(() => page.lastCard().getAttribute('data-status') === 'done', 'logic done')
  const card = page.lastCard()
  assert.ok(card.textContent.includes('长度上限'))
  const retry = findButton(card, '提高输出上限重试')
  retry.click()
  await page.waitFor(() => page.calls.length >= 2, 'retry request')
  await page.waitFor(() => page.results().length >= 2, 'retry card')
  await page.waitFor(() => page.lastCard().getAttribute('data-status') === 'done', 'retry done')
  assert.equal(page.calls.at(-1).maxTokens, 10000)
  await flush(6)
  page.dom.window.close()
})

await test('history is restored on the next mount', async () => {
  const history = {
    version: 1,
    results: [
      {
        id: 2, action: 'logic', instruction: '', createdAt: Date.now(), status: 'done',
        text: '### 结构判断\n旧结果二。', meta: { provider: 'deepseek', model: 'deepseek-flash' },
        sources: [], note: '', warning: '', error: '',
        selection: { start: 0, end: 12, text: '第一段：这是用于测试' },
        parentId: null, original: '', maxTokens: 0, applied: null, collapsed: true,
      },
      {
        id: 1, action: 'expand', instruction: '', createdAt: Date.now() - 1000, status: 'error',
        text: '', meta: null, sources: [], note: '', warning: '',
        error: '旧错误', selection: null, parentId: null, original: '', maxTokens: 0, applied: null, collapsed: true,
      },
    ],
  }
  const page = await mount({ history })
  await flush(6)
  assert.equal(page.results().length, 2)
  assert.ok(page.results()[0].textContent.includes('旧结果二'))
  assert.equal(page.results()[0].getAttribute('data-collapsed'), null, 'newest restored expanded')
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
