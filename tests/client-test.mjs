/**
 * Client-half test: mounts the real page component in jsdom, drives the
 * selection/action flow, and feeds it NDJSON frames through a mocked fetch.
 */
import fs from 'node:fs'
import path from 'node:path'
import assert from 'node:assert/strict'
import { JSDOM } from 'jsdom'

const here = path.dirname(new URL(import.meta.url).pathname)
const SOURCE = fs.readFileSync(path.join(here, '..', 'client.js'), 'utf8')

const dom = new JSDOM('<!doctype html><html><body><div id="root"></div></body></html>', {
  url: 'http://127.0.0.1:54386/',
  pretendToBeVisual: true,
})
const { window } = dom
globalThis.window = window
globalThis.document = window.document
Object.defineProperty(globalThis, 'navigator', { value: window.navigator, configurable: true })
globalThis.localStorage = window.localStorage
globalThis.HTMLElement = window.HTMLElement
globalThis.DOMParser = window.DOMParser
globalThis.Node = window.Node
globalThis.Event = window.Event
globalThis.IS_REACT_ACT_ENVIRONMENT = true

// React DOM decides at import time whether a DOM exists, so it must load after
// the jsdom globals above; otherwise every event takes the legacy IE path.
const { default: React } = await import('react')
const { createRoot } = await import('react-dom/client')
const { act } = await import('react-dom/test-utils')

/** Build the Chinese `t` the plugin would get from the locale service. */
function dictionaries() {
  const start = SOURCE.indexOf('const ZH = {')
  const mid = SOURCE.indexOf('const EN = {')
  const end = SOURCE.indexOf('/** The document the panel opens with')
  const parse = (src) => new Function(`return (${src.slice(src.indexOf('{'), src.lastIndexOf('}') + 1)})`)()
  return { ZH: parse(SOURCE.slice(start, mid)), EN: parse(SOURCE.slice(mid, end)) }
}
const { ZH } = dictionaries()

/** Instantiate one fresh module so each case starts from its own store. */
function mount({ seed, fetchImpl, catalog = DEFAULT_CATALOG, widths } = {}) {
  if (fetchImpl !== undefined) {
    globalThis.fetch = fetchImpl
    window.fetch = fetchImpl
  }
  window.localStorage.clear()
  if (seed !== undefined) window.localStorage.setItem('dsh.review-desk.v1', JSON.stringify(seed))
  if (widths !== undefined) window.localStorage.setItem('dsh.review-desk.widths.v1', JSON.stringify(widths))
  let factory
  const script = new Function('window', SOURCE)
  window.__ModuleLoader__ = { load: ({ factory: captured }) => { factory = captured } }
  script(window)
  const registry = {}
  const remote = {
    session: {
      async modelCatalog() {
        return catalog === null
          ? { ok: false, error: { code: 'model-catalog/failed', message: 'catalog unavailable' } }
          : { ok: true, value: catalog }
      },
    },
  }
  const ctx = {
    effect(fn) { fn(); return () => {} },
    locale: { register() { return () => {} }, bind: () => key => key },
    slots: {
      inject(key, callback) { callback(); return () => {} },
      register(options, Component) { registry[options.name] = { options, Component }; return () => {} },
    },
    remote,
    get(name) { return name === 'remote' ? remote : undefined },
  }
  const mod = factory(name => {
    if (name === 'react') return React
    throw new Error(`unexpected module: ${name}`)
  })
  mod.apply(ctx)
  assert.ok(registry['main'], 'main panel registered')
  assert.ok(registry['sidebar.panellist'], 'sidebar entry registered')
  assert.equal(registry['sidebar.panellist'].options.id, 'review-desk')
  assert.equal(registry['main'].options.key, 'review-desk')
  assert.equal(registry['sidebar.panellist'].options.label(), '文稿审核台')

  const container = document.createElement('div')
  container.className = 'test-root'
  document.body.appendChild(container)
  const root = createRoot(container)
  const t = key => ZH[key] ?? key
  return {
    registry,
    root,
    t,
    container,
    render: async () => {
      await act(async () => {
        root.render(React.createElement(registry['main'].Component, { t, remote }))
      })
    },
    unmount: async () => {
      await act(async () => { root.unmount() })
      container.remove()
    },
  }
}

/** A fetch that answers with a streamed NDJSON body built from frames. */
function ndjsonFetch(frames, { status = 200 } = {}) {
  return async () => {
    const body = frames.map(frame => `${JSON.stringify(frame)}\n`).join('')
    const stream = new ReadableStream({
      start(controller) {
        controller.enqueue(new TextEncoder().encode(body))
        controller.close()
      },
    })
    return new Response(stream, {
      status,
      headers: { 'content-type': 'application/x-ndjson; charset=utf-8' },
    })
  }
}


/* ------------------------------------------------------------------ *
 * Import fixtures: a stored-entry ZIP holding word/document.xml, and a
 * one-page PDF whose content stream shows a literal string.
 * ------------------------------------------------------------------ */

function zipStore(name, payload) {
  const nameBytes = new TextEncoder().encode(name)
  const data = new TextEncoder().encode(payload)
  const local = new Uint8Array(30 + nameBytes.length + data.length)
  const lv = new DataView(local.buffer)
  lv.setUint32(0, 0x04034b50, true)
  lv.setUint16(4, 20, true)
  lv.setUint16(8, 0, true)
  lv.setUint32(14, 0, true)
  lv.setUint32(18, data.length, true)
  lv.setUint32(22, data.length, true)
  lv.setUint16(26, nameBytes.length, true)
  lv.setUint16(28, 0, true)
  local.set(nameBytes, 30)
  local.set(data, 30 + nameBytes.length)

  const central = new Uint8Array(46 + nameBytes.length)
  const cv = new DataView(central.buffer)
  cv.setUint32(0, 0x02014b50, true)
  cv.setUint16(4, 20, true)
  cv.setUint16(6, 20, true)
  cv.setUint16(10, 0, true)
  cv.setUint32(20, data.length, true)
  cv.setUint32(24, data.length, true)
  cv.setUint16(28, nameBytes.length, true)
  cv.setUint32(42, 0, true)
  central.set(nameBytes, 46)

  const eocd = new Uint8Array(22)
  const ev = new DataView(eocd.buffer)
  ev.setUint32(0, 0x06054b50, true)
  ev.setUint16(8, 1, true)
  ev.setUint16(10, 1, true)
  ev.setUint32(12, central.length, true)
  ev.setUint32(16, local.length, true)

  const out = new Uint8Array(local.length + central.length + eocd.length)
  out.set(local, 0)
  out.set(central, local.length)
  out.set(eocd, local.length + central.length)
  return out
}

function buildDocx(paragraphs) {
  const body = paragraphs
    .map(text => `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`)
    .join('')
  return buildDocxBody(body)
}

function buildDocxBody(body) {
  const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${body}</w:body></w:document>`
  return zipStore('word/document.xml', xml)
}

/** The model catalog the Host Remote would hand the panel. */
const DEFAULT_CATALOG = {
  default: { provider: 'deepseek', model: 'deepseek-v4-flash', reasoningEffort: 'max' },
  groups: [{
    id: 'deepseek',
    name: 'DeepSeek',
    models: [
      {
        id: 'deepseek-v4-flash',
        name: 'DeepSeek V4 Flash',
        reasoning: { efforts: [{ id: 'low', name: '低' }, { id: 'high', name: '高' }], defaultEffort: 'low' },
      },
      { id: 'deepseek-v4-chat', name: 'DeepSeek V4 Chat' },
    ],
  }],
  failures: [],
}

/** Fetch mock for the analysis stream and the original-page render. */
function appFetch({ frames = [], pages = null, seen = [] } = {}) {
  return async (url, init) => {
    const target = String(url)
    if (target.includes('review-desk/original-pages')) {
      const status = pages === null ? 501 : 200
      return new Response(JSON.stringify(pages ?? { error: '本机未安装文档渲染引擎，无法生成原稿版式预览。' }), {
        status,
        headers: { 'content-type': 'application/json' },
      })
    }
    if (typeof init?.body === 'string') seen.push(JSON.parse(init.body))
    return ndjsonFetch(frames)()
  }
}

function buildPdf(sentence) {
  const content = `BT /F1 12 Tf 72 720 Td (${sentence}) Tj ET`
  const objects = [
    '1 0 obj << /Type /Catalog /Pages 2 0 R >> endobj\n',
    '2 0 obj << /Type /Pages /Kids [3 0 R] /Count 1 >> endobj\n',
    '3 0 obj << /Type /Page /Parent 2 0 R /Resources << /Font << /F1 5 0 R >> >> /Contents 4 0 R >> endobj\n',
    `4 0 obj << /Length ${content.length} >>\nstream\n${content}\nendstream\nendobj\n`,
    '5 0 obj << /Type /Font /Subtype /Type1 /BaseFont /Helvetica >> endobj\n',
  ]
  const body = `%PDF-1.4\n${objects.join('')}trailer << /Root 1 0 R >>\n%%EOF\n`
  return new TextEncoder().encode(body)
}

/** Choose one file in the hidden import input. */
async function importFile(page, name, bytes) {
  const input = page.container.querySelector('input[type="file"]')
  assert.ok(input, 'import input exists')
  const file = new File([bytes], name)
  Object.defineProperty(input, 'files', { value: [file], configurable: true })
  await act(async () => {
    input.dispatchEvent(new window.Event('change', { bubbles: true }))
  })
  await flush(12)
}


/** Switch the manuscript column between its editor and its rendered preview. */
async function setMode(page, mode) {
  const index = mode === 'preview' ? 1 : 0
  const segments = queryAll(page.container, '.rd-segment')
  if (segments.length > 0 && segments[index].getAttribute('data-active') === null) {
    click(segments[index])
    await flush()
  }
}


/** Fetch mock that answers successive analyze calls with queued frame sets. */
function queueFetch(queue, { pages = null, seen = [] } = {}) {
  let index = 0
  return async (url, init) => {
    const target = String(url)
    if (target.includes('review-desk/original-pages')) {
      return new Response(JSON.stringify(pages ?? { error: 'no engine' }), { status: pages === null ? 501 : 200, headers: { 'content-type': 'application/json' } })
    }
    if (typeof init?.body === 'string') seen.push(JSON.parse(init.body))
    const frames = queue[Math.min(index, queue.length - 1)]
    index += 1
    return ndjsonFetch(frames)()
  }
}

const tick = () => new Promise(resolve => setTimeout(resolve, 0))

async function flush(times = 6) {
  for (let index = 0; index < times; index += 1) {
    await act(async () => { await tick() })
  }
}

function queryAll(container, selector) {
  return [...container.querySelectorAll(selector)]
}

function click(element) {
  assert.ok(element, 'element to click exists')
  element.dispatchEvent(new window.MouseEvent('click', { bubbles: true, button: 0 }))
}

let failures = 0
async function test(name, fn) {
  try {
    await fn()
    console.log(`ok  ${name}`)
  } catch (error) {
    failures += 1
    console.error(`FAIL ${name}`)
    console.error(error)
  }
}

/* ------------------------------------------------------------------ */

await test('empty state renders and offers the sample', async () => {
  const page = mount()
  await page.render()
  assert.ok(page.container.querySelector('.rd-page'), 'page root')
  assert.equal(page.container.querySelector('.rd-title').textContent, '文稿审核台')
  assert.ok(page.container.textContent.includes('还没有文稿'))
  const buttons = queryAll(page.container, '.rd-empty-actions .rd-button').map(node => node.textContent)
  assert.deepEqual(buttons, ['载入示例', '导入文稿'])
  assert.equal(queryAll(page.container, '.rd-action').length, 4)
  assert.ok(queryAll(page.container, '.rd-action').every(node => node.disabled), 'actions start disabled')
  await page.unmount()
})

await test('sample loads, selection enables actions, logic result streams in', async () => {
  const page = mount({
    fetchImpl: ndjsonFetch([
      { type: 'meta', action: 'logic', provider: 'deepseek', model: 'deepseek-v4-flash' },
      { type: 'stage', stage: 'writing' },
      { type: 'delta', text: '### 结构判断\n本段是论证的落点。\n' },
      { type: 'delta', text: '### 逻辑问题\n- **以偏概全** —— 摘录 → 样本不足。\n' },
      { type: 'done', chars: 40 },
    ]),
  })
  await page.render()
  click(queryAll(page.container, '.rd-empty-actions .rd-button')[0])
  await flush()
  const field = page.container.querySelector('textarea.rd-text')
  assert.ok(field, 'manuscript field appears')
  assert.ok(field.value.includes('科技金融'), 'sample text loaded')
  assert.ok(page.container.querySelector('.rd-selection').textContent.includes('已选中'), 'sample preselects a passage')
  assert.ok(queryAll(page.container, '.rd-action').every(node => !node.disabled), 'preselection enables the actions')

  const start = field.value.indexOf('从实践效果看')
  const end = field.value.indexOf('同时也要看到')
  field.selectionStart = start
  field.selectionEnd = end
  await act(async () => {
    field.dispatchEvent(new window.MouseEvent('mouseup', { bubbles: true }))
  })
  assert.ok(page.container.querySelector('.rd-selection').textContent.includes('已选中'))
  const tiles = queryAll(page.container, '.rd-action')
  assert.ok(tiles.every(node => !node.disabled), 'actions enabled by the selection')

  click(tiles[1])
  await flush()
  const card = page.container.querySelector('.rd-card')
  assert.ok(card, 'result card rendered')
  assert.ok(card.textContent.includes('逻辑分析'))
  assert.ok(card.textContent.includes('本段是论证的落点'))
  assert.ok(card.querySelector('.rd-md strong'), 'bold markdown rendered')
  assert.ok(card.textContent.includes('2秒') || /\d+秒/u.test(card.textContent), 'elapsed time shown')
  assert.ok(queryAll(card, '.rd-mini').some(node => node.textContent === '复制'))
  assert.ok(queryAll(card, '.rd-mini').some(node => node.textContent === '删除'))

  click(queryAll(card, '.rd-mini').find(node => node.textContent === '删除'))
  await flush()
  assert.equal(page.container.querySelector('.rd-card'), null, 'card removed')
  await page.unmount()
})

await test('source result renders numbered references', async () => {
  const page = mount({
    fetchImpl: ndjsonFetch([
      { type: 'meta', action: 'source', provider: 'deepseek', model: 'deepseek-v4-flash' },
      { type: 'stage', stage: 'searching', detail: '科技金融 统计' },
      { type: 'sources', sources: [
        { url: 'https://www.pbc.gov.cn/a', title: '科技金融统计报告' },
        { url: 'https://example.com/b', title: '某银行年报' },
      ], note: '', claims: ['余额 42.6 亿元'] },
      { type: 'delta', text: '### 核查结论\n- **增速**：有据可查 —— 依据 [S1]。' },
      { type: 'done', chars: 10 },
    ]),
  })
  await page.render()
  click(queryAll(page.container, '.rd-empty-actions .rd-button')[0])
  await flush()
  const field = page.container.querySelector('textarea.rd-text')
  field.selectionStart = 0
  field.selectionEnd = 24
  await act(async () => { field.dispatchEvent(new window.MouseEvent('mouseup', { bubbles: true })) })
  click(queryAll(page.container, '.rd-action')[0])
  await flush()
  const sources = queryAll(page.container, '.rd-source')
  assert.equal(sources.length, 2)
  assert.equal(sources[0].getAttribute('href'), 'https://www.pbc.gov.cn/a')
  assert.equal(sources[0].getAttribute('target'), '_blank')
  assert.ok(sources[0].textContent.includes('[S1]'))
  assert.ok(sources[0].textContent.includes('pbc.gov.cn'))
  assert.ok(sources[1].textContent.includes('example.com'))
  await page.unmount()
})

await test('expansion offers a replacement and applies it to the manuscript', async () => {
  const page = mount({
    fetchImpl: ndjsonFetch([
      { type: 'meta', action: 'expand', provider: 'deepseek', model: 'deepseek-v4-flash' },
      { type: 'delta', text: '<<<TEXT\n从实践效果看，采用技术流评价的客户不良率更低。\n>>>\n' },
      { type: 'delta', text: '### 扩写说明\n- 补充了因果链。\n' },
      { type: 'done', chars: 30 },
    ]),
  })
  await page.render()
  click(queryAll(page.container, '.rd-empty-actions .rd-button')[0])
  await flush()
  const field = page.container.querySelector('textarea.rd-text')
  const start = field.value.indexOf('从实践效果看')
  const end = field.value.indexOf('同时也要看到')
  const original = field.value.slice(start, end)
  field.selectionStart = start
  field.selectionEnd = end
  await act(async () => { field.dispatchEvent(new window.MouseEvent('mouseup', { bubbles: true })) })
  click(queryAll(page.container, '.rd-action')[2])
  await flush()

  const rewrite = page.container.querySelector('.rd-rewrite-text')
  assert.ok(rewrite, 'replacement block rendered')
  assert.equal(rewrite.textContent, '从实践效果看，采用技术流评价的客户不良率更低。')
  assert.ok(!page.container.querySelector('.rd-md').textContent.includes('<<<TEXT'), 'markers stripped')

  const insert = queryAll(page.container, '.rd-mini').find(node => node.textContent === '替换原文')
  assert.ok(insert, 'replace action offered')
  click(insert)
  await flush()

  const updated = page.container.querySelector('textarea.rd-text')
  assert.ok(!updated.value.includes(original), 'original passage replaced')
  assert.ok(updated.value.includes('采用技术流评价的客户不良率更低'), 'replacement text present')
  await page.unmount()
})

await test('a failed run reports a mapped error and can retry', async () => {
  const fetchImpl = async () => new Response(JSON.stringify({ error: '未知的审核动作。' }), {
    status: 400,
    headers: { 'content-type': 'application/json' },
  })
  globalThis.fetch = fetchImpl
  window.fetch = fetchImpl
  const page = mount({ fetchImpl })
  await page.render()
  click(queryAll(page.container, '.rd-empty-actions .rd-button')[0])
  await flush()
  const field = page.container.querySelector('textarea.rd-text')
  field.selectionStart = 0
  field.selectionEnd = 10
  await act(async () => { field.dispatchEvent(new window.MouseEvent('mouseup', { bubbles: true })) })
  click(queryAll(page.container, '.rd-action')[1])
  await flush()
  const card = page.container.querySelector('.rd-card')
  assert.equal(card.getAttribute('data-status'), 'error')
  assert.ok(card.textContent.includes('未知的审核动作'), 'server detail kept')
  assert.ok(queryAll(card, '.rd-mini').some(node => node.textContent === '重试'))
  await page.unmount()
})

await test('keyboard shortcut starts a review for the selection', async () => {
  globalThis.fetch = ndjsonFetch([
    { type: 'meta', action: 'relate', provider: 'deepseek', model: 'deepseek-v4-flash' },
    { type: 'delta', text: '### 在全文中的位置\n承接上一段。' },
    { type: 'done', chars: 8 },
  ])
  window.fetch = globalThis.fetch
  const page = mount()
  await page.render()
  click(queryAll(page.container, '.rd-empty-actions .rd-button')[0])
  await flush()
  const field = page.container.querySelector('textarea.rd-text')
  field.selectionStart = 0
  field.selectionEnd = 30
  await act(async () => { field.dispatchEvent(new window.MouseEvent('mouseup', { bubbles: true })) })
  await act(async () => {
    field.dispatchEvent(new window.KeyboardEvent('keydown', { key: '4', metaKey: true, bubbles: true }))
  })
  await flush()
  const card = page.container.querySelector('.rd-card')
  assert.ok(card, 'shortcut started a run')
  assert.ok(card.textContent.includes('关联分析'))
  await page.unmount()
})

await test('timing: a running card shows its stage before text arrives', async () => {
  let release
  const gate = new Promise(resolve => { release = resolve })
  globalThis.fetch = async () => {
    const stream = new ReadableStream({
      async start(controller) {
        const encoder = new TextEncoder()
        controller.enqueue(encoder.encode(`${JSON.stringify({ type: 'meta', action: 'logic', provider: 'p', model: 'm' })}\n`))
        controller.enqueue(encoder.encode(`${JSON.stringify({ type: 'stage', stage: 'planning' })}\n`))
        await gate
        controller.enqueue(encoder.encode(`${JSON.stringify({ type: 'delta', text: '完成。' })}\n`))
        controller.enqueue(encoder.encode(`${JSON.stringify({ type: 'done', chars: 3 })}\n`))
        controller.close()
      },
    })
    return new Response(stream, { status: 200, headers: { 'content-type': 'application/x-ndjson' } })
  }
  window.fetch = globalThis.fetch
  const page = mount()
  await page.render()
  click(queryAll(page.container, '.rd-empty-actions .rd-button')[0])
  await flush()
  const field = page.container.querySelector('textarea.rd-text')
  field.selectionStart = 0
  field.selectionEnd = 12
  await act(async () => { field.dispatchEvent(new window.MouseEvent('mouseup', { bubbles: true })) })
  click(queryAll(page.container, '.rd-action')[1])
  await flush(2)
  const running = page.container.querySelector('.rd-card')
  assert.equal(running.getAttribute('data-status'), 'running')
  assert.ok(running.textContent.includes('正在规划检索') || running.textContent.includes('正在准备'), 'stage visible')
  assert.ok(queryAll(running, '.rd-mini').some(node => node.textContent === '停止'), 'stop offered')
  release()
  await flush()
  assert.equal(page.container.querySelector('.rd-card').getAttribute('data-status'), 'done')
  await page.unmount()
})

await test('a saved session restores the document and results', async () => {
  const page = mount({
    seed: {
      title: '恢复的文稿',
      text: '第一段。\n\n第二段。',
      runs: [{
        id: 'r1', action: 'logic', status: 'done', raw: '### 结构判断\n两段之间缺少过渡。',
        stage: '', model: 'deepseek/deepseek-v4-flash', sources: [], claims: [], sourceNote: '',
        startedAt: Date.now() - 4000, finishedAt: Date.now(), error: '',
        selection: { start: 0, end: 3, text: '第一段' },
      }],
    },
  })
  await page.render()
  assert.equal(page.container.querySelector('input.rd-doc-title').value, '恢复的文稿')
  assert.equal(page.container.querySelector('textarea.rd-text').value, '第一段。\n\n第二段。')
  const card = page.container.querySelector('.rd-card')
  assert.ok(card.textContent.includes('两段之间缺少过渡'))
  await page.unmount()
})


await test('imports a .docx package into the manuscript', async () => {
  const page = mount()
  await page.render()
  await importFile(page, '季度工作报告.docx', buildDocx(['第一段：这是从 Word 导入的文稿。', '第二段：科技金融实践。']))
  const preview = page.container.querySelector('.rd-preview')
  assert.ok(preview, 'imports render straight into the preview')
  assert.ok(preview.textContent.includes('第一段：这是从 Word 导入的文稿。'), preview.textContent)
  assert.equal(page.container.querySelector('input.rd-doc-title').value, '季度工作报告')
  await setMode(page, 'edit')
  const field = page.container.querySelector('textarea.rd-text')
  assert.ok(field.value.includes('\n\n第二段'), 'paragraphs kept apart in the source')
  await page.unmount()
})

await test('imports the text layer of a .pdf', async () => {
  const page = mount()
  await page.render()
  await importFile(page, 'report.pdf', buildPdf('This is the imported PDF passage.'))
  const preview = page.container.querySelector('.rd-preview')
  assert.ok(preview, 'the imported PDF opens in the preview')
  assert.ok(preview.textContent.includes('This is the imported PDF passage.'), preview.textContent)
  await setMode(page, 'edit')
  assert.ok(page.container.querySelector('textarea.rd-text').value.includes('This is the imported PDF passage.'))
  await page.unmount()
})

await test('a scanned-looking pdf reports an actionable message', async () => {
  const page = mount()
  await page.render()
  const empty = new TextEncoder().encode('%PDF-1.4\n1 0 obj << /Type /Catalog >> endobj\ntrailer << >>\n%%EOF\n')
  await importFile(page, 'scan.pdf', empty)
  assert.equal(page.container.querySelector('textarea.rd-text'), null, 'nothing imported')
  const toast = page.container.querySelector('.rd-toast')
  assert.ok(toast, 'a notice is shown')
  assert.ok(toast.textContent.includes('PDF'), toast.textContent)
  await page.unmount()
})

await test('a reviewed passage is marked in the manuscript and can be located', async () => {
  const page = mount({
    fetchImpl: ndjsonFetch([
      { type: 'meta', action: 'logic', provider: 'deepseek', model: 'deepseek-v4-flash' },
      { type: 'delta', text: '### 结构判断\n成立。' },
      { type: 'done', chars: 4 },
    ]),
  })
  await page.render()
  click(queryAll(page.container, '.rd-empty-actions .rd-button')[0])
  await flush()
  const field = page.container.querySelector('textarea.rd-text')
  const start = field.value.indexOf('科技金融的本质')
  const end = field.value.indexOf('从实践效果看')
  const passage = field.value.slice(start, end)
  field.selectionStart = start
  field.selectionEnd = end
  await act(async () => { field.dispatchEvent(new window.MouseEvent('mouseup', { bubbles: true })) })
  click(queryAll(page.container, '.rd-action')[1])
  await flush()

  const mark = page.container.querySelector('.rd-mark')
  assert.ok(mark, 'highlight layer carries a reviewed band')
  assert.equal(mark.textContent, passage)
  assert.ok(page.container.querySelector('.rd-doc-meta').textContent.includes('已审阅 1'))

  // Clear the selection, then locate the passage again from the result card.
  field.selectionStart = 0
  field.selectionEnd = 0
  await act(async () => { field.dispatchEvent(new window.MouseEvent('mouseup', { bubbles: true })) })
  assert.ok(page.container.querySelector('.rd-selection').textContent.includes('尚未选中文字'))

  const jump = queryAll(page.container, '.rd-mini').find(node => node.textContent === '定位到原文')
  assert.ok(jump, 'result card offers the locate action')
  click(jump)
  await flush()
  assert.equal(field.selectionStart, start, 'passage selected again')
  assert.equal(field.selectionEnd, end)
  assert.ok(page.container.querySelector('.rd-selection').textContent.includes('已选中'))
  await page.unmount()
})

await test('the empty selection card can select the first paragraph', async () => {
  const page = mount({ seed: { title: '文档', text: '\n\n第一段内容。\n\n第二段内容。', runs: [] } })
  await page.render()
  assert.ok(page.container.querySelector('.rd-selection').textContent.includes('尚未选中文字'))
  assert.ok(queryAll(page.container, '.rd-action').every(node => node.disabled))
  const button = queryAll(page.container, '.rd-selection .rd-mini')[0]
  assert.ok(button, 'first-paragraph action offered')
  assert.equal(button.textContent, '选中第一段')
  click(button)
  await flush()
  assert.ok(page.container.querySelector('.rd-selection').textContent.includes('已选中'))
  assert.equal(page.container.querySelector('.rd-selection .rd-selection-text').textContent, '第一段内容。')
  assert.ok(queryAll(page.container, '.rd-action').every(node => !node.disabled))
  await page.unmount()
})

await test('after an edit the stale mark and jump are dropped', async () => {
  const page = mount({
    seed: {
      title: '改动过的文稿',
      text: '第一段原文。\n\n第二段。',
      runs: [{
        id: 'r1', action: 'logic', status: 'done', raw: '### 结构判断\n可以。',
        stage: '', model: '', sources: [], claims: [], sourceNote: '', warning: '',
        startedAt: Date.now() - 3000, finishedAt: Date.now(), error: '',
        selection: { start: 3, end: 6, text: '原文。' },
      }],
    },
  })
  await page.render()
  assert.ok(page.container.querySelector('.rd-mark'), 'valid range is marked')
  const field = page.container.querySelector('textarea.rd-text')
  const setter = Object.getOwnPropertyDescriptor(window.HTMLTextAreaElement.prototype, 'value').set
  await act(async () => {
    setter.call(field, '完全不同的文字，原文已经改掉了。')
    field.dispatchEvent(new window.Event('input', { bubbles: true }))
  })
  await flush()
  assert.equal(page.container.querySelector('.rd-mark'), null, 'mark drops when the text no longer matches')
  await page.unmount()
})


await test('the body is three columns with their own headers', async () => {
  const page = mount({ seed: { title: '标题', text: '第一段内容。\n\n第二段内容。', runs: [] } })
  await page.render()
  const doc = page.container.querySelector('.rd-col-doc')
  const review = page.container.querySelector('.rd-col-review')
  const results = page.container.querySelector('.rd-col-results')
  assert.ok(doc && review && results, 'three columns rendered')
  assert.deepEqual(
    [...page.container.querySelectorAll('.rd-col-title')].map(node => node.textContent),
    ['文稿', '审查', '结果'],
  )
  assert.ok(doc.querySelector('.rd-editor'), 'manuscript in the first column')
  assert.ok(review.querySelector('.rd-action-group'), 'actions in the second column')
  assert.ok(results.querySelector('.rd-empty'), 'results in the third column')
  await page.unmount()
})

await test('preview renders markup and maps a preview selection back to source offsets', async () => {
  const source = '第一段**重点**内容。\n\n第二段内容。'
  const page = mount({ seed: { title: '预览', text: source, runs: [] } })
  await page.render()
  const toggle = queryAll(page.container, '.rd-segment')
  assert.equal(toggle[1].textContent, '预览')
  click(toggle[1])
  await flush()
  const preview = page.container.querySelector('.rd-md-preview')
  assert.ok(preview, 'rendered preview appears')
  assert.equal(page.container.querySelector('textarea.rd-text'), null, 'editor hidden in preview')
  assert.ok(preview.querySelector('strong'), 'bold markup rendered')
  assert.ok(!preview.textContent.includes('**'), 'markers are not shown')

  const spans = [...preview.querySelectorAll('[data-o]')]
  const last = spans[spans.length - 1]
  const expected = '第二段内容。'
  assert.equal(last.textContent, expected, 'last run is the second paragraph')
  assert.equal(Number(last.getAttribute('data-o')), source.indexOf(expected), 'run carries its source offset')
  const range = document.createRange()
  range.setStart(last.firstChild, 0)
  range.setEnd(last.firstChild, expected.length)
  const dom = window.getSelection()
  dom.removeAllRanges()
  dom.addRange(range)
  await act(async () => {
    preview.dispatchEvent(new window.MouseEvent('mouseup', { bubbles: true }))
  })
  const card = page.container.querySelector('.rd-selection')
  assert.ok(card.textContent.includes('已选中'), 'preview selection reaches the review column')
  assert.equal(card.querySelector('.rd-selection-text').textContent, expected)
  assert.ok(queryAll(page.container, '.rd-action').every(node => !node.disabled), 'actions enabled from the preview')
  await page.unmount()
})

await test('a typed requirement runs the custom action', async () => {
  const seen = []
  const page = mount({
    fetchImpl: appFetch({
      seen,
      frames: [
        { type: 'meta', action: 'custom', provider: 'deepseek', model: 'deepseek-v4-flash', reasoningEffort: '' },
        { type: 'delta', text: '改写后的文字。' },
        { type: 'done', chars: 6 },
      ],
    }),
  })
  await page.render()
  click(queryAll(page.container, '.rd-empty-actions .rd-button')[0])
  await flush()
  const field = page.container.querySelector('textarea.rd-text')
  field.selectionStart = 0
  field.selectionEnd = 20
  await act(async () => { field.dispatchEvent(new window.MouseEvent('mouseup', { bubbles: true })) })
  const input = page.container.querySelector('#rd-instruction')
  await act(async () => {
    const setter = Object.getOwnPropertyDescriptor(window.HTMLTextAreaElement.prototype, 'value').set
    setter.call(input, '改写成领导讲话口径，控制在 120 字内')
    input.dispatchEvent(new window.Event('input', { bubbles: true }))
  })
  await flush(2)
  const runButton = page.container.querySelector('.rd-run-custom')
  assert.ok(!runButton.disabled, 'custom run is enabled')
  click(runButton)
  await flush()

  assert.equal(seen.length, 1)
  assert.equal(seen[0].action, 'custom')
  assert.equal(seen[0].instruction, '改写成领导讲话口径，控制在 120 字内')
  const card = page.container.querySelector('.rd-card')
  assert.ok(card.textContent.includes('按要求处理'))
  assert.ok(card.textContent.includes('改写后的文字'))
  await page.unmount()
})

await test('the custom action falls back to the whole document', async () => {
  const seen = []
  const page = mount({
    seed: { title: '全文', text: '第一段。\n\n第二段。', runs: [] },
    fetchImpl: appFetch({ seen, frames: [{ type: 'done', chars: 0 }] }),
  })
  await page.render()
  const input = page.container.querySelector('#rd-instruction')
  await act(async () => {
    const setter = Object.getOwnPropertyDescriptor(window.HTMLTextAreaElement.prototype, 'value').set
    setter.call(input, '压缩成一句话')
    input.dispatchEvent(new window.Event('input', { bubbles: true }))
  })
  await flush(2)
  click(page.container.querySelector('.rd-run-custom'))
  await flush()
  assert.equal(seen[0].text, '第一段。\n\n第二段。')
  const card = page.container.querySelector('.rd-card')
  assert.ok(card.textContent.includes('全文'))
  assert.ok(![...card.querySelectorAll('.rd-mini')].some(node => node.textContent === '定位到原文'))
  assert.equal(page.container.querySelectorAll('.rd-mark').length, 0, 'no band for a whole-document run')
  await page.unmount()
})

await test('the model menu reads the Host catalog and applies a reasoning level', async () => {
  const seen = []
  const page = mount({
    seed: { title: '模型', text: '一段待审文字。', runs: [] },
    fetchImpl: appFetch({ seen, frames: [{ type: 'done', chars: 0 }] }),
  })
  await page.render()
  const button = page.container.querySelector('.rd-model-button')
  assert.ok(button.textContent.includes('模型'))
  click(button)
  await flush(4)

  const flash = queryAll(page.container, '.rd-menu-row').find(node => node.textContent.includes('DeepSeek V4 Flash'))
  assert.ok(flash, 'advertised model listed from the Remote catalog')
  click(flash)
  await flush(4)
  const highs = () => queryAll(page.container, '.rd-menu-row').filter(node => node.textContent.includes('高'))
  assert.ok(highs().length > 0, 'reasoning levels appear for the chosen model')

  // A model without levels must say so, not look like it is still loading.
  const chat = queryAll(page.container, '.rd-menu-row').find(node => node.textContent.includes('DeepSeek V4 Chat'))
  click(chat)
  await flush(2)
  const notes = [...page.container.querySelectorAll('.rd-menu-note')].map(node => node.textContent)
  assert.ok(notes.includes('该模型不提供推理等级'), `unexpected notes: ${notes.join(' | ')}`)
  assert.ok(!page.container.querySelector('.rd-model-button').textContent.includes('高'), 'no stale level on the button')

  click(queryAll(page.container, '.rd-menu-row').find(node => node.textContent.includes('DeepSeek V4 Flash')))
  await flush(4)
  click(highs()[0])
  await flush(2)
  assert.ok(page.container.querySelector('.rd-model-button').textContent.includes('高'), 'level shown on the button')

  const field = page.container.querySelector('textarea.rd-text')
  field.selectionStart = 0
  field.selectionEnd = 8
  await act(async () => { field.dispatchEvent(new window.MouseEvent('mouseup', { bubbles: true })) })
  click(queryAll(page.container, '.rd-action')[1])
  await flush()
  assert.equal(seen.length, 1)
  assert.equal(seen[0].provider, 'deepseek')
  assert.equal(seen[0].model, 'deepseek-v4-flash')
  assert.equal(seen[0].reasoningEffort, 'high')
  const card = page.container.querySelector('.rd-card')
  assert.ok(card.textContent.includes('高'), 'card records the level')
  await page.unmount()
})

await test('a failing catalog surfaces its reason instead of silence', async () => {
  const page = mount({
    seed: { title: '模型', text: '一段待审文字。', runs: [] },
    catalog: null,
  })
  await page.render()
  click(page.container.querySelector('.rd-model-button'))
  await flush(4)
  const note = page.container.querySelector('.rd-menu-note')
  assert.ok(note, 'menu shows a note')
  assert.ok(note.textContent.includes('catalog unavailable'), note.textContent)
  await page.unmount()
})

await test('an imported .docx keeps headings, bold, lists, and tables', async () => {
  const page = mount()
  await page.render()
  const body = [
    '<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>第一章 总则</w:t></w:r></w:p>',
    '<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>加粗结论</w:t></w:r><w:r><w:t>：普通说明。</w:t></w:r></w:p>',
    '<w:p><w:pPr><w:numPr/></w:pPr><w:r><w:t>第一条</w:t></w:r></w:p>',
    '<w:p><w:pPr><w:numPr/></w:pPr><w:r><w:t>第二条</w:t></w:r></w:p>',
    '<w:tbl><w:tr><w:tc><w:p><w:r><w:t>指标</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>数值</w:t></w:r></w:p></w:tc></w:tr>',
    '<w:tr><w:tc><w:p><w:r><w:t>余额</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>42.6</w:t></w:r></w:p></w:tc></w:tr></w:tbl>',
  ].join('')
  await importFile(page, '制度.docx', buildDocxBody(body))
  assert.ok(page.container.querySelector('.rd-preview'), 'document imported')
  await setMode(page, 'edit')
  const field = page.container.querySelector('textarea.rd-text')
  assert.ok(field, 'editor reachable after import')
  assert.ok(field.value.includes('# 第一章 总则'), field.value)
  assert.ok(field.value.includes('**加粗结论**：普通说明。'), field.value)
  assert.ok(field.value.includes('- 第一条\n- 第二条'), field.value)
  assert.ok(field.value.includes('| 指标 | 数值 |'), field.value)
  assert.ok(field.value.includes('| --- | --- |'), field.value)

  await setMode(page, 'preview')
  const preview = page.container.querySelector('.rd-md-preview')
  assert.equal(preview.querySelector('h1').textContent, '第一章 总则')
  assert.equal(preview.querySelector('strong').textContent, '加粗结论')
  assert.equal(preview.querySelectorAll('li').length, 2)
  assert.equal(preview.querySelectorAll('tbody tr').length, 1)
  await page.unmount()
})


/** One persisted run, complete enough for the result list. */
function seedRun(id, action, raw, extra = {}) {
  return {
    id,
    action,
    status: 'done',
    stage: '',
    raw,
    model: `deepseek/${extra.model ?? 'deepseek-v4-flash'}`,
    sources: extra.sources ?? [],
    claims: [],
    sourceNote: '',
    warning: '',
    error: '',
    effort: extra.effort ?? '',
    instruction: extra.instruction ?? '',
    wholeDocument: false,
    selection: { start: 0, end: 6, text: '第一段内容' },
    startedAt: Date.now() - 5000,
    finishedAt: Date.now() - 3000,
  }
}

await test('results fold to one open card and filter by tab', async () => {
  const page = mount({
    seed: {
      title: '结果',
      text: '第一段内容。\n\n第二段内容。',
      runs: [
        seedRun('r3', 'relate', '### 在全文中的位置\n承接上一段，与末段形成对照。'),
        seedRun('r2', 'logic', '### 结构判断\n本段论证成立。\n### 逻辑问题\n未发现明显逻辑漏洞。'),
        seedRun('r1', 'source', '### 核查结论\n- **增速**：有据可查 —— 依据 [S1]。'),
      ],
    },
  })
  await page.render()
  const cards = queryAll(page.container, '.rd-card')
  assert.equal(cards.length, 3)
  assert.equal(cards[0].getAttribute('aria-expanded') === null, true, 'header is the disclosure')
  assert.equal(cards[0].querySelector('.rd-card-body') !== null, true, 'newest card is open')
  assert.equal(cards[1].querySelector('.rd-card-body'), null, 'older card is folded')
  assert.ok(cards[1].querySelector('.rd-card-summary').textContent.includes('本段论证成立'), 'folded summary shows content')

  click(cards[1].querySelector('.rd-card-head'))
  await flush()
  assert.ok(queryAll(page.container, '.rd-card')[1].querySelector('.rd-card-body'), 'folded card expands on demand')

  const tabs = queryAll(page.container, '.rd-tab')
  assert.equal(tabs.length, 6, 'all + five actions')
  const logicTab = tabs.find(node => node.textContent.startsWith('逻辑分析'))
  click(logicTab)
  await flush()
  assert.equal(queryAll(page.container, '.rd-card').length, 1, 'tab filters the list')
  assert.ok(page.container.querySelector('.rd-card-title').textContent.includes('逻辑分析'))
  click(tabs[0])
  await flush()
  assert.equal(queryAll(page.container, '.rd-card').length, 3, 'all tab restores the list')
  await page.unmount()
})


await test('the imported original opens as a read-only layout sheet', async () => {
  const pagePayload = {
    pageCount: 2,
    dpi: 96,
    missingFonts: [],
    truncated: false,
    pages: [
      { data: 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==', width: 1, height: 1, page: 1 },
      { data: 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==', width: 1, height: 1, page: 2 },
    ],
  }
  const page = mount({ fetchImpl: appFetch({ pages: pagePayload }) })
  await page.render()
  assert.equal(page.container.querySelector('.rd-link-button'), null, 'no original before an import')
  await importFile(page, '制度.docx', buildDocx(['第一段：导入的文稿。']))
  const button = page.container.querySelector('.rd-link-button')
  assert.ok(button, 'original button appears after a document import')
  assert.equal(button.textContent, '原件')

  click(button)
  await flush(6)
  const sheet = page.container.querySelector('.rd-sheet')
  assert.ok(sheet, 'layout sheet opens')
  assert.equal(sheet.getAttribute('role'), 'dialog')
  const images = [...sheet.querySelectorAll('img')]
  assert.equal(images.length, 2)
  assert.ok(images[0].getAttribute('src').startsWith('data:image/png;base64,'))
  assert.ok(sheet.textContent.includes('2 页'), sheet.textContent)

  await act(async () => {
    document.dispatchEvent(new window.KeyboardEvent('keydown', { key: 'Escape', bubbles: true }))
  })
  assert.equal(page.container.querySelector('.rd-sheet'), null, 'Escape closes the sheet')
  await page.unmount()
})

await test('a missing render engine is reported in the sheet', async () => {
  const page = mount({ fetchImpl: appFetch({ pages: null }) })
  await page.render()
  await importFile(page, '制度.docx', buildDocx(['第一段。']))
  click(page.container.querySelector('.rd-link-button'))
  await flush(6)
  const sheet = page.container.querySelector('.rd-sheet')
  assert.ok(sheet.textContent.includes('渲染引擎'), sheet.textContent)
  await page.unmount()
})


await test('a rewrite from a custom requirement replaces the passage and can be undone', async () => {
  const page = mount({
    seed: { title: '改写', text: '第一段内容。\n\n第二段内容。', runs: [] },
    fetchImpl: appFetch({
      frames: [
        { type: 'meta', action: 'custom', provider: 'deepseek', model: 'deepseek-v4-flash', reasoningEffort: '' },
        { type: 'delta', text: '<<<TEXT\n改写后的段落。\n>>>\n### 修改说明\n- 更短。' },
        { type: 'done', chars: 20 },
      ],
    }),
  })
  await page.render()
  const field = page.container.querySelector('textarea.rd-text')
  field.selectionStart = 0
  field.selectionEnd = 6
  await act(async () => { field.dispatchEvent(new window.MouseEvent('mouseup', { bubbles: true })) })
  const input = page.container.querySelector('#rd-instruction')
  await act(async () => {
    const setter = Object.getOwnPropertyDescriptor(window.HTMLTextAreaElement.prototype, 'value').set
    setter.call(input, '压缩到一句')
    input.dispatchEvent(new window.Event('input', { bubbles: true }))
  })
  await flush(2)
  click(page.container.querySelector('.rd-run-custom'))
  await flush(6)

  const value = page.container.querySelector('textarea.rd-text').value
  assert.ok(value.includes('改写后的段落。'), value)
  assert.ok(!value.includes('第一段内容。'), 'original passage replaced')
  assert.ok(value.includes('第二段内容。'), 'rest of the document untouched')
  const undo = queryAll(page.container, '.rd-mini').find(node => node.textContent === '撤销替换')
  assert.ok(undo, 'undo action offered')
  click(undo)
  await flush()
  assert.ok(page.container.querySelector('textarea.rd-text').value.includes('第一段内容。'), 'undo restores the passage')
  await page.unmount()
})

await test('an existing result can be revised with an extra requirement', async () => {
  const seen = []
  const page = mount({
    seed: { title: '修订', text: '第一段内容。\n\n第二段内容。', runs: [] },
    fetchImpl: queueFetch([
      [
        { type: 'meta', action: 'custom', provider: 'deepseek', model: 'deepseek-v4-flash', reasoningEffort: '' },
        { type: 'delta', text: '<<<TEXT\n第一版改写。\n>>>\n### 修改说明\n- 第一版。' },
        { type: 'done', chars: 10 },
      ],
      [
        { type: 'meta', action: 'custom', provider: 'deepseek', model: 'deepseek-v4-flash', reasoningEffort: '' },
        { type: 'delta', text: '<<<TEXT\n第二版改写。\n>>>\n### 修改说明\n- 第二版。' },
        { type: 'done', chars: 10 },
      ],
    ], { seen }),
  })
  await page.render()
  const field = page.container.querySelector('textarea.rd-text')
  field.selectionStart = 0
  field.selectionEnd = 6
  await act(async () => { field.dispatchEvent(new window.MouseEvent('mouseup', { bubbles: true })) })
  const input = page.container.querySelector('#rd-instruction')
  await act(async () => {
    const setter = Object.getOwnPropertyDescriptor(window.HTMLTextAreaElement.prototype, 'value').set
    setter.call(input, '先改一版')
    input.dispatchEvent(new window.Event('input', { bubbles: true }))
  })
  await flush(2)
  click(page.container.querySelector('.rd-run-custom'))
  await flush(6)
  assert.ok(page.container.querySelector('textarea.rd-text').value.includes('第一版改写。'))

  const revise = queryAll(page.container, '.rd-mini').find(node => node.textContent === '修改')
  assert.ok(revise, 'revise action offered on a result')
  click(revise)
  await flush(2)
  const card = page.container.querySelector('.rd-card')
  const draft = card.querySelector('.rd-revise textarea')
  assert.ok(draft, 'revision input appears inside the card')
  await act(async () => {
    const setter = Object.getOwnPropertyDescriptor(window.HTMLTextAreaElement.prototype, 'value').set
    setter.call(draft, '再短一点')
    draft.dispatchEvent(new window.Event('input', { bubbles: true }))
  })
  await flush(2)
  click(card.querySelector('.rd-revise-run'))
  await flush(8)

  assert.equal(seen.length, 2)
  assert.equal(seen[1].revise, true)
  assert.equal(seen[1].instruction, '再短一点')
  assert.ok(seen[1].text.includes('第一版改写'), 'the revision starts from the last result')
  assert.ok(seen[1].original.includes('第一段内容'), 'the original passage travels with the revision')

  const cards = queryAll(page.container, '.rd-card')
  assert.equal(cards.length, 2, 'both results stay in the list')
  assert.ok(cards[0].textContent.includes('修订'), 'revision is marked')
  assert.ok(page.container.querySelector('textarea.rd-text').value.includes('第二版改写。'), 'latest rewrite replaces again')
  await page.unmount()
})


await test('the replace switch decides whether a plain answer touches the manuscript', async () => {
  const frames = [
    { type: 'meta', action: 'custom', provider: 'deepseek', model: 'deepseek-v4-flash', reasoningEffort: '' },
    { type: 'delta', text: '这段话有三个问题。\n\n说明：略。' },
    { type: 'done', chars: 10 },
  ]
  // A question is not a rewrite, so the answer stays in the results column.
  const runOnce = async (page) => {
    const field = page.container.querySelector('textarea.rd-text')
    field.selectionStart = 0
    field.selectionEnd = 6
    await act(async () => { field.dispatchEvent(new window.MouseEvent('mouseup', { bubbles: true })) })
    const input = page.container.querySelector('#rd-instruction')
    await act(async () => {
      const setter = Object.getOwnPropertyDescriptor(window.HTMLTextAreaElement.prototype, 'value').set
      setter.call(input, '帮我分析一下这段话')
      input.dispatchEvent(new window.Event('input', { bubbles: true }))
    })
    await flush(2)
    click(page.container.querySelector('.rd-run-custom'))
    await flush(8)
  }

  const page = mount({
    seed: { title: '不替换', text: '第一段内容。\n\n第二段内容。', runs: [] },
    fetchImpl: appFetch({ frames }),
  })
  await page.render()
  await runOnce(page)
  assert.ok(page.container.querySelector('textarea.rd-text').value.includes('第一段内容。'), 'an answer leaves the text alone')
  assert.ok(!queryAll(page.container, '.rd-mini').some(node => node.textContent === '撤销替换'))

  // With the switch on, the same answer is written into the passage.
  const page2 = mount({
    seed: { title: '替换', text: '第一段内容。\n\n第二段内容。', runs: [] },
    fetchImpl: appFetch({ frames }),
  })
  await page2.render()
  const check = page2.container.querySelector('.rd-check input')
  assert.ok(check, 'replace switch present')
  assert.equal(check.checked, false, 'starts off')
  await act(async () => { check.click() })
  await flush(2)
  await runOnce(page2)
  const value = page2.container.querySelector('textarea.rd-text').value
  assert.ok(value.includes('这段话有三个问题。'), value)
  assert.ok(!value.includes('说明：'), 'the explanation is not written into the manuscript')
  assert.ok(!value.includes('第一段内容。'), 'the passage was replaced')
  await page.unmount()
  await page2.unmount()
})

await test('a rewrite word in the requirement replaces without the switch', async () => {
  const page = mount({
    seed: { title: '改写', text: '第一段内容。\n\n第二段内容。', runs: [] },
    fetchImpl: appFetch({
      frames: [
        { type: 'meta', action: 'custom', provider: 'deepseek', model: 'deepseek-v4-flash', reasoningEffort: '' },
        { type: 'delta', text: '压缩后的句子。\n\n说明：更短。' },
        { type: 'done', chars: 8 },
      ],
    }),
  })
  await page.render()
  const field = page.container.querySelector('textarea.rd-text')
  field.selectionStart = 0
  field.selectionEnd = 6
  await act(async () => { field.dispatchEvent(new window.MouseEvent('mouseup', { bubbles: true })) })
  const input = page.container.querySelector('#rd-instruction')
  await act(async () => {
    const setter = Object.getOwnPropertyDescriptor(window.HTMLTextAreaElement.prototype, 'value').set
    setter.call(input, '压缩成一句')
    input.dispatchEvent(new window.Event('input', { bubbles: true }))
  })
  await flush(2)
  click(page.container.querySelector('.rd-run-custom'))
  await flush(8)
  const value = page.container.querySelector('textarea.rd-text').value
  assert.ok(value.includes('压缩后的句子。'), value)
  assert.ok(!value.includes('第一段内容。'))
  assert.ok(queryAll(page.container, '.rd-mini').some(node => node.textContent === '撤销替换'))
  await page.unmount()
})


await test('the three columns can be resized from the keyboard and reset', async () => {
  const page = mount({ seed: { title: '列宽', text: '第一段。', runs: [] } })
  await page.render()
  const body = page.container.querySelector('.rd-body')
  const separators = queryAll(page.container, '.rd-resizer')
  assert.equal(separators.length, 2, 'two separators between three columns')
  assert.equal(separators[0].getAttribute('role'), 'separator')
  assert.equal(separators[0].getAttribute('aria-orientation'), 'vertical')
  assert.equal(separators[0].getAttribute('aria-label'), '调整文稿与审查栏宽度')
  assert.equal(separators[1].getAttribute('aria-label'), '调整审查与结果栏宽度')
  assert.equal(separators[0].getAttribute('tabindex'), '0', 'keyboard reachable')

  await act(async () => {
    separators[0].dispatchEvent(new window.KeyboardEvent('keydown', { key: 'ArrowLeft', bubbles: true }))
  })
  await flush()
  const applied = body.getAttribute('style') ?? ''
  assert.ok(applied.includes('--rd-review-w'), `grid variables written: ${applied}`)
  const stored = JSON.parse(window.localStorage.getItem('dsh.review-desk.widths.v1'))
  assert.ok(stored.review >= 220 && stored.results >= 300, JSON.stringify(stored))

  await act(async () => {
    separators[0].dispatchEvent(new window.MouseEvent('dblclick', { bubbles: true }))
  })
  await flush()
  assert.equal(window.localStorage.getItem('dsh.review-desk.widths.v1'), null, 'reset clears storage')
  assert.ok(!(body.getAttribute('style') ?? '').includes('--rd-review-w'), 'reset restores the defaults')
  await page.unmount()
})

await test('saved column widths are restored on mount', async () => {
  const page = mount({
    seed: { title: '列宽', text: '第一段。', runs: [] },
    widths: { review: 340, results: 420 },
  })
  await page.render()
  const body = page.container.querySelector('.rd-body')
  assert.ok((body.getAttribute('style') ?? '').includes('--rd-review-w: 340px'), body.getAttribute('style'))
  assert.ok((body.getAttribute('style') ?? '').includes('--rd-results-w: 420px'), body.getAttribute('style'))
  assert.equal(page.container.querySelector('.rd-resizer-left').getAttribute('aria-valuenow'), '340')
  await page.unmount()
})


await test('a review package carries the manuscript and every result', async () => {
  const originalCreate = globalThis.URL.createObjectURL
  const originalRevoke = globalThis.URL.revokeObjectURL
  const originalClick = window.HTMLAnchorElement.prototype.click
  let captured = null
  globalThis.URL.createObjectURL = (blob) => { captured = blob; return 'blob:review' }
  globalThis.URL.revokeObjectURL = () => {}
  window.HTMLAnchorElement.prototype.click = function click() {}

  const page = mount({
    seed: {
      title: '导出测试',
      text: '第一段内容。\n\n第二段内容。',
      runs: [seedRun('r1', 'logic', '### 结构判断\n成立。')],
    },
  })
  await page.render()
  const exportMenu = queryAll(page.container, '.rd-toolbar .rd-button').find(node => node.textContent.startsWith('导出'))
  assert.ok(exportMenu, 'export menu exists')
  click(exportMenu)
  await flush()
  const packageItem = queryAll(page.container, '.rd-menu-row').find(node => node.textContent.includes('导出审核包'))
  assert.ok(packageItem, 'package export offered')
  click(packageItem)
  await flush(2)

  assert.ok(captured, 'a blob was produced')
  const text = await captured.text()
  const data = JSON.parse(text)
  assert.equal(data.format, 'dsh.review-desk.package')
  assert.equal(data.version, 1)
  assert.equal(data.document.title, '导出测试')
  assert.ok(data.document.text.includes('第一段内容。'))
  assert.equal(data.runs.length, 1)
  assert.equal(data.runs[0].action, 'logic')
  assert.ok(data.exportedAt.length > 10)

  globalThis.URL.createObjectURL = originalCreate
  globalThis.URL.revokeObjectURL = originalRevoke
  window.HTMLAnchorElement.prototype.click = originalClick
  await page.unmount()
})

await test('importing a package replaces the session after confirmation', async () => {
  const page = mount({ seed: { title: '旧文稿', text: '旧内容。', runs: [] } })
  await page.render()
  const input = queryAll(page.container, 'input[type="file"]')[1]
  assert.ok(input, 'package picker exists')
  const payload = {
    format: 'dsh.review-desk.package',
    version: 1,
    exportedAt: new Date().toISOString(),
    document: { title: '从另一台电脑带来', text: '新文稿第一段。\n\n新文稿第二段。' },
    runs: [seedRun('p1', 'source', '### 核查结论\n- 有据可查 [S1]。')],
    session: { instruction: '保持书面语气' },
  }
  Object.defineProperty(input, 'files', { value: [new File([JSON.stringify(payload)], 'a.review-desk.json')], configurable: true })
  await act(async () => { input.dispatchEvent(new window.Event('change', { bubbles: true })) })
  await flush(6)

  const dialog = page.container.querySelector('.rd-dialog')
  assert.ok(dialog, 'confirmation dialog appears')
  assert.ok(dialog.textContent.includes('1 条审核记录'), dialog.textContent)
  assert.ok(page.container.querySelector('textarea.rd-text').value.includes('旧内容。'), 'not applied before confirming')

  const confirm = queryAll(dialog, '.rd-button').find(node => node.textContent === '替换并导入')
  click(confirm)
  await flush(6)

  assert.equal(page.container.querySelector('input.rd-doc-title').value, '从另一台电脑带来')
  assert.ok(page.container.querySelector('textarea.rd-text').value.includes('新文稿第一段。'))
  assert.equal(queryAll(page.container, '.rd-card').length, 1, 'results came with the package')
  assert.equal(page.container.querySelector('#rd-instruction').value, '保持书面语气')
  assert.equal(page.container.querySelector('.rd-dialog'), null, 'dialog closes')
  await act(async () => { await new Promise(resolve => setTimeout(resolve, 520)) })
  const saved = JSON.parse(window.localStorage.getItem('dsh.review-desk.v1'))
  assert.ok(saved.text.includes('新文稿第一段。'), 'imported session is persisted')
  await page.unmount()
})

await test('an invalid package is refused without touching the document', async () => {
  const page = mount({ seed: { title: '保留', text: '保留的内容。', runs: [] } })
  await page.render()
  const input = queryAll(page.container, 'input[type="file"]')[1]
  Object.defineProperty(input, 'files', { value: [new File(['{"nope":1}'], 'bad.json')], configurable: true })
  await act(async () => { input.dispatchEvent(new window.Event('change', { bubbles: true })) })
  await flush(6)
  assert.equal(page.container.querySelector('.rd-dialog'), null, 'no dialog for a bad file')
  assert.ok(page.container.querySelector('textarea.rd-text').value.includes('保留的内容。'))
  const toast = page.container.querySelector('.rd-toast')
  assert.ok(toast.textContent.includes('不是有效的审核包'), toast.textContent)
  await page.unmount()
})

if (failures > 0) {
  console.error(`\n${failures} client check(s) failed`)
  process.exitCode = 1
} else {
  console.log('\nall client checks passed')
}
