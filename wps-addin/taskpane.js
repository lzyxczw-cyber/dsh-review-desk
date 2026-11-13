/**
 * Review Desk · WPS 任务窗格（v0.2）。
 *
 * 读取 WPS 选区 → 调本机 DSH 的 review-desk 桥（默认 http://127.0.0.1:47831）
 * → 流式展示审核结果 → 以「批注」或「修订替换」写回文档。
 * 结果历史保存在本机 localStorage；支持筛选、折叠、追加要求、撤销替换、定位回原文。
 */
(function () {
  'use strict'

  var DEFAULT_BRIDGE = 'http://127.0.0.1:47831'
  var BRIDGE_KEY = 'wps-review-desk.bridge'
  var SETTINGS_KEY = 'wps-review-desk.settings.v1'
  var HISTORY_KEY = 'wps-review-desk.history.v1'
  var HISTORY_LIMIT = 20
  var HISTORY_BYTES = 480000
  var VERIFY = /[?&]verify=1(?:&|$)/.test(window.location.search)
  var ACTION_LABELS = { source: '来源查证', logic: '逻辑分析', expand: '扩写', relate: '关联分析', custom: '自由要求' }
  var ACTION_ORDER = ['source', 'logic', 'expand', 'relate', 'custom']
  var STAGE_LABELS = { planning: '生成检索式…', searching: '检索资料…', writing: '撰写中…' }
  var EFFORT_OPTIONS = [
    { id: '', name: '（默认）' },
    { id: 'low', name: '低' },
    { id: 'medium', name: '中' },
    { id: 'high', name: '高' },
    { id: 'max', name: '最高' },
  ]
  var CAP_OPTIONS = [
    { value: 0, name: '（默认：按操作自动）' },
    { value: 6000, name: '6000 tokens' },
    { value: 8000, name: '8000 tokens' },
    { value: 12000, name: '12000 tokens' },
    { value: 16000, name: '16000 tokens（上限）' },
  ]

  var state = {
    bridge: DEFAULT_BRIDGE,
    settings: { provider: '', model: '', reasoningEffort: '', maxTokens: 0 },
    catalog: null,
    docName: '',
    selection: null,
    running: false,
    health: 'unknown',
    results: [],
    filter: 'all',
    seq: 0,
    historyTimer: 0,
    toastTimer: 0,
  }

  var refs = {}
  var selfTest = { started: false }

  /* ------------------------------------------------------------------ *
   * WPS 文档访问
   * ------------------------------------------------------------------ */

  function getApp() {
    try {
      if (window.Application) return window.Application
      if (typeof wps !== 'undefined' && wps.WpsApplication) return wps.WpsApplication()
    } catch (error) { /* ignore */ }
    return null
  }

  function getDoc() {
    try {
      var app = getApp()
      return app ? app.ActiveDocument : null
    } catch (error) {
      return null
    }
  }

  function readSelection() {
    var app = getApp()
    var doc = getDoc()
    state.docName = doc ? String(doc.Name || '') : ''
    if (!app || !doc) {
      state.selection = null
      return
    }
    try {
      var selection = app.Selection
      var text = String(selection.Text || '')
      state.selection = text.trim() === '' ? null : {
        start: Number(selection.Start),
        end: Number(selection.End),
        text: text,
      }
    } catch (error) {
      state.selection = null
    }
  }

  /* ------------------------------------------------------------------ *
   * 设置
   * ------------------------------------------------------------------ */

  function readSettings() {
    var settings = { provider: '', model: '', reasoningEffort: '', maxTokens: 0 }
    try {
      var raw = localStorage.getItem(SETTINGS_KEY)
      if (raw) {
        var parsed = JSON.parse(raw)
        if (parsed && typeof parsed === 'object') {
          if (typeof parsed.provider === 'string') settings.provider = parsed.provider
          if (typeof parsed.model === 'string') settings.model = parsed.model
          if (typeof parsed.reasoningEffort === 'string') settings.reasoningEffort = parsed.reasoningEffort
          if (Number.isFinite(Number(parsed.maxTokens)) && Number(parsed.maxTokens) > 0) settings.maxTokens = Number(parsed.maxTokens)
        }
      }
    } catch (error) { /* ignore */ }
    return settings
  }

  function persistSettings() {
    try { localStorage.setItem(SETTINGS_KEY, JSON.stringify(state.settings)) } catch (error) { /* ignore */ }
    renderSettingsSummary()
  }

  function applyCatalog(catalog) {
    state.catalog = catalog && typeof catalog === 'object' ? catalog : null
    renderModelOptions()
    renderEffortOptions()
    renderSettingsSummary()
  }

  function renderModelOptions() {
    if (!refs.settingModel) return
    var select = refs.settingModel
    select.innerHTML = ''
    var placeholder = document.createElement('option')
    placeholder.value = ''
    placeholder.textContent = '（默认：DSH 设置）'
    select.appendChild(placeholder)
    var catalog = state.catalog
    var providers = catalog && Array.isArray(catalog.providers) ? catalog.providers : []
    for (var provider of providers) {
      var group = document.createElement('optgroup')
      group.label = String(provider.name || provider.id)
      for (var model of (Array.isArray(provider.models) ? provider.models : [])) {
        var option = document.createElement('option')
        option.value = `${provider.id}::${model.id}`
        option.textContent = String(model.name || model.id)
        group.appendChild(option)
      }
      if (group.children.length > 0) select.appendChild(group)
    }
    var wanted = state.settings.provider && state.settings.model ? `${state.settings.provider}::${state.settings.model}` : ''
    select.value = wanted
    if (select.value !== wanted) select.value = ''
  }

  function renderEffortOptions() {
    if (!refs.settingEffort) return
    var select = refs.settingEffort
    var current = state.settings.reasoningEffort
    var options = EFFORT_OPTIONS.slice()
    var model = findSelectedModel()
    if (model && model.reasoning && Array.isArray(model.reasoning.efforts) && model.reasoning.efforts.length > 0) {
      options = [{ id: '', name: '（默认）' }].concat(model.reasoning.efforts.map(function (effort) {
        return { id: String(effort.id), name: String(effort.name || effort.id) }
      }))
    }
    select.innerHTML = ''
    for (var option of options) {
      var node = document.createElement('option')
      node.value = option.id
      node.textContent = option.name
      select.appendChild(node)
    }
    select.value = current
    if (select.value !== current) {
      state.settings.reasoningEffort = ''
      select.value = ''
    }
  }

  function findSelectedModel() {
    var catalog = state.catalog
    if (!catalog || !Array.isArray(catalog.providers)) return null
    if (!state.settings.provider || !state.settings.model) return null
    for (var provider of catalog.providers) {
      if (provider.id !== state.settings.provider) continue
      for (var model of (provider.models || [])) {
        if (model.id === state.settings.model) return model
      }
    }
    return null
  }

  function renderCapOptions() {
    if (!refs.settingCap) return
    var select = refs.settingCap
    select.innerHTML = ''
    for (var option of CAP_OPTIONS) {
      var node = document.createElement('option')
      node.value = String(option.value)
      node.textContent = option.name
      select.appendChild(node)
    }
    select.value = String(state.settings.maxTokens || 0)
  }

  function renderSettingsSummary() {
    if (!refs.settingsSummary) return
    var settings = state.settings
    var parts = []
    if (settings.provider && settings.model) {
      var model = findSelectedModel()
      parts.push(model ? `${settings.provider} / ${model.name || settings.model}` : `${settings.provider} / ${settings.model}`)
    }
    if (settings.reasoningEffort) {
      var effort = EFFORT_OPTIONS.concat((findSelectedModel()?.reasoning?.efforts || []).map(function (item) {
        return { id: String(item.id), name: String(item.name || item.id) }
      })).filter(function (item) { return item.id === settings.reasoningEffort })
      parts.push(`推理 ${effort.length > 0 ? effort[0].name : settings.reasoningEffort}`)
    }
    if (settings.maxTokens > 0) parts.push(`上限 ${settings.maxTokens}`)
    refs.settingsSummary.textContent = parts.length === 0 ? '（默认）' : `· ${parts.join(' · ')}`
  }

  /* ------------------------------------------------------------------ *
   * 后端
   * ------------------------------------------------------------------ */

  function normalizeBridge(value) {
    var url = String(value || '').trim()
    if (url === '') return DEFAULT_BRIDGE
    if (!/^https?:\/\//u.test(url)) url = `http://${url}`
    return url.replace(/\/+$/u, '')
  }

  function checkBridge() {
    setHealth('checking')
    var controller = new AbortController()
    var timer = setTimeout(function () { controller.abort() }, 3000)
    fetch(`${state.bridge}/health`, { signal: controller.signal })
      .then(function (response) { return response.json() })
      .then(function (info) {
        clearTimeout(timer)
        if (info && info.ok === true) acceptHealth()
        else setHealth('down')
      })
      .catch(function () {
        clearTimeout(timer)
        setHealth('down')
      })
  }

  function acceptHealth() {
    setHealth('ready')
    report({ kind: 'wps-health', ok: true, bridge: state.bridge, at: new Date().toISOString() })
    fetchModels()
    maybeSelfTest()
  }

  function setHealth(value) {
    state.health = value
    if (refs.connDot) {
      refs.connDot.setAttribute('data-state', value)
      refs.connDot.title = value === 'ready'
        ? '后端已连接'
        : value === 'checking' ? '正在检测后端…' : '后端不可用：请确认 DSH 正在运行'
    }
    if (refs.bridgeHint) {
      refs.bridgeHint.textContent = value === 'ready'
        ? `已连接 ${state.bridge}`
        : value === 'checking'
          ? '正在检测后端…'
          : `连不上 ${state.bridge}：请确认 DSH 正在运行，且 review-desk 处于 wps 分支并已重新加载。`
    }
  }

  function fetchModels() {
    if (!refs.settingModel) return
    fetch(`${state.bridge}/models`)
      .then(function (response) { return response.json() })
      .then(function (catalog) {
        applyCatalog(catalog)
        if (refs.settingsHint) {
          var defaultRoute = catalog && catalog.default
          refs.settingsHint.textContent = defaultRoute && defaultRoute.provider
            ? `默认：${defaultRoute.provider} / ${defaultRoute.model}${defaultRoute.reasoningEffort ? ` · 推理 ${defaultRoute.reasoningEffort}` : ''}。输出上限越高越慢、费用越高。`
            : '默认使用 DSH 的模型与推理设置；输出上限越高越慢、费用越高。'
        }
        report({ kind: 'wps-models', ok: true, providers: catalog && Array.isArray(catalog.providers) ? catalog.providers.length : 0 })
      })
      .catch(function () {
        if (refs.settingsHint) refs.settingsHint.textContent = '读取模型列表失败：只提供默认设置。'
        report({ kind: 'wps-models', ok: false })
      })
  }

  /* ------------------------------------------------------------------ *
   * 历史结果
   * ------------------------------------------------------------------ */

  function serializeResult(result) {
    return {
      id: result.id,
      action: result.action,
      instruction: result.instruction,
      createdAt: result.createdAt,
      status: result.status,
      text: result.text,
      meta: result.meta,
      sources: result.sources,
      note: result.note,
      warning: result.warning,
      error: result.error,
      selection: result.selection,
      parentId: result.parentId,
      original: result.original,
      maxTokens: result.maxTokens,
      applied: result.applied,
      collapsed: result.collapsed === true,
    }
  }

  function saveHistory() {
    if (state.historyTimer) clearTimeout(state.historyTimer)
    state.historyTimer = setTimeout(function () {
      state.historyTimer = 0
      try {
        var results = state.results.slice(0, HISTORY_LIMIT)
        var payload = JSON.stringify({ version: 1, results: results.map(serializeResult) })
        if (payload.length > HISTORY_BYTES) {
          var trimmed = JSON.parse(payload)
          trimmed.results = trimmed.results.slice(0, Math.max(3, Math.floor(results.length / 2)))
          payload = JSON.stringify(trimmed)
        }
        localStorage.setItem(HISTORY_KEY, payload)
      } catch (error) { /* history is best-effort */ }
    }, 300)
  }

  function loadHistory() {
    try {
      var raw = localStorage.getItem(HISTORY_KEY)
      if (!raw) return []
      var parsed = JSON.parse(raw)
      var list = parsed && Array.isArray(parsed.results) ? parsed.results : []
      var restored = []
      for (var item of list) {
        if (item === null || typeof item !== 'object') continue
        var result = {
          id: Number(item.id) || 0,
          action: ACTION_LABELS[item.action] ? item.action : 'custom',
          instruction: typeof item.instruction === 'string' ? item.instruction : '',
          createdAt: Number(item.createdAt) || Date.now(),
          status: item.status === 'error' || item.status === 'done' ? item.status : 'error',
          text: typeof item.text === 'string' ? item.text : '',
          stage: '',
          meta: item.meta || null,
          sources: Array.isArray(item.sources) ? item.sources : [],
          note: typeof item.note === 'string' ? item.note : '',
          warning: typeof item.warning === 'string' ? item.warning : '',
          error: item.status === 'running' ? '页面已刷新，本次结果中断' : (typeof item.error === 'string' ? item.error : ''),
          selection: item.selection && typeof item.selection === 'object'
            ? { start: Number(item.selection.start) || 0, end: Number(item.selection.end) || 0, text: String(item.selection.text || '') }
            : null,
          parentId: Number(item.parentId) || null,
          original: typeof item.original === 'string' ? item.original : '',
          maxTokens: Number(item.maxTokens) || 0,
          applied: item.applied && typeof item.applied === 'object' ? item.applied : null,
          collapsed: item.collapsed === true,
          dom: null,
        }
        restored.push(result)
        if (result.id > state.seq) state.seq = result.id
      }
      return restored.slice(0, HISTORY_LIMIT)
    } catch (error) {
      return []
    }
  }

  function clearHistory() {
    if (state.running) { toast('审核进行中，稍后再清空'); return }
    if (!window.confirm('清空全部历史结果？此操作只影响面板，不改文档。')) return
    state.results = []
    refs.results.innerHTML = ''
    renderEmpty()
    renderFilters()
    saveHistory()
    toast('已清空历史')
  }

  /* ------------------------------------------------------------------ *
   * 选区面板
   * ------------------------------------------------------------------ */

  function renderSelection() {
    var selection = state.selection
    if (refs.docName) refs.docName.textContent = state.docName
    if (refs.selectionPreview) {
      if (selection === null) {
        refs.selectionPreview.textContent = '尚未选中文字（在文档中拖动选中一段后再操作）'
        refs.selectionPreview.setAttribute('data-empty', '')
      } else {
        var preview = selection.text.length > 220 ? `${selection.text.slice(0, 220)}…` : selection.text
        refs.selectionPreview.textContent = preview
        refs.selectionPreview.removeAttribute('data-empty')
      }
    }
    if (refs.selectionMeta) {
      refs.selectionMeta.textContent = selection === null
        ? ''
        : `${selection.text.length} 字 · 位置 ${selection.start}–${selection.end}`
    }
    var disabled = selection === null || state.running
    for (var button of refs.actionButtons) button.disabled = disabled
    if (refs.btnCustom) refs.btnCustom.disabled = disabled
  }

  /* ------------------------------------------------------------------ *
   * 发起审核
   * ------------------------------------------------------------------ */

  function buildPayload(action, text, options) {
    var payload = { action: action, text: text }
    if (options.manuscript !== undefined) payload.manuscript = options.manuscript
    if (options.title) payload.title = options.title
    if (options.instruction) payload.instruction = options.instruction
    if (options.revise === true) {
      payload.revise = true
      if (options.original) payload.original = options.original
    }
    var settings = state.settings
    if (settings.provider && settings.model) {
      payload.provider = settings.provider
      payload.model = settings.model
    }
    if (settings.reasoningEffort) payload.reasoningEffort = settings.reasoningEffort
    var maxTokens = options.maxTokens !== undefined ? options.maxTokens : settings.maxTokens
    if (maxTokens > 0) payload.maxTokens = maxTokens
    return payload
  }

  function runAnalysis(action, options) {
    options = options || {}
    if (state.running) {
      toast('上一条审核还在进行')
      return null
    }
    var selection = options.selection || state.selection
    if (!selection) {
      toast('请先在文档中选中一段文字')
      return null
    }
    var doc = getDoc()
    if (!doc) {
      toast('没有打开的文档')
      return null
    }
    var manuscript = ''
    try {
      manuscript = options.manuscript !== undefined ? options.manuscript : String(doc.Content.Text || '').slice(0, 12000)
    } catch (error) { /* ignore */ }
    var title = state.docName.replace(/\.[^.]+$/u, '')
    var text = options.text !== undefined ? options.text : selection.text
    var result = {
      id: ++state.seq,
      action: action,
      instruction: options.instruction || '',
      createdAt: Date.now(),
      status: 'running',
      text: '',
      stage: '',
      meta: null,
      sources: [],
      note: '',
      warning: '',
      error: '',
      selection: selection,
      parentId: options.parentId || null,
      original: options.original || '',
      maxTokens: options.maxTokens !== undefined ? options.maxTokens : state.settings.maxTokens,
      applied: null,
      collapsed: false,
      dom: null,
    }
    state.results.unshift(result)
    collapseAllExcept(result)
    insertResultCard(result)
    renderFilters()
    saveHistory()
    setRunning(true)

    var payload = buildPayload(action, text, {
      manuscript: manuscript,
      title: title,
      instruction: options.instruction || '',
      revise: options.revise === true,
      original: options.original || '',
      maxTokens: options.maxTokens,
    })
    report({
      kind: 'wps-run',
      action: action,
      chars: selection.text.length,
      provider: payload.provider || '',
      model: payload.model || '',
      reasoningEffort: payload.reasoningEffort || '',
      maxTokens: payload.maxTokens || 0,
      at: new Date().toISOString(),
    })

    var controller = new AbortController()
    var timer = setTimeout(function () { controller.abort() }, 185000)

    fetch(`${state.bridge}/analyze`, {
      method: 'POST',
      headers: { 'content-type': 'application/json' },
      body: JSON.stringify(payload),
      signal: controller.signal,
    }).then(function (response) {
      if (!response.ok) {
        return response.json().then(function (info) {
          throw new Error(info && info.error ? info.error : `HTTP ${response.status}`)
        }).catch(function (error) {
          if (error instanceof Error && error.message.indexOf('HTTP ') === 0) throw error
          throw new Error(`HTTP ${response.status}`)
        })
      }
      return consumeStream(result, response)
    }).catch(function (error) {
      clearTimeout(timer)
      var message = error && error.name === 'AbortError'
        ? '请求超时或已取消'
        : (error && error.message) || String(error)
      if (result.status === 'running') {
        result.status = 'error'
        result.error = message
      }
      updateResultCard(result)
    }).then(function () {
      clearTimeout(timer)
      if (result.status === 'running') result.status = 'done'
      setRunning(false)
      updateResultCard(result)
      renderFilters()
      saveHistory()
      report({
        kind: 'wps-result',
        action: result.action,
        status: result.status,
        chars: result.text.length,
        warning: result.warning,
        error: result.error,
      })
      if (typeof options.onSettled === 'function') options.onSettled(result)
    })
    return result
  }

  function consumeStream(result, response) {
    if (!response.body || typeof response.body.getReader !== 'function') {
      return response.text().then(function (raw) {
        raw.split('\n').forEach(function (line) { if (line.trim() !== '') handleFrame(result, line) })
      })
    }
    var reader = response.body.getReader()
    var decoder = new TextDecoder()
    var buffer = ''
    function pump() {
      return reader.read().then(function (chunk) {
        if (chunk.done) {
          if (buffer.trim() !== '') handleFrame(result, buffer)
          return
        }
        buffer += decoder.decode(chunk.value, { stream: true })
        var index
        while ((index = buffer.indexOf('\n')) >= 0) {
          var line = buffer.slice(0, index)
          buffer = buffer.slice(index + 1)
          if (line.trim() !== '') handleFrame(result, line)
        }
        return pump()
      })
    }
    return pump()
  }

  function handleFrame(result, line) {
    var frame
    try { frame = JSON.parse(line) } catch (error) { return }
    switch (frame.type) {
      case 'meta':
        result.meta = {
          provider: frame.provider,
          model: frame.model,
          reasoningEffort: frame.reasoningEffort,
        }
        break
      case 'stage':
        result.stage = STAGE_LABELS[frame.stage] || '处理中…'
        if (frame.stage === 'searching' && frame.detail) result.stage = `检索：${frame.detail}`
        break
      case 'sources':
        result.sources = Array.isArray(frame.sources) ? frame.sources : []
        result.note = typeof frame.note === 'string' ? frame.note : ''
        break
      case 'delta':
        result.text += typeof frame.text === 'string' ? frame.text : ''
        break
      case 'warning':
        result.warning = typeof frame.message === 'string' ? frame.message : (typeof frame.warning === 'string' ? frame.warning : '')
        break
      case 'done':
        result.status = 'done'
        break
      case 'error':
        result.status = 'error'
        result.error = typeof frame.message === 'string' ? frame.message : '审核失败'
        break
      default:
        return
    }
    updateResultCard(result)
  }

  function setRunning(running) {
    state.running = running
    renderSelection()
  }

  /* ------------------------------------------------------------------ *
   * 结果卡片
   * ------------------------------------------------------------------ */

  function timeLabel(timestamp) {
    try {
      var date = new Date(timestamp)
      var now = new Date()
      var sameDay = date.toDateString() === now.toDateString()
      var hour = String(date.getHours()).padStart(2, '0')
      var minute = String(date.getMinutes()).padStart(2, '0')
      return sameDay ? `${hour}:${minute}` : `${date.getMonth() + 1}/${date.getDate()} ${hour}:${minute}`
    } catch (error) {
      return ''
    }
  }

  function parentLabel(result) {
    if (!result.parentId) return ''
    var parent = state.results.find(function (item) { return item.id === result.parentId })
    return parent ? `基于 #${parent.id} 的追加要求` : `基于 #${result.parentId} 的追加要求`
  }

  function createButton(label, handler, options) {
    var button = document.createElement('button')
    button.type = 'button'
    button.textContent = label
    if (options && options.emphasis) button.setAttribute('data-emphasis', options.emphasis)
    if (options && options.title) button.title = options.title
    button.addEventListener('click', handler)
    return button
  }

  function insertResultCard(result) {
    var container = refs.results
    if (!container) return
    removeEmpty()
    var article = document.createElement('article')
    article.className = 'rd-result'
    article.setAttribute('data-id', String(result.id))
    article.setAttribute('data-action', result.action)
    article.setAttribute('data-status', result.status)

    var head = document.createElement('header')
    head.className = 'rd-result-head'
    head.title = '点击展开 / 折叠'
    var badge = document.createElement('span')
    badge.className = 'rd-badge'
    badge.textContent = ACTION_LABELS[result.action] || result.action
    var provider = document.createElement('span')
    provider.className = 'rd-provider'
    var time = document.createElement('span')
    time.className = 'rd-time'
    time.textContent = timeLabel(result.createdAt)
    var status = document.createElement('span')
    status.className = 'rd-state'
    status.textContent = '进行中…'
    head.appendChild(badge)
    head.appendChild(provider)
    head.appendChild(time)
    head.appendChild(status)
    head.addEventListener('click', function () {
      result.collapsed = result.collapsed !== true
      article.setAttribute('data-collapsed', result.collapsed ? '1' : '0')
      saveHistory()
    })

    var body = document.createElement('div')
    body.className = 'rd-result-body'
    var parent = document.createElement('p')
    parent.className = 'rd-parent'
    parent.textContent = parentLabel(result)
    var text = document.createElement('div')
    text.className = 'rd-result-text'
    var warning = document.createElement('p')
    warning.className = 'rd-warn'
    var retry = document.createElement('div')
    retry.className = 'rd-retry'
    var refsList = document.createElement('ul')
    refsList.className = 'rd-refs'
    var revise = document.createElement('div')
    revise.className = 'rd-revise'
    var reviseInput = document.createElement('input')
    reviseInput.type = 'text'
    reviseInput.maxLength = 500
    reviseInput.placeholder = '继续修改要求，例如：再压缩到 40 字、换成更口语的表达'
    var reviseButton = createButton('执行', function () { submitRevise(result) })
    revise.appendChild(reviseInput)
    revise.appendChild(reviseButton)
    reviseInput.addEventListener('keydown', function (event) {
      if (event.key === 'Enter') {
        event.preventDefault()
        submitRevise(result)
      }
    })
    body.appendChild(parent)
    body.appendChild(text)
    body.appendChild(warning)
    body.appendChild(retry)
    body.appendChild(refsList)
    body.appendChild(revise)

    var tools = document.createElement('footer')
    tools.className = 'rd-tools'
    var locate = createButton('定位回原文', function () { locateResult(result) })
    var comment = createButton('插入批注', function () { insertComment(result) })
    var replace = createButton('替换原文（修订）', function () { replaceOriginal(result) }, { emphasis: 'replace' })
    var undo = createButton('撤销替换', function () { undoReplace(result) })
    var reviseToggle = createButton('追加要求', function () {
      var open = revise.getAttribute('data-open') === '1'
      revise.setAttribute('data-open', open ? '0' : '1')
      if (!open) reviseInput.focus()
    })
    var copy = createButton('复制', function () { copyResult(result) })
    tools.appendChild(locate)
    tools.appendChild(comment)
    tools.appendChild(replace)
    tools.appendChild(undo)
    tools.appendChild(reviseToggle)
    tools.appendChild(copy)

    article.appendChild(head)
    article.appendChild(body)
    article.appendChild(tools)
    if (result.collapsed === true) article.setAttribute('data-collapsed', '1')
    if (container.firstChild) container.insertBefore(article, container.firstChild)
    else container.appendChild(article)

    result.dom = {
      article: article, badge: badge, provider: provider, time: time, status: status,
      parent: parent, text: text, warning: warning, retry: retry, refs: refsList,
      revise: revise, reviseInput: reviseInput, toolbar: {
        locate: locate, comment: comment, replace: replace, undo: undo, revise: reviseToggle, copy: copy,
      },
    }
    updateResultCard(result)
    applyFilter(result)
  }

  function updateResultCard(result) {
    var dom = result.dom
    if (!dom) return
    var article = dom.article
    article.setAttribute('data-status', result.status)
    if (result.meta) {
      var providerText = [result.meta.provider, result.meta.model].filter(Boolean).join(' / ')
      if (result.meta.reasoningEffort) providerText += ` · ${result.meta.reasoningEffort}`
      dom.provider.textContent = providerText
    } else if (result.provider) {
      dom.provider.textContent = result.provider
    }
    if (result.status === 'running') {
      dom.status.textContent = result.stage || '进行中…'
      dom.status.setAttribute('data-running', '')
      dom.status.removeAttribute('data-error')
    } else if (result.status === 'error') {
      dom.status.textContent = '失败'
      dom.status.removeAttribute('data-running')
      dom.status.setAttribute('data-error', '')
    } else {
      dom.status.textContent = result.applied ? '已替换' : '完成'
      dom.status.removeAttribute('data-running')
      dom.status.removeAttribute('data-error')
    }
    dom.parent.textContent = parentLabel(result)
    dom.parent.style.display = dom.parent.textContent === '' ? 'none' : ''
    if (result.status === 'error') {
      dom.text.textContent = result.error || '审核失败'
    } else if (result.status === 'running') {
      dom.text.textContent = result.text
    } else {
      dom.text.innerHTML = markdownLite(result.text)
    }
    dom.warning.textContent = result.warning || ''
    dom.warning.style.display = result.warning ? '' : 'none'
    renderSources(dom.refs, result)
    renderRetry(dom.retry, result)
    var hasBody = result.text.trim() !== '' || result.status === 'error'
    dom.toolbar.comment.disabled = result.status !== 'done' || result.text.trim() === ''
    dom.toolbar.replace.disabled = !hasBody || result.status === 'running'
    dom.toolbar.replace.style.display = (result.action === 'expand' || result.action === 'custom') ? '' : 'none'
    dom.toolbar.undo.style.display = result.applied ? '' : 'none'
    dom.toolbar.undo.disabled = !result.applied
    dom.toolbar.copy.disabled = result.status === 'running' && result.text === ''
    dom.toolbar.revise.disabled = result.status !== 'done'
    dom.revise.style.display = result.status === 'done' ? '' : 'none'
    dom.toolbar.locate.style.display = result.selection ? '' : 'none'
    if (result.collapsed === true) article.setAttribute('data-collapsed', '1')
    else article.removeAttribute('data-collapsed')
    applyFilter(result)
  }

  function renderRetry(container, result) {
    container.innerHTML = ''
    var truncated = (result.warning || '').indexOf('长度上限') >= 0
    if (!truncated || result.status === 'running') return
    var button = createButton('提高输出上限重试', function () {
      var current = result.maxTokens > 0 ? result.maxTokens : defaultCap(result.action)
      var next = Math.min(16000, Math.max(current + 4000, Math.round(current * 1.5)))
      runAnalysis(result.action, {
        instruction: result.instruction,
        selection: result.selection,
        maxTokens: next,
        retryOf: result.id,
      })
      toast(`已提高上限到 ${next} tokens 重新执行`)
    })
    button.className = 'rd-retry-button'
    container.appendChild(button)
  }

  function defaultCap(action) {
    var catalog = state.catalog
    var caps = catalog && catalog.caps && catalog.caps.default ? catalog.caps.default : null
    if (caps && Number(caps[action]) > 0) return Number(caps[action])
    return action === 'expand' ? 8000 : 6000
  }

  function renderSources(list, result) {
    list.innerHTML = ''
    if (!result.sources || result.sources.length === 0) {
      if (result.note) {
        var note = document.createElement('li')
        note.textContent = result.note
        list.appendChild(note)
      }
      return
    }
    result.sources.forEach(function (source, index) {
      var item = document.createElement('li')
      item.appendChild(document.createTextNode(`[S${index + 1}] `))
      if (source.url) {
        var link = document.createElement('a')
        link.href = String(source.url)
        link.target = '_blank'
        link.rel = 'noreferrer noopener'
        link.textContent = String(source.title || source.url)
        item.appendChild(link)
      } else {
        item.appendChild(document.createTextNode(String(source.title || '')))
      }
      list.appendChild(item)
    })
  }

  function collapseAllExcept(keep) {
    for (var result of state.results) {
      if (result === keep) {
        result.collapsed = false
      } else if (result.status === 'done' && result.collapsed !== true) {
        result.collapsed = true
      }
    }
  }

  function applyFilter(result) {
    if (!result.dom) return
    var visible = state.filter === 'all' || result.action === state.filter
    result.dom.article.style.display = visible ? '' : 'none'
  }

  function renderFilters() {
    if (!refs.filterTabs) return
    var tabs = [{ key: 'all', label: '全部', count: state.results.length }]
    for (var action of ACTION_ORDER) {
      tabs.push({
        key: action,
        label: ACTION_LABELS[action],
        count: state.results.filter(function (item) { return item.action === action }).length,
      })
    }
    refs.filterTabs.innerHTML = ''
    for (var tab of tabs) {
      var button = document.createElement('button')
      button.type = 'button'
      button.className = 'rd-tab'
      if (state.filter === tab.key) button.setAttribute('data-active', '')
      button.textContent = `${tab.label}${tab.count > 0 ? ` ${tab.count}` : ''}`
      button.addEventListener('click', (function (key) {
        return function () {
          state.filter = key
          renderFilters()
          for (var result of state.results) applyFilter(result)
        }
      })(tab.key))
      refs.filterTabs.appendChild(button)
    }
    for (var result of state.results) applyFilter(result)
  }

  function removeEmpty() {
    var empty = refs.results && refs.results.querySelector('.rd-empty')
    if (empty) empty.remove()
  }

  function renderEmpty() {
    if (!refs.results || refs.results.children.length > 0) return
    var empty = document.createElement('p')
    empty.className = 'rd-empty'
    empty.textContent = '还没有结果。选中一段文字，从上面选一个动作开始。'
    refs.results.appendChild(empty)
  }

  function renderAll() {
    refs.results.innerHTML = ''
    var ordered = state.results.slice().sort(function (a, b) { return b.createdAt - a.createdAt })
    state.results = ordered
    ordered.forEach(function (result, index) {
      result.dom = null
      result.collapsed = result.status === 'done' ? index > 0 : false
    })
    // 逐条插到最前面，倒序插入以保持「最新在最上」。
    for (var index = ordered.length - 1; index >= 0; index -= 1) {
      insertResultCard(ordered[index])
    }
    renderEmpty()
    renderFilters()
  }

  /* ------------------------------------------------------------------ *
   * 结果 → 文档
   * ------------------------------------------------------------------ */

  function stripMarkers(raw) {
    return String(raw)
      .split('\n')
      .filter(function (line) { return !/^\s*<<<TEXT\s*$/u.test(line) && !/^\s*>>>\s*$/u.test(line) })
      .join('\n')
      .trim()
  }

  function replacementBody(raw) {
    var block = /<<<TEXT\s*\n([\s\S]*?)\n?\s*>>>/u.exec(raw)
    if (block !== null && block[1].trim() !== '') return block[1].trim()
    var cut = raw.search(/\n\s*###\s*(扩写说明|改写说明|修改说明|说明)/u)
    var body = cut >= 0 ? raw.slice(0, cut) : raw
    return stripMarkers(body)
  }

  function currentRange(result) {
    var doc = getDoc()
    if (!doc) return { error: '没有打开的文档' }
    var selection = result.selection
    if (!selection) return { error: '没有可用的选区' }
    var range
    try {
      range = doc.Range(selection.start, selection.end)
    } catch (error) {
      return { error: '无法定位原文范围' }
    }
    if (String(range.Text) !== selection.text) {
      return { error: '原文已变化（可能已被修改），请重新选中后再试' }
    }
    return { doc: doc, range: range }
  }

  function insertComment(result) {
    if (result.status !== 'done') { toast('结果还没生成完'); return }
    var located = currentRange(result)
    if (located.error) { toast(located.error); return }
    var text = stripMarkers(result.text)
    if (text === '') { toast('没有可写入的批注内容'); return }
    if (text.length > 3000) text = `${text.slice(0, 3000)}…`
    try {
      located.range.Comments.Add(located.range, text)
      toast('已插入批注')
      report({ kind: 'wps-comment', ok: true, chars: text.length, resultId: result.id })
    } catch (error) {
      toast(`插入批注失败：${(error && error.message) || error}`)
      report({ kind: 'wps-comment', ok: false, error: String((error && error.message) || error) })
    }
  }

  function replaceOriginal(result) {
    if (result.status !== 'done') { toast('结果还没生成完'); return }
    var located = currentRange(result)
    if (located.error) { toast(located.error); return }
    var replacement = replacementBody(result.text)
    if (replacement === '') { toast('没有可替换的正文'); return }
    var doc = located.doc
    var previous = ''
    var beforeCount = 0
    var contentText = ''
    var contentLengthBefore = 0
    var startOffset = Number(result.selection.start)
    var previousText = result.selection.text
    try {
      contentText = String(doc.Content.Text)
      contentLengthBefore = contentText.length
    } catch (error) { /* ignore */ }
    try {
      previous = String(doc.TrackRevisions)
      beforeCount = doc.Revisions.Count
      doc.TrackRevisions = true
      located.range.Text = replacement
    } catch (error) {
      try { doc.TrackRevisions = previous === 'true' } catch (ignore) { /* ignore */ }
      toast(`替换失败：${(error && error.message) || error}`)
      report({ kind: 'wps-replace', ok: false, error: String((error && error.message) || error) })
      return
    }
    try { doc.TrackRevisions = previous === 'true' } catch (ignore) { /* ignore */ }
    result.applied = {
      start: startOffset,
      end: startOffset + replacement.length,
      inserted: replacement,
      previous: previousText,
      revisionCountBefore: beforeCount,
      revisionCountAfter: doc.Revisions.Count,
      contentLengthBefore: contentLengthBefore,
      contextBefore: contentText.slice(Math.max(0, startOffset - 40), startOffset),
      contextAfter: contentText.slice(startOffset + previousText.length, startOffset + previousText.length + 40),
      tracked: previous === 'true',
      appliedAt: Date.now(),
    }
    if (result.dom) result.dom.article.setAttribute('data-status', 'done')
    updateResultCard(result)
    saveHistory()
    toast('已以修订形式替换，可在 WPS 中接受/拒绝')
    report({ kind: 'wps-replace', ok: true, chars: replacement.length, resultId: result.id })
    readSelection()
    renderSelection()
  }

  /**
   * 撤销一次替换：拒绝本次替换新增的修订，把原文还回去。
   * 以「替换前的内容长度 / 上下文」做校验；失败时回退到直接写回并如实上报。
   * @returns {{ok: boolean, mode: string, rejected: number, debug: object, error?: string}}
   */
  function undoReplace(result) {
    var doc = getDoc()
    var debug = {}
    if (!doc) { toast('没有打开的文档'); return { ok: false, mode: 'none', rejected: 0, debug: debug, error: 'no-document' } }
    var record = result.applied
    if (!record) { toast('这条结果没有替换记录'); return { ok: false, mode: 'none', rejected: 0, debug: debug, error: 'no-record' } }
    try {
      debug.contentLengthAfterReplace = String(doc.Content.Text).length
      debug.rangeInsertedText = String(doc.Range(record.start, record.end).Text).slice(0, 60)
      debug.rangePreviousText = String(doc.Range(record.start, record.start + record.previous.length).Text).slice(0, 60)
      debug.revisionsBefore = []
      var count = doc.Revisions.Count
      for (var probe = count; probe >= 1; probe -= 1) {
        try {
          var item = doc.Revisions.Item(probe)
          debug.revisionsBefore.push({
            start: Number(item.Range.Start),
            end: Number(item.Range.End),
            type: typeof item.Type === 'number' ? item.Type : -1,
          })
        } catch (error) { /* ignore */ }
      }
      debug.revisionCountBefore = count
    } catch (error) {
      debug.probeError = String((error && error.message) || error)
    }

    var mode = 'reject-revisions'
    var rejected = 0
    var failed = 0
    var delta = 0
    try {
      delta = Math.max(0, Number(record.revisionCountAfter) - Number(record.revisionCountBefore))
      var countNow = doc.Revisions.Count
      // 只处理本次替换新增的修订（最后 delta 个）；delta 不可用时按范围重叠兜底。
      for (var index = countNow; index >= 1; index -= 1) {
        var withinNew = delta > 0 ? index > countNow - delta : true
        if (!withinNew) continue
        try {
          var revision = doc.Revisions.Item(index)
          var revisionRange = revision.Range
          var overlaps = Number(revisionRange.End) > record.start - record.previous.length
            && Number(revisionRange.Start) < record.end + record.previous.length
          if (overlaps) {
            revision.Reject()
            rejected += 1
          }
        } catch (error) {
          failed += 1
        }
      }
    } catch (error) {
      failed += 1
    }
    debug.rejected = rejected
    debug.failed = failed
    debug.revisionDelta = delta

    var restored = verifyRestored(doc, record, debug)
    if (!restored && failed === 0) {
      mode = 'direct'
      var track = ''
      try {
        track = String(doc.TrackRevisions)
        doc.TrackRevisions = false
        doc.Range(record.start, record.start + record.inserted.length).Text = record.previous
      } catch (error) {
        toast(`撤销替换失败：${(error && error.message) || error}`)
        report({ kind: 'wps-undo-replace', ok: false, mode: 'direct', debug: debug })
        return { ok: false, mode: 'direct', rejected: rejected, debug: debug, error: String((error && error.message) || error) }
      } finally {
        try { doc.TrackRevisions = track === 'true' } catch (ignore) { /* ignore */ }
      }
      restored = verifyRestored(doc, record, debug)
    }
    debug.restored = restored
    if (restored) {
      result.applied = null
      if (result.dom) result.dom.article.setAttribute('data-status', 'done')
      updateResultCard(result)
      saveHistory()
      toast(mode === 'direct' ? '已写回原文（未走修订）' : '已撤销替换（拒绝相关修订）')
      report({ kind: 'wps-undo-replace', ok: true, mode: mode, rejected: rejected, debug: debug })
    } else {
      toast('未能自动撤销，请用 WPS 的撤销（⌘Z）或手动处理')
      report({ kind: 'wps-undo-replace', ok: false, mode: mode, rejected: rejected, debug: debug })
    }
    readSelection()
    renderSelection()
    return { ok: restored, mode: mode, rejected: rejected, debug: debug }
  }

  /** 校验撤销是否把文档恢复成替换前的样子。 */
  function verifyRestored(doc, record, debug) {
    try {
      var content = String(doc.Content.Text)
      debug.contentLengthAfterUndo = content.length
      debug.lengthRestored = content.length === Number(record.contentLengthBefore)
      if (record.contextBefore || record.contextAfter) {
        debug.contextRestored = content.indexOf(record.contextBefore + record.previous + record.contextAfter) >= 0
      } else {
        debug.contextRestored = false
      }
      debug.rangeRestored = String(doc.Range(record.start, record.start + record.previous.length).Text) === record.previous
      return debug.lengthRestored || debug.contextRestored || debug.rangeRestored
    } catch (error) {
      debug.verifyError = String((error && error.message) || error)
      return false
    }
  }

  function locateResult(result) {
    var doc = getDoc()
    if (!doc || !result.selection) { toast('没有可定位的选区'); return false }
    try {
      var range = doc.Range(result.selection.start, result.selection.end)
      range.Select()
      try {
        var app = getApp()
        if (app.ActiveWindow && typeof app.ActiveWindow.ScrollIntoView === 'function') {
          app.ActiveWindow.ScrollIntoView(range)
        }
      } catch (error) { /* scrolling is best-effort */ }
      toast('已定位到原文')
      readSelection()
      renderSelection()
      report({ kind: 'wps-locate', ok: true, resultId: result.id })
      return true
    } catch (error) {
      toast(`定位失败：${(error && error.message) || error}`)
      report({ kind: 'wps-locate', ok: false, error: String((error && error.message) || error) })
      return false
    }
  }

  function submitRevise(result) {
    var dom = result.dom
    if (!dom) return
    var instruction = dom.reviseInput.value.trim()
    if (instruction === '') { toast('请先写追加要求'); return }
    if (result.status !== 'done') { toast('结果还没生成完'); return }
    var body = replacementBody(result.text) || stripMarkers(result.text)
    if (body === '') { toast('这条结果没有可继续修改的正文'); return }
    var original = result.original !== '' ? result.original : (result.selection ? result.selection.text : '')
    dom.reviseInput.value = ''
    dom.revise.setAttribute('data-open', '0')
    runAnalysis('custom', {
      instruction: instruction,
      revise: true,
      text: body,
      original: original,
      selection: result.selection,
      parentId: result.id,
    })
  }

  function copyResult(result) {
    var text = result.status === 'error' ? result.error : result.text
    copyText(text)
  }

  function copyText(text) {
    if (navigator.clipboard && typeof navigator.clipboard.writeText === 'function') {
      navigator.clipboard.writeText(text).then(function () { toast('已复制') }).catch(function () { fallbackCopy(text) })
      return
    }
    fallbackCopy(text)
  }

  function fallbackCopy(text) {
    try {
      var area = document.createElement('textarea')
      area.value = text
      area.style.position = 'fixed'
      area.style.opacity = '0'
      document.body.appendChild(area)
      area.select()
      var ok = document.execCommand('copy')
      area.remove()
      toast(ok ? '已复制' : '复制失败')
    } catch (error) {
      toast('复制失败')
    }
  }

  /* ------------------------------------------------------------------ *
   * 渲染辅助
   * ------------------------------------------------------------------ */

  function escapeHtml(text) {
    return String(text)
      .replace(/&/gu, '&amp;')
      .replace(/</gu, '&lt;')
      .replace(/>/gu, '&gt;')
      .replace(/"/gu, '&quot;')
      .replace(/'/gu, '&#39;')
  }

  function inlineMarkdown(text) {
    var html = escapeHtml(text)
    html = html.replace(/\*\*([^*]+)\*\*/gu, '<strong>$1</strong>')
    html = html.replace(/`([^`]+)`/gu, '<code>$1</code>')
    html = html.replace(/\[([^\]]+)\]\(((?:https?:\/\/|\/)[^)\s]+)\)/gu, '<a href="$2" target="_blank" rel="noreferrer noopener">$1</a>')
    return html
  }

  function markdownLite(source) {
    var lines = String(source || '').replace(/\r\n?/gu, '\n').split('\n')
    var out = []
    var paragraph = []
    var flush = function () {
      if (paragraph.length === 0) return
      out.push(`<p>${paragraph.map(inlineMarkdown).join('<br>')}</p>`)
      paragraph = []
    }
    for (var line of lines) {
      var trimmed = line.trim()
      if (trimmed === '') { flush(); continue }
      if (/^(-{3,}|\*{3,})$/u.test(trimmed)) { flush(); out.push('<hr>'); continue }
      var heading = /^(#{1,6})\s+(.*)$/u.exec(trimmed)
      if (heading !== null) {
        flush()
        var level = Math.min(heading[1].length, 4)
        out.push(`<h${level}>${inlineMarkdown(heading[2])}</h${level}>`)
        continue
      }
      var bullet = /^[-*+]\s+(.*)$/u.exec(trimmed)
      if (bullet !== null) { flush(); out.push(`<div class="li">• ${inlineMarkdown(bullet[1])}</div>`); continue }
      var numbered = /^(\d+)[.)]\s+(.*)$/u.exec(trimmed)
      if (numbered !== null) { flush(); out.push(`<div class="li">${numbered[1]}. ${inlineMarkdown(numbered[2])}</div>`); continue }
      paragraph.push(line)
    }
    flush()
    return out.join('')
  }

  function toast(message) {
    if (!refs.toast) return
    refs.toast.textContent = message
    refs.toast.setAttribute('data-show', '')
    if (state.toastTimer) clearTimeout(state.toastTimer)
    state.toastTimer = setTimeout(function () { refs.toast.removeAttribute('data-show') }, 2400)
  }

  /* ------------------------------------------------------------------ *
   * 验证钩子（只在 ?verify=1 时启用，用于开发验证）
   * ------------------------------------------------------------------ */

  function maybeSelfTest() {
    if (!VERIFY || selfTest.started) return
    if (state.health !== 'ready') return
    var app = getApp()
    var doc = getDoc()
    if (!app || !doc) return
    if (!/wps-verify/i.test(String(doc.Name || ''))) return
    selfTest.started = true
    try {
      var length = String(doc.Content.Text || '').length
      doc.Range(0, Math.min(40, length)).Select()
      readSelection()
      renderSelection()
    } catch (error) {
      report({ kind: 'wps-selftest', stage: 'select', error: String((error && error.message) || error) })
      return
    }
    report({
      kind: 'wps-selftest',
      stage: 'start',
      selection: state.selection ? state.selection.text : '',
      hasCatalog: !!(state.catalog && state.catalog.providers),
      historyRestored: state.results.length,
    })
    runAnalysis('expand', {
      onSettled: function (result) {
        report({
          kind: 'wps-selftest',
          stage: 'analysis',
          status: result.status,
          chars: result.text.length,
          warning: result.warning,
          error: result.error,
        })
        if (result.status !== 'done') return
        setTimeout(function () {
          afterExpandWrite(result)
        }, 400)
      },
    })
  }

  function afterExpandWrite(result) {
    var doc = getDoc()
    var info = { comment: null, replace: null, locate: null, undo: null, save: null, revise: null, history: null }
    info.locate = { ok: locateResult(result) }
    try { info.comment = { commentsBefore: doc.Comments.Count } } catch (error) { info.comment = { error: String((error && error.message) || error) } }
    insertComment(result)
    try { info.comment.commentsAfter = doc.Comments.Count } catch (error) { info.comment.commentsAfter = -1 }
    replaceOriginal(result)
    try {
      info.replace = {
        revisions: doc.Revisions.Count,
        applied: result.applied !== null,
        appliedLength: result.applied ? result.applied.inserted.length : 0,
        contentLength: String(doc.Content.Text).length,
        rangeText: result.applied ? String(doc.Range(result.applied.start, result.applied.end).Text).slice(0, 60) : '',
        revisionRanges: [],
      }
      for (var probe = doc.Revisions.Count; probe >= 1; probe -= 1) {
        var revision = doc.Revisions.Item(probe)
        info.replace.revisionRanges.push({
          start: Number(revision.Range.Start),
          end: Number(revision.Range.End),
          type: typeof revision.Type === 'number' ? revision.Type : -1,
        })
      }
    } catch (error) {
      info.replace = { error: String((error && error.message) || error) }
    }
    try {
      var target = '/Users/xiancai/Library/Containers/com.kingsoft.wpsoffice.mac/Data/tmp/wps-review-desk-verify.docx'
      if (typeof doc.SaveAs2 === 'function') doc.SaveAs2(target)
      info.save = { path: target, name: String(doc.Name || ''), saved: doc.Saved }
    } catch (error) {
      info.save = { error: String((error && error.message) || error) }
    }
    info.undo = undoReplace(result)
    try { info.afterUndo = { comments: doc.Comments.Count, revisions: doc.Revisions.Count } } catch (error) { info.afterUndo = { error: String(error) } }
    try {
      info.history = { results: state.results.length, stored: !!localStorage.getItem(HISTORY_KEY) }
    } catch (error) {
      info.history = { error: String((error && error.message) || error) }
    }
    report({ kind: 'wps-selftest', stage: 'write', info: info })
    // 追加要求链（第二次模型调用）
    setTimeout(function () {
      var revision = runAnalysis('custom', {
        instruction: '压缩到 40 字以内，保留关键信息',
        revise: true,
        text: replacementBody(result.text),
        original: result.selection ? result.selection.text : '',
        selection: result.selection,
        parentId: result.id,
        onSettled: function (child) {
          report({
            kind: 'wps-selftest',
            stage: 'revise',
            status: child.status,
            chars: child.text.length,
            parentId: child.parentId,
            warning: child.warning,
            error: child.error,
          })
        },
      })
      if (revision === null) report({ kind: 'wps-selftest', stage: 'revise', status: 'skipped' })
    }, 600)
  }

  function report(payload) {
    if (!VERIFY) return
    try {
      payload.at = payload.at || new Date().toISOString()
      fetch('http://127.0.0.1:38999/report', {
        method: 'POST',
        headers: { 'content-type': 'application/json' },
        body: JSON.stringify(payload),
      }).catch(function () { /* collector optional */ })
    } catch (error) { /* ignore */ }
  }

  /* ------------------------------------------------------------------ *
   * 初始化
   * ------------------------------------------------------------------ */

  function init() {
    refs = {
      connDot: document.getElementById('conn-dot'),
      docName: document.getElementById('doc-name'),
      selectionPreview: document.getElementById('selection-preview'),
      selectionMeta: document.getElementById('selection-meta'),
      instruction: document.getElementById('instruction'),
      btnCustom: document.getElementById('btn-custom'),
      btnRefresh: document.getElementById('btn-refresh'),
      results: document.getElementById('results'),
      filterTabs: document.getElementById('filter-tabs'),
      btnClear: document.getElementById('btn-clear'),
      bridgeUrl: document.getElementById('bridge-url'),
      bridgeHint: document.getElementById('bridge-hint'),
      btnCheck: document.getElementById('btn-check'),
      toast: document.getElementById('toast'),
      settingModel: document.getElementById('setting-model'),
      settingEffort: document.getElementById('setting-effort'),
      settingCap: document.getElementById('setting-cap'),
      settingsSummary: document.getElementById('settings-summary'),
      settingsHint: document.getElementById('settings-hint'),
      actionButtons: Array.prototype.slice.call(document.querySelectorAll('.rd-action')),
    }

    try {
      state.bridge = normalizeBridge(localStorage.getItem(BRIDGE_KEY) || DEFAULT_BRIDGE)
    } catch (error) {
      state.bridge = DEFAULT_BRIDGE
    }
    if (refs.bridgeUrl) refs.bridgeUrl.value = state.bridge
    state.settings = readSettings()
    renderCapOptions()
    if (refs.settingModel) {
      refs.settingModel.addEventListener('change', function () {
        var value = refs.settingModel.value
        if (value === '') {
          state.settings.provider = ''
          state.settings.model = ''
        } else {
          var parts = value.split('::')
          state.settings.provider = parts[0] || ''
          state.settings.model = parts[1] || ''
        }
        renderEffortOptions()
        persistSettings()
      })
    }
    if (refs.settingEffort) {
      refs.settingEffort.addEventListener('change', function () {
        state.settings.reasoningEffort = refs.settingEffort.value
        persistSettings()
      })
    }
    if (refs.settingCap) {
      refs.settingCap.addEventListener('change', function () {
        state.settings.maxTokens = Number(refs.settingCap.value) || 0
        persistSettings()
      })
    }

    for (var button of refs.actionButtons) {
      button.addEventListener('click', (function (action) {
        return function () { runAnalysis(action) }
      })(button.getAttribute('data-action')))
    }
    if (refs.btnCustom) {
      refs.btnCustom.addEventListener('click', function () {
        runAnalysis('custom', { instruction: refs.instruction.value.trim() })
      })
    }
    if (refs.instruction) {
      refs.instruction.addEventListener('keydown', function (event) {
        if (event.key === 'Enter') {
          event.preventDefault()
          runAnalysis('custom', { instruction: refs.instruction.value.trim() })
        }
      })
    }
    if (refs.btnRefresh) {
      refs.btnRefresh.addEventListener('click', function () {
        readSelection()
        renderSelection()
        toast(state.selection === null ? '没有选中文字' : `已选中 ${state.selection.text.length} 字`)
      })
    }
    if (refs.bridgeUrl) {
      refs.bridgeUrl.addEventListener('change', function () {
        state.bridge = normalizeBridge(refs.bridgeUrl.value)
        refs.bridgeUrl.value = state.bridge
        try { localStorage.setItem(BRIDGE_KEY, state.bridge) } catch (error) { /* ignore */ }
        checkBridge()
      })
    }
    if (refs.btnCheck) refs.btnCheck.addEventListener('click', checkBridge)
    if (refs.btnClear) refs.btnClear.addEventListener('click', clearHistory)

    // 恢复历史
    state.results = loadHistory()
    if (state.results.length > 0) renderAll()
    else renderEmpty()
    renderFilters()
    for (var result of state.results) {
      if (result.applied) result.dom && updateResultCard(result)
    }

    renderSettingsSummary()
    checkBridge()
    readSelection()
    renderSelection()
    setInterval(function () {
      readSelection()
      renderSelection()
      maybeSelfTest()
    }, 1500)

    report({
      kind: 'wps-taskpane',
      version: '0.2',
      loaded: true,
      hasApp: !!getApp(),
      docName: state.docName,
      bridge: state.bridge,
      historyRestored: state.results.length,
      at: new Date().toISOString(),
    })
  }

  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', init)
  else init()
})()
