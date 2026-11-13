/**
 * Review Desk · WPS 任务窗格。
 *
 * 读取 WPS 选区 → 调本机 DSH 的 review-desk 桥（默认 http://127.0.0.1:47831）
 * → 流式展示审核结果 → 以「批注」或「修订替换」写回文档。
 */
(function () {
  'use strict'

  var DEFAULT_BRIDGE = 'http://127.0.0.1:47831'
  var BRIDGE_KEY = 'wps-review-desk.bridge'
  var VERIFY = /[?&]verify=1(?:&|$)/.test(window.location.search)
  var ACTION_LABELS = { source: '来源查证', logic: '逻辑分析', expand: '扩写', relate: '关联分析', custom: '自由要求' }
  var STAGE_LABELS = { planning: '生成检索式…', searching: '检索资料…', writing: '撰写中…' }

  var state = {
    bridge: DEFAULT_BRIDGE,
    docName: '',
    selection: null,
    running: false,
    health: 'unknown',
    results: [],
    seq: 0,
    toastTimer: 0,
  }

  var refs = {}

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

  function runAnalysis(action, instruction, onSettled) {
    if (state.running) return null
    var selection = state.selection
    if (selection === null) {
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
      manuscript = String(doc.Content.Text || '').slice(0, 12000)
    } catch (error) { /* ignore */ }
    var title = state.docName.replace(/\.[^.]+$/u, '')

    var result = {
      id: ++state.seq,
      action: action,
      instruction: instruction || '',
      selection: selection,
      text: '',
      meta: null,
      stage: '',
      status: 'running',
      sources: [],
      note: '',
      warning: '',
      error: '',
      dom: null,
    }
    state.results.unshift(result)
    insertResultCard(result)
    setRunning(true)

    var payload = { action: action, text: selection.text, manuscript: manuscript, title: title }
    if (instruction) payload.instruction = instruction
    report({ kind: 'wps-run', action: action, chars: selection.text.length, at: new Date().toISOString() })

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
      report({
        kind: 'wps-result',
        action: result.action,
        status: result.status,
        chars: result.text.length,
        error: result.error,
      })
      if (typeof onSettled === 'function') onSettled(result)
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
        result.meta = { provider: frame.provider, model: frame.model }
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

  function insertResultCard(result) {
    var container = refs.results
    if (!container) return
    var empty = container.querySelector('.rd-empty')
    if (empty) empty.remove()

    var article = document.createElement('article')
    article.className = 'rd-result'
    article.setAttribute('data-id', String(result.id))
    article.setAttribute('data-status', 'running')

    var head = document.createElement('header')
    head.className = 'rd-result-head'
    var badge = document.createElement('span')
    badge.className = 'rd-badge'
    badge.textContent = ACTION_LABELS[result.action] || result.action
    var provider = document.createElement('span')
    provider.className = 'rd-provider'
    var status = document.createElement('span')
    status.className = 'rd-state'
    status.textContent = '进行中…'
    status.setAttribute('data-running', '')
    head.appendChild(badge)
    head.appendChild(provider)
    head.appendChild(status)

    var body = document.createElement('div')
    body.className = 'rd-result-body'
    var text = document.createElement('div')
    text.className = 'rd-result-text'
    body.appendChild(text)
    var warning = document.createElement('p')
    warning.className = 'rd-warn'
    body.appendChild(warning)
    var refsList = document.createElement('ul')
    refsList.className = 'rd-refs'
    body.appendChild(refsList)

    var tools = document.createElement('footer')
    tools.className = 'rd-tools'
    var commentButton = document.createElement('button')
    commentButton.type = 'button'
    commentButton.textContent = '插入批注'
    commentButton.addEventListener('click', function () { insertComment(result) })
    var replaceButton = document.createElement('button')
    replaceButton.type = 'button'
    replaceButton.setAttribute('data-emphasis', 'replace')
    replaceButton.textContent = '替换原文（修订）'
    replaceButton.addEventListener('click', function () { replaceOriginal(result) })
    var copyButton = document.createElement('button')
    copyButton.type = 'button'
    copyButton.textContent = '复制结果'
    copyButton.addEventListener('click', function () { copyResult(result) })
    tools.appendChild(commentButton)
    tools.appendChild(replaceButton)
    tools.appendChild(copyButton)

    article.appendChild(head)
    article.appendChild(body)
    article.appendChild(tools)
    container.insertBefore(article, container.firstChild)

    result.dom = { article: article, provider: provider, status: status, text: text, warning: warning, refs: refsList, comment: commentButton, replace: replaceButton, copy: copyButton }
    updateResultCard(result)
  }

  function updateResultCard(result) {
    var dom = result.dom
    if (!dom) return
    dom.article.setAttribute('data-status', result.status)
    if (result.meta) {
      dom.provider.textContent = [result.meta.provider, result.meta.model].filter(Boolean).join(' / ')
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
      dom.status.textContent = '完成'
      dom.status.removeAttribute('data-running')
      dom.status.removeAttribute('data-error')
    }
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
    var hasResult = result.text.trim() !== '' || result.status === 'error'
    dom.comment.disabled = !hasResult
    dom.replace.disabled = !hasResult
    dom.copy.disabled = !hasResult
    dom.replace.style.display = (result.action === 'expand' || result.action === 'custom') ? '' : 'none'
    if (refs.results) {
      var empty = refs.results.querySelector('.rd-empty')
      if (empty) empty.remove()
    }
  }

  function renderSources(list, result) {
    list.innerHTML = ''
    if (!result.sources || result.sources.length === 0) {
      if (result.note) {
        var item = document.createElement('li')
        item.textContent = result.note
        list.appendChild(item)
      }
      return
    }
    result.sources.forEach(function (source, index) {
      var item = document.createElement('li')
      var label = `[S${index + 1}] `
      item.appendChild(document.createTextNode(label))
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
    if (result.status === 'running') { toast('结果还没生成完'); return }
    var located = currentRange(result)
    if (located.error) { toast(located.error); return }
    var text = stripMarkers(result.text)
    if (text === '') { toast('没有可写入的批注内容'); return }
    if (text.length > 3000) text = `${text.slice(0, 3000)}…`
    try {
      located.range.Comments.Add(located.range, text)
      toast('已插入批注')
      report({ kind: 'wps-comment', ok: true, chars: text.length })
    } catch (error) {
      toast(`插入批注失败：${(error && error.message) || error}`)
      report({ kind: 'wps-comment', ok: false, error: String((error && error.message) || error) })
    }
  }

  function replaceOriginal(result) {
    if (result.status === 'running') { toast('结果还没生成完'); return }
    var located = currentRange(result)
    if (located.error) { toast(located.error); return }
    var replacement = replacementBody(result.text)
    if (replacement === '') { toast('没有可替换的正文'); return }
    var doc = located.doc
    var previous
    try {
      previous = doc.TrackRevisions
      doc.TrackRevisions = true
      located.range.Text = replacement
    } catch (error) {
      toast(`替换失败：${(error && error.message) || error}`)
      report({ kind: 'wps-replace', ok: false, error: String((error && error.message) || error) })
      return
    } finally {
      try { doc.TrackRevisions = previous } catch (error) { /* ignore */ }
    }
    toast('已以修订形式替换，可在 WPS 中接受/拒绝')
    report({ kind: 'wps-replace', ok: true, chars: replacement.length })
    readSelection()
    renderSelection()
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

  var selfTest = { started: false }

  /** 验证文档里自动跑一遍：选中 → 分析 → 批注 → 修订替换 → 上报。 */
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
    report({ kind: 'wps-selftest', stage: 'start', selection: state.selection ? state.selection.text : '' })
    runAnalysis('expand', '', function (result) {
      report({ kind: 'wps-selftest', stage: 'analysis', status: result.status, chars: result.text.length, error: result.error })
      if (result.status !== 'done') return
      setTimeout(function () {
        insertComment(result)
        replaceOriginal(result)
        var info = {}
        try {
          var active = getDoc()
          info.comments = active.Comments.Count
          info.revisions = active.Revisions.Count
          info.saved = active.Saved
          var target = '/Users/xiancai/Library/Containers/com.kingsoft.wpsoffice.mac/Data/tmp/wps-review-desk-verify.docx'
          if (typeof active.SaveAs2 === 'function') active.SaveAs2(target)
          info.exported = target
          info.exportedName = String(active.Name || '')
          info.exportedSaved = active.Saved
        } catch (error) {
          info.error = String((error && error.message) || error)
        }
        report({ kind: 'wps-selftest', stage: 'write', info: info })
      }, 400)
    })
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
      bridgeUrl: document.getElementById('bridge-url'),
      bridgeHint: document.getElementById('bridge-hint'),
      btnCheck: document.getElementById('btn-check'),
      toast: document.getElementById('toast'),
      actionButtons: Array.prototype.slice.call(document.querySelectorAll('.rd-action')),
    }

    try {
      state.bridge = normalizeBridge(localStorage.getItem(BRIDGE_KEY) || DEFAULT_BRIDGE)
    } catch (error) {
      state.bridge = DEFAULT_BRIDGE
    }
    if (refs.bridgeUrl) refs.bridgeUrl.value = state.bridge

    refs.actionButtons.forEach(function (button) {
      button.addEventListener('click', function () { runAnalysis(button.getAttribute('data-action')) })
    })
    if (refs.btnCustom) {
      refs.btnCustom.addEventListener('click', function () {
        runAnalysis('custom', refs.instruction.value.trim())
      })
    }
    if (refs.instruction) {
      refs.instruction.addEventListener('keydown', function (event) {
        if (event.key === 'Enter') {
          event.preventDefault()
          runAnalysis('custom', refs.instruction.value.trim())
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
      loaded: true,
      hasApp: !!getApp(),
      docName: state.docName,
      bridge: state.bridge,
      at: new Date().toISOString(),
    })
  }

  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', init)
  else init()
})()
