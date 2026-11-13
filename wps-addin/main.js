/**
 * Review Desk · WPS 加载项主入口。
 *
 * 只负责功能区与任务窗格的开合、位置与宽度；审稿界面在 taskpane.html / taskpane.js。
 * 任务窗格默认停靠在右侧，并保留至少 560px 宽度以便两栏布局。
 */
(function () {
  'use strict'

  var REPORT = 'http://127.0.0.1:38999/report'
  var MIN_PANE_WIDTH = 560
  var PREFERRED_PANE_WIDTH = 760
  var state = { paneId: null, ribbonUI: null, autoOpened: false }

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

  function baseUrl() {
    try {
      var href = document.location.toString()
      return href.substring(0, href.lastIndexOf('/'))
    } catch (error) {
      return ''
    }
  }

  /** 验证文档自动带上 verify=1，让面板把自检事件发到本地 collector。 */
  function isVerifyDoc() {
    var doc = getDoc()
    return doc !== null && /wps-verify/i.test(String(doc.Name || ''))
  }

  function paneUrl() {
    return baseUrl() + '/taskpane.html' + (isVerifyDoc() ? '?verify=1' : '')
  }

  /** 把任务窗格停靠到右侧，并尽量保住两栏布局需要的宽度。 */
  function applyPaneLayout(pane) {
    try { pane.MinWidth = MIN_PANE_WIDTH } catch (error) { /* read-only on some builds */ }
    try { pane.DockPosition = 2 } catch (error) { /* 2 = dock right */ }
    var widen = function () {
      try {
        if (Number(pane.Width) < MIN_PANE_WIDTH) pane.Width = PREFERRED_PANE_WIDTH
      } catch (error) { /* width is best-effort */ }
    }
    widen()
    setTimeout(widen, 300)
    setTimeout(widen, 1200)
  }

  function reportPane(pane) {
    if (!isVerifyDoc()) return
    try {
      var info = { dock: Number(pane.DockPosition), width: Number(pane.Width), height: Number(pane.Height), visible: pane.Visible }
      fetch(REPORT, {
        method: 'POST',
        headers: { 'content-type': 'application/json' },
        body: JSON.stringify(Object.assign({ kind: 'wps-pane' }, info, { at: new Date().toISOString() })),
      }).catch(function () { /* collector optional */ })
    } catch (error) { /* ignore */ }
  }

  function openPane(toggle) {
    var app = getApp()
    if (app === null) return null
    if (state.paneId !== null) {
      try {
        var existing = app.GetTaskPane(state.paneId)
        if (existing) {
          existing.Visible = toggle === true ? !existing.Visible : true
          applyPaneLayout(existing)
          reportPane(existing)
          return { id: state.paneId, visible: existing.Visible }
        }
      } catch (error) { /* task pane was disposed; create a new one */ }
      state.paneId = null
    }
    try {
      var pane = app.CreateTaskPane(paneUrl())
      state.paneId = pane.ID
      pane.Visible = true
      applyPaneLayout(pane)
      reportPane(pane)
      setTimeout(function () { reportPane(pane) }, 1500)
      return { id: state.paneId, visible: true }
    } catch (error) {
      console.error('[review-desk] failed to open task pane:', error)
      return null
    }
  }

  function openWhenReady(attempt) {
    attempt = attempt || 0
    if (getDoc() !== null) {
      if (!state.autoOpened) {
        state.autoOpened = true
        openPane(false)
      }
      return
    }
    if (attempt < 20) setTimeout(function () { openWhenReady(attempt + 1) }, 500)
  }

  function OnAddinLoad(ribbonUI) {
    state.ribbonUI = ribbonUI || null
    try {
      if (window.Application && ribbonUI) window.Application.ribbonUI = ribbonUI
    } catch (error) { /* ignore */ }
    setTimeout(function () { openWhenReady(0) }, 1000)
    return true
  }

  function OnAction(control) {
    var id = control && control.Id
    if (id === 'btnOpenPane') openPane(true)
    return true
  }

  window.OnAddinLoad = OnAddinLoad
  window.OnAction = OnAction
  window.OnGetEnabled = function () { return true }
  window.OnGetVisible = function () { return true }
  window.GetImage = function () { return '' }
  window.GetUrlPath = baseUrl
})()
