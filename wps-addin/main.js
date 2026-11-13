/**
 * Review Desk · WPS 加载项主入口。
 *
 * 只负责功能区与任务窗格的开合；真正的审稿界面在 taskpane.html / taskpane.js。
 * 加载项启动后，如果已经打开了文档，会自动展开一次审稿台（之后可用功能区按钮切换）。
 */
(function () {
  'use strict'

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

  function openPane(toggle) {
    var app = getApp()
    if (app === null) return null
    if (state.paneId !== null) {
      try {
        var existing = app.GetTaskPane(state.paneId)
        if (existing) {
          existing.Visible = toggle === true ? !existing.Visible : true
          return { id: state.paneId, visible: existing.Visible }
        }
      } catch (error) { /* task pane was disposed; create a new one */ }
      state.paneId = null
    }
    try {
      var pane = app.CreateTaskPane(paneUrl())
      state.paneId = pane.ID
      pane.DockPosition = app.Enum ? app.Enum.JSKsoEnum_msoCTPDockPositionRight : 2
      pane.Visible = true
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
