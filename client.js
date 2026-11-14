window.__ModuleLoader__.load({
  id: '@local/review-desk',
  factory(require) {
    const React = require('react')
    const h = React.createElement
    const { useCallback, useEffect, useMemo, useRef, useState } = React

    /* ------------------------------------------------------------------ *
     * Panel identity and dictionaries
     * ------------------------------------------------------------------ */

    /** Id shared by the sidebar entry and the main panel it opens. */
    const PANEL_ID = 'review-desk'
    /** Build tag, surfaced in the page attribute and the console. */
    const BUILD = '1.2.1'
    /** Locale namespace owned by this plugin. */
    const NS = 'reviewDesk'
    /** Authenticated route the Host half registered. */
    const ANALYZE_ROUTE = 'api/review-desk/analyze'
    /** Original-page rendering route. */
    const PAGES_ROUTE = 'api/review-desk/original-pages'
    /** Installs / inspects the WPS add-in on this machine. */
    const WPS_ROUTE = 'api/review-desk/wps-addin'
    /** Persistence key; a version bump discards older shapes. */
    const STORAGE_KEY = 'dsh.review-desk.v1'

    const ZH = {
      panel: '审悦',
      title: '审悦',
      subtitle: '选中左侧文字，在右侧完成查证、逻辑、扩写与关联。',
      importFile: '导入文稿',
      importPackage: '导入审核包',
      exportPackage: '导出审核包',
      importMenu: '导入',
      exportMenu: '导出',
      packageHint: '一个 JSON 文件带走文稿、全部审核记录与当前要求；不含原始 .docx/.pdf 文件。',
      packageExportDone: '审核包已导出',
      packageImportTitle: '导入审核包',
      packageImportBody: '将用包内的文稿和 {runs} 条审核记录替换当前内容，此操作不能撤销。',
      packageImportConfirm: '替换并导入',
      packageImportCancel: '取消',
      packageEmpty: '这个包没有任何文稿内容，可能已损坏。',
      packageInvalid: '不是有效的审核包文件。',
      packageTooLarge: '审核包过大（上限约 8 MB），请先精简结果。',
      packageFresh: '若旧包来自更早版本，记录仍会保留，但部分细节可能缺省。',
      exportMd: '导出 Markdown',
      importedPackage: '审核包已导入',
      exportFile: '导出',
      wpsMenu: 'WPS 版',
      wpsInstall: '安装 WPS 加载项',
      wpsInstallDesc: '把审悦装进本机 WPS 文字',
      wpsStatus: '查看安装状态',
      wpsUninstall: '卸载 WPS 加载项',
      wpsInstalled: 'WPS 加载项已安装：{path}。重启 WPS 后在「审悦」功能区打开审稿台。',
      wpsUninstalled: '已卸载 WPS 加载项。',
      wpsStatusOk: 'WPS 加载项已安装（v{version}）。',
      wpsNotInstalled: '尚未安装 WPS 加载项。',
      wpsFailed: 'WPS 加载项操作失败：{message}',
      copyAll: '复制全文',
      clear: '清空',
      clearArmed: '确认清空',
      titlePlaceholder: '未命名文稿',
      metaChars: '字',
      metaParagraphs: '段',
      metaMinutes: '分钟阅读',
      emptyTitle: '还没有文稿',
      emptyBody: '粘贴 Markdown 或纯文本，或导入 .docx、.pdf；也可以先看看示例。',
      loadSample: '载入示例',
      selectionLabel: '已选中',
      selectionEmpty: '尚未选中文字',
      selectionEmptyBody: '在左侧文稿中拖动选中一段文字，右侧功能随即启用。',
      selectionTooLong: '选择较长，本次仅审核前 6000 字。',
      needSelection: '请先选中要审核的文字。',
      busy: '正在处理上一次请求，请稍候。',
      actionSource: '来源查证',
      actionSourceDesc: '核查数据、时间与引用的出处',
      actionLogic: '逻辑分析',
      actionLogicDesc: '检查论证链条与自相矛盾',
      actionExpand: '扩写',
      actionExpandDesc: '补足背景、机制与依据后入稿',
      actionRelate: '关联分析',
      actionRelateDesc: '看它与全文其他部分的关系',
      results: '结果',
      resultsEmptyTitle: '结果会显示在这里',
      resultsEmptyBody: '四种分析都基于当前选中的文字；{mod}1–4 可直接触发。',
      stop: '停止',
      retry: '重试',
      remove: '删除',
      copy: '复制',
      copied: '已复制到剪贴板',
      insert: '替换原文',
      inserted: '已替换到文稿中',
      appended: '原文已改动，已追加到文稿末尾',
      sources: '参考资料',
      stagePlanning: '正在规划检索…',
      stageSearching: '正在检索',
      stageWriting: '正在撰写结果…',
      stagePreparing: '正在准备…',
      cancelled: '已停止',
      failed: '未能完成',
      seconds: '秒',
      expandBlock: '可替换正文',
      instruction: '补充要求',
      instructionPlaceholder: '可选：告诉审核人你的关注点',
      sampleNote: '示例文本仅用于演示，内容不代表任何真实机构。',
      fileTooLarge: '文件超过 2 MB，请拆分后再导入。',
      fileFailed: '无法读取该文件。',
      cleared: '已清空',
      copiedAll: '全文已复制',
      selectionHint: '选中文字后可用',
      manuscript: '文稿正文',
      errService: '审核服务尚未就绪，请重启应用后再试。',
      errAuth: '登录状态已失效，请刷新页面。',
      errNetwork: '无法连接到审核服务，请检查本地服务是否在运行。',
      announceDone: '已完成',
      announceFailed: '未完成',
      importing: '正在解析…',
      imported: '已导入',
      importHint: '支持 .docx、.pdf、.md、.txt；正文导入后会替换当前文稿',
      importEmpty: '这个文件没有可用的文字。',
      importTooLarge: '文件过大：docx ≤ 20 MB、PDF ≤ 15 MB、文本 ≤ 2 MB。',
      importPdfNoText: '无法从这份 PDF 提取文字。它可能是扫描件，或使用了不受支持的字体编码；可以先用 Word 另存为 .docx 再导入。',
      importLegacyDoc: '暂不支持旧版 .doc，请另存为 .docx 后导入。',
      importUnsupported: '不支持的文件类型，请选择 .docx、.pdf、.md 或 .txt。',
      importNoDeflate: '当前运行环境不支持文档解压，无法解析 .docx 与 .pdf。',
      importFailed: '文件解析失败，请确认文件未损坏。',
      importEncrypted: '读不出来：这份文稿可能是加了密码的 Word 文档，或旧版 .doc 改了扩展名。请在 Word 里取消密码、另存为 .docx 后再导入。',
      importZipVariant: '这个 .docx 的压缩结构不标准（zip64 或非常规压缩）。用 Word 打开后「另存为 → .docx」再试一次。',
      importStructure: '文件内容不完整：找不到 word/document.xml。可能是转换工具导出的变体，请用 Word 另存为一份新的 .docx。',
      reviewedCount: '已审阅',
      jumpTo: '定位到原文',
      markStale: '原文已改动，找不到这一段。',
      selectFirst: '先选中文字',
      selectFirstAction: '选中第一段',
      reviewedBlocks: '已审阅',
      markResults: '这条划线有 {count} 条结果，点击可查看',
      markJumped: '已跳到对应结果',
      tabAll: '全部',
      expandCard: '展开',
      collapseCard: '收起',
      emptyFilter: '这一类还没有结果',
      resizeLeft: '调整文稿与审查栏宽度',
      resizeRight: '调整审查与结果栏宽度',
      resizeHint: '拖动调整，双击恢复默认',
      colDoc: '文稿',
      colReview: '审查',
      colResults: '结果',
      modeEdit: '编辑',
      modeSource: '源码',
      modePreview: '预览',
      blockEditHint: '点任意位置直接改字，点别处自动保存；表格单元格同样可直接编辑，Enter 换行，格式用「源码」调整。',
      blockCommit: '完成',
      blockCancel: '取消',
      imageAlt: '图片',
      imageBlocked: '图片无法加载',
      modeName: '显示方式',
      viewOriginal: '原件',
      revise: '修改',
      reviseLabel: '追加修改要求',
      revisePlaceholder: '例如：再压缩到 100 字，语气更硬一点',
      reviseRun: '应用修改',
      reviseEmpty: '先写下修改要求。',
      cancel: '取消',
      replaceIntent: '改写后替换原文',
      replaceIntentHint: '要求里出现「改写/压缩/改成…」等字样时会自动替换；打开这个开关则一律替换。可用「撤销替换」恢复。',
      replacedNotice: '已替换原文，可撤销',
      appendedNotice: '原文已改动，已追加到末尾',
      undoReplace: '撤销替换',
      undoDone: '已恢复原文',
      undoStale: '原文已经变化，无法撤销',
      revisionBadge: '修订',
      originalNote: '只读：原稿用于核对版式，编辑请切到「编辑」',
      pagesLoading: '正在生成原稿版式…',
      pagesUnit: '页',
      pagesFirst: '仅显示前 8 页',
      pageLabel: '第 {page} 页',
      closeSheet: '关闭',
      originalMissing: '已重新载入的文稿没有原始文件，请再次导入。',
      actionCustom: '按要求处理',
      actionCustomDesc: '直接按你写的要求修改或改写',
      customHint: '未选中文字时按全文处理；上方四项分析会把它当作补充说明。',
      customPlaceholder: '例如：改写成领导讲话口径，控制在 200 字内，保留全部数据',
      customRun: '执行要求',
      customEmpty: '先写下你的处理要求。',
      customDone: '已按你的要求处理',
      customInstruction: '要求',
      modelLabel: '模型',
      modelFollow: '跟随默认',
      modelLoading: '正在读取模型…',
      modelEmpty: '没有可用的模型',
      modelHint: '更多模型在「设置 → 模型」中配置，支持 OpenAI 兼容的第三方接口。',
      modelError: '模型列表读取失败',
      reasoning: '推理等级',
      reasoningDefault: '默认',
      reasoningLoading: '正在读取推理等级…',
      reasoningNone: '该模型不提供推理等级',
      reasoningFallback: '所选推理等级「{effort}」该模型不接受，本次改用默认等级。',
      wholeDocument: '全文',
      importFormatKept: '已按原文档结构转换为 Markdown（标题、列表、加粗、表格）。',
    }

    const EN = {
      panel: 'ShenYue',
      title: 'ShenYue',
      subtitle: 'Select text on the left; verify, examine, expand, and connect it on the right.',
      importFile: 'Import',
      importPackage: 'Import review package',
      exportPackage: 'Export review package',
      importMenu: 'Import',
      exportMenu: 'Export',
      packageHint: 'One JSON file with the manuscript, every result, and the current instruction. The original .docx/.pdf is not included.',
      packageExportDone: 'Review package exported',
      packageImportTitle: 'Import review package',
      packageImportBody: 'It replaces the current manuscript and results with the {runs} records in this file. This cannot be undone.',
      packageImportConfirm: 'Replace and import',
      packageImportCancel: 'Cancel',
      packageEmpty: 'The package holds no manuscript; it may be damaged.',
      packageInvalid: 'This is not a valid review package.',
      packageTooLarge: 'The package is over about 8 MB; trim the results first.',
      packageFresh: 'Older packages still import, but a few details may be missing.',
      exportMd: 'Export Markdown',
      importedPackage: 'Review package imported',
      exportFile: 'Export',
      wpsMenu: 'WPS edition',
      wpsInstall: 'Install WPS add-in',
      wpsInstallDesc: 'Install ShenYue into local WPS Writer',
      wpsStatus: 'Check install status',
      wpsUninstall: 'Uninstall WPS add-in',
      wpsInstalled: 'WPS add-in installed at {path}. Restart WPS and open the review pane from the ShenYue ribbon tab.',
      wpsUninstalled: 'WPS add-in removed.',
      wpsStatusOk: 'WPS add-in installed (v{version}).',
      wpsNotInstalled: 'The WPS add-in is not installed yet.',
      wpsFailed: 'WPS add-in action failed: {message}',
      copyAll: 'Copy all',
      clear: 'Clear',
      clearArmed: 'Confirm clear',
      titlePlaceholder: 'Untitled document',
      metaChars: 'chars',
      metaParagraphs: 'paragraphs',
      metaMinutes: 'min read',
      emptyTitle: 'No document yet',
      emptyBody: 'Paste Markdown or plain text, import a .docx or .pdf, or load the sample.',
      loadSample: 'Load sample',
      selectionLabel: 'Selected',
      selectionEmpty: 'Nothing selected',
      selectionEmptyBody: 'Drag across a passage on the left to enable the actions.',
      selectionTooLong: 'Long selection: only the first 6000 characters are reviewed.',
      needSelection: 'Select the text to review first.',
      busy: 'The previous request is still running.',
      actionSource: 'Source check',
      actionSourceDesc: 'Verify data, dates, and citations',
      actionLogic: 'Logic review',
      actionLogicDesc: 'Test the argument and contradictions',
      actionExpand: 'Expand',
      actionExpandDesc: 'Add background, mechanism, evidence',
      actionRelate: 'Connection',
      actionRelateDesc: 'Relate it to the rest of the document',
      results: 'Results',
      resultsEmptyTitle: 'Results appear here',
      resultsEmptyBody: 'All four reviews use the current selection; {mod}1–4 run them directly.',
      stop: 'Stop',
      retry: 'Retry',
      remove: 'Delete',
      copy: 'Copy',
      copied: 'Copied to clipboard',
      insert: 'Replace text',
      inserted: 'Replaced in the document',
      appended: 'The passage changed; appended at the end',
      sources: 'References',
      stagePlanning: 'Planning searches…',
      stageSearching: 'Searching',
      stageWriting: 'Writing the result…',
      stagePreparing: 'Preparing…',
      cancelled: 'Stopped',
      failed: 'Not completed',
      seconds: 's',
      expandBlock: 'Replacement text',
      instruction: 'Extra instruction',
      instructionPlaceholder: 'Optional: what should the reviewer focus on?',
      sampleNote: 'Sample text only; it describes no real institution.',
      fileTooLarge: 'The file exceeds 2 MB; split it and import again.',
      fileFailed: 'This file could not be read.',
      cleared: 'Cleared',
      copiedAll: 'Document copied',
      selectionHint: 'Available once text is selected',
      manuscript: 'Document body',
      errService: 'The review service is not ready; restart the app and try again.',
      errAuth: 'Your session expired; reload the page.',
      errNetwork: 'Cannot reach the review service. Check that the local host is running.',
      announceDone: ' finished',
      announceFailed: ' did not finish',
      importing: 'Reading…',
      imported: 'Imported',
      importHint: 'Supports .docx, .pdf, .md, .txt; the import replaces the current document',
      importEmpty: 'This file has no usable text.',
      importTooLarge: 'File too large: docx ≤ 20 MB, PDF ≤ 15 MB, text ≤ 2 MB.',
      importPdfNoText: 'No text layer could be read from this PDF. It may be a scan or use an unsupported font encoding; save it as .docx and import that.',
      importLegacyDoc: 'Legacy .doc is not supported; save it as .docx first.',
      importUnsupported: 'Unsupported file type. Choose .docx, .pdf, .md, or .txt.',
      importNoDeflate: 'This runtime cannot inflate documents, so .docx and .pdf cannot be read.',
      importFailed: 'The file could not be parsed. Check that it is intact.',
      importEncrypted: 'Cannot read this file: it may be a password-protected Word document, or a legacy .doc renamed to .docx. Remove the password and save as .docx, then import again.',
      importZipVariant: 'This .docx uses an unusual archive layout (zip64 or another compression method). Open it in Word and save a fresh copy as .docx.',
      importStructure: 'The file is incomplete: word/document.xml is missing. It may be a converter variant; re-save it from Word as a new .docx.',
      reviewedCount: 'Reviewed',
      jumpTo: 'Locate in text',
      markStale: 'The passage changed; it cannot be located.',
      selectFirst: 'Select text first',
      selectFirstAction: 'Select first paragraph',
      reviewedBlocks: 'Reviewed',
      markResults: '{count} result(s) cover this passage; click to view',
      markJumped: 'Jumped to the matching results',
      tabAll: 'All',
      expandCard: 'Expand',
      collapseCard: 'Collapse',
      emptyFilter: 'Nothing in this category yet',
      resizeLeft: 'Resize the manuscript and review columns',
      resizeRight: 'Resize the review and results columns',
      resizeHint: 'Drag to resize, double-click to reset',
      colDoc: 'Manuscript',
      colReview: 'Review',
      colResults: 'Results',
      modeEdit: 'Edit',
      modeSource: 'Source',
      modePreview: 'Preview',
      blockEditHint: 'Click anywhere to edit in place; it saves when you click away. Table cells edit the same way, Enter adds a line, and formatting lives in Source.',
      blockCommit: 'Done',
      blockCancel: 'Cancel',
      imageAlt: 'Image',
      imageBlocked: 'Image failed to load',
      modeName: 'View mode',
      viewOriginal: 'Original',
      revise: 'Revise',
      reviseLabel: 'Additional instruction',
      revisePlaceholder: 'For example: cut it to 100 words, firmer tone',
      reviseRun: 'Apply',
      reviseEmpty: 'Write the instruction first.',
      cancel: 'Cancel',
      replaceIntent: 'Replace the passage',
      replaceIntentHint: 'A requirement that reads as a rewrite replaces the selection automatically; this switch makes every instruction replace. Undo from the result card.',
      replacedNotice: 'The passage was replaced; undo is available',
      appendedNotice: 'The passage changed, so the text was appended',
      undoReplace: 'Undo replace',
      undoDone: 'Original restored',
      undoStale: 'The text changed, so it cannot be undone',
      revisionBadge: 'Revision',
      originalNote: 'Read-only layout reference; switch to Edit to change the text',
      pagesLoading: 'Rendering the original layout…',
      pagesUnit: 'pages',
      pagesFirst: 'first 8 pages only',
      pageLabel: 'Page {page}',
      closeSheet: 'Close',
      originalMissing: 'The reloaded document has no original file; import it again.',
      actionCustom: 'Follow instructions',
      actionCustomDesc: 'Rewrite the passage exactly as you describe',
      customHint: 'With no selection this applies to the whole document; the four reviews use it as an extra instruction.',
      customPlaceholder: 'For example: rewrite in an executive tone, under 200 words, keep every figure',
      customRun: 'Run instructions',
      customEmpty: 'Write the instruction first.',
      customDone: 'Instructions applied',
      customInstruction: 'Instruction',
      modelLabel: 'Model',
      modelFollow: 'Profile default',
      modelLoading: 'Reading models…',
      modelEmpty: 'No models available',
      modelHint: 'Add more under Settings → Models; OpenAI-compatible gateways are supported.',
      modelError: 'Could not read the model list',
      reasoning: 'Reasoning',
      reasoningDefault: 'Default',
      reasoningLoading: 'Reading reasoning levels…',
      reasoningNone: 'This model exposes no reasoning levels',
      reasoningFallback: 'The model rejected the "{effort}" level; this run used the default.',
      wholeDocument: 'Whole document',
      importFormatKept: 'Converted with the document structure kept (headings, lists, bold, tables).',
    }

    /** The document the panel opens with; also the demo content. */
    const SAMPLE_TITLE = '区域性银行科技金融实践：以园区科技企业服务为例'
    const SAMPLE_TEXT = [
      '近年来，科技金融已成为区域性银行转型的重要方向。截至 2023 年末，我国科技型中小企业贷款余额同比增长 21.9%，高于各项贷款增速 11.3 个百分点，科技金融服务覆盖面持续扩大。',
      '我行立足本地产业园区，围绕科技型企业"轻资产、重研发、长周期"的特点，构建了以知识产权质押、研发贷、投贷联动为主的产品体系。2023 年累计服务科技型企业 1,286 户，科技型企业贷款余额 42.6 亿元，较年初增长 38%，不良率 0.31%。',
      '科技金融的本质是用专业能力识别技术价值，而不是简单地降低授信门槛。因此，我们建立了"技术流"评价体系，将专利质量、研发投入强度、核心团队稳定性纳入授信模型。',
      '从实践效果看，凡是采用了"技术流"评价的客户，其贷款不良率显著低于其他客户，这说明技术评价体系是有效的。下一步，我行将把该体系推广到全部对公客户，并计划在三年内实现科技型企业贷款余额翻两番。',
      '同时也要看到，科技金融的收益实现依赖投贷联动的落地，而我行在股权投资方面的牌照与人才储备仍然不足，短期内难以形成完整的收益闭环。',
    ].join('\n\n')
    const SAMPLE_SELECTION = {
      start: SAMPLE_TEXT.indexOf('从实践效果看'),
      end: SAMPLE_TEXT.indexOf('同时也要看到'),
    }

    /* ------------------------------------------------------------------ *
     * Small utilities
     * ------------------------------------------------------------------ */

    let uid = 0
    function nextId() {
      uid += 1
      return `r${Date.now().toString(36)}${uid}`
    }

    function charCount(text) {
      return text.replace(/\s+/gu, '').length
    }

    function paragraphCount(text) {
      const parts = text.split(/\n\s*\n/u).map(part => part.trim()).filter(part => part !== '')
      return parts.length === 0 && text.trim() !== '' ? 1 : parts.length
    }

    /** The modifier the platform actually uses, for every visible shortcut. */
    const MOD_KEY = /mac|iphone|ipad|ipod/iu.test(navigator.platform || navigator.userAgent) ? '⌘' : 'Ctrl'

    /**
     * Whether a typed requirement asks for a rewrite rather than an answer.
     * Deliberately conservative: a miss only leaves the text untouched, while a
     * hit can be undone from the result card.
     */
    const REWRITE_WORDS = /(改写|重写|改成|改为|换一种|换为|换成|换作|压缩|精简|扩写|润色|调整|翻译|纠错|校对|缩写|扩到|压到|控制在|删到|改短|改长|口语化|书面化|正式化|口径)/u

    function looksLikeRewrite(requirement) {
      return REWRITE_WORDS.test(requirement)
    }

    /** Keep a destructive action armed for a few seconds. */
    function useArmed(timeoutMs = 3200) {
      const [armed, setArmed] = useState(false)
      const timer = useRef(0)
      useEffect(() => () => { if (timer.current) clearTimeout(timer.current) }, [])
      const arm = useCallback(() => {
        if (armed) return true
        setArmed(true)
        timer.current = setTimeout(() => setArmed(false), timeoutMs)
        return false
      }, [armed, timeoutMs])
      const disarm = useCallback(() => {
        if (timer.current) clearTimeout(timer.current)
        setArmed(false)
      }, [])
      return [armed, arm, disarm]
    }

    /** Copy text through the async clipboard, falling back to a temporary field. */
    async function copyText(text) {
      try {
        if (navigator.clipboard?.writeText) {
          await navigator.clipboard.writeText(text)
          return true
        }
      } catch {
        /* fall through to the legacy path */
      }
      try {
        const field = document.createElement('textarea')
        field.value = text
        field.setAttribute('readonly', '')
        field.style.position = 'fixed'
        field.style.top = '-1000px'
        field.style.opacity = '0'
        document.body.appendChild(field)
        field.select()
        const ok = document.execCommand('copy')
        document.body.removeChild(field)
        return ok
      } catch {
        return false
      }
    }

    function hostOf(url) {
      try {
        return new URL(url).hostname.replace(/^www\./u, '')
      } catch {
        return url
      }
    }

    function elapsedSeconds(run) {
      const end = run.finishedAt ?? Date.now()
      return Math.max(1, Math.round((end - run.startedAt) / 1000))
    }

    /** One plain line that stands in for a folded result. */
    function summaryOf(run, t) {
      if (run.status === 'error') return run.error === '' ? t('failed') : run.error
      const body = run.action === 'expand' ? parseExpansion(run.raw).body : ''
      const text = body === '' ? stripMarkers(run.raw) : body
      const lines = text.split('\n')
      const clean = value => value.replace(/^[-*+]\s+/u, '').replace(/^\d+[.)]\s+/u, '').trim()
      // A heading alone says nothing about the result; prefer the first real line.
      const line = lines
        .filter(value => !/^#{1,6}\s/u.test(value.trim()))
        .map(clean)
        .find(value => value !== '')
        ?? lines.map(value => value.replace(/^#{1,6}\s+/u, '')).map(clean).find(value => value !== '')
      if (line === undefined) return ''
      return line
        .replace(/\*\*/gu, '')
        .replace(/`/gu, '')
        .replace(/\[([^\]]+)\]\([^)]+\)/gu, '$1')
    }

    /** Split an expansion into the replaceable body and the notes after it. */
    function parseExpansion(raw) {
      const open = raw.indexOf('<<<TEXT')
      const close = raw.lastIndexOf('>>>')
      if (open < 0 || close < 0 || close < open) {
        return { body: '', notes: stripMarkers(raw) }
      }
      return {
        body: raw.slice(open + '<<<TEXT'.length, close).trim(),
        notes: stripMarkers(raw.slice(close + 3)).trim(),
      }
    }

    /**
     * The text a result would put back into the manuscript. Results that follow
     * the block contract are unambiguous; a forced rewrite falls back to the
     * prose before its explanation section.
     */
    function rewriteBody(raw, force) {
      const parsed = parseExpansion(raw)
      if (parsed.body !== '') return parsed.body
      if (!force) return ''
      const cut = raw.search(/\n\s*(#{1,6}\s*)?(修改说明|变更说明|说明)[:：]?/u)
      return (cut > 0 ? raw.slice(0, cut) : raw).trim()
    }

    function stripMarkers(raw) {
      return raw
        .split('\n')
        .filter(line => !/^\s*<<<TEXT\s*$/u.test(line) && !/^\s*>>>\s*$/u.test(line))
        .join('\n')
        .trim()
    }

    /* ------------------------------------------------------------------ *
     * Document import: plain text, .docx (ZIP + OOXML), .pdf (text layer)
     * ------------------------------------------------------------------ */

    /** Per-format byte ceilings; parsing happens on the main thread. */
    const IMPORT_LIMITS = { text: 2 * 1024 * 1024, docx: 20 * 1024 * 1024, pdf: 15 * 1024 * 1024 }

    const DOCX_PART = 'word/document.xml'

    /** Latin-1 view of bytes for structural scanning; chunked to stay small. */
    function latin1(bytes) {
      const chunk = 0x8000
      const parts = []
      for (let index = 0; index < bytes.length; index += chunk) {
        parts.push(String.fromCharCode.apply(null, bytes.subarray(index, index + chunk)))
      }
      return parts.join('')
    }

    /** Inflate one zlib or raw-deflate stream through the platform codec. */
    async function inflate(bytes, mode) {
      if (typeof DecompressionStream === 'undefined') throw new Error('deflate-unsupported')
      const stream = new Blob([bytes]).stream().pipeThrough(new DecompressionStream(mode))
      return new Uint8Array(await new Response(stream).arrayBuffer())
    }

    function u16(view, at) { return view.getUint16(at, true) }
    function u32(view, at) { return view.getUint32(at, true) }

    /** Read one entry of a ZIP container (the OOXML package format). */
    async function readZipEntry(bytes, wanted) {
      const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength)
      const length = bytes.length
      let eocd = -1
      for (let at = length - 22; at >= 0 && at >= length - 65_557; at -= 1) {
        if (u32(view, at) === 0x0605_4b50) { eocd = at; break }
      }
      if (eocd < 0) throw new Error('zip-invalid')
      const count = u16(view, eocd + 10)
      let cursor = u32(view, eocd + 16)
      for (let entry = 0; entry < count; entry += 1) {
        if (u32(view, cursor) !== 0x0201_4b50) throw new Error('zip-invalid')
        const method = u16(view, cursor + 10)
        const compressed = u32(view, cursor + 20)
        const nameLength = u16(view, cursor + 28)
        const extraLength = u16(view, cursor + 30)
        const commentLength = u16(view, cursor + 32)
        const localOffset = u32(view, cursor + 42)
        const name = latin1(bytes.subarray(cursor + 46, cursor + 46 + nameLength))
        cursor += 46 + nameLength + extraLength + commentLength
        if (name.toLowerCase() !== wanted.toLowerCase()) continue
        if (compressed === 0xffff_ffff) throw new Error('zip64-unsupported')
        if (u32(view, localOffset) !== 0x0403_4b50) throw new Error('zip-invalid')
        const nameAtLocal = u16(view, localOffset + 26)
        const extraAtLocal = u16(view, localOffset + 28)
        const start = localOffset + 30 + nameAtLocal + extraAtLocal
        const raw = bytes.subarray(start, start + compressed)
        if (method === 0) return raw
        if (method === 8) return inflate(raw, 'deflate-raw')
        throw new Error('zip-method')
      }
      throw new Error('zip-missing')
    }

    /** WordprocessingML namespace of the parts this importer reads. */
    const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'

    function tagOf(node) {
      return node.localName ?? String(node.nodeName).replace(/^.*:/u, '')
    }

    function elementKids(node) {
      return [...node.childNodes].filter(child => child.nodeType === 1)
    }

    function wordAttr(node, name) {
      return node.getAttributeNS(W_NS, name) ?? node.getAttribute(`w:${name}`) ?? ''
    }

    /** Run-level text of one paragraph, with the emphasis each run carries. */
    function docxInline(paragraph) {
      const pieces = []
      const visit = (node, style) => {
        for (const child of elementKids(node)) {
          const tag = tagOf(child)
          if (tag === 'r') {
            const next = { ...style }
            for (const prop of elementKids(child)) {
              if (tagOf(prop) !== 'rPr') continue
              for (const flag of elementKids(prop)) {
                const value = wordAttr(flag, 'val')
                const on = value !== '0' && value !== 'false'
                if (tagOf(flag) === 'b' && on) next.bold = true
                if (tagOf(flag) === 'i' && on) next.italic = true
                if (tagOf(flag) === 'strike' && on) next.strike = true
              }
            }
            for (const runChild of elementKids(child)) {
              const runTag = tagOf(runChild)
              if (runTag === 't') pieces.push({ text: runChild.textContent ?? '', ...next })
              else if (runTag === 'br') pieces.push({ text: '\n', ...next })
              else if (runTag === 'tab') pieces.push({ text: '  ', ...next })
            }
            continue
          }
          if (tag === 'pPr' || tag === 'rPr') continue
          visit(child, style)
        }
      }
      visit(paragraph, { bold: false, italic: false, strike: false })
      return pieces
    }

    /** Join neighbouring runs that carry the same emphasis. */
    function mergePieces(pieces) {
      const out = []
      for (const piece of pieces) {
        const last = out[out.length - 1]
        if (last !== undefined
          && last.bold === piece.bold
          && last.italic === piece.italic
          && last.strike === piece.strike) {
          last.text += piece.text
          continue
        }
        out.push({ ...piece })
      }
      return out
    }

    function decorated(text, style) {
      if (text === '' || text === '\n') return text
      const lead = /^\s*/u.exec(text)[0]
      const trail = /\s*$/u.exec(text)[0]
      let core = text.slice(lead.length, text.length - trail.length)
      if (core === '') return text
      if (style.strike) core = `~~${core}~~`
      if (style.italic) core = `*${core}*`
      if (style.bold) core = `**${core}**`
      return `${lead}${core}${trail}`
    }

    /** One paragraph as Markdown, with headings and list markers kept. */
    function docxParagraph(paragraph) {
      let level = 0
      let list = false
      const props = elementKids(paragraph).find(child => tagOf(child) === 'pPr')
      if (props !== undefined) {
        for (const child of elementKids(props)) {
          const tag = tagOf(child)
          if (tag === 'pStyle') {
            const value = wordAttr(child, 'val')
            const heading = /^heading\s*([1-6])$/iu.exec(value) ?? /^([1-6])$/u.exec(value)
            if (heading !== null) level = Number(heading[1])
            else if (/^title$/iu.test(value)) level = 1
            else if (/^subtitle$/iu.test(value)) level = 2
          }
          if (tag === 'numPr') list = true
        }
      }
      const pieces = mergePieces(docxInline(paragraph))
      let text = ''
      for (const piece of pieces) text += decorated(piece.text, piece)
      if (text.trim() === '' && !list) return ''
      const body = text.replace(/^\s+/u, '').replace(/\s+$/u, '')
      if (level > 0) return `${'#'.repeat(level)} ${body}`
      if (list) return `- ${body}`
      return body
    }

    /** One table as a Markdown pipe table; cell paragraphs join with a space. */
    function docxTable(table) {
      const rows = []
      for (const row of elementKids(table)) {
        if (tagOf(row) !== 'tr') continue
        const cells = []
        for (const cell of elementKids(row)) {
          if (tagOf(cell) !== 'tc') continue
          const lines = []
          for (const paragraph of cell.getElementsByTagNameNS('*', 'p')) {
            const line = docxParagraph(paragraph).replace(/^#{1,6}\s+/u, '').replace(/^-\s+/u, '')
            if (line !== '') lines.push(line)
          }
          cells.push(lines.join(' ').replace(/\|/gu, '\\|'))
        }
        if (cells.length > 0) rows.push(cells)
      }
      if (rows.length === 0) return ''
      const width = Math.max(...rows.map(row => row.length))
      const pad = (row) => {
        const copy = row.slice()
        while (copy.length < width) copy.push('')
        return copy
      }
      const lines = rows.map(row => `| ${pad(row).join(' | ')} |`)
      lines.splice(1, 0, `| ${Array.from({ length: width }, () => '---').join(' | ')} |`)
      return lines.join('\n')
    }

    /** The document body as Markdown, preserving headings, lists, and tables. */
    function docxMarkdown(doc) {
      const body = doc.getElementsByTagNameNS('*', 'body')[0]
      if (body === undefined) throw new Error('docx-invalid')
      const blocks = []
      for (const child of elementKids(body)) {
        const tag = tagOf(child)
        if (tag === 'p') {
          const line = docxParagraph(child)
          if (line !== '') blocks.push({ kind: /^-\s/u.test(line) ? 'list' : 'block', text: line })
          continue
        }
        if (tag === 'tbl') {
          const table = docxTable(child)
          if (table !== '') blocks.push({ kind: 'block', text: table })
        }
      }
      let out = ''
      for (const [index, block] of blocks.entries()) {
        const previous = blocks[index - 1]
        const glue = previous !== undefined && previous.kind === 'list' && block.kind === 'list' ? '\n' : '\n\n'
        out += (index === 0 ? '' : glue) + block.text
      }
      return out.replace(/\n{3,}/gu, '\n\n').trim()
    }

    /**
     * Read a .docx package: Markdown text plus the document title Word stored.
     * @returns {{ text: string, title: string }} the converted body and title.
     */
    async function extractDocx(file) {
      if (typeof DecompressionStream === 'undefined') throw new Error('deflate-unsupported')
      const bytes = new Uint8Array(await file.arrayBuffer())
      const documentXml = new TextDecoder('utf-8').decode(await readZipEntry(bytes, DOCX_PART))
      const doc = new DOMParser().parseFromString(documentXml, 'application/xml')
      if (doc.getElementsByTagName('parsererror').length > 0 || doc.documentElement.nodeName === 'parsererror') {
        throw new Error('docx-invalid')
      }
      const text = docxMarkdown(doc)
      if (text.trim() === '') throw new Error('doc-empty')
      let title = ''
      try {
        const coreXml = new TextDecoder('utf-8').decode(await readZipEntry(bytes, 'docProps/core.xml'))
        const core = new DOMParser().parseFromString(coreXml, 'application/xml')
        title = (core.getElementsByTagNameNS('*', 'title')[0]?.textContent ?? '').trim()
      } catch {
        /* the title part is optional */
      }
      return { text, title }
    }

    /** One PDF string: literal escapes, octal codes, and line continuations. */
    function readPdfLiteral(source, start) {
      let index = start + 1
      let depth = 1
      let out = ''
      while (index < source.length) {
        const char = source[index]
        if (char === '\\') {
          const next = source[index + 1]
          if (next === 'n') out += '\n'
          else if (next === 'r') out += '\r'
          else if (next === 't') out += '\t'
          else if (next === 'b') out += '\b'
          else if (next === 'f') out += '\f'
          else if (next === '(' || next === ')' || next === '\\') out += next
          else if (next === '\n') { /* continuation */ }
          else if (next === '\r') { if (source[index + 2] === '\n') index += 1 }
          else if (next >= '0' && next <= '7') {
            let octal = next
            let step = 2
            while (step < 4 && source[index + step] >= '0' && source[index + step] <= '7') {
              octal += source[index + step]
              step += 1
            }
            out += String.fromCharCode(parseInt(octal, 8) & 0xff)
            index += step - 1
          } else out += next
          index += 2
          continue
        }
        if (char === '(') depth += 1
        if (char === ')') {
          depth -= 1
          if (depth === 0) return { value: out, next: index + 1 }
        }
        out += char
        index += 1
      }
      return { value: out, next: index }
    }

    function readPdfHex(source, start) {
      const end = source.indexOf('>', start)
      if (end < 0) return { value: '', next: source.length }
      const digits = source.slice(start + 1, end).replace(/[^0-9A-Fa-f]/gu, '')
      let out = ''
      for (let index = 0; index + 1 < digits.length; index += 2) {
        out += String.fromCharCode(parseInt(digits.slice(index, index + 2), 16))
      }
      return { value: out, next: end + 1 }
    }

    /** `beginbfchar` / `beginbfrange` entries of one ToUnicode CMap. */
    function readCMap(source, map) {
      const hex = (text) => String.fromCharCode(...text.match(/.{1,4}/gu).map(part => parseInt(part, 16) & 0xffff))
      const bfchar = /beginbfchar([\s\S]*?)endbfchar/gu
      for (const block of source.matchAll(bfchar)) {
        const rows = /<([0-9A-Fa-f]+)>\s*<([0-9A-Fa-f]+)>/gu
        for (const row of block[1].matchAll(rows)) {
          const key = row[1].toLowerCase()
          map.set(key, hex(row[2]))
        }
      }
      const bfrange = /beginbfrange([\s\S]*?)endbfrange/gu
      for (const block of source.matchAll(bfrange)) {
        const rows = /<([0-9A-Fa-f]+)>\s*<([0-9A-Fa-f]+)>\s*(<([0-9A-Fa-f]+)>|\[([^\]]*)\])/gu
        for (const row of block[1].matchAll(rows)) {
          const from = parseInt(row[1], 16)
          const to = parseInt(row[2], 16)
          if (row[4] !== undefined) {
            const first = parseInt(row[4], 16)
            for (let code = from; code <= to && code - from < 512; code += 1) {
              map.set(code.toString(16).padStart(row[1].length, '0'), String.fromCharCode(first + code - from))
            }
          } else if (row[5] !== undefined) {
            const items = row[5].match(/<([0-9A-Fa-f]+)>/gu) ?? []
            items.forEach((item, offset) => {
              const code = from + offset
              if (code > to) return
              map.set(code.toString(16).padStart(row[1].length, '0'), hex(item.slice(1, -1)))
            })
          }
        }
      }
    }

    /** Decode one shown string through the collected code map. */
    function decodePdfString(raw, map, twoByte) {
      let out = ''
      if (twoByte) {
        for (let index = 0; index + 1 < raw.length; index += 2) {
          const code = ((raw.charCodeAt(index) << 8) | raw.charCodeAt(index + 1)).toString(16).padStart(4, '0')
          out += map.get(code) ?? ''
        }
      } else {
        for (let index = 0; index < raw.length; index += 1) {
          const code = raw.charCodeAt(index).toString(16).padStart(2, '0')
          out += map.get(code) ?? raw[index]
        }
      }
      return out
    }

    const PDF_OP = /^[A-Za-z'"*]+/u
    const PDF_NUM = /^[-+]?\d*\.?\d+/u

    /** Split one content stream into strings, numbers, arrays, and operators. */
    function tokenizeContent(source) {
      const tokens = []
      let index = 0
      while (index < source.length) {
        const char = source[index]
        if (char === '(') {
          const read = readPdfLiteral(source, index)
          tokens.push({ t: 'str', v: read.value })
          index = read.next
          continue
        }
        if (char === '<' && source[index + 1] === '<') { index += 2; continue }
        if (char === '<') {
          const read = readPdfHex(source, index)
          tokens.push({ t: 'str', v: read.value })
          index = read.next
          continue
        }
        if (char === '>') { index += 1; continue }
        if (char === '[') { tokens.push({ t: 'open' }); index += 1; continue }
        if (char === ']') { tokens.push({ t: 'close' }); index += 1; continue }
        if (char === '%') {
          const stop = source.indexOf('\n', index)
          index = stop < 0 ? source.length : stop + 1
          continue
        }
        if (/\s/u.test(char)) { index += 1; continue }
        const number = PDF_NUM.exec(source.slice(index))
        if (number !== null) {
          tokens.push({ t: 'num', v: Number(number[0]) })
          index += number[0].length
          continue
        }
        const operator = PDF_OP.exec(source.slice(index))
        if (operator !== null) {
          tokens.push({ t: 'op', v: operator[0] })
          index += operator[0].length
          continue
        }
        index += 1
      }
      return tokens
    }

    /**
     * Text pieces and line breaks of one content stream. Each break carries the
     * baseline gap the positioning operator introduced, which separates wrapped
     * lines from real paragraph breaks.
     */
    function pdfContentPieces(source, map, twoByte) {
      const pieces = []
      let operands = []
      let arrayDepth = 0
      let array = null
      let fontSize = 12
      let leading = 14
      let y = 0
      let lastY = null
      let hasText = false

      const breakTo = (gap) => {
        if (hasText || lastY !== null) pieces.push({ kind: 'break', gap, size: fontSize })
        hasText = false
      }
      const show = (raw) => {
        const text = decodePdfString(raw, map, twoByte)
        if (text === '') return
        pieces.push({ kind: 'text', text })
        hasText = true
      }
      const lastString = () => {
        for (let index = operands.length - 1; index >= 0; index -= 1) {
          if (operands[index].t === 'str') return operands[index].v
        }
        return null
      }
      const numbers = () => operands.filter(op => op.t === 'num').map(op => op.v)

      for (const token of tokenizeContent(source)) {
        if (token.t === 'open') { arrayDepth += 1; array = []; continue }
        if (token.t === 'close') {
          arrayDepth -= 1
          if (arrayDepth === 0 && array !== null) { operands.push({ t: 'array', v: array }); array = null }
          continue
        }
        if (arrayDepth > 0) { if (array !== null) array.push(token); continue }
        if (token.t !== 'op') { operands.push(token); continue }

        switch (token.v) {
          case 'Tf': {
            const sizes = numbers()
            if (sizes.length > 0 && sizes[sizes.length - 1] > 0) fontSize = sizes[sizes.length - 1]
            break
          }
          case 'TL': {
            const values = numbers()
            if (values.length > 0) leading = Math.abs(values[values.length - 1]) || leading
            break
          }
          case 'Tm': {
            const values = numbers()
            if (values.length >= 6) {
              const next = values[5]
              breakTo(lastY === null ? 0 : lastY - next)
              y = next
              lastY = next
            }
            break
          }
          case 'Td':
          case 'TD': {
            const values = numbers()
            if (values.length >= 2) {
              const next = y + values[1]
              breakTo(lastY === null ? 0 : lastY - next)
              y = next
              lastY = next
              if (token.v === 'TD') leading = Math.abs(values[1]) || leading
            }
            break
          }
          case 'T*': {
            y -= leading
            breakTo(leading)
            lastY = y
            break
          }
          case 'Tj': {
            const raw = lastString()
            if (raw !== null) show(raw)
            break
          }
          case 'TJ': {
            const list = operands.find(op => op.t === 'array')
            if (list !== undefined) {
              for (const item of list.v) {
                if (item.t === 'str') show(item.v)
                else if (item.t === 'num' && item.v <= -140) show(' ')
              }
            }
            break
          }
          case "'":
          case '"': {
            y -= leading
            breakTo(leading)
            lastY = y
            const raw = lastString()
            if (raw !== null) show(raw)
            break
          }
          case 'BT': {
            if (lastY === null) lastY = 0
            break
          }
          default:
            break
        }
        operands = []
      }
      return pieces
    }

    /** Join pieces into prose: wrapped lines merge, paragraph gaps stay blank lines. */
    function pdfProse(pieces) {
      const isCjk = (char) => /[\u3000-\u303f\u3400-\u4dbf\u4e00-\u9fff\uf900-\ufaff\uff00-\uffef]/u.test(char)
      let out = ''
      let pending = null
      for (const piece of pieces) {
        if (piece.kind === 'break') {
          const size = piece.size > 0 ? piece.size : 12
          const gap = piece.gap
          if (gap <= size * 0.35) continue
          pending = gap > size * 1.55 ? 'para' : 'line'
          continue
        }
        if (out !== '') {
          if (pending === 'para') out += '\n\n'
          else if (pending === 'line') {
            const tail = out[out.length - 1]
            const head = piece.text[0] ?? ''
            if (tail === '-' && /^[a-z]/u.test(head)) out = out.slice(0, -1)
            else if (/[A-Za-z0-9]$/u.test(tail) && /^[A-Za-z0-9]/u.test(head)) out += ' '
            else if (!isCjk(tail) && !isCjk(head) && head !== '') out += ' '
          }
        }
        out += piece.text
        pending = null
      }
      return out
    }

    /** ASCII-plus-CJK ratio; low values mean glyph codes never mapped. */
    function readableRatio(text) {
      if (text === '') return 0
      let printable = 0
      let total = 0
      for (const char of text) {
        if (/\s/u.test(char)) continue
        total += 1
        const code = char.codePointAt(0)
        if (code >= 0x20
          && !(code >= 0x7f && code <= 0x9f)
          && !(code >= 0xe000 && code <= 0xf8ff)
          && !(code >= 0x2000 && code <= 0x200f)) printable += 1
      }
      return total === 0 ? 0 : printable / total
    }

    async function extractPdf(file) {
      if (typeof DecompressionStream === 'undefined') throw new Error('deflate-unsupported')
      const bytes = new Uint8Array(await file.arrayBuffer())
      const source = latin1(bytes)
      if (!source.startsWith('%PDF')) throw new Error('pdf-invalid')

      // 1. Every object stream, in file order, inflated when it is deflated.
      const streams = []
      const pattern = /stream\r?\n/gu
      let total = 0
      for (const match of source.matchAll(pattern)) {
        const start = match.index + match[0].length
        const end = source.indexOf('endstream', start)
        if (end < 0) continue
        const dictionary = source.slice(Math.max(0, match.index - 800), match.index)
        const raw = bytes.subarray(start, end)
        total += raw.length
        if (total > 40 * 1024 * 1024) break
        let decoded
        if (/\/FlateDecode/u.test(dictionary) && !/\/Filter\s*\[/u.test(dictionary)) {
          try {
            decoded = latin1(await inflate(raw, 'deflate'))
          } catch {
            continue
          }
        } else if (/\/Filter/u.test(dictionary)) {
          continue
        } else {
          decoded = latin1(raw)
        }
        streams.push({ text: decoded, dictionary })
      }

      // 2. Glyph-to-Unicode maps first: they apply to the content streams.
      const map = new Map()
      for (const stream of streams) {
        if (stream.text.includes('beginbfchar') || stream.text.includes('beginbfrange')) readCMap(stream.text, map)
      }
      const twoByte = map.size > 0 && [...map.keys()].some(key => key.length >= 4)

      // 3. Content streams: those that show text rather than declare fonts/pages.
      const parts = []
      for (const stream of streams) {
        if (!/\b(?:Tj|TJ)\b/u.test(stream.text)) continue
        if (/beginbfchar|begincmap|\/Type\s*\/Font|\/Subtype\s*\/Type0/u.test(stream.text)) continue
        const prose = pdfProse(pdfContentPieces(stream.text, map, twoByte))
        if (prose.trim() !== '') parts.push(prose)
      }

      const text = parts
        .join('\n\n')
        .replace(/\r/gu, '')
        .replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/gu, '')
        .replace(/[ \t]+\n/gu, '\n')
        .replace(/\n{3,}/gu, '\n\n')
        .trim()
      if (text.length < 20 || readableRatio(text) < 0.6) throw new Error('pdf-no-text-layer')
      return text
    }

    /**
     * Import one chosen file.
     * @param {File} file - the picked file.
     * @returns {Promise<{ text: string, title: string }>} Markdown body and any title the file carries.
     */
    /** Container magic of the two Word formats this importer can meet. */
    async function magicOf(file) {
      const head = new Uint8Array(await file.slice(0, 8).arrayBuffer())
      if (head.length >= 8 && head[0] === 0xd0 && head[1] === 0xcf && head[2] === 0x11 && head[3] === 0xe0) return 'ole'
      if (head.length >= 2 && head[0] === 0x50 && head[1] === 0x4b) return 'zip'
      return 'other'
    }

    async function extractDocument(file) {
      const extension = (file.name.split('.').pop() ?? '').toLowerCase()
      const magic = await magicOf(file)
      if (extension === 'docx' || extension === 'docm') {
        // A password-protected (or renamed legacy) document is not a ZIP at all.
        if (magic === 'ole') throw new Error('docx-encrypted')
        if (magic === 'other') throw new Error('docx-invalid')
        if (file.size > IMPORT_LIMITS.docx) throw new Error('too-large')
        return extractDocx(file)
      }
      if (extension === 'pdf') {
        if (file.size > IMPORT_LIMITS.pdf) throw new Error('too-large')
        return { text: await extractPdf(file), title: '' }
      }
      if (extension === 'doc') throw new Error('legacy-doc')
      if (extension === 'txt' || extension === 'md' || extension === 'markdown' || file.type.startsWith('text/')) {
        if (file.size > IMPORT_LIMITS.text) throw new Error('too-large')
        return { text: (await file.text()).replace(/\r\n/gu, '\n'), title: '' }
      }
      throw new Error('unsupported')
    }

    /** One localized, actionable message per import failure. */
    function importMessage(error, t) {
      switch (error instanceof Error ? error.message : '') {
        case 'too-large': return t('importTooLarge')
        case 'pdf-no-text-layer': return t('importPdfNoText')
        case 'legacy-doc': return t('importLegacyDoc')
        case 'unsupported': return t('importUnsupported')
        case 'deflate-unsupported': return t('importNoDeflate')
        case 'docx-encrypted': return t('importEncrypted')
        case 'zip-invalid':
        case 'zip-missing':
        case 'zip-method':
        case 'zip64-unsupported':
          return t('importZipVariant')
        case 'docx-invalid':
        case 'zip-missing':
          return t('importStructure')
        case 'pdf-invalid':
        case 'doc-empty':
          return t('importFailed')
        default:
          return t('importFailed')
      }
    }

    /* ------------------------------------------------------------------ *
     * Rendered preview of the Markdown the importer produced
     * ------------------------------------------------------------------ */

    /** Inline run of one line, carrying the source range its visible text came from. */
    function inlineRuns(line, base) {
      const runs = []
      const pattern = /(!\[[^\]]*\]\([^)\s]+\)|\*\*[^*]+\*\*|__[^_]+__|~~[^~]+~~|`[^`]+`|\*[^*\n]+\*|_[^_\n]+_|\[[^\]]+\]\([^)\s]+\))/gu
      let cursor = 0
      const plain = (end) => {
        if (end <= cursor) return
        runs.push({ text: line.slice(cursor, end), start: base + cursor, end: base + end })
      }
      for (const match of line.matchAll(pattern)) {
        plain(match.index)
        const raw = match[0]
        const start = base + match.index
        const end = start + raw.length
        if (raw.startsWith('![')) {
          const image = /^!\[([^\]]*)\]\(([^)\s]+)\)$/u.exec(raw)
          if (image !== null) {
            runs.push({ text: '', start, end, image: image[2], alt: image[1] })
          }
        } else if (raw.startsWith('**') || raw.startsWith('__')) {
          runs.push({ text: raw.slice(2, -2), start: start + 2, end: end - 2, bold: true })
        } else if (raw.startsWith('~~')) {
          runs.push({ text: raw.slice(2, -2), start: start + 2, end: end - 2, strike: true })
        } else if (raw.startsWith('`')) {
          runs.push({ text: raw.slice(1, -1), start: start + 1, end: end - 1, code: true })
        } else if (raw.startsWith('*') || raw.startsWith('_')) {
          runs.push({ text: raw.slice(1, -1), start: start + 1, end: end - 1, italic: true })
        } else {
          const link = /^\[([^\]]+)\]\(([^)\s]+)\)$/u.exec(raw)
          if (link !== null) runs.push({ text: link[1], start: start + 1, end: start + 1 + link[1].length, href: link[2] })
        }
        cursor = match.index + raw.length
      }
      plain(line.length)
      return runs
    }

    /** Cut runs at reviewed-block edges so each piece can carry its own band. */
    function splitRunsAtMarks(runs, marks) {
      if (marks.length === 0) return runs.map(run => ({ ...run, mark: -1 }))
      const out = []
      for (const run of runs) {
        let cursor = run.start
        for (const [index, mark] of marks.entries()) {
          if (mark.end <= cursor || mark.start >= run.end) continue
          const from = Math.max(cursor, mark.start)
          const to = Math.min(run.end, mark.end)
          if (from > cursor) {
            out.push({ ...run, text: run.text.slice(cursor - run.start, from - run.start), start: cursor, end: from, mark: -1 })
          }
          out.push({ ...run, text: run.text.slice(from - run.start, to - run.start), start: from, end: to, mark: index })
          cursor = to
        }
        if (cursor < run.end) {
          out.push({ ...run, text: run.text.slice(cursor - run.start), start: cursor, end: run.end, mark: -1 })
        }
      }
      return out
    }

    /** The actions behind one merged mark, in the order they first appear. */
    function markActions(mark) {
      if (mark === undefined || !Array.isArray(mark.runs)) return []
      const seen = []
      for (const run of mark.runs) {
        if (Object.hasOwn(ACTION_COLORS, run.action) && !seen.includes(run.action)) seen.push(run.action)
      }
      return seen
    }

    /** Underline stack: the newest action paints the outermost band. */
    function markShadow(mark) {
      const actions = markActions(mark)
      if (actions.length === 0) return 'inset 0 -2px 0 var(--rd-act-custom)'
      return actions
        .slice()
        .reverse()
        .map((action, index) => `inset 0 -${index * 2 + 1}px 0 var(--rd-act-${action})`)
        .join(', ')
    }

    /** Tint and underline for one mark, shared by the editor band and the rendered runs. */
    function markStyle(mark) {
      const actions = markActions(mark)
      const style = { boxShadow: markShadow(mark) }
      style.background = actions.length === 1
        ? `color-mix(in srgb, var(--rd-act-${actions[0]}) 11%, transparent)`
        : 'color-mix(in srgb, var(--dsw-alias-label-primary) 6%, transparent)'
      return style
    }

    function runsToNodes(runs, keyPrefix, marks, markTitle) {
      return splitRunsAtMarks(runs, marks)
        .filter(run => run.text !== '')
        .map((run, index) => {
          const props = {
            key: `${keyPrefix}-r${index}`,
            'data-o': run.start,
            className: [run.italic && run.bold ? 'rd-em' : '', run.strike ? 'rd-del' : '',
              run.mark >= 0 ? 'rd-mark' : ''].filter(Boolean).join(' ') || undefined,
            ...(run.mark >= 0 ? { 'data-mark': run.mark } : {}),
            ...(run.mark >= 0 && markTitle !== undefined ? { title: markTitle(run.mark) } : {}),
            ...(run.mark >= 0 ? { style: markStyle(marks[run.mark]) } : {}),
          }
          if (run.href !== undefined) {
            return h('a', { ...props, href: run.href, target: '_blank', rel: 'noreferrer noopener' }, run.text)
          }
          const tag = run.code ? 'code' : run.bold ? 'strong' : run.italic ? 'em' : run.strike ? 'del' : 'span'
          return h(tag, props, run.text)
        })
    }

    /** Split table cells while keeping each cell's source offset. */
    function tableCells(line, base) {
      const cells = []
      let start = 0
      let text = ''
      const push = (end) => { cells.push({ text: text.trim(), start: base + start + text.length - text.trimStart().length }); text = '' ; start = end + 1 }
      for (let index = 0; index < line.length; index += 1) {
        if (line[index] === '\\') { text += line[index + 1] ?? ''; index += 1; continue }
        if (line[index] === '|') { push(index); continue }
        text += line[index]
      }
      if (text.trim() !== '') cells.push({ text: text.trim(), start: base + line.length - text.length + text.length - text.trimStart().length })
      return cells.filter(cell => cell.text !== '')
    }

    /** Block structure of the source text, with a source range for every line. */
    function parseBlocks(text) {
      const lines = []
      let offset = 0
      for (const line of text.split('\n')) {
        lines.push({ text: line, start: offset })
        offset += line.length + 1
      }
      const blocks = []
      let index = 0
      const isSpace = (value) => value.trim() === ''
      const isImage = (value) => /^!\[[^\]]*\]\([^)\s]+\)$/u.test(value.trim())
      const startsBlock = (value) => /^ {0,3}(#{1,6}\s|[-*+]\s|\d+[.)]\s|>|-{3,}$)/u.test(value)
        || /^\s*\|/u.test(value)
        || isImage(value)
      const lineEnd = (line) => line.start + line.text.length

      while (index < lines.length) {
        const line = lines[index]
        if (isSpace(line.text)) { index += 1; continue }

        if (/^ {0,3}(-{3,}|\*{3,})$/u.test(line.text)) {
          blocks.push({ kind: 'hr', start: line.start, end: lineEnd(line) })
          index += 1
          continue
        }

        if (isImage(line.text)) {
          const image = /^!\[([^\]]*)\]\(([^)\s]+)\)$/u.exec(line.text.trim())
          blocks.push({
            kind: 'image',
            src: image[2],
            alt: image[1],
            start: line.start + line.text.indexOf('!'),
            end: lineEnd(line),
          })
          index += 1
          continue
        }

        const heading = /^(#{1,6})\s+(.*)$/u.exec(line.text)
        if (heading !== null) {
          blocks.push({
            kind: 'heading',
            level: heading[1].length,
            start: line.start,
            end: lineEnd(line),
            lines: [{ text: heading[2], start: line.start + line.text.length - heading[2].length }],
          })
          index += 1
          continue
        }

        if (/^\s*\|/u.test(line.text)) {
          const rows = []
          let separator = null
          const tableStart = line.start
          let tableEnd = lineEnd(line)
          while (index < lines.length && /^\s*\|/u.test(lines[index].text)) {
            const raw = lines[index]
            tableEnd = lineEnd(raw)
            if (separator === null && /^\s*\|[\s:|-]+\|\s*$/u.test(raw.text)) {
              separator = rows.length
              index += 1
              continue
            }
            rows.push({ cells: tableCells(raw.text, raw.start), header: separator === null })
            index += 1
          }
          if (rows.length > 0) blocks.push({ kind: 'table', rows, start: tableStart, end: tableEnd })
          continue
        }

        const item = /^(\s*)([-*+]|\d+[.)])\s+(.*)$/u.exec(line.text)
        if (item !== null) {
          const ordered = /\d/u.test(item[2])
          const items = []
          const listStart = line.start
          let listEnd = lineEnd(line)
          while (index < lines.length) {
            const next = /^(\s*)([-*+]|\d+[.)])\s+(.*)$/u.exec(lines[index].text)
            if (next === null) break
            listEnd = lineEnd(lines[index])
            items.push({
              text: next[3],
              start: lines[index].start + lines[index].text.length - next[3].length,
            })
            index += 1
          }
          blocks.push({ kind: ordered ? 'ol' : 'ul', items, start: listStart, end: listEnd })
          continue
        }

        if (/^\s*>/u.test(line.text)) {
          const quote = []
          const quoteStart = line.start
          let quoteEnd = lineEnd(line)
          while (index < lines.length && /^\s*>/u.test(lines[index].text)) {
            const raw = lines[index].text.replace(/^\s*>\s?/u, '')
            quoteEnd = lineEnd(lines[index])
            quote.push({ text: raw, start: lines[index].start + lines[index].text.length - raw.length })
            index += 1
          }
          blocks.push({ kind: 'quote', lines: quote, start: quoteStart, end: quoteEnd })
          continue
        }

        const paragraph = []
        const paragraphStart = line.start
        let paragraphEnd = lineEnd(line)
        while (index < lines.length && !isSpace(lines[index].text) && !startsBlock(lines[index].text)) {
          paragraphEnd = lineEnd(lines[index])
          paragraph.push({ text: lines[index].text, start: lines[index].start })
          index += 1
        }
        if (paragraph.length > 0) blocks.push({ kind: 'p', lines: paragraph, start: paragraphStart, end: paragraphEnd })
        else index += 1
      }
      return blocks
    }

    /** Render one block; every visible character keeps its source offset. */
    function renderBlock(block, blockIndex, marks, markTitle) {
      const key = `b${blockIndex}`
      const lineNodes = (lines, prefix) => lines.flatMap((line, lineIndex) => {
        const nodes = runsToNodes(inlineRuns(line.text, line.start), `${prefix}-${lineIndex}`, marks, markTitle)
        return lineIndex === 0 ? nodes : [h('br', { key: `${prefix}-br${lineIndex}` }), ...nodes]
      })

      switch (block.kind) {
        case 'hr':
          return h('hr', { key })
        case 'image':
          return h('figure', { key, className: 'rd-image' },
            h('img', { src: block.src, alt: block.alt, loading: 'lazy' }),
            block.alt === '' ? null : h('figcaption', null, block.alt))
        case 'heading': {
          const level = Math.min(block.level, 4)
          return h(`h${level}`, { key }, lineNodes(block.lines, key))
        }
        case 'quote':
          return h('blockquote', { key }, lineNodes(block.lines, key))
        case 'table': {
          const head = block.rows.filter(row => row.header)
          const body = block.rows.filter(row => !row.header)
          const cell = (value, cellIndex) => h('td', { key: `c${cellIndex}` },
            runsToNodes(inlineRuns(value.text, value.start), `${key}-${cellIndex}`, marks, markTitle))
          return h('div', { key, className: 'rd-table-wrap' },
            h('table', null,
              head.length === 0 ? null : h('thead', null, head.map((row, rowIndex) => h('tr', { key: `h${rowIndex}` },
                row.cells.map((value, cellIndex) => h('th', { key: `c${cellIndex}` },
                  runsToNodes(inlineRuns(value.text, value.start), `${key}-h${rowIndex}-${cellIndex}`, marks, markTitle)))))),
              body.length === 0 ? null : h('tbody', null, body.map((row, rowIndex) => h('tr', { key: `b${rowIndex}` },
                row.cells.map((value, cellIndex) => cell(value, cellIndex)))))))
        }
        case 'ul':
        case 'ol':
          return h(block.kind, { key }, block.items.map((item, itemIndex) => h('li', { key: `i${itemIndex}` },
            runsToNodes(inlineRuns(item.text, item.start), `${key}-${itemIndex}`, marks, markTitle))))
        default:
          return h('p', { key }, lineNodes(block.lines, key))
      }
    }

    /** Markdown for the inline content of one rendered node. */
    function inlineMarkdown(node) {
      let out = ''
      for (const child of node.childNodes) {
        if (child.nodeType === 3) {
          out += child.nodeValue ?? ''
          continue
        }
        if (child.nodeType !== 1) continue
        const tag = child.tagName.toLowerCase()
        if (tag === 'br') out += '\n'
        else if (tag === 'strong' || tag === 'b') out += `**${inlineMarkdown(child)}**`
        else if (tag === 'em' || tag === 'i') out += `*${inlineMarkdown(child)}*`
        else if (tag === 'del' || tag === 's' || tag === 'strike') out += `~~${inlineMarkdown(child)}~~`
        else if (tag === 'code') out += `\`${child.textContent ?? ''}\``
        else if (tag === 'a') out += `[${inlineMarkdown(child)}](${child.getAttribute('href') ?? ''})`
        else if (tag === 'img') out += `![${child.getAttribute('alt') ?? ''}](${child.getAttribute('src') ?? ''})`
        else out += inlineMarkdown(child)
      }
      return out
    }

    /** Markdown for one pipe table, keeping the first row as the header. */
    function tableMarkdown(table) {
      const rows = [...table.querySelectorAll('tr')]
      if (rows.length === 0) return ''
      const cellsOf = row => [...row.children].filter(cell => cell.tagName === 'TH' || cell.tagName === 'TD')
      const cellText = (cell) => inlineMarkdown(cell).replace(/\n+/gu, ' ').replace(/\|/gu, '\\|').trim()
      const lines = rows.map(row => `| ${cellsOf(row).map(cellText).join(' | ')} |`)
      const width = Math.max(1, cellsOf(rows[0]).length)
      const separator = `| ${new Array(width).fill('---').join(' | ')} |`
      return [lines[0], separator, ...lines.slice(1)].join('\n')
    }

    /**
     * Markdown for the block rendered inside `element`, read back from the DOM
     * after an in-place edit. Null when the block carries no editable text.
     */
    function serializeElement(element) {
      const node = element.firstElementChild
      if (node === null) return null
      const tag = node.tagName.toLowerCase()
      if (tag === 'figure' || tag === 'hr') return null
      if (tag === 'table') return tableMarkdown(node)
      if (tag === 'div') {
        const table = node.querySelector('table')
        return table === null ? null : tableMarkdown(table)
      }
      if (tag === 'ul' || tag === 'ol') {
        const ordered = tag === 'ol'
        return [...node.children]
          .filter(child => child.tagName === 'LI')
          .map((item, index) => `${ordered ? `${index + 1}.` : '-'} ${inlineMarkdown(item).replace(/\n+$/u, '')}`)
          .join('\n')
      }
      if (tag === 'blockquote') {
        return inlineMarkdown(node)
          .replace(/\n+$/u, '')
          .split('\n')
          .map(line => (line === '' ? '>' : `> ${line}`))
          .join('\n')
      }
      if (/^h[1-6]$/u.test(tag)) {
        const level = Math.min(Number(tag.slice(1)), 6)
        return `${'#'.repeat(level)} ${inlineMarkdown(node).replace(/\n+$/u, '')}`
      }
      if (tag === 'p') return inlineMarkdown(node).replace(/\n+$/u, '')
      return null
    }

    /** Insert a soft line break at the caret without letting the browser split the block. */
    function insertLineBreak() {
      try {
        if (document.execCommand('insertLineBreak')) return
      } catch {
        /* fall through to the range path */
      }
      const selection = window.getSelection()
      if (selection === null || selection.rangeCount === 0) return
      const range = selection.getRangeAt(0)
      range.deleteContents()
      const br = document.createElement('br')
      range.insertNode(br)
      range.setStartAfter(br)
      range.collapse(true)
      selection.removeAllRanges()
      selection.addRange(range)
    }

    /** Insert pasted text as plain text, never as foreign markup. */
    function insertPlainText(text) {
      try {
        if (document.execCommand('insertText', false, text)) return
      } catch {
        /* fall through to the range path */
      }
      const selection = window.getSelection()
      if (selection === null || selection.rangeCount === 0) return
      const range = selection.getRangeAt(0)
      range.deleteContents()
      const node = document.createTextNode(text)
      range.insertNode(node)
      range.setStartAfter(node)
      range.collapse(true)
      selection.removeAllRanges()
      selection.addRange(range)
    }

    /** The rendered document, read-only; every visible character keeps its source offset. */
    function renderPreview(text, marks, markTitle) {
      return h('div', { className: 'rd-md-preview' },
        parseBlocks(text).map((block, index) => renderBlock(block, index, marks, markTitle)))
    }

    /** Source offset of one selection edge inside the rendered preview. */
    function previewOffset(root, node, offset, edge) {
      if (node === root) {
        const spans = [...root.querySelectorAll('[data-o]')]
        if (spans.length === 0) return null
        if (edge === 'start') return Number(spans[Math.min(offset, spans.length - 1)].getAttribute('data-o'))
        const last = spans[Math.min(offset, spans.length) - 1]
        return Number(last.getAttribute('data-o')) + (last.textContent ?? '').length
      }
      let element = node.nodeType === Node.TEXT_NODE ? node.parentElement : node
      while (element !== null && element !== root && element.getAttribute?.('data-o') === null) {
        element = element.parentElement
      }
      if (element === null || element === root || element.getAttribute?.('data-o') === null) return null
      const base = Number(element.getAttribute('data-o'))
      if (node.nodeType === Node.TEXT_NODE) {
        let before = 0
        for (const child of element.childNodes) {
          if (child === node) break
          before += (child.textContent ?? '').length
        }
        return base + before + offset
      }
      return edge === 'start' ? base : base + (element.textContent ?? '').length
    }

    /* ------------------------------------------------------------------ *
     * Reviewed passages
     * ------------------------------------------------------------------ */

    /**
     * Merge the ranges that still match the current text into the blocks the
     * manuscript paints behind the field.
     * @param runs - review runs, newest first.
     * @param text - the current manuscript.
     * @returns merged, ascending, non-overlapping ranges with their runs.
     */
    function reviewedMarks(runs, text) {
      const ranges = []
      for (const run of runs) {
        if (run.wholeDocument === true) continue
        const selection = run.selection
        if (selection === null || selection === undefined) continue
        const { start, end, text: passage } = selection
        if (!Number.isInteger(start) || !Number.isInteger(end)) continue
        if (end <= start || end > text.length) continue
        if (text.slice(start, end) !== passage) continue
        ranges.push({ start, end, runs: [run] })
      }
      ranges.sort((a, b) => a.start - b.start)
      const merged = []
      for (const range of ranges) {
        const last = merged[merged.length - 1]
        if (last !== undefined && range.start <= last.end) {
          last.end = Math.max(last.end, range.end)
          last.runs.push(...range.runs)
        } else {
          merged.push({ ...range })
        }
      }
      return merged
    }

    /** The manuscript with one transparent span per reviewed block. */
    function renderMarked(text, marks) {
      const nodes = []
      let cursor = 0
      marks.forEach((mark, index) => {
        if (mark.start > cursor) nodes.push(text.slice(cursor, mark.start))
        nodes.push(h('span', {
          key: `mark-${index}`,
          className: 'rd-mark',
          'data-mark': index,
          style: markStyle(mark),
        }, text.slice(mark.start, mark.end)))
        cursor = mark.end
      })
      if (cursor < text.length) nodes.push(text.slice(cursor))
      return nodes
    }

    /* ------------------------------------------------------------------ *
     * Markdown rendering (no library, escaped by construction)
     * ------------------------------------------------------------------ */

    const INLINE = /(\*\*[^*]+\*\*|`[^`]+`|\[[^\]]+\]\((?:https?:\/\/|\/)[^)\s]+\)|~~[^~]+~~)/u

    function renderInline(text, keyPrefix) {
      const nodes = []
      const parts = text.split(INLINE)
      for (const [index, part] of parts.entries()) {
        if (part === '') continue
        const key = `${keyPrefix}-${index}`
        if (part.startsWith('**') && part.endsWith('**') && part.length > 4) {
          nodes.push(h('strong', { key }, part.slice(2, -2)))
          continue
        }
        if (part.startsWith('`') && part.endsWith('`') && part.length > 2) {
          nodes.push(h('code', { key }, part.slice(1, -1)))
          continue
        }
        if (part.startsWith('~~') && part.endsWith('~~') && part.length > 4) {
          nodes.push(h('del', { key }, part.slice(2, -2)))
          continue
        }
        const link = /^\[([^\]]+)\]\(([^)\s]+)\)$/u.exec(part)
        if (link !== null) {
          nodes.push(h('a', {
            key,
            href: link[2],
            target: '_blank',
            rel: 'noreferrer noopener',
          }, link[1]))
          continue
        }
        nodes.push(part)
      }
      return nodes
    }

    /** Render the subset of Markdown the prompts ask for. */
    function renderMarkdown(source, keyPrefix = 'md') {
      const lines = source.replace(/\r\n/gu, '\n').split('\n')
      const blocks = []
      let paragraph = []
      let list = null
      let ordered = false
      let quote = []

      const flushParagraph = () => {
        if (paragraph.length === 0) return
        const text = paragraph.join('\n')
        blocks.push(h('p', { key: `${keyPrefix}-p${blocks.length}` }, renderInline(text, `${keyPrefix}-p${blocks.length}`)))
        paragraph = []
      }
      const flushList = () => {
        if (list === null) return
        const key = `${keyPrefix}-l${blocks.length}`
        blocks.push(h(ordered ? 'ol' : 'ul', { key }, list.map((item, index) => (
          h('li', { key: `${key}-${index}` }, renderInline(item, `${key}-${index}`))
        ))))
        list = null
      }
      const flushQuote = () => {
        if (quote.length === 0) return
        const key = `${keyPrefix}-q${blocks.length}`
        blocks.push(h('blockquote', { key }, renderInline(quote.join(' '), key)))
        quote = []
      }
      const flushAll = () => {
        flushParagraph()
        flushList()
        flushQuote()
      }

      for (const line of lines) {
        const trimmed = line.trim()
        if (trimmed === '') {
          flushAll()
          continue
        }
        if (/^-{3,}$/u.test(trimmed) || /^\*{3,}$/u.test(trimmed)) {
          flushAll()
          blocks.push(h('hr', { key: `${keyPrefix}-hr${blocks.length}` }))
          continue
        }
        const heading = /^(#{1,6})\s+(.*)$/u.exec(trimmed)
        if (heading !== null) {
          flushAll()
          const level = heading[1].length
          blocks.push(h(level <= 2 ? 'h3' : 'h4', { key: `${keyPrefix}-h${blocks.length}` }, renderInline(heading[2], `${keyPrefix}-h${blocks.length}`)))
          continue
        }
        const bullet = /^[-*+]\s+(.*)$/u.exec(trimmed)
        if (bullet !== null) {
          flushParagraph()
          flushQuote()
          if (list === null || ordered) {
            flushList()
            list = []
            ordered = false
          }
          list.push(bullet[1])
          continue
        }
        const numbered = /^\d+[.)]\s+(.*)$/u.exec(trimmed)
        if (numbered !== null) {
          flushParagraph()
          flushQuote()
          if (list === null || !ordered) {
            flushList()
            list = []
            ordered = true
          }
          list.push(numbered[1])
          continue
        }
        if (trimmed.startsWith('>')) {
          flushParagraph()
          flushList()
          quote.push(trimmed.replace(/^>\s?/u, ''))
          continue
        }
        flushList()
        flushQuote()
        paragraph.push(trimmed)
      }
      flushAll()
      return blocks
    }

    /* ------------------------------------------------------------------ *
     * Workspace state, shared across mounts and persisted locally
     * ------------------------------------------------------------------ */

    const listeners = new Set()

    function blankRun(action, selection, extra = {}) {
      return {
        id: nextId(),
        action,
        status: 'running',
        stage: 'preparing',
        raw: '',
        model: '',
        sources: [],
        sourceNote: '',
        claims: [],
        selection: { start: selection.start, end: selection.end, text: selection.text },
        startedAt: Date.now(),
        finishedAt: 0,
        error: '',
        warning: '',
        instruction: typeof extra.instruction === 'string' ? extra.instruction : '',
        wholeDocument: extra.wholeDocument === true,
        effort: typeof extra.effort === 'string' ? extra.effort : '',
        revise: extra.revise === true,
        forceReplace: extra.forceReplace === true,
        parentId: typeof extra.parentId === 'string' ? extra.parentId : '',
        original: typeof extra.original === 'string' ? extra.original : '',
        replaced: null,
      }
    }

    function readStored() {
      const empty = { title: '', text: '', runs: [], activeRunId: '' }
      try {
        const raw = localStorage.getItem(STORAGE_KEY)
        if (raw === null) return empty
        const parsed = JSON.parse(raw)
        if (parsed === null || typeof parsed !== 'object') return empty
        return {
          title: typeof parsed.title === 'string' ? parsed.title : '',
          text: typeof parsed.text === 'string' ? parsed.text : '',
          runs: Array.isArray(parsed.runs)
            ? parsed.runs
              .filter(run => run !== null && typeof run === 'object' && typeof run.id === 'string')
              .map(run => ({
                ...run,
                status: run.status === 'running' ? 'cancelled' : run.status,
                finishedAt: run.status === 'running' ? Date.now() : (run.finishedAt ?? Date.now()),
                selection: run.selection ?? { start: 0, end: 0, text: '' },
                sources: Array.isArray(run.sources) ? run.sources : [],
                claims: Array.isArray(run.claims) ? run.claims : [],
              }))
              .slice(0, 12)
            : [],
          activeRunId: '',
        }
      } catch {
        return empty
      }
    }

    /** Persisted per-browser model choice; `null` follows the profile default. */
    const MODEL_KEY = 'dsh.review-desk.model.v1'

    function readModelChoice() {
      try {
        const raw = localStorage.getItem(MODEL_KEY)
        if (raw === null) return null
        const parsed = JSON.parse(raw)
        if (parsed === null || typeof parsed !== 'object') return null
        const provider = typeof parsed.provider === 'string' ? parsed.provider : ''
        const model = typeof parsed.model === 'string' ? parsed.model : ''
        const reasoningEffort = typeof parsed.reasoningEffort === 'string' ? parsed.reasoningEffort : ''
        if (provider === '' && reasoningEffort === '') return null
        if ((provider === '') !== (model === '')) return null
        return { provider, model, reasoningEffort }
      } catch {
        return null
      }
    }

    /** Wire fields for the chosen route; an empty provider keeps the profile default. */
    function routePayload(choice) {
      if (choice === null) return {}
      const payload = {}
      if (choice.provider !== '') {
        payload.provider = choice.provider
        payload.model = choice.model
      }
      if (choice.reasoningEffort !== '') payload.reasoningEffort = choice.reasoningEffort
      return payload
    }

    const REPLACE_KEY = 'dsh.review-desk.replace.v1'

    function readReplaceIntent() {
      try {
        return localStorage.getItem(REPLACE_KEY) === 'on'
      } catch {
        return false
      }
    }

    function saveReplaceIntent(value) {
      try {
        localStorage.setItem(REPLACE_KEY, value ? 'on' : 'off')
      } catch {
        /* storage is best-effort */
      }
    }

    function saveModelChoice(choice) {
      try {
        if (choice === null) localStorage.removeItem(MODEL_KEY)
        else localStorage.setItem(MODEL_KEY, JSON.stringify(choice))
      } catch {
        /* storage is best-effort */
      }
    }

    /** Column geometry: review and results widths in px; null follows the CSS defaults. */
    const WIDTHS_KEY = 'dsh.review-desk.widths.v1'

    const COLUMN_LIMITS = {
      review: { min: 220, max: 420, initial: 296 },
      results: { min: 300, max: 560, initial: 344 },
      document: 360,
    }

    function readWidths() {
      try {
        const raw = localStorage.getItem(WIDTHS_KEY)
        if (raw === null) return null
        const parsed = JSON.parse(raw)
        if (parsed === null || typeof parsed !== 'object') return null
        const review = Number(parsed.review)
        const results = Number(parsed.results)
        if (!Number.isFinite(review) || !Number.isFinite(results)) return null
        if (review < COLUMN_LIMITS.review.min || results < COLUMN_LIMITS.results.min) return null
        return { review: Math.round(review), results: Math.round(results) }
      } catch {
        return null
      }
    }

    function saveWidths(widths) {
      try {
        if (widths === null) localStorage.removeItem(WIDTHS_KEY)
        else localStorage.setItem(WIDTHS_KEY, JSON.stringify(widths))
      } catch {
        /* storage is best-effort */
      }
    }

    /** Clamp a candidate geometry so the document column keeps its minimum. */
    function fitWidths(candidate, bodyWidth) {
      const budget = bodyWidth - COLUMN_LIMITS.document
      let review = Math.max(COLUMN_LIMITS.review.min, Math.min(COLUMN_LIMITS.review.max, Math.round(candidate.review)))
      let results = Math.max(COLUMN_LIMITS.results.min, Math.min(COLUMN_LIMITS.results.max, Math.round(candidate.results)))
      if (review + results > budget) {
        const overflow = review + results - budget
        const reduceReview = Math.min(overflow, review - COLUMN_LIMITS.review.min)
        review -= reduceReview
        results -= Math.min(overflow - reduceReview, results - COLUMN_LIMITS.results.min)
      }
      return { review, results }
    }

    /** Model catalog of this Host generation, read once when the menu opens. */
    let catalogCache = null

    /** Adapt the Host's model catalog to what the menu renders. */
    function catalogView(value) {
      return {
        default: {
          provider: typeof value?.default?.provider === 'string' ? value.default.provider : '',
          model: typeof value?.default?.model === 'string' ? value.default.model : '',
          reasoningEffort: typeof value?.default?.reasoningEffort === 'string' ? value.default.reasoningEffort : '',
        },
        routes: (Array.isArray(value?.groups) ? value.groups : []).map(group => ({
          provider: String(group.id),
          name: typeof group.name === 'string' && group.name !== '' ? group.name : String(group.id),
          models: (Array.isArray(group.models) ? group.models : []).map(model => ({
            id: String(model.id),
            name: typeof model.name === 'string' && model.name !== '' ? model.name : String(model.id),
            reasoning: model.reasoning === undefined || model.reasoning === null
              ? null
              : {
                efforts: (Array.isArray(model.reasoning.efforts) ? model.reasoning.efforts : [])
                  .filter(effort => typeof effort?.id === 'string' && effort.id !== '')
                  .map(effort => ({
                    id: effort.id,
                    name: typeof effort.name === 'string' && effort.name !== '' ? effort.name : effort.id,
                    description: typeof effort.description === 'string' ? effort.description : '',
                  })),
                defaultEffort: typeof model.reasoning.defaultEffort === 'string' ? model.reasoning.defaultEffort : '',
              },
          })),
        })),
        failures: (Array.isArray(value?.failures) ? value.failures : []).map(failure => ({
          name: typeof failure?.name === 'string' ? failure.name : String(failure?.id ?? ''),
          message: typeof failure?.message === 'string' ? failure.message : '',
        })),
      }
    }

    /**
     * Reasoning options of the route a choice resolves to.
     * @returns `undefined` while the catalog or the route is unknown, `null` when
     * the model advertises no levels, and the metadata otherwise.
     */
    function effortsFor(data, choice) {
      if (data === null) return undefined
      const explicit = choice !== null && choice.provider !== ''
      const provider = explicit ? choice.provider : data.default.provider
      const model = explicit ? choice.model : data.default.model
      const route = data.routes.find(entry => entry.provider === provider)
      const found = route?.models.find(entry => entry.id === model)
      if (found === undefined) return undefined
      return found.reasoning ?? null
    }

    /** Portable review package: one file carries the manuscript and every result. */
    const PACKAGE_FORMAT = 'dsh.review-desk.package'
    const PACKAGE_VERSION = 1
    const PACKAGE_LIMIT = 8 * 1024 * 1024

    /** Normalise one stored or imported run to the fields the panel renders. */
    function normalizeRun(run) {
      if (run === null || typeof run !== 'object' || typeof run.id !== 'string') return null
      const selection = run.selection !== null && typeof run.selection === 'object' ? run.selection : {}
      return {
        ...run,
        status: run.status === 'running' ? 'cancelled' : (typeof run.status === 'string' ? run.status : 'done'),
        finishedAt: typeof run.finishedAt === 'number' ? run.finishedAt : Date.now(),
        startedAt: typeof run.startedAt === 'number' ? run.startedAt : Date.now(),
        stage: '',
        selection: {
          start: Number.isInteger(selection.start) ? selection.start : 0,
          end: Number.isInteger(selection.end) ? selection.end : 0,
          text: typeof selection.text === 'string' ? selection.text : '',
        },
        sources: Array.isArray(run.sources) ? run.sources : [],
        claims: Array.isArray(run.claims) ? run.claims : [],
        replaced: run.replaced !== null && typeof run.replaced === 'object' ? run.replaced : null,
      }
    }

    /** Build the JSON body of an export from the live session. */
    function buildPackage(snapshot) {
      return {
        format: PACKAGE_FORMAT,
        version: PACKAGE_VERSION,
        exportedAt: new Date().toISOString(),
        app: { plugin: 'review-desk', build: BUILD },
        document: { title: snapshot.title, text: snapshot.text },
        runs: snapshot.runs.map(run => ({ ...run, stage: '' })),
        session: { instruction: snapshot.instruction ?? '' },
      }
    }

    /**
     * Validate one parsed package.
     * @returns {{ title: string, text: string, runs: object[], instruction: string } | { error: string }}
     */
    function readPackage(parsed) {
      if (parsed === null || typeof parsed !== 'object') return { error: 'invalid' }
      if (parsed.format !== PACKAGE_FORMAT) return { error: 'invalid' }
      if (typeof parsed.version !== 'number' || parsed.version > PACKAGE_VERSION) return { error: 'invalid' }
      const document = parsed.document
      if (document === null || typeof document !== 'object' || typeof document.text !== 'string') return { error: 'invalid' }
      if (document.text.trim() === '') return { error: 'empty' }
      const runs = (Array.isArray(parsed.runs) ? parsed.runs : []).map(normalizeRun).filter(Boolean).slice(0, 60)
      return {
        title: typeof document.title === 'string' ? document.title : '',
        text: document.text,
        runs,
        instruction: typeof parsed.session?.instruction === 'string' ? parsed.session.instruction : '',
      }
    }

    let state = readStored()
    let saveTimer = 0

    function persist() {
      if (saveTimer) clearTimeout(saveTimer)
      saveTimer = setTimeout(() => {
        saveTimer = 0
        try {
          const payload = { title: state.title, text: state.text, runs: state.runs.slice(0, 12) }
          let serialized = JSON.stringify(payload)
          if (serialized.length > 380_000) {
            payload.runs = payload.runs.slice(0, 3)
            serialized = JSON.stringify(payload)
          }
          if (serialized.length > 1_000_000) return
          localStorage.setItem(STORAGE_KEY, serialized)
        } catch {
          /* storage is best-effort */
        }
      }, 400)
    }

    function notify() {
      for (const listener of listeners) listener()
    }

    const store = {
      get: () => state,
      subscribe(listener) {
        listeners.add(listener)
        return () => { listeners.delete(listener) }
      },
      update(patch) {
        state = { ...state, ...patch }
        notify()
        persist()
      },
      /** Mutate one run in place, then publish a fresh snapshot. */
      patchRun(id, patch, options = {}) {
        const run = state.runs.find(item => item.id === id)
        if (run === undefined) return
        Object.assign(run, patch)
        state = { ...state, runs: state.runs.slice() }
        notify()
        if (options.persist !== false) persist()
      },
      addRun(run) {
        state = { ...state, runs: [run, ...state.runs].slice(0, 20), activeRunId: run.id }
        notify()
        persist()
      },
      removeRun(id) {
        state = {
          ...state,
          runs: state.runs.filter(item => item.id !== id),
          activeRunId: state.activeRunId === id ? '' : state.activeRunId,
        }
        notify()
        persist()
      },
    }

    const useStore = React.useSyncExternalStore === undefined
      ? function useStoreFallback() {
        const [snapshot, setSnapshot] = useState(state)
        useEffect(() => store.subscribe(() => setSnapshot(state)), [])
        return snapshot
      }
      : function useStoreSync() {
        return React.useSyncExternalStore(store.subscribe, store.get, store.get)
      }

    /* ------------------------------------------------------------------ *
     * Host transport
     * ------------------------------------------------------------------ */

    const running = new Map()

    function analyzeUrl() {
      try {
        return new URL(ANALYZE_ROUTE, document.baseURI).toString()
      } catch {
        return `/${ANALYZE_ROUTE}`
      }
    }

    /** Read an NDJSON response body line by line. */
    async function readFrames(response, onFrame) {
      const reader = response.body?.getReader?.()
      const decoder = new TextDecoder()
      const dispatch = (line) => {
        const trimmed = line.trim()
        if (trimmed === '') return
        try {
          onFrame(JSON.parse(trimmed))
        } catch {
          /* a partial or non-JSON line is dropped */
        }
      }
      if (reader === undefined) {
        for (const line of (await response.text()).split('\n')) dispatch(line)
        return
      }
      let buffer = ''
      for (;;) {
        const { done, value } = await reader.read()
        if (done) break
        buffer += decoder.decode(value, { stream: true })
        let index = buffer.indexOf('\n')
        while (index >= 0) {
          dispatch(buffer.slice(0, index))
          buffer = buffer.slice(index + 1)
          index = buffer.indexOf('\n')
        }
      }
      buffer += decoder.decode()
      dispatch(buffer)
    }

    let flushTimer = 0

    function scheduleFlush(id) {
      if (flushTimer) return
      flushTimer = setTimeout(() => {
        flushTimer = 0
        store.patchRun(id, {}, { persist: false })
      }, 70)
    }

    function applyFrame(runId, frame) {
      switch (frame?.type) {
        case 'meta':
          store.patchRun(runId, {
            model: frame.model === '' ? '' : `${frame.provider}/${frame.model}`,
          }, { persist: false })
          break
        case 'stage':
          store.patchRun(runId, { stage: typeof frame.stage === 'string' ? frame.stage : '', stageDetail: frame.detail ?? '' }, { persist: false })
          break
        case 'sources':
          store.patchRun(runId, {
            sources: Array.isArray(frame.sources) ? frame.sources : [],
            sourceNote: typeof frame.note === 'string' ? frame.note : '',
            claims: Array.isArray(frame.claims) ? frame.claims : [],
          }, { persist: false })
          break
        case 'delta':
          if (typeof frame.text === 'string') {
            const run = state.runs.find(item => item.id === runId)
            if (run === undefined) return
            run.raw += frame.text
            scheduleFlush(runId)
          }
          break
        case 'warning':
          store.patchRun(runId, {
            warning: typeof frame.message === 'string' ? frame.message : '',
            warningCode: typeof frame.code === 'string' ? frame.code : '',
            warningEffort: typeof frame.effort === 'string' ? frame.effort : '',
          }, { persist: false })
          break
        case 'done':
          store.patchRun(runId, { status: 'done', finishedAt: Date.now(), stage: '' })
          break
        case 'error':
          store.patchRun(runId, {
            status: 'error',
            error: typeof frame.message === 'string' ? frame.message : '',
            finishedAt: Date.now(),
            stage: '',
          })
          break
        default:
          break
      }
    }

    /** Keep every other run's range pointing at the same characters. */
    function shiftSelections(runs, from, delta, exceptId) {
      if (delta === 0) return runs
      return runs.map(run => {
        if (run.id === exceptId) return run
        const selection = run.selection
        if (selection === null || selection === undefined) return run
        if (!Number.isInteger(selection.start) || selection.start < from) return run
        return { ...run, selection: { ...selection, start: selection.start + delta, end: selection.end + delta } }
      })
    }

    /**
     * Write one result body over the passage it came from, keeping both the
     * replaced original and the new range on the run so the card can undo it.
     * @returns the new range, or null when nothing was written.
     */
    function replacePassage(runId, body, { appendWhenStale }) {
      const run = state.runs.find(item => item.id === runId)
      if (run === undefined || body === '') return null
      const current = state.text
      const selection = run.selection
      let start = -1
      let previous = ''
      const glue = current === '' || current.endsWith('\n') ? '' : '\n\n'
      if (run.wholeDocument === true) {
        start = 0
        previous = current
      } else if (Number.isInteger(selection.start)
        && selection.end <= current.length
        && current.slice(selection.start, selection.end) === selection.text) {
        start = selection.start
        previous = selection.text
      } else if (appendWhenStale) {
        start = current.length + glue.length
        previous = ''
      }
      if (start < 0) return null
      const prefix = appendWhenStale && previous === '' && start > 0 ? glue : ''
      const delta = body.length + prefix.length - previous.length
      const from = prefix === '' ? start : start - prefix.length
      const next = current.slice(0, from) + prefix + body + current.slice(from + previous.length)
      const runs = shiftSelections(state.runs, from + previous.length, delta, runId).map(item => (
        item.id === runId
          ? {
            ...item,
            replaced: { start: from + prefix.length, previous, inserted: body },
            // Remember the first original so later revisions keep a true reference.
            original: item.original !== '' && item.original !== undefined ? item.original : previous,
            selection: { start: from + prefix.length, end: from + prefix.length + body.length, text: body },
          }
          : item
      ))
      state = { ...state, text: next, runs }
      notify()
      persist()
      return { start: from + prefix.length, end: from + prefix.length + body.length }
    }

    /** A custom or revision result that carries a replacement block rewrites its passage. */
    function autoReplace(runId) {
      const run = state.runs.find(item => item.id === runId)
      if (run === undefined || run.replaced !== null && run.replaced !== undefined) return
      if (run.status !== 'done') return
      const forced = run.forceReplace === true || run.revise === true
      if (run.action !== 'custom' && run.revise !== true) return
      const body = rewriteBody(run.raw, forced)
      if (body === '') return
      replacePassage(runId, body, { appendWhenStale: false })
    }

    /** Put the manuscript back the way it was before one replacement. */
    function undoReplace(runId) {
      const run = state.runs.find(item => item.id === runId)
      const record = run?.replaced
      if (record === null || record === undefined) return false
      const current = state.text
      if (current.slice(record.start, record.start + record.inserted.length) !== record.inserted) return false
      const delta = record.previous.length - record.inserted.length
      const next = current.slice(0, record.start) + record.previous + current.slice(record.start + record.inserted.length)
      const runs = shiftSelections(state.runs, record.start + record.inserted.length, delta, runId).map(item => (
        item.id === runId
          ? {
            ...item,
            replaced: null,
            selection: { start: record.start, end: record.start + record.previous.length, text: record.previous },
          }
          : item
      ))
      state = { ...state, text: next, runs }
      notify()
      persist()
      return true
    }

    async function executeRun(runId, payload) {
      const controller = new AbortController()
      running.set(runId, controller)
      try {
        const response = await fetch(analyzeUrl(), {
          method: 'POST',
          credentials: 'include',
          headers: { 'content-type': 'application/json' },
          body: JSON.stringify(payload),
          signal: controller.signal,
        })
        if (!response.ok) {
          let detail = `HTTP ${response.status}`
          try {
            const body = await response.json()
            if (typeof body?.error === 'string') detail = body.error
          } catch {
            /* keep the status text */
          }
          const failure = new Error(detail)
          if (response.status === 404) failure.code = 'service'
          else if (response.status === 401 || response.status === 403) failure.code = 'auth'
          throw failure
        }
        await readFrames(response, frame => applyFrame(runId, frame))
        const run = state.runs.find(item => item.id === runId)
        if (run !== undefined && run.status === 'running') {
          const partial = run.raw.trim() !== ''
          store.patchRun(runId, {
            status: partial ? 'done' : 'error',
            error: partial ? '' : '结果流意外中断，请重试。',
            finishedAt: Date.now(),
            stage: '',
          })
        }
        autoReplace(runId)
      } catch (error) {
        const aborted = error?.name === 'AbortError'
        const code = typeof error?.code === 'string'
          ? error.code
          : (error instanceof TypeError ? 'network' : '')
        store.patchRun(runId, {
          status: aborted ? 'cancelled' : 'error',
          error: aborted ? '' : (error instanceof Error ? error.message : String(error)),
          errorCode: aborted ? '' : code,
          finishedAt: Date.now(),
          stage: '',
        })
      } finally {
        running.delete(runId)
        if (state.activeRunId === runId) {
          state = { ...state, activeRunId: '' }
          notify()
        }
      }
    }

    /* ------------------------------------------------------------------ *
     * Icons — stroke glyphs sized to the text they sit beside
     * ------------------------------------------------------------------ */

    function glyph(paths, size) {
      return h('svg', {
        width: size,
        height: size,
        viewBox: '0 0 24 24',
        fill: 'none',
        stroke: 'currentColor',
        strokeWidth: 1.6,
        strokeLinecap: 'round',
        strokeLinejoin: 'round',
        'aria-hidden': true,
        focusable: false,
      }, paths.map((path, index) => h(path.tag ?? 'path', { key: index, ...path })))
    }

    function IconSource({ size = 16 }) {
      return glyph([
        { tag: 'circle', cx: 10.5, cy: 10.5, r: 6.5 },
        { d: 'M15.6 15.6 21 21' },
        { d: 'M8 10.5h5' },
      ], size)
    }

    function IconLogic({ size = 16 }) {
      return glyph([
        { tag: 'circle', cx: 5.5, cy: 6, r: 2.2 },
        { tag: 'circle', cx: 5.5, cy: 18, r: 2.2 },
        { tag: 'circle', cx: 19, cy: 12, r: 2.2 },
        { d: 'M7.7 6.6h3.1a4 4 0 0 1 4 4v.5' },
        { d: 'M7.7 17.4h3.1a4 4 0 0 0 4-4v-.5' },
        { d: 'M14.8 12h2' },
      ], size)
    }

    function IconExpand({ size = 16 }) {
      return glyph([
        { d: 'M4 6.5h9' },
        { d: 'M4 12h7' },
        { d: 'M4 17.5h9' },
        { d: 'M16.5 8.5v7' },
        { d: 'M13 12h7' },
      ], size)
    }

    function IconRelate({ size = 16 }) {
      return glyph([
        { d: 'M10 8.5H7.2a3.5 3.5 0 0 0 0 7H10' },
        { d: 'M14 8.5h2.8a3.5 3.5 0 0 1 0 7H14' },
        { d: 'M8.4 12h7.2' },
      ], size)
    }

    function IconCustom({ size = 16 }) {
      return glyph([
        { d: 'M5 6.5h8' },
        { d: 'M5 12h6' },
        { d: 'M5 17.5h5' },
        { d: 'M17.5 3.5l1.3 3.2 3.2 1.3-3.2 1.3-1.3 3.2-1.3-3.2L13 8l3.2-1.3z' },
        { d: 'M16 15.5l1 2.2 2.2 1-2.2 1-1 2.2-1-2.2-2.2-1 2.2-1z' },
      ], size)
    }

    function IconPanel({ size = 18 }) {
      return glyph([
        { d: 'M7.2 3.6h6.1L18 8.2v11.1a1.4 1.4 0 0 1-1.4 1.4H7.2a1.4 1.4 0 0 1-1.4-1.4V5a1.4 1.4 0 0 1 1.4-1.4Z' },
        { d: 'M13.3 3.6v4.8h4.7' },
        { d: 'M8.9 14.2l2.1 2.1 4.2-4.5' },
      ], size)
    }

    function IconStop({ size = 13 }) {
      return h('svg', {
        width: size, height: size, viewBox: '0 0 24 24', 'aria-hidden': true, focusable: false,
      }, h('rect', { x: 6, y: 6, width: 12, height: 12, rx: 2.4, fill: 'currentColor' }))
    }

    /** Order shared by the action rows and the ⌘1–⌘4 shortcuts. */
    const ACTION_ORDER = ['source', 'logic', 'expand', 'relate']
    /** Result tabs: the four analyses plus free-form instructions. */
    const ACTION_TABS = [...ACTION_ORDER, 'custom']

    /** Underline palette: one colour per action, shared by marks, tabs, and cards. */
    const ACTION_COLORS = { source: true, logic: true, expand: true, relate: true, custom: true }

    const ACTION_META = {
      source: { icon: IconSource, label: 'actionSource', desc: 'actionSourceDesc', shortcut: '1' },
      logic: { icon: IconLogic, label: 'actionLogic', desc: 'actionLogicDesc', shortcut: '2' },
      expand: { icon: IconExpand, label: 'actionExpand', desc: 'actionExpandDesc', shortcut: '3' },
      relate: { icon: IconRelate, label: 'actionRelate', desc: 'actionRelateDesc', shortcut: '4' },
      custom: { icon: IconCustom, label: 'actionCustom', desc: 'actionCustomDesc', shortcut: '' },
    }

    /* ------------------------------------------------------------------ *
     * Styles, rendered with the page so unmounting removes them
     * ------------------------------------------------------------------ */

    const CSS = `
.rd-page {
  --rd-act-source: #1d6ff2;
  --rd-act-logic: #9333ea;
  --rd-act-expand: #0d9f6e;
  --rd-act-relate: #e2530b;
  --rd-act-custom: #6b7280;
  --rd-inset: clamp(20px, 2.4vw, 32px);
  --rd-col-pad: clamp(14px, 1.4vw, 20px);
  --rd-radius-card: 14px;
  --rd-radius-control: 9px;
  position: relative;
  display: flex;
  flex-direction: column;
  height: 100%;
  min-height: 0;
  overflow: hidden;
  background: var(--dsw-alias-bg-base);
  color: var(--dsw-alias-label-primary);
  font-family: -apple-system, BlinkMacSystemFont, "SF Pro Text", "PingFang SC", "Hiragino Sans GB", "Helvetica Neue", Arial, sans-serif;
  font-size: 14px;
  line-height: 20px;
  -webkit-font-smoothing: antialiased;
  text-rendering: optimizeLegibility;
}
.rd-page *, .rd-page *::before, .rd-page *::after { box-sizing: border-box; }
.rd-page :focus-visible { outline: 2px solid var(--dsw-alias-brand-primary); outline-offset: 2px; border-radius: 7px; }

/* Header ------------------------------------------------------------- */
.rd-header {
  flex: none;
  display: flex;
  align-items: flex-start;
  justify-content: space-between;
  gap: 16px;
  padding: 26px var(--rd-inset) 14px;
}
html[data-platform='darwin'] .rd-header { padding-top: calc(26px + var(--dsh-frame-top-clearance, 0px)); }
.rd-heading { min-width: 0; }
.rd-title {
  margin: 0;
  font-size: 20px;
  line-height: 28px;
  font-weight: 600;
  letter-spacing: -0.01em;
}
.rd-subtitle {
  margin: 3px 0 0;
  font-size: 13px;
  line-height: 20px;
  color: var(--dsw-alias-label-secondary);
}
.rd-toolbar {
  flex: 1 1 auto;
  min-width: 0;
  display: flex;
  flex-wrap: wrap;
  align-items: center;
  justify-content: flex-end;
  gap: 8px;
  padding-top: 2px;
}
.rd-button {
  display: inline-flex;
  align-items: center;
  gap: 6px;
  height: 30px;
  padding: 0 11px;
  border: 0.5px solid var(--dsw-alias-border-l1);
  border-radius: var(--rd-radius-control);
  background: var(--dsw-alias-bg-layer-1);
  color: var(--dsw-alias-label-primary);
  font: inherit;
  font-size: 13px;
  cursor: pointer;
  white-space: nowrap;
  transition: background-color 120ms ease, border-color 120ms ease;
}
.rd-button:hover { background: color-mix(in srgb, var(--dsw-alias-label-primary) 6%, transparent); }
.rd-button[data-variant='primary'] {
  border-color: transparent;
  background: var(--dsw-alias-brand-primary);
  color: var(--dsw-alias-label-primary-foreground);
}
.rd-button[data-variant='primary']:hover { filter: brightness(1.06); }
.rd-button[data-danger] { color: var(--dsw-alias-state-error-primary); }
.rd-button:disabled { opacity: 0.5; cursor: not-allowed; }

/* Body: three columns ------------------------------------------------ */
.rd-body {
  flex: 1;
  min-height: 0;
  overflow: hidden;
  display: grid;
  grid-template-columns:
    minmax(0, 1fr)
    var(--rd-review-w, clamp(252px, 20vw, 296px))
    var(--rd-results-w, clamp(344px, 27vw, 428px));
  grid-template-rows: minmax(0, 1fr);
  border-top: 0.5px solid var(--dsw-alias-border-l2);
  position: relative;
}
.rd-col {
  min-width: 0;
  min-height: 0;
  display: flex;
  flex-direction: column;
  overflow: hidden;
  background: var(--dsw-alias-bg-base);
}
.rd-col-review, .rd-col-results {
  /* A slightly recessed surface separates tools from the manuscript paper. */
  background: color-mix(in srgb, var(--dsw-alias-label-primary) 3.5%, var(--dsw-alias-bg-base));
  border-left: 1px solid var(--dsw-alias-border-l2);
}
.rd-col-head {
  flex: none;
  display: flex;
  align-items: center;
  gap: 8px;
  min-height: 38px;
  padding: 0 var(--rd-col-pad);
  border-bottom: 0.5px solid var(--dsw-alias-border-l2);
  background: color-mix(in srgb, var(--dsw-alias-label-primary) 3%, transparent);
}
.rd-col-title {
  font-size: 11.5px;
  line-height: 16px;
  font-weight: 600;
  letter-spacing: 0.04em;
  color: var(--dsw-alias-label-secondary);
}
.rd-col-meta {
  margin-left: auto;
  display: inline-flex;
  align-items: center;
  gap: 8px;
  font-size: 11.5px;
  line-height: 16px;
  color: var(--dsw-alias-label-secondary);
}
.rd-segmented {
  display: inline-flex;
  gap: 2px;
  padding: 2px;
  border-radius: 7px;
  background: color-mix(in srgb, var(--dsw-alias-label-primary) 6%, transparent);
}
.rd-segment {
  height: 20px;
  padding: 0 8px;
  border: 0;
  border-radius: 5px;
  background: transparent;
  color: var(--dsw-alias-label-secondary);
  font: inherit;
  font-size: 11px;
  cursor: pointer;
}
.rd-segment[data-active] {
  background: var(--dsw-alias-bg-layer-1);
  color: var(--dsw-alias-label-primary);
  box-shadow: 0 0 0 0.5px var(--dsw-alias-border-l2);
}
.rd-segment:disabled { opacity: 0.45; cursor: not-allowed; }

/* A quiet text button for column headers. */
.rd-link-button {
  height: 22px;
  padding: 0 7px;
  border: 0;
  border-radius: 6px;
  background: transparent;
  color: var(--dsw-alias-label-secondary);
  font: inherit;
  font-size: 11.5px;
  cursor: pointer;
}
.rd-link-button:hover { background: color-mix(in srgb, var(--dsw-alias-label-primary) 7%, transparent); color: var(--dsw-alias-label-primary); }

/* Original-layout sheet ------------------------------------------------ */
.rd-sheet {
  position: absolute;
  inset: 0;
  z-index: 50;
  display: flex;
  flex-direction: column;
  background: var(--dsw-alias-bg-base);
  animation: rd-sheet-in 180ms ease-out;
}
@keyframes rd-sheet-in { from { opacity: 0; transform: translateY(6px); } }
.rd-sheet-head {
  flex: none;
  display: flex;
  align-items: center;
  gap: 12px;
  padding: 12px var(--rd-inset);
  border-bottom: 0.5px solid var(--dsw-alias-border-l2);
}
.rd-sheet-heading { min-width: 0; }
.rd-sheet-title { margin: 0; font-size: 15px; line-height: 22px; font-weight: 600; }
.rd-sheet-note { margin: 2px 0 0; font-size: 11.5px; line-height: 16px; color: var(--dsw-alias-label-secondary); }
.rd-sheet-meta { margin-left: auto; font-size: 12px; line-height: 16px; color: var(--dsw-alias-label-secondary); font-variant-numeric: tabular-nums; }
.rd-sheet-close {
  flex: none;
  width: 30px;
  height: 30px;
  border: 0;
  border-radius: 8px;
  background: transparent;
  color: var(--dsw-alias-label-secondary);
  font: inherit;
  font-size: 14px;
  cursor: pointer;
}
.rd-sheet-close:hover { background: color-mix(in srgb, var(--dsw-alias-label-primary) 7%, transparent); color: var(--dsw-alias-label-primary); }
.rd-sheet-body {
  flex: 1;
  min-height: 0;
  overflow-y: auto;
  overscroll-behavior: contain;
  display: flex;
  flex-direction: column;
  align-items: center;
  gap: 18px;
  padding: 24px var(--rd-inset) 32px;
  background: color-mix(in srgb, var(--dsw-alias-label-primary) 5%, var(--dsw-alias-bg-base));
}
.rd-sheet-status { margin: 40px 0; font-size: 13px; line-height: 20px; color: var(--dsw-alias-label-secondary); }
.rd-sheet-error { color: var(--dsw-alias-state-error-primary); max-width: 34em; text-align: center; }
.rd-sheet-page {
  flex: 0 0 auto;
  width: min(100%, 720px);
  margin: 0;
  background: #fff;
  border-radius: 3px;
  box-shadow: 0 2px 18px rgba(0, 0, 0, 0.16);
  overflow: hidden;
}
.rd-sheet-page img { display: block; width: 100%; height: auto; }
.rd-sheet-page figcaption {
  padding: 6px 10px;
  font-size: 11px;
  line-height: 15px;
  color: #6b7280;
  background: #f6f7f9;
  border-top: 1px solid #e5e7eb;
}

/* Manuscript column --------------------------------------------------- */
.rd-doc-body {
  flex: 1;
  min-height: 0;
  display: flex;
  flex-direction: column;
  padding: 22px var(--rd-col-pad) 18px;
}
.rd-doc-inner {
  flex: 1;
  min-height: 0;
  width: 100%;
  margin: 0;
  display: flex;
  flex-direction: column;
}
.rd-doc-title {
  width: 100%;
  border: 0;
  outline: 0;
  padding: 0;
  background: transparent;
  color: inherit;
  font: inherit;
  font-size: 19px;
  line-height: 27px;
  font-weight: 600;
  letter-spacing: -0.01em;
}
.rd-doc-title::placeholder { color: var(--dsw-alias-label-tertiary); opacity: 1; }
.rd-doc-meta {
  margin: 6px 0 0;
  font-size: 12px;
  line-height: 16px;
  color: var(--dsw-alias-label-secondary);
  display: flex;
  flex-wrap: wrap;
  align-items: center;
  gap: 6px;
}
.rd-dot { width: 3px; height: 3px; border-radius: 50%; background: currentColor; opacity: 0.5; }
.rd-meta-mark { color: var(--dsw-alias-label-primary); }

/* The manuscript field and its highlight layer share every metric, so the
   reviewed bands sit exactly behind the glyphs the textarea draws. */
.rd-editor { position: relative; flex: 1; min-height: 0; display: flex; }
.rd-text {
  position: relative;
  z-index: 1;
  flex: 1;
  min-height: 0;
  width: 100%;
  margin: 0;
  padding: 20px 0 24px;
  border: 0;
  outline: 0;
  resize: none;
  background: transparent;
  color: inherit;
  font: inherit;
  font-size: 16px;
  line-height: 1.85;
  letter-spacing: 0.012em;
  white-space: pre-wrap;
  overflow-wrap: break-word;
  scrollbar-width: none;
}
.rd-text::-webkit-scrollbar { width: 0; height: 0; display: none; }
.rd-text::placeholder { color: var(--dsw-alias-label-tertiary); opacity: 1; }
.rd-highlights {
  position: absolute;
  inset: 0;
  overflow: hidden;
  pointer-events: none;
  z-index: 0;
}
.rd-highlights-inner {
  position: relative;
  padding: 20px 0 24px;
  font: inherit;
  font-size: 16px;
  line-height: 1.85;
  letter-spacing: 0.012em;
  white-space: pre-wrap;
  overflow-wrap: break-word;
  color: transparent;
  -webkit-text-fill-color: transparent;
}
.rd-highlights-inner .rd-mark {
  border-radius: 4px;
  background: color-mix(in srgb, var(--dsw-alias-brand-primary) 10%, transparent);
  box-shadow: inset 0 -1px 0 color-mix(in srgb, var(--dsw-alias-brand-primary) 40%, transparent);
}

/* Action colour dots -------------------------------------------------- */
.rd-act-dot {
  display: inline-block;
  flex: none;
  width: 6px;
  height: 6px;
  margin-right: 5px;
  border-radius: 50%;
  vertical-align: 1px;
  box-shadow: 0 0 0 0.5px color-mix(in srgb, #000 14%, transparent);
}
.rd-tab .rd-act-dot { width: 5px; height: 5px; margin-right: 4px; }

/* Block editing ------------------------------------------------------- */
.rd-blocks { position: relative; }
.rd-block-hint {
  margin: 0 0 10px;
  font-size: 11.5px;
  line-height: 17px;
  color: var(--dsw-alias-label-secondary);
}
.rd-block {
  position: relative;
  padding: 1px 8px 1px 10px;
  margin: 0 -8px 2px -10px;
  border-radius: 8px;
  cursor: text;
  transition: background-color 120ms ease;
}
.rd-block:hover { background: color-mix(in srgb, var(--dsw-alias-label-primary) 4%, transparent); }
.rd-block > :first-child { margin-top: 0; }
.rd-block > :last-child { margin-bottom: 0; }
.rd-block[contenteditable='true'] { outline: none; }
.rd-block[contenteditable='true']:focus {
  background: color-mix(in srgb, var(--dsw-alias-brand-primary) 6%, transparent);
  box-shadow: inset 0 0 0 1px color-mix(in srgb, var(--dsw-alias-brand-primary) 26%, transparent);
}
.rd-block[contenteditable='false'] { cursor: default; }
.rd-block table { cursor: text; }
.rd-image { margin: 14px 0; }
.rd-image img { display: block; max-width: 100%; height: auto; border-radius: 8px; }
.rd-image figcaption {
  margin-top: 6px;
  font-size: 11.5px;
  line-height: 17px;
  color: var(--dsw-alias-label-secondary);
  text-align: center;
}

/* Rendered preview ---------------------------------------------------- */
.rd-preview {
  flex: 1;
  min-width: 0;
  min-height: 0;
  overflow-y: auto;
  overscroll-behavior: contain;
  padding: 14px 0 28px;
  font-size: 16px;
  line-height: 1.85;
  letter-spacing: 0.012em;
}
.rd-preview:focus-visible { outline: 2px solid var(--dsw-alias-brand-primary); outline-offset: 4px; border-radius: 6px; }
.rd-md-preview { color: inherit; overflow-wrap: anywhere; }
.rd-md-preview > :first-child { margin-top: 0; }
.rd-md-preview h1 { margin: 18px 0 10px; font-size: 23px; line-height: 32px; font-weight: 600; letter-spacing: -0.01em; }
.rd-md-preview h2 { margin: 20px 0 9px; font-size: 19px; line-height: 28px; font-weight: 600; letter-spacing: -0.005em; }
.rd-md-preview h3 { margin: 18px 0 8px; font-size: 17px; line-height: 25px; font-weight: 600; }
.rd-md-preview h4 { margin: 16px 0 6px; font-size: 16px; line-height: 24px; font-weight: 600; }
.rd-md-preview p { margin: 10px 0; }
.rd-md-preview ul, .rd-md-preview ol { margin: 10px 0; padding-left: 24px; }
.rd-md-preview li { margin: 4px 0; }
.rd-md-preview blockquote {
  margin: 12px 0;
  padding-left: 14px;
  border-left: 2px solid var(--dsw-alias-border-l2);
  color: var(--dsw-alias-label-secondary);
}
.rd-md-preview hr { margin: 18px 0; border: 0; border-top: 0.5px solid var(--dsw-alias-border-l2); }
.rd-md-preview em { font-style: italic; }
.rd-md-preview del { opacity: 0.65; }
.rd-md-preview code {
  padding: 1px 5px;
  border-radius: 5px;
  background: color-mix(in srgb, var(--dsw-alias-label-primary) 6%, transparent);
  font-family: ui-monospace, SFMono-Regular, "SF Mono", Menlo, monospace;
  font-size: 0.92em;
}
.rd-md-preview a {
  color: var(--dsw-alias-label-primary);
  text-decoration: underline;
  text-underline-offset: 2px;
  text-decoration-color: color-mix(in srgb, var(--dsw-alias-label-primary) 45%, transparent);
}
.rd-md-preview .rd-mark {
  border-radius: 4px;
  background: color-mix(in srgb, var(--dsw-alias-brand-primary) 8%, transparent);
  box-shadow: inset 0 -1px 0 color-mix(in srgb, var(--dsw-alias-brand-primary) 40%, transparent);
  cursor: pointer;
}
.rd-md-preview .rd-mark:hover {
  background: color-mix(in srgb, var(--dsw-alias-brand-primary) 24%, transparent);
}
.rd-table-wrap { margin: 14px 0; max-width: 100%; overflow-x: auto; }
.rd-md-preview table { border-collapse: collapse; width: 100%; font-size: 14px; line-height: 21px; }
.rd-md-preview th, .rd-md-preview td {
  border: 0.5px solid var(--dsw-alias-border-l2);
  padding: 6px 9px;
  text-align: left;
  vertical-align: top;
}
.rd-md-preview th {
  background: color-mix(in srgb, var(--dsw-alias-label-primary) 4%, transparent);
  font-weight: 600;
}

/* Review column ------------------------------------------------------- */
.rd-review-body {
  flex: 1;
  min-width: 0;
  min-height: 0;
  overflow-y: scroll;
  overscroll-behavior: contain;
  scrollbar-gutter: stable;
  padding: 16px var(--rd-col-pad) 20px;
  display: flex;
  flex-direction: column;
  gap: 14px;
}
.rd-selection {
  flex: 0 0 auto;
  border-radius: 12px;
  border: 0.5px solid var(--dsw-alias-border-l2);
  background: color-mix(in srgb, var(--dsw-alias-label-primary) 3%, transparent);
  padding: 11px 13px 12px;
}
.rd-selection[data-empty] { background: transparent; border-style: dashed; }
.rd-selection-head {
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: 8px;
  font-size: 12px;
  line-height: 16px;
  color: var(--dsw-alias-label-secondary);
}
.rd-selection-label { display: inline-flex; align-items: center; gap: 6px; font-weight: 500; }
.rd-badge-dot {
  width: 6px; height: 6px; border-radius: 50%;
  background: var(--dsw-alias-brand-primary);
}
.rd-selection-text {
  margin: 7px 0 0;
  font-size: 13px;
  line-height: 20px;
  display: -webkit-box;
  -webkit-line-clamp: 4;
  -webkit-box-orient: vertical;
  overflow: hidden;
  white-space: pre-wrap;
}
.rd-selection-empty { margin: 7px 0 0; font-size: 12.5px; line-height: 18px; color: var(--dsw-alias-label-secondary); }
.rd-inline-note { margin: 7px 0 0; font-size: 11.5px; line-height: 16px; color: var(--dsw-alias-state-warn-primary); }

.rd-action-group {
  flex: 0 0 auto;
  border: 0.5px solid var(--dsw-alias-border-l2);
  border-radius: 12px;
  overflow: hidden;
  background: var(--dsw-alias-bg-layer-1);
}
.rd-action {
  display: grid;
  grid-template-columns: 18px minmax(0, 1fr) auto;
  grid-template-rows: auto auto;
  align-items: center;
  column-gap: 10px;
  row-gap: 1px;
  width: 100%;
  padding: 10px 12px;
  border: 0;
  background: transparent;
  color: inherit;
  font: inherit;
  text-align: left;
  cursor: pointer;
  transition: background-color 120ms ease;
}
.rd-action + .rd-action { border-top: 0.5px solid var(--dsw-alias-border-l1); }
.rd-action:hover:not(:disabled) { background: color-mix(in srgb, var(--dsw-alias-label-primary) 5%, transparent); }
.rd-action:focus-visible { outline-offset: -2px; }
.rd-action:active:not(:disabled) { background: color-mix(in srgb, var(--dsw-alias-label-primary) 8%, transparent); }
.rd-action:disabled { cursor: not-allowed; }
.rd-action-icon { grid-row: 1 / span 2; display: inline-flex; color: var(--dsw-alias-label-primary); }
.rd-action-label { font-size: 13.5px; line-height: 20px; font-weight: 500; }
.rd-action:disabled .rd-action-label { color: var(--dsw-alias-label-secondary); }
.rd-action-kbd { font-size: 11px; line-height: 16px; color: var(--dsw-alias-label-secondary); font-variant-numeric: tabular-nums; }
.rd-action-desc { grid-column: 2 / span 2; font-size: 11.5px; line-height: 16px; color: var(--dsw-alias-label-secondary); }

.rd-instruction-block { flex: 0 0 auto; display: flex; flex-direction: column; gap: 6px; }
.rd-instruct-label { font-size: 12px; line-height: 16px; color: var(--dsw-alias-label-secondary); }
.rd-instruct-input {
  width: 100%;
  min-height: 64px;
  padding: 8px 10px;
  border: 0.5px solid var(--dsw-alias-border-l2);
  border-radius: var(--rd-radius-control);
  background: var(--dsw-alias-bg-layer-1);
  color: inherit;
  font: inherit;
  font-size: 12.5px;
  line-height: 19px;
  resize: vertical;
}
.rd-instruct-input::placeholder { color: var(--dsw-alias-label-tertiary); opacity: 1; }
.rd-run-custom {
  display: inline-flex;
  align-items: center;
  justify-content: center;
  gap: 8px;
  height: 34px;
  border: 0;
  border-radius: var(--rd-radius-control);
  background: var(--dsw-alias-brand-primary);
  color: var(--dsw-alias-label-primary-foreground);
  font: inherit;
  font-size: 13px;
  font-weight: 500;
  cursor: pointer;
}
.rd-run-custom:disabled { opacity: 0.45; cursor: not-allowed; }
.rd-run-custom-key { font-size: 11px; opacity: 0.75; }
.rd-instruction-hint { margin: 0; font-size: 11px; line-height: 16px; color: var(--dsw-alias-label-secondary); }
.rd-check {
  display: inline-flex;
  align-items: center;
  gap: 6px;
  font-size: 12px;
  line-height: 18px;
  color: var(--dsw-alias-label-primary);
  cursor: pointer;
}
.rd-check input {
  width: 14px;
  height: 14px;
  margin: 0;
  accent-color: var(--dsw-alias-brand-primary);
  cursor: pointer;
}

/* Scrollbars: always visible, theme-aware ------------------------------ */
.rd-review-body,
.rd-results-body,
.rd-preview,
.rd-sheet-body {
  scrollbar-width: thin;
  scrollbar-color: color-mix(in srgb, var(--dsw-alias-label-primary) 26%, transparent) transparent;
}
.rd-review-body::-webkit-scrollbar,
.rd-results-body::-webkit-scrollbar,
.rd-preview::-webkit-scrollbar,
.rd-sheet-body::-webkit-scrollbar {
  width: 10px;
  height: 10px;
}
.rd-review-body::-webkit-scrollbar-track,
.rd-results-body::-webkit-scrollbar-track,
.rd-preview::-webkit-scrollbar-track,
.rd-sheet-body::-webkit-scrollbar-track {
  background: transparent;
}
.rd-review-body::-webkit-scrollbar-thumb,
.rd-results-body::-webkit-scrollbar-thumb,
.rd-preview::-webkit-scrollbar-thumb,
.rd-sheet-body::-webkit-scrollbar-thumb {
  border: 3px solid transparent;
  border-radius: 999px;
  background: color-mix(in srgb, var(--dsw-alias-label-primary) 26%, transparent);
  background-clip: padding-box;
}
.rd-review-body::-webkit-scrollbar-thumb:hover,
.rd-results-body::-webkit-scrollbar-thumb:hover,
.rd-preview::-webkit-scrollbar-thumb:hover,
.rd-sheet-body::-webkit-scrollbar-thumb:hover {
  background: color-mix(in srgb, var(--dsw-alias-label-primary) 40%, transparent);
  background-clip: padding-box;
}

.rd-revise {
  margin: 12px -1px -2px;
  padding-top: 10px;
  border-top: 0.5px solid var(--dsw-alias-border-l1);
  display: flex;
  flex-direction: column;
  gap: 7px;
}
.rd-revise-label { font-size: 11.5px; line-height: 16px; color: var(--dsw-alias-label-secondary); }
.rd-revise-actions { display: flex; align-items: center; gap: 8px; }
.rd-revise-run { height: 30px; padding: 0 14px; font-size: 12.5px; }

/* Column separators --------------------------------------------------- */
.rd-resizer {
  position: absolute;
  top: 0;
  bottom: 0;
  width: 11px;
  margin-left: -5.5px;
  border: 0;
  padding: 0;
  background: transparent;
  cursor: col-resize;
  touch-action: none;
  z-index: 6;
  outline: none;
}
.rd-resizer-left {
  left: calc(100% - var(--rd-results-w, clamp(344px, 27vw, 428px)) - var(--rd-review-w, clamp(252px, 20vw, 296px)));
}
.rd-resizer-right {
  left: calc(100% - var(--rd-results-w, clamp(344px, 27vw, 428px)));
}
.rd-resizer::after {
  content: '';
  position: absolute;
  top: 0;
  bottom: 0;
  left: 5px;
  width: 1px;
  background: transparent;
  transition: background-color 120ms ease, opacity 120ms ease;
}
.rd-resizer:hover::after,
.rd-resizer[data-active]::after {
  background: var(--dsw-alias-label-primary);
  opacity: 0.32;
}
.rd-resizer:focus-visible::after {
  left: 4px;
  width: 2px;
  background: var(--dsw-alias-brand-primary);
  opacity: 0.85;
}
.rd-page[data-resizing] { user-select: none; }
.rd-page[data-resizing] .rd-resizer { pointer-events: none; }

/* Result tabs --------------------------------------------------------- */
.rd-tabs {
  flex: none;
  display: flex;
  flex-wrap: wrap;
  gap: 4px;
  padding: 8px var(--rd-col-pad);
  border-bottom: 0.5px solid var(--dsw-alias-border-l1);
}
.rd-tabs::-webkit-scrollbar { display: none; }
.rd-tab {
  flex: none;
  display: inline-flex;
  align-items: baseline;
  gap: 5px;
  height: 24px;
  padding: 0 9px;
  border: 0.5px solid transparent;
  border-radius: 999px;
  background: transparent;
  color: var(--dsw-alias-label-secondary);
  font: inherit;
  font-size: 12px;
  cursor: pointer;
  white-space: nowrap;
}
.rd-tab:hover:not(:disabled) { background: color-mix(in srgb, var(--dsw-alias-label-primary) 6%, transparent); }
.rd-tab[data-active] {
  border-color: var(--dsw-alias-border-l2);
  background: var(--dsw-alias-bg-base);
  color: var(--dsw-alias-label-primary);
  font-weight: 500;
}
.rd-tab:disabled { opacity: 0.4; cursor: not-allowed; }
.rd-tab-count { font-size: 10.5px; color: inherit; opacity: 0.65; font-variant-numeric: tabular-nums; }

/* Results column ------------------------------------------------------ */
.rd-results-body {
  flex: 1;
  min-width: 0;
  min-height: 0;
  overflow-x: hidden;
  overflow-y: scroll;
  overscroll-behavior: contain;
  scrollbar-gutter: stable;
  display: flex;
  flex-direction: column;
  gap: 12px;
  padding: 16px var(--rd-col-pad) 24px;
  scroll-padding-top: 12px;
}

/* Toolbar menus ------------------------------------------------------- */
.rd-tmenu { position: relative; }
.rd-tmenu-caret { font-size: 10px; line-height: 16px; color: var(--dsw-alias-label-secondary); }
.rd-tmenu-list {
  top: calc(100% + 6px);
  right: 0;
  left: auto;
  width: 248px;
  z-index: 40;
}

/* Import confirmation ------------------------------------------------- */
.rd-mask {
  position: absolute;
  inset: 0;
  z-index: 60;
  display: flex;
  align-items: center;
  justify-content: center;
  padding: 24px;
  background: color-mix(in srgb, #000 32%, transparent);
}
.rd-dialog {
  width: min(100%, 430px);
  padding: 20px 22px 18px;
  border: 0.5px solid var(--dsw-alias-border-l2);
  border-radius: 14px;
  background: var(--dsw-alias-bg-overlay);
  box-shadow: 0 18px 50px rgba(0, 0, 0, 0.28);
}
.rd-dialog-title { margin: 0 0 8px; font-size: 15px; line-height: 22px; font-weight: 600; }
.rd-dialog-body { margin: 0 0 6px; font-size: 13px; line-height: 20px; color: var(--dsw-alias-label-primary); }
.rd-dialog-note { margin: 0 0 16px; font-size: 11.5px; line-height: 17px; color: var(--dsw-alias-label-secondary); }
.rd-dialog-actions { display: flex; justify-content: flex-end; gap: 8px; }

/* Model menu ---------------------------------------------------------- */
.rd-model { position: relative; }
.rd-model-button {
  display: inline-flex;
  align-items: baseline;
  gap: 6px;
  max-width: 230px;
  height: 30px;
  padding: 0 10px;
  border: 0.5px solid var(--dsw-alias-border-l2);
  border-radius: var(--rd-radius-control);
  background: var(--dsw-alias-bg-layer-1);
  color: inherit;
  font: inherit;
  cursor: pointer;
}
.rd-model-button:hover { background: color-mix(in srgb, var(--dsw-alias-label-primary) 6%, transparent); }
.rd-model-caption { flex: none; font-size: 11.5px; color: var(--dsw-alias-label-secondary); }
.rd-model-name { min-width: 0; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; font-size: 12.5px; font-weight: 500; }
.rd-model-effort {
  flex: none;
  max-width: 92px;
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
  font-size: 11px;
  line-height: 16px;
  color: var(--dsw-alias-label-secondary);
}
.rd-model-caret { flex: none; font-size: 10px; line-height: 16px; color: var(--dsw-alias-label-secondary); }
.rd-menu {
  position: absolute;
  top: calc(100% + 6px);
  right: 0;
  z-index: 30;
  width: 292px;
  max-height: 64vh;
  display: flex;
  flex-direction: column;
  border: 0.5px solid var(--dsw-alias-border-l2);
  border-radius: 12px;
  background: var(--dsw-alias-bg-overlay);
  box-shadow: 0 12px 32px rgba(0, 0, 0, 0.18);
  overflow: hidden;
}
.rd-menu-scroll { overflow-y: auto; padding: 6px; }
.rd-menu-group {
  margin: 8px 8px 4px;
  font-size: 10.5px;
  line-height: 15px;
  letter-spacing: 0.04em;
  color: var(--dsw-alias-label-secondary);
}
.rd-menu-row {
  position: relative;
  display: flex;
  align-items: baseline;
  gap: 8px;
  width: 100%;
  padding: 7px 8px 7px 26px;
  border: 0;
  border-radius: 7px;
  background: transparent;
  color: inherit;
  font: inherit;
  text-align: left;
  cursor: pointer;
}
.rd-menu-row:hover { background: color-mix(in srgb, var(--dsw-alias-label-primary) 6%, transparent); }
.rd-menu-row[data-checked]::before {
  content: '✓';
  position: absolute;
  left: 9px;
  font-size: 12px;
  color: var(--dsw-alias-label-primary);
}
.rd-menu-label { font-size: 13px; line-height: 19px; }
.rd-menu-sub {
  margin-left: auto;
  font-size: 10.5px;
  line-height: 16px;
  color: var(--dsw-alias-label-secondary);
  font-variant-numeric: tabular-nums;
  white-space: nowrap;
}
.rd-menu-note { margin: 6px 8px; font-size: 12px; line-height: 18px; color: var(--dsw-alias-label-secondary); }
.rd-menu-hint {
  margin: 0;
  padding: 8px 12px;
  border-top: 0.5px solid var(--dsw-alias-border-l1);
  font-size: 10.5px;
  line-height: 15px;
  color: var(--dsw-alias-label-secondary);
}

/* Result cards ------------------------------------------------------- */
.rd-card {
  flex: 0 0 auto;
  min-width: 0;
  max-width: 100%;
  border: 0.5px solid var(--dsw-alias-border-l2);
  border-radius: var(--rd-radius-card);
  background: var(--dsw-alias-bg-layer-1);
  overflow: hidden;
}
.rd-card[data-status='error'] { border-color: color-mix(in srgb, var(--dsw-alias-state-error-primary) 45%, var(--dsw-alias-border-l1)); }
.rd-card[data-highlighted] {
  border-color: var(--dsw-alias-brand-primary);
  box-shadow: 0 0 0 2px color-mix(in srgb, var(--dsw-alias-brand-primary) 32%, transparent);
}
.rd-card[data-highlighted] .rd-card-head {
  background: color-mix(in srgb, var(--dsw-alias-brand-primary) 9%, transparent);
}
.rd-card-head {
  display: flex;
  align-items: center;
  gap: 7px;
  width: 100%;
  padding: 9px 10px 9px 8px;
  border: 0;
  border-bottom: 0.5px solid var(--dsw-alias-border-l1);
  background: transparent;
  color: inherit;
  font: inherit;
  text-align: left;
  cursor: pointer;
}
.rd-card-head:hover { background: color-mix(in srgb, var(--dsw-alias-label-primary) 4%, transparent); }
.rd-card[data-folded] .rd-card-head { border-bottom: 0; }
.rd-card-chevron {
  flex: none;
  width: 12px;
  font-size: 9px;
  line-height: 16px;
  color: var(--dsw-alias-label-secondary);
}
.rd-card-chip {
  flex: none;
  padding: 1px 6px;
  border-radius: 5px;
  background: color-mix(in srgb, var(--dsw-alias-label-primary) 7%, transparent);
  font-size: 10.5px;
  line-height: 15px;
  color: var(--dsw-alias-label-secondary);
  white-space: nowrap;
}
.rd-card-summary {
  display: -webkit-box;
  width: 100%;
  margin: 0;
  padding: 0 12px 11px 27px;
  border: 0;
  background: transparent;
  font: inherit;
  font-size: 12.5px;
  line-height: 19px;
  color: var(--dsw-alias-label-secondary);
  text-align: left;
  -webkit-line-clamp: 2;
  -webkit-box-orient: vertical;
  overflow: hidden;
  cursor: pointer;
}
.rd-card-icon { flex: none; display: inline-flex; color: var(--dsw-alias-brand-primary); }
.rd-card-title { font-size: 12.5px; line-height: 18px; font-weight: 600; }
.rd-card-meta {
  margin-left: auto;
  display: inline-flex;
  align-items: center;
  gap: 5px;
  font-size: 11px;
  line-height: 16px;
  color: var(--dsw-alias-label-secondary);
  font-variant-numeric: tabular-nums;
}
.rd-card-body { padding: 11px 13px 13px; min-width: 0; overflow-wrap: anywhere; }
.rd-card-body .rd-md { font-size: 13.5px; line-height: 21px; }
.rd-card-note {
  margin: 0 0 8px;
  font-size: 11.5px;
  line-height: 17px;
  color: var(--dsw-alias-label-secondary);
}
.rd-card-error {
  margin: 0;
  font-size: 12.5px;
  line-height: 19px;
  color: var(--dsw-alias-state-error-primary);
  overflow-wrap: anywhere;
}
.rd-card-foot {
  display: flex;
  flex-wrap: wrap;
  align-items: center;
  gap: 8px;
  padding: 6px 8px;
  border-top: 0.5px solid var(--dsw-alias-border-l1);
}
.rd-mini {
  display: inline-flex;
  align-items: center;
  gap: 5px;
  height: 30px;
  padding: 0 10px;
  border: 0;
  border-radius: 7px;
  background: transparent;
  color: var(--dsw-alias-label-secondary);
  font: inherit;
  font-size: 12px;
  cursor: pointer;
}
.rd-mini:hover { background: var(--dsw-alias-bg-layer-2); color: var(--dsw-alias-label-primary); }
.rd-mini[data-danger]:hover { color: var(--dsw-alias-state-error-primary); }
.rd-mini[data-primary] { color: var(--dsw-alias-label-primary); font-weight: 600; }
.rd-mini-spacer { margin-left: auto; }

.rd-md { min-width: 0; max-width: 100%; color: inherit; overflow-wrap: anywhere; }
.rd-md > :first-child { margin-top: 0; }
.rd-md > :last-child { margin-bottom: 0; }
.rd-md h3 { margin: 15px 0 6px; font-size: 13.5px; line-height: 20px; font-weight: 600; }
.rd-md h4 { margin: 13px 0 5px; font-size: 13px; line-height: 19px; font-weight: 600; }
.rd-md p { margin: 8px 0; }
.rd-md ul, .rd-md ol { margin: 7px 0; padding-left: 19px; }
.rd-md li { margin: 3px 0; }
.rd-md li > p { margin: 0; }
.rd-md strong { font-weight: 600; }
.rd-md em { font-style: italic; }
.rd-md del { opacity: 0.6; }
.rd-md code {
  padding: 1px 5px;
  border-radius: 5px;
  background: var(--dsw-alias-bg-layer-2);
  font-family: ui-monospace, SFMono-Regular, "SF Mono", Menlo, monospace;
  font-size: 0.92em;
}
.rd-md a {
  color: var(--dsw-alias-label-primary);
  text-decoration: underline;
  text-underline-offset: 2px;
  text-decoration-color: color-mix(in srgb, var(--dsw-alias-label-primary) 45%, transparent);
}
.rd-md a:hover { text-decoration-color: currentColor; }
.rd-md blockquote {
  margin: 8px 0;
  padding-left: 11px;
  border-left: 2px solid var(--dsw-alias-border-l2);
  color: var(--dsw-alias-label-secondary);
}
.rd-md hr { margin: 12px 0; border: 0; border-top: 0.5px solid var(--dsw-alias-border-l1); }

.rd-rewrite {
  margin: 0 0 10px;
  padding: 11px 12px;
  border-radius: 10px;
  background: color-mix(in srgb, var(--dsw-alias-label-primary) 4%, transparent);
  border: 0.5px solid var(--dsw-alias-border-l2);
}
.rd-rewrite-label {
  display: block;
  margin-bottom: 6px;
  font-size: 11px;
  line-height: 15px;
  letter-spacing: 0.02em;
  color: var(--dsw-alias-label-secondary);
}
.rd-rewrite-text { margin: 0; font-size: 13.5px; line-height: 22px; white-space: pre-wrap; }

.rd-sources { margin-top: 12px; border-top: 0.5px solid var(--dsw-alias-border-l1); padding-top: 9px; }
.rd-sources-title { font-size: 11.5px; line-height: 16px; color: var(--dsw-alias-label-secondary); }
.rd-source {
  display: flex;
  align-items: baseline;
  gap: 7px;
  min-width: 0;
  padding: 7px 0;
  color: var(--dsw-alias-label-primary);
  text-decoration: none;
  font-size: 12.5px;
  line-height: 18px;
}
.rd-source:hover .rd-source-title { text-decoration: underline; }
.rd-source-index {
  flex: none;
  width: 17px;
  font-size: 10.5px;
  line-height: 15px;
  color: var(--dsw-alias-label-secondary);
  font-variant-numeric: tabular-nums;
}
.rd-source-title { flex: 1 1 auto; min-width: 0; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
.rd-source-host {
  flex: 0 1 auto;
  max-width: 45%;
  margin-left: auto;
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
  font-size: 11px;
  color: var(--dsw-alias-label-secondary);
}

.rd-caret::after {
  content: '';
  display: inline-block;
  width: 2px;
  height: 1em;
  margin-left: 2px;
  vertical-align: -0.14em;
  background: currentColor;
  animation: rd-blink 1.1s steps(2, start) infinite;
}
@keyframes rd-blink { 50% { opacity: 0; } }

.rd-empty {
  display: flex;
  flex-direction: column;
  align-items: center;
  gap: 4px;
  padding: 30px 18px;
  text-align: center;
  color: var(--dsw-alias-label-secondary);
}
.rd-empty-title { font-size: 13px; line-height: 20px; font-weight: 500; color: var(--dsw-alias-label-primary); }
.rd-empty-body { margin: 0; font-size: 12px; line-height: 18px; max-width: 30em; }
.rd-empty-actions { display: flex; gap: 8px; margin-top: 10px; }

.rd-sr {
  position: absolute;
  width: 1px;
  height: 1px;
  margin: -1px;
  padding: 0;
  overflow: hidden;
  clip: rect(0 0 0 0);
  white-space: nowrap;
  border: 0;
}
.rd-toast {
  position: absolute;
  left: 50%;
  bottom: 22px;
  transform: translateX(-50%);
  padding: 8px 15px;
  border-radius: 999px;
  border: 0.5px solid var(--dsw-alias-border-l1);
  background: var(--dsw-alias-bg-overlay);
  color: var(--dsw-alias-label-primary);
  font-size: 12.5px;
  line-height: 18px;
  box-shadow: 0 6px 22px rgba(0, 0, 0, 0.14);
  z-index: 4;
  animation: rd-rise 160ms ease-out;
}
@keyframes rd-rise { from { opacity: 0; transform: translate(-50%, 6px); } }

@media (prefers-reduced-motion: reduce) {
  .rd-page * { animation: none !important; transition: none !important; }
}
@media (pointer: coarse) {
  .rd-button { height: 44px; }
  .rd-action { padding: 12px; }
  .rd-mini { height: 44px; }
  .rd-text { font-size: 17px; }
  .rd-run-custom { height: 44px; }
}
@media (max-width: 1180px) {
  .rd-resizer { display: none; }
  .rd-body {
    grid-template-columns: minmax(0, 1fr) clamp(320px, 34vw, 396px);
    grid-template-rows: auto minmax(0, 1fr);
  }
  .rd-col-doc { grid-column: 1; grid-row: 1 / span 2; }
  .rd-col-review { grid-column: 2; grid-row: 1; max-height: 52vh; border-top: 0; }
  .rd-col-results { grid-column: 2; grid-row: 2; border-left: 0.5px solid var(--dsw-alias-border-l2); border-top: 0.5px solid var(--dsw-alias-border-l2); }
}
@media (max-width: 900px) {
  .rd-body { grid-template-columns: minmax(0, 1fr); grid-template-rows: minmax(0, 1.05fr) auto minmax(0, 1fr); }
  .rd-col-doc { grid-column: 1; grid-row: 1; }
  .rd-col-review { grid-column: 1; grid-row: 2; border-left: 0; border-top: 0.5px solid var(--dsw-alias-border-l2); max-height: 44vh; }
  .rd-col-results { grid-column: 1; grid-row: 3; border-left: 0; border-top: 0.5px solid var(--dsw-alias-border-l2); }
  .rd-doc-body { padding: 16px var(--rd-col-pad) 12px; }
}
@media (max-width: 720px) {
  .rd-header { flex-direction: column; align-items: stretch; gap: 10px; }
  .rd-toolbar { padding-top: 0; flex-wrap: wrap; }
  .rd-title { font-size: 18px; line-height: 25px; }
}
`

    /* ------------------------------------------------------------------ *
     * Components
     * ------------------------------------------------------------------ */

    function InlineButton({ label, onClick, disabled, variant, danger, title }) {
      return h('button', {
        type: 'button',
        className: 'rd-button',
        onClick,
        disabled: disabled === true,
        title,
        'data-variant': variant,
        'data-danger': danger === true ? '' : undefined,
      }, label)
    }

    function MiniButton({ children, onClick, primary, danger, title, style }) {
      return h('button', {
        type: 'button',
        className: 'rd-mini',
        onClick,
        title,
        'data-primary': primary === true ? '' : undefined,
        'data-danger': danger === true ? '' : undefined,
        style,
      }, children)
    }

    function SelectionCard({ selection, t, disabled, canSelectFirst, onSelectFirst }) {
      const hasSelection = selection.text.trim() !== ''
      const length = charCount(selection.text)
      return h('section', {
        className: 'rd-selection',
        'data-empty': hasSelection ? undefined : '',
        'aria-label': t('selectionLabel'),
      },
      h('div', { className: 'rd-selection-head' },
        h('span', { className: 'rd-selection-label' },
          hasSelection ? h('span', { className: 'rd-badge-dot', 'aria-hidden': true }) : null,
          hasSelection ? t('selectionLabel') : t('selectionEmpty')),
        hasSelection ? h('span', null, `${length} ${t('metaChars')}`) : null),
      hasSelection
        ? h('p', { className: 'rd-selection-text' }, selection.text)
        : h('div', null,
          h('p', { className: 'rd-selection-empty' }, t('selectionEmptyBody')),
          canSelectFirst
            ? h(MiniButton, { onClick: onSelectFirst, primary: true, style: { marginTop: '8px' } }, t('selectFirstAction'))
            : null),
      hasSelection && length > 6000
        ? h('p', { className: 'rd-inline-note' }, t('selectionTooLong'))
        : null)
    }

    function ActionRow({ action, t, disabled, onRun }) {
      const meta = ACTION_META[action]
      const Icon = meta.icon
      return h('button', {
        type: 'button',
        className: 'rd-action',
        disabled,
        onClick: () => onRun(action),
        title: disabled ? t('selectFirst') : meta.shortcut === '' ? t(meta.label) : `${t(meta.label)} · ${MOD_KEY}${meta.shortcut}`,
      },
      h('span', { className: 'rd-action-icon' }, h(Icon, { size: 16 })),
      h('span', { className: 'rd-action-label' }, t(meta.label)),
      meta.shortcut === '' ? null : h('span', { className: 'rd-action-kbd', 'aria-hidden': true }, `${MOD_KEY}${meta.shortcut}`),
      h('span', { className: 'rd-action-desc' }, disabled ? t('selectFirst') : t(meta.desc)))
    }

    /** Compact segmented control for the manuscript column. */
    function ModeToggle({ t, mode, onChange, disabled }) {
      return h('div', { className: 'rd-segmented', role: 'group', 'aria-label': t('modeName') },
        [['edit', t('modeEdit')], ['source', t('modeSource')], ['preview', t('modePreview')]].map(([value, label]) => h('button', {
          key: value,
          type: 'button',
          className: 'rd-segment',
          'data-active': mode === value ? '' : undefined,
          'aria-pressed': mode === value,
          disabled,
          onClick: () => onChange(value),
        }, label)))
    }

    /** A compact toolbar menu: label plus a list of secondary actions. */
    function ToolbarMenu({ label, items, disabled }) {
      const [open, setOpen] = useState(false)
      const rootRef = useRef(null)

      useEffect(() => {
        if (!open) return undefined
        const dismiss = (event) => {
          if (rootRef.current !== null && !rootRef.current.contains(event.target)) setOpen(false)
        }
        const escape = (event) => { if (event.key === 'Escape') setOpen(false) }
        document.addEventListener('mousedown', dismiss)
        document.addEventListener('keydown', escape)
        return () => {
          document.removeEventListener('mousedown', dismiss)
          document.removeEventListener('keydown', escape)
        }
      }, [open])

      return h('div', { className: 'rd-tmenu', ref: rootRef },
        h('button', {
          type: 'button',
          className: 'rd-button',
          'aria-haspopup': 'menu',
          'aria-expanded': open,
          disabled: disabled === true,
          onClick: () => setOpen(value => !value),
        }, label, h('span', { className: 'rd-tmenu-caret', 'aria-hidden': true }, '⌄')),
        open
          ? h('div', { className: 'rd-menu rd-tmenu-list', role: 'menu' },
            items.map(item => h('button', {
              key: item.key,
              type: 'button',
              role: 'menuitem',
              className: 'rd-menu-row',
              onClick: () => {
                setOpen(false)
                item.onSelect()
              },
            },
            h('span', { className: 'rd-menu-label' }, item.label),
            item.description === undefined ? null : h('span', { className: 'rd-menu-sub' }, item.description))))
          : null)
    }

    /**
     * Model chooser: the profile default, every advertised route, and the
     * reasoning levels of whichever route is in force.
     */
    function ModelMenu({ t, choice, status, data, message, onOpen, onPickModel, onPickEffort }) {
      const [open, setOpen] = useState(false)
      const rootRef = useRef(null)

      useEffect(() => {
        if (!open) return undefined
        const dismiss = (event) => {
          if (rootRef.current !== null && !rootRef.current.contains(event.target)) setOpen(false)
        }
        const escape = (event) => { if (event.key === 'Escape') setOpen(false) }
        document.addEventListener('mousedown', dismiss)
        document.addEventListener('keydown', escape)
        return () => {
          document.removeEventListener('mousedown', dismiss)
          document.removeEventListener('keydown', escape)
        }
      }, [open])

      const explicit = choice !== null && choice.provider !== ''
      const effortId = choice?.reasoningEffort ?? ''

      let label = t('modelFollow')
      if (explicit) {
        const route = data?.routes?.find(entry => entry.provider === choice.provider)
        label = route?.models?.find(model => model.id === choice.model)?.name ?? choice.model
      } else if (data?.default?.model) {
        const route = data.routes?.find(entry => entry.provider === data.default.provider)
        label = route?.models?.find(model => model.id === data.default.model)?.name ?? data.default.model
      }
      const reasoning = data === null ? undefined : effortsFor(data, choice)
      const efforts = reasoning?.efforts ?? []
      const effortName = (() => {
        if (efforts.length === 0 || reasoning === null || reasoning === undefined) return ''
        const wanted = effortId === '' ? reasoning.defaultEffort : effortId
        return efforts.find(effort => effort.id === wanted)?.name ?? wanted
      })()

      const toggle = () => {
        const next = !open
        setOpen(next)
        if (next) onOpen()
      }

      const rows = []
      if (status === 'loading') {
        rows.push(h('p', { key: 'loading', className: 'rd-menu-note' }, t('modelLoading')))
      } else if (status === 'error') {
        rows.push(h('p', { key: 'error', className: 'rd-menu-note' }, `${t('modelError')}${message === '' ? '' : `：${message}`}`))
      } else if (data !== null) {
        rows.push(h('button', {
          key: 'follow',
          type: 'button',
          className: 'rd-menu-row',
          'data-checked': explicit ? undefined : '',
          onClick: () => onPickModel(null),
        },
        h('span', { className: 'rd-menu-label' }, t('modelFollow')),
        h('span', { className: 'rd-menu-sub' }, data.default.provider === '' ? '' : `${data.default.provider} / ${data.default.model}`)))
        for (const route of data.routes) {
          rows.push(h('p', { key: `h-${route.provider}`, className: 'rd-menu-group' }, route.name))
          for (const model of route.models) {
            const value = `${route.provider}/${model.id}`
            rows.push(h('button', {
              key: value,
              type: 'button',
              className: 'rd-menu-row',
              'data-checked': explicit && `${choice.provider}/${choice.model}` === value ? '' : undefined,
              onClick: () => onPickModel({ provider: route.provider, model: model.id }),
            },
            h('span', { className: 'rd-menu-label' }, model.name),
            model.name === model.id ? null : h('span', { className: 'rd-menu-sub' }, model.id)))
          }
        }
        if (data.routes.length === 0) rows.push(h('p', { key: 'empty', className: 'rd-menu-note' }, t('modelEmpty')))
      }

      // Reasoning levels belong to the route in force, so they follow the choice.
      let effortRows = []
      if (open && data !== null) {
        if (reasoning === undefined) {
          effortRows.push(h('p', { key: 'effort-loading', className: 'rd-menu-note' }, t('reasoningLoading')))
        } else if (reasoning === null || efforts.length === 0) {
          effortRows.push(h('p', { key: 'effort-none', className: 'rd-menu-note' }, t('reasoningNone')))
        } else {
          effortRows.push(h('button', {
            key: 'effort-default',
            type: 'button',
            className: 'rd-menu-row',
            'data-checked': effortId === '' ? '' : undefined,
            onClick: () => onPickEffort(''),
          },
          h('span', { className: 'rd-menu-label' }, t('reasoningDefault')),
          h('span', { className: 'rd-menu-sub' }, (() => {
            const followsDefault = choice === null || choice.provider === ''
            const id = followsDefault && data.default.reasoningEffort !== ''
              ? data.default.reasoningEffort
              : reasoning.defaultEffort
            return id === '' ? '' : efforts.find(effort => effort.id === id)?.name ?? id
          })())))
          for (const effort of efforts) {
            effortRows.push(h('button', {
              key: `effort-${effort.id}`,
              type: 'button',
              className: 'rd-menu-row',
              'data-checked': effortId === effort.id ? '' : undefined,
              title: effort.description,
              onClick: () => onPickEffort(effort.id),
            },
            h('span', { className: 'rd-menu-label' }, effort.name),
            h('span', { className: 'rd-menu-sub' }, effort.id)))
          }
        }
      }

      return h('div', { className: 'rd-model', ref: rootRef },
        h('button', {
          type: 'button',
          className: 'rd-model-button',
          'aria-haspopup': 'menu',
          'aria-expanded': open,
          onClick: toggle,
          title: explicit ? `${choice.provider} / ${choice.model}` : t('modelFollow'),
        },
        h('span', { className: 'rd-model-caption' }, t('modelLabel')),
        h('span', { className: 'rd-model-name' }, label),
        effortName === '' ? null : h('span', { className: 'rd-model-effort' }, effortName),
        h('span', { className: 'rd-model-caret', 'aria-hidden': true }, '⌄')),
        open
          ? h('div', { className: 'rd-menu', role: 'menu' },
            h('div', { className: 'rd-menu-scroll' },
              h('p', { className: 'rd-menu-group' }, t('modelLabel')),
              rows,
              data !== null && data.failures.length > 0
                ? h('p', { key: 'failures', className: 'rd-menu-note' },
                  data.failures.map(failure => `${failure.name}: ${failure.message}`).join('；'))
                : null,
              effortRows.length === 0
                ? null
                : [h('p', { key: 'effort-head', className: 'rd-menu-group' }, t('reasoning')),
                  ...effortRows]),
            h('p', { className: 'rd-menu-hint' }, t('modelHint')))
          : null)
    }

    function Sources({ sources, note, t }) {
      if (sources.length === 0) {
        return note === '' ? null : h('p', { className: 'rd-card-note', style: { marginTop: '10px' } }, note)
      }
      return h('div', { className: 'rd-sources' },
        h('div', { className: 'rd-sources-title' }, t('sources')),
        sources.slice(0, 8).map((source, index) => h('a', {
          key: `${source.url}-${index}`,
          className: 'rd-source',
          href: source.url,
          target: '_blank',
          rel: 'noreferrer noopener',
        },
        h('span', { className: 'rd-source-index' }, `[S${index + 1}]`),
        h('span', { className: 'rd-source-title' }, source.title ?? hostOf(source.url)),
        h('span', { className: 'rd-source-host' }, hostOf(source.url)))))
    }

    function ResultCard({ run, t, expanded, highlighted, onToggle, onStop, onRetry, onRemove, onCopy, onInsert, onJump, onUndo, onRevise }) {
      const [revising, setRevising] = useState(false)
      const [draft, setDraft] = useState('')
      const meta = ACTION_META[run.action] ?? ACTION_META.logic
      const Icon = meta.icon
      const streaming = run.status === 'running'
      const cancelled = run.status === 'cancelled'
      const failed = run.status === 'error'
      const rewritable = run.action === 'expand' || run.action === 'custom' || run.revise === true
      const expansion = rewritable && run.raw !== '' ? parseExpansion(run.raw) : null
      const notes = expansion === null ? stripMarkers(run.raw) : expansion.notes
      const summary = summaryOf(run, t)

      let body
      if (streaming && run.raw === '') {
        const stage = run.stage === 'searching'
          ? `${t('stageSearching')}${run.stageDetail ? `：${run.stageDetail}` : '…'}`
          : run.stage === 'planning'
            ? t('stagePlanning')
            : run.stage === 'writing'
              ? t('stageWriting')
              : t('stagePreparing')
        body = h('p', { className: 'rd-card-note', role: 'status', 'aria-live': 'polite' }, stage)
      } else if (failed) {
        const mapped = run.errorCode === 'service'
          ? t('errService')
          : run.errorCode === 'auth'
            ? t('errAuth')
            : run.errorCode === 'network'
              ? t('errNetwork')
              : ''
        body = h('p', { className: 'rd-card-error' }, mapped !== '' ? mapped : (run.error === '' ? t('failed') : run.error))
      } else {
        body = h('div', null,
          expansion !== null && expansion.body !== ''
            ? h('div', { className: 'rd-rewrite' },
              h('span', { className: 'rd-rewrite-label' }, t('expandBlock')),
              h('p', {
                className: streaming ? 'rd-rewrite-text rd-caret' : 'rd-rewrite-text',
              }, expansion.body))
            : null,
          h('div', { className: 'rd-md' }, renderMarkdown(notes, run.id)))
      }

      const showSources = run.action === 'source' && (run.sources.length > 0 || run.sourceNote !== '')

      return h('article', {
        className: 'rd-card',
        'data-status': run.status,
        'data-run': run.id,
        'data-highlighted': highlighted ? '' : undefined,
        'data-folded': expanded ? undefined : '',
      },
      h('button', {
        type: 'button',
        className: 'rd-card-head',
        'aria-expanded': expanded,
        onClick: () => onToggle(run.id),
        title: expanded ? t('collapseCard') : t('expandCard'),
      },
      h('span', { className: 'rd-card-chevron', 'aria-hidden': true }, expanded ? '▾' : '▸'),
      h('span', { className: 'rd-card-icon' }, h(Icon, { size: 14 })),
      h('span', {
        className: 'rd-act-dot',
        'aria-hidden': true,
        style: { background: `var(--rd-act-${Object.hasOwn(ACTION_COLORS, run.action) ? run.action : 'custom'})` },
      }),
      h('span', { className: 'rd-card-title' }, t(meta.label)),
      run.revise === true ? h('span', { className: 'rd-card-chip' }, t('revisionBadge')) : null,
      run.wholeDocument === true ? h('span', { className: 'rd-card-chip' }, t('wholeDocument')) : null,
      run.effort === undefined || run.effort === '' ? null : h('span', { className: 'rd-card-chip' }, run.effort),
      h('span', { className: 'rd-card-meta' },
        streaming ? '···' : cancelled ? t('cancelled') : `${elapsedSeconds(run)}${t('seconds')}`)),

      expanded
        ? h('div', { className: 'rd-card-body' },
          run.instruction !== '' && run.action !== 'custom'
            ? h('p', { className: 'rd-card-note' }, `${t('instruction')}: ${run.instruction}`)
            : null,
          (() => {
            const warning = run.warningCode === 'reasoning-fallback'
              ? t('reasoningFallback').replace('{effort}', run.warningEffort ?? '')
              : run.warning
            return warning === '' || warning === undefined ? null : h('p', { className: 'rd-card-note' }, warning)
          })(),
          body,
          showSources ? h(Sources, { sources: run.sources, note: run.sourceNote, t }) : null,
          revising
            ? h('div', { className: 'rd-revise' },
              h('label', { className: 'rd-revise-label', htmlFor: `revise-${run.id}` }, t('reviseLabel')),
              h('textarea', {
                id: `revise-${run.id}`,
                className: 'rd-instruct-input',
                rows: 2,
                value: draft,
                placeholder: t('revisePlaceholder'),
                onChange: event => setDraft(event.target.value),
                onKeyDown: (event) => {
                  if ((event.metaKey || event.ctrlKey) && event.key === 'Enter') {
                    event.preventDefault()
                    if (draft.trim() !== '') {
                      onRevise(run, draft)
                      setRevising(false)
                      setDraft('')
                    }
                  }
                },
              }),
              h('div', { className: 'rd-revise-actions' },
                h('button', {
                  type: 'button',
                  className: 'rd-run-custom rd-revise-run',
                  disabled: draft.trim() === '',
                  onClick: () => {
                    onRevise(run, draft)
                    setRevising(false)
                    setDraft('')
                  },
                }, t('reviseRun')),
                h(MiniButton, { onClick: () => { setRevising(false); setDraft('') } }, t('cancel'))))
            : null)
        : h('button', {
          type: 'button',
          className: 'rd-card-summary',
          onClick: () => onToggle(run.id),
          title: t('expandCard'),
        }, summary === '' ? t('stagePreparing') : summary),

      expanded || streaming
        ? h('footer', { className: 'rd-card-foot' },
          run.selection.text.trim() === '' || run.wholeDocument === true
            ? null
            : h(MiniButton, { onClick: () => onJump(run), title: t('jumpTo') }, t('jumpTo')),
          streaming
            ? h(MiniButton, { onClick: () => onStop(run.id) }, h(IconStop, { size: 11 }), t('stop'))
            : null,
          !streaming && !failed
            ? h(MiniButton, { onClick: () => setRevising(value => !value), primary: revising }, t('revise'))
            : null,
          !streaming && !failed
            ? h(MiniButton, { onClick: () => onCopy(run), primary: run.action === 'expand' }, t('copy'))
            : null,
          !streaming && !failed && expansion !== null && expansion.body !== '' && (run.replaced === null || run.replaced === undefined)
            ? h(MiniButton, { onClick: () => onInsert(run, expansion.body), primary: true }, t('insert'))
            : null,
          !streaming && run.replaced !== null && run.replaced !== undefined
            ? h(MiniButton, { onClick: () => onUndo(run), primary: true }, t('undoReplace'))
            : null,
          failed || cancelled
            ? h(MiniButton, { onClick: () => onRetry(run), primary: true }, t('retry'))
            : null,
          h('span', { className: 'rd-mini-spacer' }),
          h(MiniButton, { onClick: () => onRemove(run.id), danger: true }, t('remove')))
        : null)
    }

        function EmptyResults({ t }) {
      return h('div', { className: 'rd-empty' },
        h('p', { className: 'rd-empty-title' }, t('resultsEmptyTitle')),
        h('p', { className: 'rd-empty-body' }, t('resultsEmptyBody').split('{mod}').join(MOD_KEY)))
    }

    function EmptyManuscript({ t, onSample, onImport }) {
      return h('div', { className: 'rd-empty', style: { flex: 1 } },
        h('p', { className: 'rd-empty-title' }, t('emptyTitle')),
        h('p', { className: 'rd-empty-body' }, t('emptyBody')),
        h('div', { className: 'rd-empty-actions' },
          h(InlineButton, { label: t('loadSample'), onClick: onSample, variant: 'primary' }),
          h(InlineButton, { label: t('importFile'), onClick: onImport })))
    }

    /* ------------------------------------------------------------------ *
     * The page
     * ------------------------------------------------------------------ */

    function ReviewDeskPage(props) {
      const snapshot = useStore()
      const t = props.t
      const [selection, setSelection] = useState({ start: 0, end: 0, text: '' })
      const [instruction, setInstruction] = useState('')
      const [importing, setImporting] = useState(false)
      const [viewMode, setViewMode] = useState('edit')
      const [modelChoice, setModelChoice] = useState(readModelChoice)
      const [catalog, setCatalog] = useState(catalogCache)
      const [catalogStatus, setCatalogStatus] = useState(catalogCache === null ? 'idle' : 'ready')
      const [catalogMessage, setCatalogMessage] = useState('')
      const [replaceIntent, setReplaceIntent] = useState(readReplaceIntent)
      const [widths, setWidths] = useState(readWidths)
      const [pendingPackage, setPendingPackage] = useState(null)
      const [highlight, setHighlight] = useState({ ids: [], nonce: 0 })
      const [frozen, setFrozen] = useState(null)
      const [dragging, setDragging] = useState('')
      const [filter, setFilter] = useState('all')
      const [expanded, setExpanded] = useState(() => new Set())
      const [original, setOriginal] = useState(null)
      const [pages, setPages] = useState(null)
      const [sheetOpen, setSheetOpen] = useState(false)
      const [notice, setNotice] = useState('')
      const [announcement, setAnnouncement] = useState('')
      const wasRunning = useRef(false)
      const [clearArmed, armClear, disarmClear] = useArmed()
      const pageRef = useRef(null)
      const bodyRef = useRef(null)
      const dragRef = useRef(null)
      const highlightTimer = useRef(0)
      const textRef = useRef(null)
      const previewRef = useRef(null)
      const blocksRef = useRef(null)
      const scrollRef = useRef(null)
      const overlayRef = useRef(null)
      const innerRef = useRef(null)
      const fileRef = useRef(null)
      const packageRef = useRef(null)
      const noticeTimer = useRef(0)
      const pendingFocus = useRef(null)

      const showNotice = useCallback((message) => {
        setNotice(message)
        if (noticeTimer.current) clearTimeout(noticeTimer.current)
        noticeTimer.current = setTimeout(() => setNotice(''), 2400)
      }, [])
      useEffect(() => () => { if (noticeTimer.current) clearTimeout(noticeTimer.current) }, [])

      const busy = snapshot.activeRunId !== ''
      const hasSelection = selection.text.trim() !== ''
      const runs = snapshot.runs

      // Ranges that still match the text, merged for one band per block.
      const marks = useMemo(() => reviewedMarks(runs, snapshot.text), [runs, snapshot.text])
      const viewText = frozen?.text ?? snapshot.text
      const viewMarks = frozen?.marks ?? marks
      const blocks = useMemo(() => parseBlocks(viewText), [viewText])
      const markTitle = useCallback(
        index => t('markResults').replace('{count}', String(marks[index]?.runs.length ?? 0)),
        [marks, t],
      )

      // Only the newest result opens; the rest stay folded until asked.
      useEffect(() => {
        const newest = runs[0]
        if (newest !== undefined) setExpanded(new Set([newest.id]))
      }, [runs[0]?.id])

      // An automatic replacement moves the selection onto the new text.
      useEffect(() => {
        const newest = runs[0]
        const record = newest?.replaced
        if (record === null || record === undefined) return
        setSelection(current => (
          current.start === record.start
            ? { start: record.start, end: record.start + record.inserted.length, text: record.inserted }
            : current
        ))
      }, [runs[0]?.replaced])

      /**
       * Bring one reviewed passage's results into view: clear the filter, open
       * them, mark them, and scroll the first one into place.
       */
      const revealRuns = useCallback((ids) => {
        const unique = [...new Set(ids)].filter(id => typeof id === 'string' && id !== '')
        if (unique.length === 0) return
        setFilter('all')
        setExpanded((current) => {
          const next = new Set(current)
          for (const id of unique) next.add(id)
          return next
        })
        setHighlight(current => ({ ids: unique, nonce: current.nonce + 1 }))
        if (highlightTimer.current) clearTimeout(highlightTimer.current)
        highlightTimer.current = setTimeout(() => setHighlight({ ids: [], nonce: 0 }), 2600)
      }, [])

      useEffect(() => () => {
        if (highlightTimer.current) clearTimeout(highlightTimer.current)
      }, [])

      // Scroll after the filtered list has rendered.
      useEffect(() => {
        if (highlight.ids.length === 0) return
        const body = scrollRef.current
        if (body === null) return
        const first = body.querySelector(`[data-run="${highlight.ids[0]}"]`)
        if (first === null) return
        const top = first.getBoundingClientRect().top - body.getBoundingClientRect().top + body.scrollTop - 12
        const reduce = window.matchMedia?.('(prefers-reduced-motion: reduce)').matches === true
        body.scrollTo?.({ top: Math.max(0, top), behavior: reduce ? 'auto' : 'smooth' })
      }, [highlight])

      /** A click (not a drag) on a reviewed passage opens its results. */
      const revealAtOffset = useCallback((offset) => {
        const mark = marks.find(entry => offset >= entry.start && offset < entry.end)
        if (mark === undefined) return
        revealRuns(mark.runs.map(run => run.id))
      }, [marks, revealRuns])

      const onPreviewClick = useCallback((event) => {
        const target = typeof event.target?.closest === 'function' ? event.target.closest('[data-mark]') : null
        if (target === null) return
        const index = Number(target.getAttribute('data-mark'))
        const mark = Number.isInteger(index) ? marks[index] : undefined
        if (mark === undefined) return
        const dom = window.getSelection()
        if (dom !== null && !dom.isCollapsed) return
        revealRuns(mark.runs.map(run => run.id))
      }, [marks, revealRuns])

      /**
       * Editing happens inside the rendered block: focus freezes the rendered
       * tree so streamed results cannot rewrite the DOM under the caret, and
       * blur serialises that block back to Markdown.
       */
      const beginBlockEdit = useCallback(() => {
        if (frozen !== null) return
        setFrozen({ text: store.get().text, marks })
      }, [frozen, marks])

      const finishBlockEdit = useCallback((element) => {
        const current = frozen
        setFrozen(null)
        if (current === null || element === null) return
        const index = Number(element.getAttribute('data-block'))
        const block = parseBlocks(current.text)[index]
        if (block === undefined) return
        const markdown = serializeElement(element)
        if (markdown === null) return
        if (markdown === current.text.slice(block.start, block.end)) return
        store.update({ text: current.text.slice(0, block.start) + markdown + current.text.slice(block.end) })
      }, [frozen])

      // A replace or an import during an edit abandons the frozen tree.
      useEffect(() => { setFrozen(null) }, [snapshot.text])

      const toggleRun = useCallback((id) => {
        setExpanded((current) => {
          const next = new Set(current)
          if (next.has(id)) next.delete(id)
          else next.add(id)
          return next
        })
      }, [])

      const tabs = useMemo(() => {
        const list = [{ key: 'all', label: t('tabAll'), count: runs.length }]
        for (const action of ACTION_TABS) {
          list.push({
            key: action,
            label: t(ACTION_META[action].label),
            count: runs.filter(run => run.action === action).length,
          })
        }
        return list
      }, [runs, t])

      const visibleRuns = useMemo(
        () => (filter === 'all' ? runs : runs.filter(run => run.action === filter)),
        [filter, runs],
      )

      /** Name of the reasoning level the next run will use, for the cards. */
      const effortLabel = useMemo(() => {
        const reasoning = effortsFor(catalog, modelChoice)
        if (reasoning === null || reasoning === undefined || reasoning.efforts.length === 0) return ''
        const wanted = modelChoice?.reasoningEffort ?? ''
        const followsDefault = modelChoice === null || modelChoice.provider === ''
        const fallback = followsDefault && catalog.default.reasoningEffort !== ''
          ? catalog.default.reasoningEffort
          : reasoning.defaultEffort
        const id = wanted === '' ? fallback : wanted
        return reasoning.efforts.find(effort => effort.id === id)?.name ?? id
      }, [catalog, modelChoice])

      // The highlight layer scrolls and lines up with the field, without state.
      const syncScroll = useCallback((top) => {
        if (overlayRef.current !== null) overlayRef.current.scrollTop = top
      }, [])
      useEffect(() => {
        syncScroll(textRef.current?.scrollTop ?? 0)
      }, [snapshot.text, marks, syncScroll])

      const bodyWidth = useCallback(() => bodyRef.current?.clientWidth ?? 0, [])

      /** Current geometry, falling back to the CSS defaults when never resized. */
      const currentWidths = useCallback(() => {
        const body = bodyRef.current
        if (body === null) return { review: COLUMN_LIMITS.review.initial, results: COLUMN_LIMITS.results.initial }
        const reviewColumn = body.querySelector('.rd-col-review')
        const resultsColumn = body.querySelector('.rd-col-results')
        return {
          review: reviewColumn?.getBoundingClientRect().width ?? COLUMN_LIMITS.review.initial,
          results: resultsColumn?.getBoundingClientRect().width ?? COLUMN_LIMITS.results.initial,
        }
      }, [])

      /** Paint a geometry without re-rendering the whole page. */
      const paintWidths = useCallback((next) => {
        const body = bodyRef.current
        if (body === null) return
        body.style.setProperty('--rd-review-w', `${next.review}px`)
        body.style.setProperty('--rd-results-w', `${next.results}px`)
      }, [])

      const commitWidths = useCallback((next) => {
        const fitted = fitWidths(next, bodyWidth())
        setWidths(fitted)
        saveWidths(fitted)
        paintWidths(fitted)
      }, [bodyWidth, paintWidths])

      const startDrag = useCallback((side, event) => {
        if (event.button !== undefined && event.button !== 0) return
        const start = currentWidths()
        dragRef.current = { side, start, startX: event.clientX }
        setDragging(side)
        event.currentTarget.setPointerCapture?.(event.pointerId)
        event.preventDefault()
      }, [currentWidths])

      const moveDrag = useCallback((event) => {
        const drag = dragRef.current
        if (drag === null) return
        const delta = event.clientX - drag.startX
        const next = drag.side === 'left'
          ? { review: drag.start.review - delta, results: drag.start.results }
          : { review: drag.start.review, results: drag.start.results - delta }
        paintWidths(fitWidths(next, bodyWidth()))
      }, [bodyWidth, paintWidths])

      const endDrag = useCallback((event) => {
        const drag = dragRef.current
        if (drag === null) return
        dragRef.current = null
        setDragging('')
        event?.currentTarget?.releasePointerCapture?.(event.pointerId)
        const next = fitWidths(currentWidths(), bodyWidth())
        commitWidths(next)
      }, [bodyWidth, commitWidths, currentWidths])

      /** Arrow keys move the separator; the drag maths is shared with the pointer. */
      const nudge = useCallback((side, delta) => {
        const start = widths ?? currentWidths()
        const next = side === 'left'
          ? { review: start.review + delta, results: start.results }
          : { review: start.review, results: start.results + delta }
        commitWidths(next)
      }, [commitWidths, currentWidths, widths])

      const resetWidths = useCallback(() => {
        setWidths(null)
        saveWidths(null)
        const body = bodyRef.current
        if (body !== null) {
          body.style.removeProperty('--rd-review-w')
          body.style.removeProperty('--rd-results-w')
        }
      }, [])

      // React owns the committed geometry; a drag writes the same variables
      // directly on the node, so streamed re-renders cannot revert mid-drag.
      useEffect(() => {
        const body = bodyRef.current
        if (body === null) return
        if (widths === null) {
          body.style.removeProperty('--rd-review-w')
          body.style.removeProperty('--rd-results-w')
          return
        }
        body.style.setProperty('--rd-review-w', `${widths.review}px`)
        body.style.setProperty('--rd-results-w', `${widths.results}px`)
      }, [widths])

      // Dragging anywhere keeps the cursor consistent and stops text selection.
      useEffect(() => {
        if (dragging === '') return undefined
        const move = (event) => moveDrag(event)
        const up = (event) => endDrag(event)
        window.addEventListener('pointermove', move)
        window.addEventListener('pointerup', up)
        window.addEventListener('pointercancel', up)
        return () => {
          window.removeEventListener('pointermove', move)
          window.removeEventListener('pointerup', up)
          window.removeEventListener('pointercancel', up)
        }
      }, [dragging, endDrag, moveDrag])

      // A narrower window must not push the document column below its minimum.
      useEffect(() => {
        const body = bodyRef.current
        if (body === null || widths === null || typeof ResizeObserver === 'undefined') return undefined
        const observer = new ResizeObserver(() => {
          const fitted = fitWidths(widths, body.clientWidth)
          if (fitted.review !== widths.review || fitted.results !== widths.results) {
            setWidths(fitted)
            saveWidths(fitted)
          }
        })
        observer.observe(body)
        return () => observer.disconnect()
      }, [widths])

      // Fill exactly the space between the panel top and the window bottom, so
      // the columns always have a bounded height and their panes can scroll.
      useEffect(() => {
        const element = pageRef.current
        if (element === null) return undefined
        let frame = 0
        const fit = () => {
          frame = 0
          const top = element.getBoundingClientRect().top
          const parent = element.parentElement
          const parentHeight = parent === null ? Number.POSITIVE_INFINITY : parent.clientHeight
          // Never grow past the window remainder or past a shorter host container.
          const available = Math.round(Math.min(window.innerHeight - top, parentHeight > 200 ? parentHeight : Number.POSITIVE_INFINITY))
          if (available > 200) element.style.height = `${available}px`
          // Opt-in layout diagnostics: localStorage['dsh.review-desk.debug'] = '1'.
          try {
            if (localStorage.getItem('dsh.review-desk.debug') === '1') {
              console.info('[review-desk] layout', {
                build: BUILD,
                top: Math.round(top),
                window: window.innerHeight,
                parent: parent === null ? null : parent.clientHeight,
                applied: element.style.height,
                body: element.querySelector('.rd-results-body')?.clientHeight ?? null,
                bodyContent: element.querySelector('.rd-results-body')?.scrollHeight ?? null,
              })
            }
          } catch {
            /* diagnostics are optional */
          }
        }
        const schedule = () => {
          if (frame !== 0) return
          frame = window.requestAnimationFrame(fit)
        }
        fit()
        window.addEventListener('resize', schedule)
        const observer = typeof ResizeObserver === 'undefined' ? null : new ResizeObserver(schedule)
        observer?.observe(document.body)
        return () => {
          window.removeEventListener('resize', schedule)
          observer?.disconnect()
          if (frame !== 0) window.cancelAnimationFrame(frame)
        }
      }, [])

      // Announce the end of a run once, without narrating every streamed delta.
      useEffect(() => {
        const runningNow = runs.some(run => run.status === 'running')
        if (wasRunning.current && !runningNow) {
          const latest = runs[0]
          setAnnouncement(latest === undefined
            ? t('announceDone')
            : `${t(ACTION_META[latest.action]?.label ?? 'results')}${latest.status === 'error' ? t('announceFailed') : t('announceDone')}`)
        }
        wasRunning.current = runningNow
      }, [runs, t])

      const syncSelection = useCallback((target) => {
        if (target === null || target === undefined) return
        const start = target.selectionStart ?? 0
        const end = target.selectionEnd ?? 0
        const value = target.value ?? ''
        const text = start === end ? '' : value.slice(start, end)
        setSelection(current => (
          current.start === start && current.end === end && current.text === text
            ? current
            : { start, end, text }
        ))
      }, [])

      /**
       * Start one review for the current selection, or for an explicit range a
       * keyboard shortcut read straight off the field.
       */
      const startRun = useCallback((action, override) => {
        const document = snapshot.text
        let range = override ?? selection
        if (action === 'custom' && range.text.trim() === '') {
          if (document.trim() === '') {
            showNotice(t('needSelection'))
            return
          }
          range = { start: 0, end: document.length, text: document }
        }
        if (range.text.trim() === '') {
          showNotice(t('needSelection'))
          return
        }
        if (action === 'custom' && instruction.trim() === '') {
          showNotice(t('customEmpty'))
          return
        }
        if (busy) {
          showNotice(t('busy'))
          return
        }
        setSelection(range)
        const run = blankRun(action, range, {
          instruction: instruction.trim(),
          wholeDocument: action === 'custom' && range.text === document,
          effort: effortLabel,
          forceReplace: action === 'custom' && (replaceIntent || looksLikeRewrite(instruction)),
        })
        store.addRun(run)
        const reduce = window.matchMedia?.('(prefers-reduced-motion: reduce)').matches === true
        scrollRef.current?.scrollTo?.({ top: 0, behavior: reduce ? 'auto' : 'smooth' })
        void executeRun(run.id, {
          action,
          text: range.text,
          manuscript: snapshot.text,
          title: snapshot.title,
          instruction,
          ...routePayload(modelChoice),
        })
      }, [busy, effortLabel, instruction, modelChoice, replaceIntent, selection, showNotice, snapshot.text, snapshot.title, t])

      /** Bring one reviewed passage back into view and select it. */
      const jumpTo = useCallback((run) => {
        const passage = run.selection
        if (passage === null || passage === undefined) return
        const current = store.get().text
        let { start, end, text: target } = passage
        if (current.slice(start, end) !== target) {
          const found = target === '' ? -1 : current.indexOf(target)
          if (found < 0) {
            showNotice(t('markStale'))
            return
          }
          start = found
          end = found + target.length
        }
        const index = marks.findIndex(mark => start >= mark.start && start < mark.end)
        const reduce = window.matchMedia?.('(prefers-reduced-motion: reduce)').matches === true
        if (viewMode !== 'source') {
          const root = viewMode === 'edit' ? blocksRef.current : previewRef.current
          const element = index < 0 ? null : root?.querySelector(`[data-mark="${index}"]`)
          element?.scrollIntoView?.({ block: 'center', behavior: reduce ? 'auto' : 'smooth' })
          setSelection({ start, end, text: current.slice(start, end) })
          return
        }
        const field = textRef.current
        if (field === null) return
        field.focus({ preventScroll: true })
        field.setSelectionRange(start, end)
        const element = index < 0 ? null : overlayRef.current?.querySelector(`[data-mark="${index}"]`)
        if (element !== null && element !== undefined) {
          const top = Math.max(0, element.offsetTop - field.clientHeight * 0.3)
          field.scrollTop = top
          syncScroll(top)
        }
        setSelection({ start, end, text: current.slice(start, end) })
      }, [marks, showNotice, syncScroll, t, viewMode])

      /** Select the first paragraph, so the actions become available. */
      const selectFirstBlock = useCallback(() => {
        const field = textRef.current
        const text = store.get().text
        if (text === '') return
        const match = /[^\n\s][\s\S]*?(?=\n\s*\n|$)/u.exec(text)
        if (match === null) return
        const start = match.index
        const end = start + match[0].replace(/\s+$/u, '').length
        if (field !== null) {
          field.focus({ preventScroll: true })
          field.setSelectionRange(start, end)
        }
        setSelection({ start, end, text: text.slice(start, end) })
      }, [])

      const stopRun = useCallback((id) => {
        const controller = running.get(id)
        if (controller !== undefined) controller.abort()
      }, [])

      const retryRun = useCallback((run) => {
        store.removeRun(run.id)
        const fresh = blankRun(run.action, run.selection, {
          instruction: run.instruction,
          wholeDocument: run.wholeDocument,
          effort: run.effort,
        })
        store.addRun(fresh)
        void executeRun(fresh.id, {
          action: run.action,
          text: run.selection.text,
          manuscript: snapshot.text,
          title: snapshot.title,
          instruction: run.instruction,
          ...routePayload(modelChoice),
        })
      }, [modelChoice, snapshot.text, snapshot.title])

      const copyRun = useCallback(async (run) => {
        const expansion = run.action === 'expand' ? parseExpansion(run.raw) : null
        const text = expansion !== null && expansion.body !== '' ? expansion.body : stripMarkers(run.raw)
        const ok = await copyText(text)
        showNotice(ok ? t('copied') : t('failed'))
      }, [showNotice, t])

      const insertRun = useCallback((run, body) => {
        const current = store.get().text
        const matches = run.wholeDocument === true
          || (Number.isInteger(run.selection.start)
            && run.selection.end <= current.length
            && current.slice(run.selection.start, run.selection.end) === run.selection.text)
        const range = replacePassage(run.id, body, { appendWhenStale: true })
        if (range === null) {
          showNotice(t('failed'))
          return
        }
        pendingFocus.current = range
        setViewMode('source')
        setSelection({ ...range, text: body })
        showNotice(matches ? t('replacedNotice') : t('appendedNotice'))
      }, [showNotice, t])

      const undoRunReplace = useCallback((run) => {
        if (!undoReplace(run.id)) {
          showNotice(t('undoStale'))
          return
        }
        const latest = store.get().runs.find(item => item.id === run.id)
        if (latest !== undefined && latest.selection !== undefined) setSelection({ ...latest.selection })
        showNotice(t('undoDone'))
      }, [showNotice, t])

      /** Continue a result with one more instruction. */
      const reviseRun = useCallback((parent, requirement) => {
        const trimmed = requirement.trim()
        if (trimmed === '') {
          showNotice(t('reviseEmpty'))
          return
        }
        if (busy) {
          showNotice(t('busy'))
          return
        }
        const body = parseExpansion(parent.raw).body
        const source = body !== '' ? body : stripMarkers(parent.raw)
        if (source.trim() === '') {
          showNotice(t('failed'))
          return
        }
        // Keep the first original as the fact reference across a revision chain.
        const original = parent.original !== '' && parent.original !== undefined
          ? parent.original
          : parent.selection.text
        const run = blankRun('custom', parent.selection, {
          instruction: trimmed,
          wholeDocument: parent.wholeDocument === true,
          revise: true,
          parentId: parent.id,
          original,
          effort: effortLabel,
        })
        store.addRun(run)
        const reduce = window.matchMedia?.('(prefers-reduced-motion: reduce)').matches === true
        scrollRef.current?.scrollTo?.({ top: 0, behavior: reduce ? 'auto' : 'smooth' })
        void executeRun(run.id, {
          action: 'custom',
          text: source,
          manuscript: snapshot.text,
          title: snapshot.title,
          instruction: trimmed,
          revise: true,
          original,
          ...routePayload(modelChoice),
        })
      }, [busy, effortLabel, modelChoice, showNotice, snapshot.text, snapshot.title, t])

      // Restore focus and the inserted range after a replacement.
      useEffect(() => {
        const range = pendingFocus.current
        if (range === null || textRef.current === null) return
        pendingFocus.current = null
        const field = textRef.current
        field.focus()
        field.setSelectionRange(range.start, range.end)
        syncSelection(field)
      }, [snapshot.text, viewMode, syncSelection])

      /** Switching back to editing restores the passage the user had selected. */
      const changeMode = useCallback((mode) => {
        if (mode === viewMode) return
        if (mode === 'source' && selection.text !== '') {
          pendingFocus.current = { start: selection.start, end: selection.end }
        }
        setViewMode(mode)
      }, [selection, viewMode])

      /** Read a text selection made in a rendered view back as source offsets. */
      const readViewSelection = useCallback((root) => {
        if (root === null || root === undefined) return
        const dom = window.getSelection()
        if (dom === null || dom.rangeCount === 0 || dom.isCollapsed) return
        const range = dom.getRangeAt(0)
        if (!root.contains(range.startContainer) || !root.contains(range.endContainer)) return
        const start = previewOffset(root, range.startContainer, range.startOffset, 'start')
        const end = previewOffset(root, range.endContainer, range.endOffset, 'end')
        if (start === null || end === null || end <= start) return
        const text = store.get().text.slice(start, end)
        setSelection(current => (
          current.start === start && current.end === end && current.text === text
            ? current
            : { start, end, text }
        ))
      }, [])

      /** Read this Host generation's model catalog, the same one the composer uses. */
      const openModels = useCallback(async () => {
        if (catalogCache !== null) {
          setCatalog(catalogCache)
          setCatalogStatus('ready')
          return
        }
        if (catalogStatus === 'loading') return
        setCatalogStatus('loading')
        setCatalogMessage('')
        try {
          const remote = props.remote
          if (remote === undefined || remote === null) throw new Error('Remote is unavailable')
          const response = await remote.session.modelCatalog()
          if (response === null || typeof response !== 'object' || response.ok !== true) {
            throw new Error(response?.error?.message ?? 'model catalog rejected')
          }
          const view = catalogView(response.value)
          if (view.routes.length === 0) throw new Error('no model route is available')
          catalogCache = view
          setCatalog(view)
          setCatalogStatus('ready')
        } catch (error) {
          setCatalogStatus('error')
          setCatalogMessage(error instanceof Error ? error.message : String(error))
        }
      }, [catalogStatus, props.remote])

      const pickModel = useCallback((choice) => {
        const next = choice === null
          ? null
          : { provider: choice.provider, model: choice.model, reasoningEffort: '' }
        setModelChoice(next)
        saveModelChoice(next)
      }, [])

      const pickEffort = useCallback((effort) => {
        const base = modelChoice ?? { provider: '', model: '', reasoningEffort: '' }
        const next = { ...base, reasoningEffort: effort }
        if (next.provider === '' && next.reasoningEffort === '') {
          setModelChoice(null)
          saveModelChoice(null)
          return
        }
        setModelChoice(next)
        saveModelChoice(next)
      }, [modelChoice])

      const onKeyDown = useCallback((event) => {
        const meta = event.metaKey || event.ctrlKey
        if (!meta || event.altKey) return
        const index = ['1', '2', '3', '4'].indexOf(event.key)
        if (index < 0) return
        const action = ACTION_ORDER[index]
        event.preventDefault()
        const field = event.currentTarget
        const start = field.selectionStart ?? 0
        const end = field.selectionEnd ?? 0
        startRun(action, {
          start,
          end,
          text: start === end ? '' : field.value.slice(start, end),
        })
      }, [startRun])

      // The same shortcuts work while focus is outside a text field, so the
      // hint does not promise something that only works in the manuscript.
      useEffect(() => {
        const onGlobalKey = (event) => {
          if (!(event.metaKey || event.ctrlKey) || event.altKey) return
          const index = ['1', '2', '3', '4'].indexOf(event.key)
          if (index < 0) return
          const target = event.target
          if (target === textRef.current) return
          const tag = target?.tagName
          if (tag === 'INPUT' || tag === 'TEXTAREA' || tag === 'SELECT' || target?.isContentEditable === true) return
          event.preventDefault()
          startRun(ACTION_ORDER[index])
        }
        window.addEventListener('keydown', onGlobalKey)
        return () => window.removeEventListener('keydown', onGlobalKey)
      }, [startRun])

      /** Render the imported file's pages once, then show them as a sheet. */
      const openOriginal = useCallback(async () => {
        if (original === null) {
          showNotice(t('originalMissing'))
          return
        }
        setSheetOpen(true)
        if (pages !== null && pages.status !== 'error') return
        setPages({ status: 'loading', pages: [], pageCount: 0, truncated: false, error: '' })
        try {
          const extension = (original.name.split('.').pop() ?? '').toLowerCase()
          const url = new URL(PAGES_ROUTE, document.baseURI)
          url.searchParams.set('ext', extension)
          const response = await fetch(url.toString(), {
            method: 'POST',
            credentials: 'include',
            headers: { 'content-type': 'application/octet-stream' },
            body: original.file,
          })
          let data = null
          try {
            data = await response.json()
          } catch {
            data = null
          }
          if (!response.ok) throw new Error(data?.error ?? `HTTP ${response.status}`)
          setPages({
            status: 'ready',
            pages: Array.isArray(data?.pages) ? data.pages : [],
            pageCount: typeof data?.pageCount === 'number' ? data.pageCount : 0,
            truncated: data?.truncated === true,
            error: '',
          })
        } catch (error) {
          setPages({
            status: 'error',
            pages: [],
            pageCount: 0,
            truncated: false,
            error: error instanceof Error ? error.message : String(error),
          })
        }
      }, [original, pages, showNotice, t])

      // Escape closes the layout sheet.
      useEffect(() => {
        if (!sheetOpen) return undefined
        const close = (event) => { if (event.key === 'Escape') setSheetOpen(false) }
        document.addEventListener('keydown', close)
        return () => document.removeEventListener('keydown', close)
      }, [sheetOpen])

      const onPickFile = useCallback(async (event) => {
        const file = event.target.files?.[0]
        event.target.value = ''
        if (file === undefined) return
        setImporting(true)
        try {
          const imported = await extractDocument(file)
          if (imported.text.trim() === '') {
            showNotice(t('importEmpty'))
            return
          }
          const fallback = file.name.replace(/\.[^.]+$/u, '').trim()
          const name = imported.title === '' ? fallback : imported.title
          store.update({ text: imported.text, title: name === '' ? store.get().title : name })
          setSelection({ start: 0, end: 0, text: '' })
          const extension = (file.name.split('.').pop() ?? '').toLowerCase()
          setOriginal(extension === 'docx' || extension === 'docm' || extension === 'doc' || extension === 'pdf'
            ? { name: file.name, file }
            : null)
          setPages(null)
          setViewMode('edit')
          showNotice(extension === 'docx' || extension === 'docm' || extension === 'pdf'
            ? t('importFormatKept')
            : t('imported'))
        } catch (error) {
          const code = error instanceof Error ? error.message : String(error)
          try { console.warn('[review-desk] import failed:', code, error) } catch { /* console is optional */ }
          showNotice(importMessage(error, t))
        } finally {
          setImporting(false)
        }
      }, [showNotice, t])

      const loadSample = useCallback(() => {
        setOriginal(null)
        setPages(null)
        store.update({ text: SAMPLE_TEXT, title: SAMPLE_TITLE })
        setSelection({ ...SAMPLE_SELECTION, text: SAMPLE_TEXT.slice(SAMPLE_SELECTION.start, SAMPLE_SELECTION.end) })
        pendingFocus.current = SAMPLE_SELECTION
      }, [])

      const exportFile = useCallback(() => {
        const current = store.get()
        const name = (current.title.trim() === '' ? 'manuscript' : current.title.trim()).replace(/[\\/:*?"<>|]/gu, '-')
        const blob = new Blob([`# ${current.title.trim()}\n\n${current.text}\n`], { type: 'text/markdown;charset=utf-8' })
        const url = URL.createObjectURL(blob)
        const anchor = document.createElement('a')
        anchor.href = url
        anchor.download = `${name}.md`
        anchor.click()
        setTimeout(() => URL.revokeObjectURL(url), 1000)
      }, [])

      /** One file carrying the manuscript and every result. */
      const exportPackageFile = useCallback(() => {
        const current = store.get()
        const text = JSON.stringify(buildPackage({ ...current, instruction }), null, 2)
        if (text.length > PACKAGE_LIMIT) {
          showNotice(t('packageTooLarge'))
          return
        }
        const base = (current.title.trim() === '' ? 'review' : current.title.trim()).replace(/[\\/:*?"<>|]/gu, '-')
        const stamp = new Date().toISOString().slice(0, 10)
        const blob = new Blob([text], { type: 'application/json;charset=utf-8' })
        const url = URL.createObjectURL(blob)
        const anchor = document.createElement('a')
        anchor.href = url
        anchor.download = `${base}-${stamp}.review-desk.json`
        anchor.click()
        setTimeout(() => URL.revokeObjectURL(url), 1000)
        showNotice(t('packageExportDone'))
      }, [instruction, showNotice, t])

      const onPickPackage = useCallback(async (event) => {
        const file = event.target.files?.[0]
        event.target.value = ''
        if (file === undefined) return
        if (file.size > PACKAGE_LIMIT) {
          showNotice(t('packageTooLarge'))
          return
        }
        try {
          const parsed = JSON.parse(await file.text())
          const result = readPackage(parsed)
          if (result.error === 'empty') {
            showNotice(t('packageEmpty'))
            return
          }
          if (result.error !== undefined) {
            showNotice(t('packageInvalid'))
            return
          }
          setPendingPackage(result)
        } catch {
          showNotice(t('packageInvalid'))
        }
      }, [showNotice, t])

      const confirmImportPackage = useCallback(() => {
        if (pendingPackage === null) return
        setOriginal(null)
        setPages(null)
        setSelection({ start: 0, end: 0, text: '' })
        setFilter('all')
        setExpanded(new Set())
        setInstruction(pendingPackage.instruction)
        store.update({
          title: pendingPackage.title,
          text: pendingPackage.text,
          runs: pendingPackage.runs,
          activeRunId: '',
        })
        setPendingPackage(null)
        setViewMode('edit')
        pendingFocus.current = null
        showNotice(t('importedPackage'))
      }, [pendingPackage, showNotice, t])

      // Escape closes the import confirmation.
      useEffect(() => {
        if (pendingPackage === null) return undefined
        const cancel = (event) => { if (event.key === 'Escape') setPendingPackage(null) }
        document.addEventListener('keydown', cancel)
        return () => document.removeEventListener('keydown', cancel)
      }, [pendingPackage])

      const onClear = useCallback(() => {
        if (!armClear()) return
        store.update({ text: '', title: '', runs: [], activeRunId: '' })
        setSelection({ start: 0, end: 0, text: '' })
        setOriginal(null)
        setPages(null)
        disarmClear()
        showNotice(t('cleared'))
      }, [armClear, disarmClear, showNotice, t])

      /** Install / inspect / remove the WPS add-in on this machine. */
      const runWpsAddin = useCallback(async (action) => {
        try {
          const response = await fetch(new URL(WPS_ROUTE, document.baseURI).toString(), {
            method: 'POST',
            credentials: 'include',
            headers: { 'content-type': 'application/json' },
            body: JSON.stringify({ action }),
          })
          const info = await response.json().catch(() => ({}))
          if (!response.ok) {
            throw new Error(typeof info?.error === 'string' ? info.error : `HTTP ${response.status}`)
          }
          if (action === 'install') {
            showNotice(t('wpsInstalled').replace('{path}', info.target ?? ''))
          } else if (action === 'uninstall') {
            showNotice(t('wpsUninstalled'))
          } else {
            showNotice(info.installed ? t('wpsStatusOk').replace('{version}', info.version ?? '') : t('wpsNotInstalled'))
          }
        } catch (error) {
          showNotice(t('wpsFailed').replace('{message}', error instanceof Error ? error.message : String(error)))
        }
      }, [showNotice, t])

      const chars = charCount(snapshot.text)
      const paragraphs = paragraphCount(snapshot.text)
      const minutes = Math.max(1, Math.round(chars / 400))
      const metaText = snapshot.text.trim() === ''
        ? t('subtitle')
        : `${chars} ${t('metaChars')} · ${paragraphs} ${t('metaParagraphs')} · ${t('metaMinutes')} ${minutes}`

      return h('section', {
        className: 'rd-page',
        ref: pageRef,
        'data-plugin-panel': '',
        'data-build': BUILD,
        'data-resizing': dragging === '' ? undefined : '',
        'aria-label': t('title'),
      },
        h('style', null, CSS),

        h('header', { className: 'rd-header', 'data-window-drag': '' },
          h('div', { className: 'rd-heading' },
            h('h1', { className: 'rd-title' }, t('title')),
            h('p', { className: 'rd-subtitle' }, metaText)),
          h('div', { className: 'rd-toolbar' },
            h(ModelMenu, {
              t,
              choice: modelChoice,
              status: catalogStatus,
              data: catalog,
              message: catalogMessage,
              onOpen: openModels,
              onPickModel: pickModel,
              onPickEffort: pickEffort,
            }),
            h(ToolbarMenu, {
              label: importing ? t('importing') : t('importMenu'),
              disabled: importing,
              items: [
                { key: 'file', label: t('importFile'), description: '.docx · .pdf · .md · .txt', onSelect: () => fileRef.current?.click() },
                { key: 'package', label: t('importPackage'), description: '.json', onSelect: () => packageRef.current?.click() },
              ],
            }),
            h(ToolbarMenu, {
              label: t('exportMenu'),
              disabled: snapshot.text.trim() === '',
              items: [
                { key: 'md', label: t('exportMd'), description: '.md', onSelect: exportFile },
                { key: 'package', label: t('exportPackage'), description: '.json', onSelect: exportPackageFile },
              ],
            }),
            h(ToolbarMenu, {
              label: t('wpsMenu'),
              items: [
                { key: 'wps-install', label: t('wpsInstall'), description: t('wpsInstallDesc'), onSelect: () => runWpsAddin('install') },
                { key: 'wps-status', label: t('wpsStatus'), onSelect: () => runWpsAddin('status') },
                { key: 'wps-uninstall', label: t('wpsUninstall'), onSelect: () => runWpsAddin('uninstall') },
              ],
            }),
            h(InlineButton, {
              label: t('copyAll'),
              disabled: snapshot.text.trim() === '',
              onClick: async () => {
                const ok = await copyText(store.get().text)
                showNotice(ok ? t('copiedAll') : t('failed'))
              },
            }),
            h(InlineButton, {
              label: clearArmed ? t('clearArmed') : t('clear'),
              onClick: onClear,
              danger: clearArmed,
              disabled: snapshot.text.trim() === '' && runs.length === 0,
            }),
            h('input', {
              ref: fileRef,
              type: 'file',
              accept: '.txt,.md,.markdown,.docx,.pdf,text/plain,text/markdown,application/pdf,application/vnd.openxmlformats-officedocument.wordprocessingml.document',
              style: { display: 'none' },
              onChange: onPickFile,
            }),
            h('input', {
              ref: packageRef,
              type: 'file',
              accept: '.json,application/json',
              style: { display: 'none' },
              onChange: onPickPackage,
            }))),

        h('div', { className: 'rd-body', ref: bodyRef },

          // 文稿 -----------------------------------------------------------------
          h('section', { className: 'rd-col rd-col-doc', 'aria-label': t('colDoc') },
            h('header', { className: 'rd-col-head' },
              h('span', { className: 'rd-col-title' }, t('colDoc')),
              h('div', { className: 'rd-col-meta' },
                original === null
                  ? null
                  : h('button', {
                    type: 'button',
                    className: 'rd-link-button',
                    onClick: openOriginal,
                    title: t('originalNote'),
                  }, t('viewOriginal')),
                h(ModeToggle, {
                  t,
                  mode: viewMode,
                  onChange: changeMode,
                  disabled: snapshot.text.trim() === '',
                }))),
            h('div', { className: 'rd-doc-body' },
              h('div', { className: 'rd-doc-inner' },
                h('input', {
                  className: 'rd-doc-title',
                  value: snapshot.title,
                  placeholder: t('titlePlaceholder'),
                  'aria-label': t('titlePlaceholder'),
                  onChange: event => store.update({ title: event.target.value }),
                }),
                h('div', { className: 'rd-doc-meta' },
                  h('span', null, `${chars} ${t('metaChars')}`),
                  h('span', { className: 'rd-dot', 'aria-hidden': true }),
                  h('span', null, `${paragraphs} ${t('metaParagraphs')}`),
                  h('span', { className: 'rd-dot', 'aria-hidden': true }),
                  h('span', null, `${minutes} ${t('metaMinutes')}`),
                  marks.length === 0
                    ? null
                    : [h('span', { key: 'rd', className: 'rd-dot', 'aria-hidden': true }),
                      h('span', { key: 'rc', className: 'rd-meta-mark' }, `${t('reviewedBlocks')} ${marks.length}`)],
                  snapshot.text === SAMPLE_TEXT
                    ? [h('span', { key: 'd', className: 'rd-dot', 'aria-hidden': true }),
                      h('span', { key: 'n' }, t('sampleNote'))]
                    : null),
                snapshot.text.trim() === ''
                  ? h(EmptyManuscript, {
                    t,
                    onSample: loadSample,
                    onImport: () => fileRef.current?.click(),
                  })
                  : viewMode === 'preview'
                    ? h('div', {
                      className: 'rd-preview',
                      ref: previewRef,
                      tabIndex: 0,
                      role: 'document',
                      'aria-label': t('manuscript'),
                      onMouseUp: event => readViewSelection(event.currentTarget),
                      onKeyUp: event => readViewSelection(event.currentTarget),
                      onClick: onPreviewClick,
                    }, renderPreview(snapshot.text, marks, markTitle))
                    : viewMode === 'source'
                      ? h('div', { className: 'rd-editor' },
                        h('div', { className: 'rd-highlights', ref: overlayRef, 'aria-hidden': true },
                          h('div', { className: 'rd-highlights-inner', ref: innerRef }, renderMarked(snapshot.text, marks))),
                        h('textarea', {
                          ref: textRef,
                          className: 'rd-text',
                          value: snapshot.text,
                          spellCheck: false,
                          'aria-label': t('manuscript'),
                          placeholder: t('emptyBody'),
                          onChange: event => store.update({ text: event.target.value }),
                          onSelect: event => syncSelection(event.currentTarget),
                          onKeyUp: event => syncSelection(event.currentTarget),
                          onMouseUp: event => syncSelection(event.currentTarget),
                          onScroll: event => syncScroll(event.currentTarget.scrollTop),
                          onKeyDown,
                          onClick: (event) => {
                            const field = event.currentTarget
                            if (field.selectionStart !== field.selectionEnd) return
                            revealAtOffset(field.selectionStart ?? 0)
                          },
                        }))
                      : h('div', {
                        className: 'rd-preview rd-blocks',
                        ref: blocksRef,
                        tabIndex: 0,
                        role: 'document',
                        'aria-label': t('manuscript'),
                        onMouseUp: event => readViewSelection(event.currentTarget),
                        onKeyUp: event => readViewSelection(event.currentTarget),
                        onClick: onPreviewClick,
                      },
                      h('p', { className: 'rd-block-hint' }, t('blockEditHint')),
                      blocks.map((block, index) => {
                        const editable = block.kind !== 'image' && block.kind !== 'hr'
                        return h('div', {
                          key: `b${index}`,
                          className: 'rd-block',
                          'data-block': index,
                          contentEditable: editable,
                          suppressContentEditableWarning: true,
                          spellCheck: editable,
                          ...(editable
                            ? {
                              onFocus: beginBlockEdit,
                              onBlur: event => finishBlockEdit(event.currentTarget),
                              onKeyDown: (event) => {
                                if (event.key !== 'Enter' || event.metaKey || event.ctrlKey || event.altKey) return
                                event.preventDefault()
                                insertLineBreak()
                              },
                              onPaste: (event) => {
                                event.preventDefault()
                                const text = event.clipboardData?.getData('text/plain') ?? ''
                                if (text !== '') insertPlainText(text)
                              },
                            }
                            : {}),
                        }, renderBlock(block, index, viewMarks, markTitle))
                      }))))),


          // 审查 -----------------------------------------------------------------
          h('section', { className: 'rd-col rd-col-review', 'aria-label': t('colReview') },
            h('header', { className: 'rd-col-head' },
              h('span', { className: 'rd-col-title' }, t('colReview'))),
            h('div', { className: 'rd-review-body' },
              h(SelectionCard, {
                selection,
                t,
                disabled: busy,
                canSelectFirst: !hasSelection && snapshot.text.trim() !== '',
                onSelectFirst: selectFirstBlock,
              }),
              h('div', { className: 'rd-action-group' },
                ACTION_ORDER.map(action => h(ActionRow, {
                  key: action,
                  action,
                  t,
                  disabled: !hasSelection || busy,
                  onRun: startRun,
                }))),
              h('div', { className: 'rd-instruction-block' },
                h('label', { className: 'rd-instruct-label', htmlFor: 'rd-instruction' }, t('customInstruction')),
                h('textarea', {
                  id: 'rd-instruction',
                  className: 'rd-instruct-input',
                  rows: 3,
                  value: instruction,
                  placeholder: t('customPlaceholder'),
                  onChange: event => setInstruction(event.target.value),
                  onKeyDown: (event) => {
                    if ((event.metaKey || event.ctrlKey) && event.key === 'Enter') {
                      event.preventDefault()
                      startRun('custom')
                    }
                  },
                }),
                h('button', {
                  type: 'button',
                  className: 'rd-run-custom',
                  'data-variant': 'primary',
                  disabled: busy || instruction.trim() === '' || (!hasSelection && snapshot.text.trim() === ''),
                  onClick: () => startRun('custom'),
                  title: `${t('customRun')} · ${MOD_KEY}↩`,
                }, t('customRun'), h('span', { className: 'rd-run-custom-key', 'aria-hidden': true }, `${MOD_KEY}↩`)),
                h('label', { className: 'rd-check', title: t('replaceIntentHint') },
                  h('input', {
                    type: 'checkbox',
                    checked: replaceIntent,
                    onChange: (event) => {
                      const next = event.target.checked
                      setReplaceIntent(next)
                      saveReplaceIntent(next)
                    },
                  }),
                  h('span', null, t('replaceIntent'))),
                h('p', { className: 'rd-instruction-hint' }, t('customHint'))))),

          // 结果 -----------------------------------------------------------------
          h('section', { className: 'rd-col rd-col-results', 'aria-label': t('colResults') },
            h('header', { className: 'rd-col-head' },
              h('span', { className: 'rd-col-title' }, t('colResults')),
              h('div', { className: 'rd-col-meta' },
                h('span', null, String(runs.length)))),
            h('div', { className: 'rd-tabs', role: 'tablist', 'aria-label': t('colResults') },
              tabs.map(tab => h('button', {
                key: tab.key,
                type: 'button',
                role: 'tab',
                className: 'rd-tab',
                'data-active': filter === tab.key ? '' : undefined,
                'aria-selected': filter === tab.key,
                disabled: tab.key !== 'all' && tab.count === 0,
                onClick: () => setFilter(tab.key),
              },
              tab.key === 'all' ? null : h('span', {
                className: 'rd-act-dot',
                'aria-hidden': true,
                style: { background: `var(--rd-act-${tab.key})` },
              }),
              tab.label,
              h('span', { className: 'rd-tab-count' }, String(tab.count))))),
            h('div', { className: 'rd-results-body', ref: scrollRef, 'aria-busy': busy },
              visibleRuns.length === 0
                ? (runs.length === 0
                  ? h(EmptyResults, { t })
                  : h('div', { className: 'rd-empty' },
                    h('p', { className: 'rd-empty-title' }, t('emptyFilter'))))
                : visibleRuns.map(run => h(ResultCard, {
                  key: run.id,
                  run,
                  t,
                  expanded: expanded.has(run.id),
                  highlighted: highlight.ids.includes(run.id),
                  onToggle: toggleRun,
                  onStop: stopRun,
                  onRetry: retryRun,
                  onRemove: id => store.removeRun(id),
                  onCopy: copyRun,
                  onInsert: insertRun,
                  onJump: jumpTo,
                  onUndo: undoRunReplace,
                  onRevise: reviseRun,
                })))),

          /**
           * Column separators. They are absolutely positioned outside the grid
           * flow, so dragging repaints the tracks without remounting the page.
           */
          h('div', {
            className: 'rd-resizer rd-resizer-left',
            role: 'separator',
            'aria-orientation': 'vertical',
            'aria-label': t('resizeLeft'),
            'aria-valuemin': COLUMN_LIMITS.review.min,
            'aria-valuemax': COLUMN_LIMITS.review.max,
            'aria-valuenow': Math.round(widths?.review ?? COLUMN_LIMITS.review.initial),
            tabIndex: 0,
            title: t('resizeHint'),
            'data-active': dragging === 'left' ? '' : undefined,
            onPointerDown: event => startDrag('left', event),
            onDoubleClick: resetWidths,
            onKeyDown: (event) => {
              const step = event.shiftKey ? 48 : 16
              if (event.key === 'ArrowLeft') { event.preventDefault(); nudge('left', step) }
              else if (event.key === 'ArrowRight') { event.preventDefault(); nudge('left', -step) }
              else if (event.key === 'Home') { event.preventDefault(); resetWidths() }
            },
          }),
          h('div', {
            className: 'rd-resizer rd-resizer-right',
            role: 'separator',
            'aria-orientation': 'vertical',
            'aria-label': t('resizeRight'),
            'aria-valuemin': COLUMN_LIMITS.results.min,
            'aria-valuemax': COLUMN_LIMITS.results.max,
            'aria-valuenow': Math.round(widths?.results ?? COLUMN_LIMITS.results.initial),
            tabIndex: 0,
            title: t('resizeHint'),
            'data-active': dragging === 'right' ? '' : undefined,
            onPointerDown: event => startDrag('right', event),
            onDoubleClick: resetWidths,
            onKeyDown: (event) => {
              const step = event.shiftKey ? 48 : 16
              if (event.key === 'ArrowLeft') { event.preventDefault(); nudge('right', step) }
              else if (event.key === 'ArrowRight') { event.preventDefault(); nudge('right', -step) }
              else if (event.key === 'Home') { event.preventDefault(); resetWidths() }
            },
          })),

        pendingPackage !== null
          ? h('div', { className: 'rd-mask', role: 'presentation' },
            h('div', { className: 'rd-dialog', role: 'dialog', 'aria-modal': 'true', 'aria-labelledby': 'rd-import-title' },
              h('h2', { id: 'rd-import-title', className: 'rd-dialog-title' }, t('packageImportTitle')),
              h('p', { className: 'rd-dialog-body' },
                t('packageImportBody').replace('{runs}', String(pendingPackage.runs.length))),
              h('p', { className: 'rd-dialog-note' }, t('packageHint')),
              h('div', { className: 'rd-dialog-actions' },
                h('button', {
                  type: 'button',
                  className: 'rd-button',
                  onClick: () => setPendingPackage(null),
                }, t('packageImportCancel')),
                h('button', {
                  type: 'button',
                  className: 'rd-button',
                  'data-variant': 'primary',
                  autoFocus: true,
                  onClick: confirmImportPackage,
                }, t('packageImportConfirm')))))
          : null,

        sheetOpen
          ? h('div', { className: 'rd-sheet', role: 'dialog', 'aria-modal': 'true', 'aria-label': t('viewOriginal') },
            h('header', { className: 'rd-sheet-head' },
              h('div', { className: 'rd-sheet-heading' },
                h('h2', { className: 'rd-sheet-title' }, t('viewOriginal')),
                h('p', { className: 'rd-sheet-note' }, t('originalNote'))),
              h('span', { className: 'rd-sheet-meta' },
                pages !== null && pages.status === 'ready'
                  ? `${pages.pageCount} ${t('pagesUnit')}${pages.truncated ? ` · ${t('pagesFirst')}` : ''}`
                  : ''),
              h('button', {
                type: 'button',
                className: 'rd-sheet-close',
                onClick: () => setSheetOpen(false),
                'aria-label': t('closeSheet'),
                title: t('closeSheet'),
              }, '✕')),
            h('div', { className: 'rd-sheet-body' },
              pages === null || pages.status === 'loading'
                ? h('p', { className: 'rd-sheet-status' }, t('pagesLoading'))
                : pages.status === 'error'
                  ? h('p', { className: 'rd-sheet-status rd-sheet-error' }, pages.error)
                  : pages.pages.map((page, index) => h('figure', { key: index, className: 'rd-sheet-page' },
                    h('img', {
                      src: `data:image/png;base64,${page.data}`,
                      alt: t('pageLabel').replace('{page}', String(page.page === 0 ? index + 1 : page.page)),
                      loading: 'lazy',
                    }),
                    h('figcaption', null, t('pageLabel').replace('{page}', String(page.page === 0 ? index + 1 : page.page)))))))
          : null,

        h('span', { className: 'rd-sr', role: 'status', 'aria-live': 'polite' }, announcement),
        notice === '' ? null : h('div', { className: 'rd-toast', role: 'status' }, notice))
    }

    /* ------------------------------------------------------------------ *
     * Registration
     * ------------------------------------------------------------------ */

    const inject = ['slots', 'locale', 'remote', 'remote.session']

    function apply(ctx) {
      try {
        console.info(`[review-desk] build ${BUILD}`)
      } catch {
        /* the console is optional */
      }
      ctx.effect(
        () => ctx.locale.register(NS, { zh: ZH, en: EN }),
        'review-desk: dictionaries',
      )
      const bound = ctx.locale.bind(NS)
      const t = (key) => {
        try {
          const value = bound(key)
          if (typeof value === 'string' && value !== '' && value !== key) return value
        } catch {
          /* fall back to the source dictionary */
        }
        return ZH[key] ?? key
      }

      ctx.effect(
        () => ctx.slots.inject('main', () => ctx.slots.register({
          name: 'main',
          key: PANEL_ID,
          inject: () => ({ t, remote: ctx.remote }),
        }, ReviewDeskPage)),
        'review-desk: main panel',
      )

      ctx.effect(
        () => ctx.slots.inject('sidebar.panellist', () => ctx.slots.register({
          name: 'sidebar.panellist',
          id: PANEL_ID,
          order: 20,
          label: () => t('panel'),
        }, IconPanel)),
        'review-desk: sidebar entry',
      )
    }

    return { inject, apply }
  },
})
