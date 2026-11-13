/**
 * Review Desk — host half.
 *
 * One authenticated fetch route (`/api/review-desk/analyze`) streams the result
 * of one manuscript-review operation as newline-delimited JSON frames, and one
 * agent tool exposes the same operations to the model. Both callers share
 * {@link analyze}: the route streams its frames, the tool collects them.
 *
 * The model route is the profile's default selection, read at request time.
 * Source verification asks the configured web provider for evidence first and
 * hands the numbered sources to the model, so the answer cites material it was
 * actually given instead of recalled links.
 */

import { readFile } from 'node:fs/promises'

/** Authenticated route this half owns on the shared `/api` channel. */
const ROUTE_PATH = '/api/review-desk/analyze'

/** Renders the original document's pages, for the faithful layout view. */
const PAGES_PATH = '/api/review-desk/original-pages'
const ENGINE_PATH = '/api/review-desk/engine'

/** Bounds of one original-page render; totals stay inside one JSON response. */
const PAGE_RENDER = {
  maxImages: 8,
  maxPixels: 9_000_000,
  maxDimension: 1_280,
  maxBytes: 6 * 1024 * 1024,
  sourceBytes: 30 * 1024 * 1024,
  timeoutMs: 90_000,
}

/** Operations the panel and the tool expose. `custom` runs a free-form requirement. */
const ACTIONS = new Set(['source', 'logic', 'expand', 'relate', 'custom'])

/** Upper bounds applied before any model request is prepared. */
const LIMITS = {
  text: 6_000,
  manuscript: 12_000,
  instruction: 500,
  queries: 3,
  claims: 5,
  sources: 8,
  evidenceChars: 5_000,
}

/** Output-token caps per operation; generous enough for the required structure plus a reasoning block. */
const MAX_TOKENS = { source: 4_000, logic: 4_000, expand: 5_000, relate: 4_000, custom: 4_000 }

/** Planner cap: one small JSON object, with room for a reasoning block first. */
const PLANNER_MAX_TOKENS = 1_200

/** End-to-end deadline for one operation, including web search. */
const DEADLINE_MS = 180_000

/** Release marker wrapping the rewrite an expansion offers for replacement. */
const TEXT_OPEN = '<<<TEXT'
const TEXT_CLOSE = '>>>'

const COMMON_RULES = [
  '你是一位严谨的中文文稿审核专家，服务于政府、金融机构与企业研究院的报告写作场景。',
  '只输出审核结果本身：不要寒暄，不要复述任务，不要解释你的做法。',
  '默认使用简体中文与 Markdown 输出，严格遵循给定的章节结构。',
  '所有判断都必须能在待审文字或给定资料中找到依据；不确定就写不确定，绝不编造数据、来源、机构或引文。',
].join('\n')

const ACTION_PROMPTS = {
  source: {
    system: [
      '你是严谨的事实核查员。你会收到一段待核查的文字，以及联网检索到的资料条目（编号 S1、S2……）。',
      '',
      '核查范围：数据、时间、机构名称、政策与法规、人物职务、引用与出处、可被外部验证的事件。',
      '',
      '规则：',
      '1. 只依据给定的资料条目与公认常识判断。资料没有覆盖到的，明确写“未找到可核实的公开来源”。',
      '2. 每条判断后用 [S1] 的形式标注依据；没有依据的判断不得标注编号。',
      '3. 判断只有四种：有据可查 / 部分一致 / 相互矛盾 / 无法核实。',
      '4. 不要复述整段原文，只提取可核查的要点，每条不超过 30 字。',
      '',
      '输出结构（严格遵守）：',
      '### 核查结论',
      '- **要点**：判断 —— 依据 [S1]。',
      '### 需要留意的表述',
      '- 原文表述 → 建议改法（只在确有风险时给出；没有风险就写“未发现需要修改的表述”。）',
      '### 检索覆盖',
      '一句话说明本次检索覆盖到什么程度、还缺什么。',
    ].join('\n'),
  },
  logic: {
    system: [
      '你是资深编辑，擅长论证结构诊断。',
      '',
      '检查这些方面：论点与论据是否匹配、推理链条是否完整、有无循环论证、以偏概全、因果倒置、概念偷换、量词与程度失当、前后自相矛盾、结论是否超出证据。',
      '',
      '输出结构（严格遵守）：',
      '### 结构判断',
      '2–3 句：这段文字在整段论证中承担什么作用，是否成立。',
      '### 逻辑问题',
      '按严重程度排序。每条写成：**问题类型** —— 原文摘录（≤20 字）→ 为什么不成立 → 修改方向。',
      '若没有实质问题，写“未发现明显逻辑漏洞”，并说明理由。',
      '### 收紧建议',
      '2–4 条可直接执行的修改：改哪一句、怎么改。',
    ].join('\n'),
  },
  expand: {
    system: [
      '你是擅长研究报告与公文文体的资深撰稿人。',
      '',
      '任务：把用户给出的段落扩写为更充分、可直接入稿的文字，篇幅约为原文的 2–3 倍。保持原意、语气与第一人称/第三人称口径，不改变任何事实、数字、机构名与结论强度。',
      '',
      '扩写时补足：背景与由来、机制与路径、依据与先例、具体做法、影响与边界。不得虚构数据、引文、机构、时间或案例；需要作者补充真实数据的地方，在说明中单独指出，不要写进正文。',
      '',
      '输出结构（严格遵守，标记行必须原样出现且各占一行）：',
      TEXT_OPEN,
      '（扩写后的正文：保持原有段落结构，可直接替换原文，不要出现任何标记或说明）',
      TEXT_CLOSE,
      '### 扩写说明',
      '- 补了什么、为什么补。',
      '- 哪些地方需要作者补充真实数据或引文。',
    ].join('\n'),
  },
  custom: {
    system: [
      '【输出格式（最高优先级，先判断再动笔）】',
      'A. 作者在让你改这段文字（改写、改成、压缩、精简、扩写、润色、换口径、调语气、纠错等）时，必须按下面格式输出：',
      '<<<TEXT',
      '（改好的完整文字，可直接替换原段落；不要带标题、不要带“以下是”之类的引导语）',
      '>>>',
      '### 修改说明',
      '- 改了哪些地方、为什么。',
      'B. 作者在提问、分析或核查（有没有问题、是否成立、哪里不对、为什么、帮我检查等）时，正常回答，不要输出 <<<TEXT 标记。',
      '拿不准时按 A 处理，并在修改说明里补一句“如需其他形式请说明”。',
      '',
      '【其他要求】',
      '1. 【任务要求】就是本次任务本身，必须完整满足它的每一个限定（篇幅、口径、文体、格式、人称、受众）。',
      '2. 只处理【待审文字】；需要结合上下文时参考【全文】，但不要改动未被要求的部分。',
      '3. 要求含糊时，按最合理的一种理解直接产出成品；不要先反问。',
      '4. 要求与文字无关或无法完成时，直接说明原因并给出可执行建议，不要硬编。',
      '5. 除非要求指定了其他格式，用 Markdown 输出。',
    ].join('\n'),
  },
  revise: {
    system: [
      '你是资深文稿编辑，正在修订一段已经写好的文字。',
      '',
      '输入包含【任务要求】【当前文字】（需要修订的稿子）与可选的【原始段落】（事实参照）。',
      '',
      '规则：',
      '1. 在【当前文字】的基础上落实【任务要求】，保留原有事实、数字与专有名词，除非要求明确改动。',
      '2. 【原始段落】只作事实参照，不要退回它的写法，也不要把说明写进正文。',
      '3. 输出形式固定：改好的完整文字放在 <<<TEXT 与 >>> 之间（两个标记各占一行），随后用「### 修改说明」写清本次改了什么。',
      '4. 不要输出“以下是”，不要让成品依赖上下文才能读懂。',
    ].join('\n'),
  },
  relate: {
    system: [
      '你是负责统稿的资深编辑。你会收到【全文】与【选中段落】。',
      '',
      '任务：分析选中段落与全文其他部分的关联，服务统稿决策。',
      '',
      '输出结构（严格遵守）：',
      '### 在全文中的位置',
      '这段承担什么功能，与前文、后文分别是什么关系。',
      '### 关联点',
      '- 呼应：与哪一处（摘录 ≤15 字）形成支撑、递进或对照。',
      '- 冲突：与哪一处表述不一致、重复或口径不同；没有就写“未发现”。',
      '- 缺口：全文缺少什么，使这段显得孤立。',
      '### 统稿建议',
      '2–4 条：保留、前移、合并、删除或补写，具体到段落与语句。',
    ].join('\n'),
  },
}

/** Frame type guards kept small: the client trusts only the frame `type`. */
function isAbort(error) {
  return error !== null
    && typeof error === 'object'
    && (error.name === 'AbortError' || error.code === 'ABORT_ERR')
}

function messageOf(error) {
  if (error instanceof Error) return error.message
  if (typeof error === 'string') return error
  try {
    return JSON.stringify(error)
  } catch {
    return String(error)
  }
}

function clampText(value, limit) {
  if (typeof value !== 'string') return ''
  const trimmed = value.trim()
  return trimmed.length > limit ? `${trimmed.slice(0, limit)}\n……（内容过长，已截断）` : trimmed
}

/** Read one request body as the documented payload, rejecting unknown actions. */
function readPayload(raw) {
  if (raw === null || typeof raw !== 'object' || Array.isArray(raw)) {
    throw new Error('请求格式不正确：需要 JSON 对象。')
  }
  const action = raw.action
  if (typeof action !== 'string' || !ACTIONS.has(action)) {
    throw new Error('未知的审核动作。')
  }
  const revise = raw.revise === true
  if (revise && action !== 'custom') throw new Error('修订请求必须使用 custom 动作。')
  const text = clampText(raw.text, LIMITS.text)
  if (text === '') throw new Error('请先选取需要审核的文字。')
  const instruction = clampText(raw.instruction, LIMITS.instruction)
  if (action === 'custom' && instruction === '') {
    throw new Error('请写下你的处理要求，例如“改写成领导讲话口径，控制在 200 字内”。')
  }
  return {
    action,
    text,
    manuscript: clampText(raw.manuscript, LIMITS.manuscript),
    instruction,
    title: clampText(raw.title, 200),
    provider: typeof raw.provider === 'string' ? raw.provider.slice(0, 100) : '',
    model: typeof raw.model === 'string' ? raw.model.slice(0, 200) : '',
    reasoningEffort: typeof raw.reasoningEffort === 'string' ? raw.reasoningEffort.slice(0, 40) : '',
    revise: raw.revise === true,
    original: clampText(raw.original, LIMITS.text),
    /** Evidence gathered before the model call; filled in by {@link analyze}. */
    sources: [],
    /** Why source checking produced no evidence, when it did not. */
    searchNote: '',
  }
}

/** One numbered evidence list; the model cites these identifiers verbatim. */
function evidenceBlock(sources, note) {
  if (sources.length === 0) {
    const reason = typeof note === 'string' && note !== '' ? `（原因：${note}）` : ''
    return `（本次未获得联网资料${reason}，只能依据公认常识判断，不得给出任何来源编号。）`
  }
  const rows = []
  let used = 0
  for (const [index, source] of sources.entries()) {
    const label = `S${index + 1}`
    const summary = source.summary ?? source.snippet ?? ''
    const line = `[${label}] ${source.title ?? '(无标题)'} — ${summary.slice(0, 400)}（${source.url}）`
    if (used + line.length > LIMITS.evidenceChars) break
    used += line.length
    rows.push(line)
  }
  return rows.join('\n')
}

/** Compose the single user message one operation sends. */
function composeUserMessage(request, sources) {
  const parts = []
  if (request.title !== '') parts.push(`【文稿标题】${request.title}`)
  if (request.action === 'custom' || request.revise) {
    parts.push(`【任务要求】\n${request.instruction === '' ? '（作者没有填写要求：请对这段文字做一次全面的审稿，指出问题并给出修改建议。）' : request.instruction}`)
  }
  if (request.revise) {
    parts.push(`【当前文字】\n${request.text}`)
    if (request.original !== '') parts.push(`【原始段落（事实参照）】\n${request.original}`)
  } else {
    parts.push(`【待审文字】\n${request.text}`)
  }

  if ((request.action === 'custom' || request.revise) && request.manuscript !== '') {
    parts.push(`【全文（仅作上下文参考）】\n${request.manuscript}`)
  }
  if (request.action === 'relate') {
    parts.push(`【全文】\n${request.manuscript === '' ? '（调用方未提供全文，请只依据【待审文字】判断关联，并说明缺少全文。）' : request.manuscript}`)
  }
  if (request.action === 'source') {
    parts.push(`【检索到的资料】\n${evidenceBlock(sources, request.searchNote)}`)
  }
  if (request.instruction !== '' && request.action !== 'custom') {
    parts.push(`【作者补充要求】\n${request.instruction}`)
  }
  return parts.join('\n\n')
}

/**
 * The route one request uses: the panel's explicit choice when it still names a
 * registered provider, otherwise the profile default.
 * @param ctx - context carrying `llm` and `agentDefaultModel`.
 * @param request - validated payload; `provider` and `model` are optional.
 * @returns the provider and model to call.
 * @throws {Error} when the chosen provider is no longer registered.
 */
function resolveRoute(ctx, request) {
  const effort = typeof request.reasoningEffort === 'string' ? request.reasoningEffort : ''
  if (typeof request.provider === 'string' && request.provider !== ''
    && typeof request.model === 'string' && request.model !== '') {
    const known = ctx.llm.listProviders().some(entry => entry.id === request.provider)
    if (!known) throw new Error(`所选模型通道「${request.provider}」已不可用，请重新选择模型。`)
    return { provider: request.provider, model: request.model, reasoningEffort: effort }
  }
  const selection = ctx.agentDefaultModel.currentSelection()
  const provider = typeof selection?.provider === 'string' ? selection.provider : ''
  const model = typeof selection?.model === 'string' ? selection.model : ''
  if (provider === '' || model === '') {
    throw new Error('尚未配置默认模型，请在「设置 → 模型」中选择一个模型后重试。')
  }
  // Following the default keeps its reasoning level unless the panel overrides it.
  const fallback = typeof selection?.reasoningEffort === 'string' ? selection.reasoningEffort : ''
  return { provider, model, reasoningEffort: effort === '' ? fallback : effort }
}

/**
 * One prepared call as `{ text }` pieces and, at the end, `{ warning }` when the
 * provider stopped at the output cap.
 */
async function* streamText(ctx, request, signal) {
  const { provider, model, reasoningEffort } = resolveRoute(ctx, request)

  let prepared
  try {
    prepared = await ctx.llm.prepareCall(
      {
        provider,
        model,
        maxTokens: MAX_TOKENS[request.action],
        ...(reasoningEffort === '' ? {} : { reasoningEffort }),
      },
      signal,
    )
  } catch (error) {
    // A level the model no longer accepts must not take the whole review down.
    if (reasoningEffort === '' || isAbort(error)) throw error
    prepared = await ctx.llm.prepareCall(
      { provider, model, maxTokens: MAX_TOKENS[request.action] },
      signal,
    )
    yield { warning: { code: 'reasoning-fallback', effort: reasoningEffort } }
  }

  const messages = [{ role: 'user', content: [{ type: 'text', text: composeUserMessage(request, request.sources) }] }]

  const prompt = request.revise ? ACTION_PROMPTS.revise : ACTION_PROMPTS[request.action]
  for await (const chunk of prepared.stream({
    ...prepared.config,
    system: `${COMMON_RULES}\n\n${prompt.system}`,
    messages,
    signal,
  })) {
    if (chunk.type === 'text-delta') {
      yield { text: chunk.text }
      continue
    }
    if (chunk.type === 'finish') {
      const reason = chunk.reason
      if (reason?.kind === 'error') {
        throw new Error(reason.failure?.message ?? '模型请求失败。')
      }
      if (reason?.kind === 'aborted') {
        const error = new Error(reason.failure?.message ?? '请求已取消。')
        error.name = 'AbortError'
        throw error
      }
      if (reason?.kind === 'max-tokens') {
        yield { warning: '输出达到长度上限，结果可能不完整；可以缩短选中文字后重试。' }
      }
    }
  }
}

/** Ask the model for search queries and the claims worth checking. */
async function planSearch(ctx, request, signal) {
  const system = [
    '你是事实核查的检索规划员。',
    '从待核查文字中提取 2–3 条中文检索式，用于在搜索引擎里核实其中的关键事实；检索式要含主体、时间、数字等限定词，不要照抄整句。',
    '同时列出 2–5 条可核查的事实性陈述，每条不超过 30 字。',
    '只输出 JSON，不要代码块，不要解释：{"queries":["…"],"claims":["…"]}',
  ].join('\n')

  const input = [request.title === '' ? '' : `【文稿标题】${request.title}`, `【待核查文字】\n${request.text}`]
    .filter(part => part !== '')
    .join('\n\n')

  let route
  try {
    route = resolveRoute(ctx, request)
  } catch {
    return { queries: [], claims: [] }
  }

  const prepared = await ctx.llm.prepareCall(
    {
      provider: route.provider,
      model: route.model,
      maxTokens: PLANNER_MAX_TOKENS,
      ...(route.reasoningEffort === '' ? {} : { reasoningEffort: route.reasoningEffort }),
    },
    signal,
    (controls, captured) => {
      const least = captured?.reasoning?.efforts?.[0]
      return least === undefined ? controls : { ...controls, reasoningEffort: least.id }
    },
  )
  let raw = ''
  for await (const chunk of prepared.stream({
    ...prepared.config,
    system,
    messages: [{ role: 'user', content: [{ type: 'text', text: input }] }],
    signal,
  })) {
    if (chunk.type === 'text-delta') raw += chunk.text
    if (chunk.type === 'finish') {
      if (chunk.reason?.kind === 'error') throw new Error(chunk.reason.failure?.message ?? '检索规划失败。')
      if (chunk.reason?.kind === 'aborted') return { queries: [], claims: [] }
    }
  }

  const start = raw.indexOf('{')
  const end = raw.lastIndexOf('}')
  if (start < 0 || end <= start) return { queries: [], claims: [] }
  try {
    const parsed = JSON.parse(raw.slice(start, end + 1))
    const pick = (value, limit) => Array.isArray(value)
      ? value.filter(item => typeof item === 'string' && item.trim() !== '').slice(0, limit).map(item => item.trim())
      : []
    return { queries: pick(parsed.queries, LIMITS.queries), claims: pick(parsed.claims, LIMITS.claims) }
  } catch {
    return { queries: [], claims: [] }
  }
}

/** Run each query through the configured web provider, deduplicated by URL. */
async function gatherSources(ctx, queries, signal, emit) {
  const web = ctx.get('web')
  if (web === undefined) return { sources: [], note: '未配置联网检索能力' }
  const seen = new Set()
  const sources = []
  let failure = ''
  for (const query of queries) {
    signal.throwIfAborted()
    emit({ type: 'stage', stage: 'searching', detail: query })
    let result
    try {
      result = await web.search({ query, maxResults: 4 }, signal)
    } catch (error) {
      if (isAbort(error)) throw error
      if (failure === '') failure = messageOf(error)
      continue
    }
    for (const source of result?.sources ?? []) {
      if (typeof source?.url !== 'string' || source.url === '' || seen.has(source.url)) continue
      seen.add(source.url)
      sources.push({
        url: source.url,
        ...(typeof source.title === 'string' && source.title !== '' ? { title: source.title } : {}),
        ...(typeof source.snippet === 'string' && source.snippet !== '' ? { snippet: source.snippet } : {}),
        ...(typeof source.summary === 'string' && source.summary !== '' ? { summary: source.summary } : {}),
        ...(typeof source.publishedAt === 'string' && source.publishedAt !== '' ? { publishedAt: source.publishedAt } : {}),
      })
      if (sources.length >= LIMITS.sources) return { sources, note: null }
    }
  }
  if (sources.length > 0) return { sources, note: null }
  return {
    sources,
    note: failure === ''
      ? '本次检索没有返回可用资料'
      : `联网检索失败：${failure}`,
  }
}

/**
 * One review operation as protocol frames.
 * @param {object} ctx - context carrying `llm`, `agentDefaultModel`, and optional `web`.
 * @param {object} request - validated payload.
 * @param {AbortSignal} signal - cancellation for the whole operation.
 * @param {(frame: object) => void} emit - frame sink.
 */
async function analyze(ctx, request, signal, emit) {
  const route = resolveRoute(ctx, request)
  emit({
    type: 'meta',
    action: request.action,
    provider: route.provider,
    model: route.model,
    reasoningEffort: route.reasoningEffort,
  })

  request.sources = []
  if (request.action === 'source') {
    emit({ type: 'stage', stage: 'planning' })
    const plan = await planSearch(ctx, request, signal)
    if (plan.queries.length > 0) {
      const gathered = await gatherSources(ctx, plan.queries, signal, emit)
      request.sources = gathered.sources
      request.searchNote = gathered.note ?? ''
      emit({ type: 'sources', sources: gathered.sources, note: gathered.note, claims: plan.claims, queries: plan.queries })
    } else {
      request.searchNote = '未能生成检索式'
      emit({ type: 'sources', sources: [], note: request.searchNote, claims: [], queries: [] })
    }
    emit({ type: 'stage', stage: 'writing' })
  }

  let text = ''
  for await (const piece of streamText(ctx, request, signal)) {
    if (typeof piece.text === 'string') {
      text += piece.text
      emit({ type: 'delta', text: piece.text })
      continue
    }
    if (typeof piece.warning === 'string') emit({ type: 'warning', message: piece.warning })
    else if (typeof piece.warning === 'object' && piece.warning !== null) {
      emit({ type: 'warning', ...piece.warning })
    }
  }
  emit({ type: 'done', chars: text.length })
  return text
}

/** Collect one operation into a single result for the agent tool. */
async function analyzeCollect(ctx, request, signal) {
  const frames = []
  const text = await analyze(ctx, request, signal, frame => frames.push(frame))
  const sources = frames.find(frame => frame.type === 'sources')?.sources ?? []
  const warning = frames.find(frame => frame.type === 'warning')?.message ?? ''
  const note = frames.find(frame => frame.type === 'sources')?.note ?? ''
  const body = text.trim() === ''
    ? `本次没有产生审核结果。${warning === '' ? '' : warning}${request.action === 'source' && note !== '' ? `（${note}）` : ''}`
    : text
  return { text: body, sources, warning }
}

/** One NDJSON response whose producer pushes frames until it settles. */
function frameResponse(produce) {
  const encoder = new TextEncoder()
  let open = true
  const stream = new ReadableStream({
    async start(controller) {
      const emit = (frame) => {
        if (!open) return
        try {
          controller.enqueue(encoder.encode(`${JSON.stringify(frame)}\n`))
        } catch {
          open = false
        }
      }
      try {
        await produce(emit)
      } catch {
        /* the producer reports its own failures as frames */
      } finally {
        open = false
        try {
          controller.close()
        } catch {
          /* the client already disconnected */
        }
      }
    },
    cancel() {
      open = false
    },
  })
  return new Response(stream, {
    status: 200,
    headers: {
      'content-type': 'application/x-ndjson; charset=utf-8',
      'cache-control': 'no-store',
      'x-content-type-options': 'nosniff',
    },
  })
}

function jsonResponse(value, status) {
  return new Response(JSON.stringify(value), {
    status,
    headers: { 'content-type': 'application/json; charset=utf-8', 'cache-control': 'no-store' },
  })
}

/**
 * Handle one analyze request on the shared authenticated channel.
 * @param {object} ctx - context carrying the model services.
 * @param {Request} fetchRequest - the authenticated browser request.
 * @returns {Promise<Response>} a streamed NDJSON result or a JSON error.
 */
async function handleAnalyzeRequest(ctx, fetchRequest) {
  if (fetchRequest.method !== 'POST') {
    return new Response(null, { status: 405, headers: { allow: 'POST' } })
  }
  let payload
  try {
    payload = await fetchRequest.json()
  } catch {
    return jsonResponse({ error: '请求体不是有效的 JSON。' }, 400)
  }
  let request
  try {
    request = readPayload(payload)
  } catch (error) {
    return jsonResponse({ error: messageOf(error) }, 400)
  }

  const timeout = AbortSignal.timeout(DEADLINE_MS)
  const signal = AbortSignal.any([fetchRequest.signal, timeout])
  return frameResponse(async (emit) => {
    try {
      await analyze(ctx, request, signal, emit)
    } catch (error) {
      // The browser already left: nothing can receive a frame.
      if (fetchRequest.signal.aborted) return
      emit({
        type: 'error',
        message: timeout.aborted
          ? '本次审核超时（3 分钟），请缩短选中文字后重试。'
          : messageOf(error),
      })
    }
  })
}

/** Render the first pages of an uploaded document as PNG data URLs. */
async function handlePagesRequest(ctx, request) {
  if (request.method !== 'POST') {
    return new Response(null, { status: 405, headers: { allow: 'POST' } })
  }
  const officeKit = ctx.get('officeKit')
  if (officeKit === undefined) {
    return jsonResponse({ error: '本机未安装文档渲染引擎，无法生成原稿版式预览。' }, 501)
  }
  const url = new URL(request.url)
  const extension = (url.searchParams.get('ext') ?? '').toLowerCase()
  if (!['docx', 'doc', 'pdf'].includes(extension)) {
    return jsonResponse({ error: '只支持 .docx、.doc 与 .pdf 的原稿预览。' }, 400)
  }
  let bytes
  try {
    const body = new Uint8Array(await request.arrayBuffer())
    if (body.length === 0) return jsonResponse({ error: '文件为空。' }, 400)
    if (body.length > PAGE_RENDER.sourceBytes) {
      return jsonResponse({ error: '文件超过 30 MB，暂不支持原稿预览。' }, 413)
    }
    bytes = body
  } catch {
    return jsonResponse({ error: '无法读取上传的文件。' }, 400)
  }

  const version = String(bytes.length)
  const source = {
    key: `review-desk-original-${extension}`,
    version,
    bytes: bytes.length,
    async read(_signal, maxBytes) {
      if (bytes.length > maxBytes) throw new Error('文件超出渲染上限。')
      return { bytes, version }
    },
  }

  const signal = AbortSignal.any([request.signal, AbortSignal.timeout(PAGE_RENDER.timeoutMs)])
  try {
    const result = await officeKit.renderImages({
      extension,
      source,
      maxImages: PAGE_RENDER.maxImages,
      maxPixels: PAGE_RENDER.maxPixels,
      maxDimension: PAGE_RENDER.maxDimension,
      maxBytes: PAGE_RENDER.maxBytes,
    }, signal)
    return jsonResponse({
      pageCount: result.pageCount,
      dpi: result.dpi,
      missingFonts: result.missingFonts,
      truncated: result.pageCount > result.images.length,
      pages: result.images.map(image => ({
        data: Buffer.from(image.png).toString('base64'),
        width: image.width,
        height: image.height,
        page: image.page ?? 0,
      })),
    }, 200)
  } catch (error) {
    if (isAbort(error)) return jsonResponse({ error: '原稿渲染超时或已取消。' }, 504)
    return jsonResponse({ error: `原稿渲染失败：${messageOf(error)}` }, 500)
  }
}

/**
 * Browser engine assets shipped with the plugin. Names are fixed, so the route
 * can serve only these files from this package.
 */
const ENGINE_ASSETS = {
  'superdoc.js': { file: 'superdoc.js', type: 'text/javascript; charset=utf-8' },
  'superdoc.css': { file: 'superdoc.css', type: 'text/css; charset=utf-8' },
}

async function handleEngineRequest(request) {
  if (request.method !== 'GET') return new Response(null, { status: 405, headers: { allow: 'GET' } })
  const name = new URL(request.url).searchParams.get('name') ?? ''
  const asset = Object.hasOwn(ENGINE_ASSETS, name) ? ENGINE_ASSETS[name] : undefined
  if (asset === undefined) return new Response('not found', { status: 404 })
  try {
    const bytes = await readFile(new URL(`./vendor/superdoc/${asset.file}`, import.meta.url))
    return new Response(bytes, {
      status: 200,
      headers: {
        'content-type': asset.type,
        'content-length': String(bytes.length),
        'cache-control': 'public, max-age=3600',
      },
    })
  } catch (error) {
    return new Response(`engine asset unavailable: ${messageOf(error)}`, { status: 500 })
  }
}

/** Register the routes and the agent tool on a context that has the model services. */
function install(ctx) {
  ctx.effect(
    () => ctx.connection.fetch.register({
      path: ROUTE_PATH,
      methods: ['POST'],
      requestBody: 'buffered',
      fetch: request => handleAnalyzeRequest(ctx, request),
    }),
    'review-desk: analyze route',
  )

  ctx.effect(
    () => ctx.connection.fetch.register({
      path: PAGES_PATH,
      methods: ['POST'],
      requestBody: 'streaming',
      fetch: request => handlePagesRequest(ctx, request),
    }),
    'review-desk: original pages route',
  )

  ctx.effect(
    () => ctx.connection.fetch.register({
      path: ENGINE_PATH,
      methods: ['GET'],
      fetch: request => handleEngineRequest(request),
    }),
    'review-desk: browser engine assets',
  )

  ctx.effect(
    () => ctx.tools.register({
      name: 'review_manuscript',
      description: [
        '审核一段中文文稿：action=source 核查事实来源，action=logic 分析论证逻辑，',
        'action=expand 扩写为可直接入稿的文字，action=relate 分析与全文其他部分的关联，',
        'action=custom 按 instruction 里的具体要求处理（改写口径、压缩、调整语气等）。',
        '结果与「文稿审核台」面板中同一操作完全一致；用户明确要求审核文稿时使用它。',
      ].join(''),
      parameters: {
        type: 'object',
        properties: {
          action: {
            type: 'string',
            enum: ['source', 'logic', 'expand', 'relate', 'custom'],
            description: '审核类型；custom 表示按 instruction 处理这段文字。',
          },
          text: { type: 'string', description: '待审核的原文片段，尽量完整，不要改写。' },
          manuscript: { type: 'string', description: '文稿全文，relate 需要它来判断段落间的关联。' },
          instruction: { type: 'string', description: '用户的具体要求；custom 必填，其他类型作为补充说明。' },
        },
        required: ['action', 'text'],
        additionalProperties: false,
      },
      output: {
        schema: {
          type: 'object',
          properties: {
            action: { type: 'string' },
            text: { type: 'string' },
            sources: {
              type: 'array',
              items: {
                type: 'object',
                properties: {
                  url: { type: 'string' },
                  title: { type: 'string' },
                  snippet: { type: 'string' },
                },
                required: ['url'],
              },
            },
          },
          required: ['action', 'text'],
        },
        render: (_args, value) => [{ type: 'text', text: typeof value?.text === 'string' ? value.text : '' }],
      },
      async execute(args, exec) {
        const request = readPayload({ ...args, title: '' })
        const signal = AbortSignal.any([exec.signal, AbortSignal.timeout(DEADLINE_MS)])
        const result = await analyzeCollect(ctx, request, signal)
        return { action: request.action, text: result.text, sources: result.sources }
      },
      isConcurrencySafe: () => true,
    }),
    'review-desk: review_manuscript tool',
  )
}

/** Hard dependencies: the authenticated channel and the tool registry. */
export const inject = ['connection', 'tools']

/**
 * Mount the review desk's host capability.
 * @param {object} ctx - plugin context.
 */
export function apply(ctx) {
  ctx.inject(['llm', 'agentDefaultModel'], install)
}
