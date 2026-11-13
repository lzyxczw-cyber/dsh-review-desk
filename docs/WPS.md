# WPS 加载项方向（`wps` 分支）

> 实验分支，不与 `main` 合并。目标：让「文稿审核台」直接跑在 WPS 文字里，
> 用 WPS 原生文档模型解决「导入即 Markdown、格式丢失」的问题。

## 为什么是 WPS

主分支把 Markdown 当作唯一事实源，导入 `.docx` 会丢字体、字号、对齐、页眉页脚、图片等版式。
WPS 加载项路线下，文档始终是 WPS 打开的真 `.docx`：

- 格式、字体、段落、对齐、撤销/重做 = WPS 原生能力，零实现成本；
- 审稿结果以「批注」或「修订」写回，符合审阅习惯，导出即原格式；
- AI 面板（任务窗格）只负责「选中 → 审核 → 结果」，不碰排版。

## 架构

```
WPS 文字（真 .docx）
├── wps-addin/  加载项：功能区 + 任务窗格（taskpane.html/js）
│      │  选中文字 → POST /analyze（NDJSON 流式）
│      ▼
│   本机回环桥  http://127.0.0.1:47831   （由 review-desk Host 半启动）
│      │  复用 analyze()：提示词、默认模型、web 检索
│      ▼
└── DSH 进程（review-desk 插件）
```

为什么需要桥：WPS 加载项运行在 WPS 里，无法使用 DSH 的 owner-authenticated `/api` 通道；
桥把同一套 `analyze()` 以回环 NDJSON 暴露出来，模型凭据与检索能力继续留在 DSH 内。

## 桥接 API

`index.js` 的 `startWpsBridge(ctx)` 在插件 `install()` 时启动：

- `GET /health` → `{"ok":true,"service":"review-desk-wps-bridge","bridge":"1"}`
- `GET /models` → 默认模型/推理等级、可选模型（含推理等级）、各动作输出上限
- `POST /analyze`（JSON，字段同 DSH 面板请求）→ `application/x-ndjson` 流，帧类型：
  `meta` / `stage` / `sources` / `delta` / `warning` / `done` / `error`
- 请求可选字段：`provider` / `model` / `reasoningEffort` / `maxTokens`（1000–16000，0 或省略用动作默认）
- 默认输出上限：来源/逻辑/关联/要求 6000 tokens，扩写 8000 tokens
- CORS：`access-control-allow-origin: *`（回环，无需凭据）
- Body 上限 4 MB，单次审核超时 180 s
- 环境变量：`REVIEW_DESK_WPS_BRIDGE=0` 关闭；`REVIEW_DESK_WPS_BRIDGE_PORT` 改端口

请求示例：

```sh
curl -N http://127.0.0.1:47831/analyze \
  -H 'content-type: application/json' \
  -d '{"action":"logic","text":"待审核的一段文字","title":"示例文稿"}'
```

## 验证结论（2026-11-14，WPS 12.1.26046 / macOS）

已在本机用独立 WPS 实例验证通过（不影响用户正在编辑的文档）：

| 能力 | 结果 |
|---|---|
| 加载项加载 / ribbon / 任务窗格 | ✅ |
| 读选区、`Range` 定位、替换 | ✅ |
| `TrackRevisions` 修订（产生 `<w:ins>`/`<w:del>`，可接受） | ✅ |
| 批注 `Range.Comments.Add`（读取用 `Comment.Range.Text`） | ✅ |
| 字体 / 段落 / 对齐 / 查找 | ✅ |
| 撤销 / 重做（`Document.Undo` / `Document.Redo`） | ✅ |
| `SaveAs2` 另存 docx、`ExportAsFixedFormat` 导出 PDF | ✅（沙盒内路径） |
| 任务窗格 → 本机 HTTP（fetch，CORS） | ✅ |
| **v0.2 端到端**：扩写 695 字 → 定位 → 批注 683 字 → 修订替换 365 字 → **撤销替换（拒绝 2 个修订，正文长度/上下文/范围全部还原）** → 追加要求 240 字 | ✅ |
| **v0.2 面板工作流**（jsdom）：历史持久化/筛选/折叠、追加要求链、输出触顶重试、模型/推理/上限设置随请求发送 | ✅ |

验证产物：`wps-review-demo/evidence/v0.2/`（report-v0.2.jsonl、wps-review-desk-verify-v0.2.docx，含 `<w:ins>`×7、`<w:del>`×5、批注）。

## 面板工作流（v0.2）

- 结果历史保存在本机 localStorage（最多 20 条），支持按动作筛选、折叠/展开、清空；
- 任意结果可「追加要求」继续修改（custom + revise），形成修订链；
- 替换原文以 `TrackRevisions` 写入；「撤销替换」会拒绝覆盖该范围的修订并还原原文（失败时提示用 WPS 撤销）；
- 「定位回原文」选中并滚动到文档中的原文位置；
- 顶部设置可选择 DSH 里的模型通道、推理等级和输出上限；触顶时提供「提高输出上限重试」。

## 已知风险

1. **Mac 分发**：仅支持本地 `publish.xml` 侧载；WPS Mac 在线部署（58890）不可用，正式分发需评估 WPS 开放平台（Windows 为主）。
2. **桥接安全**：只监听 `127.0.0.1`、无鉴权；任何本机进程都能调用（会消耗模型额度）。后续可加随机 token（写入 `publish.xml` 与加载项配置）。
3. **版本差异**：WPS 各版本对 `TrackRevisions` / `Comments` / 事件的支持可能有差异，需要兼容性测试。
4. **加载项只在启动时加载**：迭代/更新需要重启 WPS。
5. **任务窗格内核**：CEF Chrome 104，现代框架可用但需回归。

## 后续路线

1. 结果工作流（历史、追加要求、撤销替换）与审核包；
2. 加载项设置项（模型选择、桥接 token、快捷动作）；
3. 打包分发（Windows 在线部署 / Mac 侧载安装器 / 开放平台上架）；
4. 与 DSH 插件共享结果格式（审核记录 JSON），实现「WPS 审稿 → DSH 继续」。
