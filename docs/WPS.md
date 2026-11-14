# WPS 加载项方向（`wps` 分支）

> 实验分支，不与 `main` 合并。目标：让「审悦」直接跑在 WPS 文字里，
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

- `GET /health` → `{"ok":true,"service":"review-desk-wps-bridge","bridge":"1","auth":"token","pairing":true,"exportDir":"…"}`
- `GET /pair` → 自动生成的访问令牌（需 `x-review-desk-pair: 1` 且回环来源；面板启动时自动配对）
- `GET /models` → 默认模型/推理等级、可选模型（含推理等级）、各动作输出上限
- `POST /export` → `{filename, content}`，写入导出目录并返回 `{ok,path,filename,bytes}`（拒绝路径穿越、同名自动加序号）
- `GET /addin/...` → 由桥直接提供加载项页面（安装后 WPS 从这里加载；只监听回环）
- `POST /analyze`（JSON，字段同 DSH 面板请求）→ `application/x-ndjson` 流，帧类型：
  `meta` / `stage` / `sources` / `delta` / `warning` / `done` / `error`
- 请求可选字段：`provider` / `model` / `reasoningEffort` / `maxTokens`（1000–16000，0 或省略用动作默认）
- 默认输出上限：来源/逻辑/关联/要求 6000 tokens，扩写 8000 tokens
- **来源限制**：只接受回环来源（`http://127.0.0.1[:port]`、`http://localhost[:port]`、打包后的 `null` 来源）或非浏览器客户端；其他 Origin 一律 403
- **令牌默认开启且固定**：首次启动生成令牌并写入 `~/.dsh/review-desk/wps-bridge-token`（0600），DSH 重启复用同一令牌；面板通过 `/pair` 自动配对；`/models`、`/analyze`、`/export` 需带 `x-review-desk-token` 头（`/health` 保持开放用于连通性检测）
- Body 上限 4 MB（导出 8 MB），单次审核超时 180 s
- 环境变量：`REVIEW_DESK_WPS_BRIDGE=0` 关闭；`REVIEW_DESK_WPS_BRIDGE_PORT` 改端口；`REVIEW_DESK_WPS_TOKEN` 指定固定令牌；`REVIEW_DESK_WPS_TOKEN_FILE` 改令牌文件；`REVIEW_DESK_WPS_BRIDGE_AUTH=open` 关闭校验；`REVIEW_DESK_WPS_EXPORT_DIR` 改导出目录；`REVIEW_DESK_WPS_JSADDONS_DIR` 改加载项目录（默认按平台：macOS `~/Library/Containers/com.kingsoft.wpsoffice.mac/Data/.kingsoft/wps/jsaddons`、Windows `%APPDATA%/kingsoft/wps/jsaddons`、Linux `~/.local/share/Kingsoft/wps/jsaddons`）

请求示例：

```sh
curl -N http://127.0.0.1:47831/analyze \
  -H 'content-type: application/json' \
  -d '{"action":"logic","text":"待审核的一段文字","title":"示例文稿"}'
```

## 安装与分发（v1.0）

装好 DSH 版（review-desk 插件）后，安装 WPS 版只需要写一条 `publish.xml` 登记；**页面由 DSH 桥的 `http://127.0.0.1:47831/addin/` 提供**，DSH 保持运行即可，不再需要 `wpsjs debug` 开发服务器。

- **DSH 面板**：顶栏「WPS 版 → 安装 WPS 加载项」（`POST /api/review-desk/wps-addin`，`action=install|status|uninstall`）；
- **命令行**：`node scripts/install-wps-addin.mjs install`（`pnpm wps:install`），`status` / `uninstall` 同理；
- **Agent**：`install_wps_addin` 工具（`action=install|status|uninstall`）。

安装后写出的条目形如：

```xml
<jspluginonline name="wps-review-desk" type="wps"
  url="http://127.0.0.1:47831/addin/" debug="" enable="enable_dev"
  install="null" version="1.0.0"/>
```

要点：

1. 必须用 `jspluginonline` + `enable_dev`（WPS 不接受 `jsplugin` 的 http URL；离线文件目录会因 `file://` 页面无法 fetch 回环桥而不可用）；
2. 完全退出 WPS 再打开；首次加载 WPS 会弹「允许运行」确认，点允许（本机测试已通过）；更新版本后会再弹一次「加载项已被修改」；
3. 卸载：面板/命令行 `uninstall`，或删除 `publish.xml` 里的条目；
4. **Windows**：安装器同样可用（Node 运行），写 `%APPDATA%/kingsoft/wps/jsaddons/publish.xml`；WPS 安装后重启即可。未在 Windows 实机验证；
5. **在线部署**：WPS 开放平台的在线包需要开发者账号与 HTTPS 域名，本仓库暂未提交（Mac 上 WPS 在线部署端口不可用）。

### 一次性安装（DSH 侧）

安装 DSH 插件后，任选其一即可完成 WPS 版安装：

```sh
# 在 review-desk 插件目录
node scripts/install-wps-addin.mjs install
```

或在 DSH 面板点「WPS 版 → 安装 WPS 加载项」，或让 agent 调 `install_wps_addin`。

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
| **v0.3 两栏 + 右侧停靠**：任务窗格 `DockPosition=2`、宽 720px；左栏 300px / 右栏 420px（flex row），窄窗自动堆叠 | ✅ |
| **v0.3 端到端**：扩写 720 字 → 定位并高亮 → 批注 708 字 → 修订替换 231 字 → 撤销替换（拒绝 2 个修订，正文还原）→ 追加要求 326 字 | ✅ |
| **v0.3 安全**：外部 Origin 返回 403、回环 Origin 反射、`REVIEW_DESK_WPS_TOKEN` 令牌校验（bridge 测试） | ✅ |
| **v0.4 令牌默认开启**：桥自动生成令牌、面板 `/pair` 自动配对；无令牌 401、`/pair` 需配对头与回环来源 | ✅ |
| **v0.4 审核包 / 报告导出**：导出 `dsh.review-desk.package`（format/version/文档/runs 校验通过）与 Markdown 报告，经本机桥写入导出目录 | ✅ |
| **v0.4 整篇批量审阅**：整篇按段落切分（本次 2 段），逐段完成并生成汇总结果；面板进度条与取消（jsdom 验证） | ✅ |
| **v0.4 左栏宽度**：拖拽调宽并持久化（jsdom 验证；真机侧栏 360px 生效） | ✅ |
| **v1.0 安装 / 分发**：DSH 面板/CLI/agent 安装 → `publish.xml`（jspluginonline）→ WPS 从 `http://127.0.0.1:47831/addin/` 加载（`location` 校验）→ 自动配对；无需开发服务器 | ✅ |
| **v1.0 令牌固定**：令牌写入 `~/.dsh/review-desk/wps-bridge-token`；插件重载两次后 `/pair` 返回同一令牌 | ✅ |
| **v1.0 设置 / 导出上移**：模型、分段、连接收进右上角「设置」弹层；导出变成顶栏「导出」菜单（含单条结果、导入审核包） | ✅ |
| **v1.0 审核包导入**：导入自身导出的审核包，历史新增 5 条并显示「与当前文档一致」 | ✅ |
| **v1.0 批量续跑 / 分段可选**：取消后「继续」从断点续跑；600/900/1200/1600 字/段可配置（jsdom） | ✅ |
| **v1.0 单条导出**：导出「当前结果」Markdown（文件名带“选中结果”） | ✅ |
| **v1.0 端到端**：扩写 → 定位 → 批注 → 修订替换 → 撤销（拒绝 2 修订，正文还原）→ 追加要求 → 整篇分段（2 段）→ 导出包/报告 → 导入 5 条 → 导出单条 | ✅ |

验证产物：`wps-review-demo/evidence/v0.2/`、`v0.3/`、`v0.4/`、`v0.5/`、`v1.0/`（报告、导出的审核包/报告/单条结果、导出的 docx、面板截图）。

## 面板工作流（v1.0）

- **两栏任务窗格**：默认停靠右侧（`DockPosition=2`），最小宽度 560px；左栏顶栏右上角是「导出」与「设置」两个弹层入口，往下是选区 / 执行 / 历史列表，右栏为选中结果详情；窗格窄于 640px 时自动堆叠；
- 结果历史保存在本机 localStorage（最多 100 条），支持按动作筛选、关键词搜索、清空；
- 任意结果可「追加要求」继续修改（custom + revise），形成修订链；
- 替换原文以 `TrackRevisions` 写入；「撤销替换」会拒绝覆盖该范围的修订并还原原文（失败时提示用 WPS 撤销）；
- 「定位回原文」选中并滚动到文档中的原文位置，并短暂高亮；
- 设置弹层里可选择 DSH 里的模型通道、推理等级、输出上限、分段长度（600/900/1200/1600 字/段）、长文自动分段开关，以及桥地址 / 令牌 / 检测连接；触顶时提供「提高输出上限重试」；
- **导出菜单**：审核包（`dsh.review-desk.package`）、Markdown 报告、**当前选中结果**、复制审核包、**导入审核包**；文件经本机桥写入导出目录；
- **长文分段审阅**：选中 >1200 字时自动切段逐段审核并汇总（可在设置关闭）；
- **整篇批量审阅**：整篇按段落切分顺序审核，进度条 + 可取消；取消后「继续」从断点续跑；结束后生成本地汇总；
- **左栏宽度可拖拽**（双击复位、自动记忆）；
- 快捷键：⌘/Ctrl+Enter 执行、⌘/Ctrl+K 搜索、Esc 退出输入。

## 已知风险

1. **分发**：Mac 已用 `publish.xml` 侧载 + DSH 桥页面跑通；Windows 安装器路径已实现但未实机验证；WPS 在线部署需要开放平台账号与 HTTPS 域名。
2. **桥接安全**：只监听 `127.0.0.1`；令牌默认开启且持久化（`~/.dsh/review-desk/wps-bridge-token`），`/pair` 只接受回环来源；本机其他进程仍可调用（会消耗模型额度）。
3. **版本差异**：WPS 各版本对 `TrackRevisions` / `Comments` / 事件的支持可能有差异，需要兼容性测试。
4. **加载项只在启动时加载**：迭代/更新需要重启 WPS；DSH 未运行时 WPS 里不会出现审悦功能区。
5. **任务窗格内核**：CEF Chrome 104，现代框架可用但需回归。

## 后续路线

1. Windows 实机验证与打包安装器（含 WPS 安装路径探测）；
2. WPS 开放平台在线部署（HTTPS + 开发者账号）；
3. DSH 面板里的「安装 / 更新 / 卸载」状态页（当前为菜单动作）；
4. 批量任务断点持久化到 localStorage（重启 WPS 后继续）。
