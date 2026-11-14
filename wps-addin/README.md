# 审悦 · WPS 加载项（v1.0）

在 WPS 文字里直接审稿：选中一段文字 → 四项分析 / 自由要求 → 结果以**批注**或**修订替换**写回当前文档。文档始终是 WPS 里的真 `.docx`，格式、字体、段落、撤销重做全部原生保留。

> 这是 `review-desk` 仓库 `wps` 分支的实验功能，不与 `main` 合并；DSH 插件（审悦面板 / `review_manuscript` 工具）照常可用。

## 组成

```
review-desk/
├── index.js        # wps 分支新增：本机回环桥（127.0.0.1:47831，NDJSON 流式 + /models）
└── wps-addin/      # WPS 加载项（本目录）
    ├── main.js         # 加载项入口：功能区 + 任务窗格开合
    ├── ribbon.xml      # 「文稿审核」功能区
    ├── taskpane.html   # 审稿面板
    ├── taskpane.js     # 选区 / 分析 / 批注 / 修订替换 / 历史工作流
    └── taskpane.css
```

## 功能（v1.0）

- **两栏任务窗格**：默认停靠右侧；左栏设置 / 选区 / 执行 / 历史列表，右栏为选中结果的详情与写回操作。窗格较窄（<640px）时自动堆叠为一栏；左栏宽度可拖拽调整（双击复位，自动记忆）；
- **分析**：来源查证 / 逻辑分析 / 扩写 / 关联分析 / 自由要求，流式显示；
- **模型 / 推理 / 输出设置**：任务窗格顶部「模型 / 推理 / 输出设置」，可选择 DSH 里可用的模型通道、推理等级、输出上限（默认 / 6000 / 8000 / 12000 / 16000 tokens），设置保存在本机；
- **截断处理**：输出触顶时卡片显示提示和「提高输出上限重试」按钮，一键加大上限重跑；
- **批注 / 修订替换**：插入批注；替换以 WPS 修订形式写入（可接受/拒绝/⌘Z），替换成功后可「撤销替换」（自动拒绝相关修订并还原原文）；
- **结果历史**：自动保存在本机浏览器（最多 100 条），重开面板可继续查看；按动作筛选、关键词搜索；
- **追加要求**：任意结果可继续提要求（如"再压缩到 40 字"），形成修订链，子结果标注「基于 #N」；
- **定位回原文**：选中文档中对应原文、滚动定位并短暂高亮；
- **快捷键**：⌘/Ctrl+Enter 执行（自由要求优先 / 否则重跑选中结果）、⌘/Ctrl+K 搜索、Esc 退出输入；
- **审核包 / 报告导出**：一键导出与 DSH 审悦兼容的 `.review-desk.json` 审核包，或 Markdown 审稿报告；文件由本机桥写入 DSH 导出目录（默认 `~/Documents/ReviewDeskExports`，可用 `REVIEW_DESK_WPS_EXPORT_DIR` 改）；
- **审核包导入**：`导出` 菜单 → 导入审核包，把 DSH 或 WPS 导出的 JSON 结果重新载入面板继续审；
- **长文分段审阅**：选中文字超过 1200 字（可在设置关闭）时自动按段落/句子切段、逐段审核并汇总；设置里可选 600/900/1200/1600 字/段；
- **整篇批量审阅**：按段落切分整篇文档，顺序审阅（最多 40 段 / 6 万字），带进度条；取消后可「继续」从断点续跑；结束后生成本地汇总结果；
- **连接设置**：可改本机桥地址（默认 `127.0.0.1:47831`）与访问令牌；
- **安全加固（默认开启）**：桥只接受回环来源（`127.0.0.1` / `localhost` / 打包后的 `file://`）；令牌自动生成、面板自动配对（DSH 重启后自动更新）；也可用 `REVIEW_DESK_WPS_TOKEN` 指定固定令牌，或 `REVIEW_DESK_WPS_BRIDGE_AUTH=open` 关闭校验。

## 前置条件

1. DSH 正在运行，且 profile 里安装的 `@local/review-desk` 指向本仓库的 `wps` 分支（本机已是本地目录链接安装）；
2. 桥接默认监听 `http://127.0.0.1:47831`：
   - `GET /health` → `{"ok":true,"service":"review-desk-wps-bridge","auth":"token","pairing":true,"exportDir":"…"}`
   - `GET /pair`（需 `x-review-desk-pair: 1` 且回环来源）→ 自动生成的访问令牌
   - `GET /models` → 默认模型、可选模型/推理等级、各动作输出上限
  - `GET /addin/...` → 由桥提供加载项页面（安装后 WPS 从这里加载）
   - `POST /export` → 把审核包 / 报告写进导出目录
3. 可用 `REVIEW_DESK_WPS_BRIDGE=0` 关闭桥、`REVIEW_DESK_WPS_BRIDGE_PORT` 改端口；
   令牌默认自动生成（面板自动配对）；`REVIEW_DESK_WPS_TOKEN=<随机串>` 指定固定令牌；`REVIEW_DESK_WPS_BRIDGE_AUTH=open` 关闭校验；
   导出目录：`REVIEW_DESK_WPS_EXPORT_DIR`（默认 `~/Documents/ReviewDeskExports`）。

## 开发调试

```sh
cd wps-addin
wpsjs debug          # 启动本地服务（3889）并打开 WPS；会自动写入 publish.xml
```

或只起服务、手动开一个 WPS 实例：

```sh
wpsjs debug -s
open -n -a /Applications/wpsoffice.app /path/to/test.docx
```

> 加载项只在 WPS **启动时**加载：改代码后重启 WPS 才会更新（dev server 热更新只重载页面）。
> `wpsjs debug` 会激活/启动 WPS；如果你正在 WPS 里编辑文档，先保存。

## 测试

```sh
cd tests
npm install
npm test                       # host + bridge + client + WPS 面板（jsdom）
```

面板的 jsdom 测试覆盖：两栏骨架、选区、流式结果、批注、修订替换、撤销替换（拒绝修订）、定位高亮、追加要求链、筛选与搜索、键盘快捷键、令牌自动配对与随请求发送、错误提示、历史持久化、模型设置、截断重试、审核包/报告导出、整篇批量审阅与取消、左栏拖拽持久化。

## 验证文档

文件名包含 `wps-verify` 时，加载项会自动打开 `taskpane.html?verify=1`：
面板会自选一段文字、跑一次「扩写」、写入批注、做一次修订替换并另存 docx、撤销替换、再跑一次「追加要求」，把过程 POST 到
`http://127.0.0.1:38999/report`，供本地验证脚本收集。普通文档不会触发任何自动操作。

## 已知限制

- 后端必须是本机 DSH：桥接只监听回环地址、无鉴权（见 `docs/WPS.md` 安全说明）；
- 历史结果保存在本机浏览器；换机器/清缓存会丢；
- 未做审核包导入导出、多文档管理；
- WPS Mac 只能本地侧载 `publish.xml`，在线部署（58890）不可用，正式分发需另行设计；
- 任务窗格 CEF 为 Chrome 104，React 18 可用，但更旧版本 WPS 需实测。
