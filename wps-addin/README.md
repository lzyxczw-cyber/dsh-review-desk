# 文稿审核台 · WPS 加载项（wps 分支）

在 WPS 文字里直接审稿：选中一段文字 → 四项分析 / 自由要求 → 结果以**批注**或**修订替换**写回当前文档。文档始终是 WPS 里的真 `.docx`，格式、字体、段落、撤销重做全部原生保留。

> 这是 `review-desk` 仓库 `wps` 分支的实验功能，不与 `main` 合并；DSH 插件（文稿审核台面板 / `review_manuscript` 工具）照常可用。

## 组成

```
review-desk/
├── index.js        # wps 分支新增：本机回环桥（127.0.0.1:47831，NDJSON 流式）
└── wps-addin/      # WPS 加载项（本目录）
    ├── main.js         # 加载项入口：功能区 + 任务窗格开合
    ├── ribbon.xml      # 「文稿审核」功能区
    ├── taskpane.html   # 审稿面板
    ├── taskpane.js     # 选区 / 分析 / 批注 / 修订替换
    └── taskpane.css
```

## 前置条件

1. DSH 正在运行，且 profile 里安装的 `@local/review-desk` 指向本仓库的 `wps` 分支（本机已是本地目录链接安装）；
2. 桥接默认监听 `http://127.0.0.1:47831`（`curl http://127.0.0.1:47831/health` 应返回
   `{"ok":true,"service":"review-desk-wps-bridge"}`）；
3. 可用 `REVIEW_DESK_WPS_BRIDGE=0` 关闭桥、`REVIEW_DESK_WPS_BRIDGE_PORT` 改端口。

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

## 使用

1. 工具栏（或自动弹出）打开「文稿审核」→ 任务窗格；
2. 在文档里选中一段文字，选择「来源查证 / 逻辑分析 / 扩写 / 关联分析」，或输入自由要求后执行；
3. 结果卡片：
   - **插入批注**：把结果写为该选区的 WPS 批注；
   - **替换原文（修订）**：扩写 / 自由要求结果显示此按钮；替换以「修订」形式写入，可在 WPS 中接受 / 拒绝，或 ⌘Z 撤销；
   - **复制结果**：复制 Markdown 文本。
4. 连接设置里可以改后端地址（默认 `127.0.0.1:47831`）。

## 已知限制（MVP）

- 后端必须是本机 DSH：桥接只监听回环地址、无鉴权（见 `docs/WPS.md` 安全说明）；
- 未做模型选择、历史结果持久化、审核包、导入/导出流程（后续版本）；
- WPS Mac 只能本地侧载 `publish.xml`，在线部署（58890）不可用，正式分发需另行设计；
- 任务窗格 CEF 为 Chrome 104，React 18 可用，但更旧版本 WPS 需实测。

## 验证文档

文件名包含 `wps-verify` 时，加载项会自动打开 `taskpane.html?verify=1`：
面板会自选一段文字、跑一次「逻辑分析」、插入批注并做一次修订替换，把过程 POST 到
`http://127.0.0.1:38999/report`，供本地验证脚本收集。普通文档不会触发任何自动操作。
