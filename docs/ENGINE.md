# `engine` 分支：把 Word 排版引擎接进来（实验）

> 这个分支**不打算直接合并**。先在本地试，好用再谈合并与许可。

## 这个分支做了什么

1. **接入了 SuperDoc**（DOCX 引擎，浏览器内渲染真实 Word 版式：分页、表格、页眉页脚、字体与缩进）。
   - 引擎文件随插件放在 `vendor/superdoc/`（`superdoc.min.js` 1.3 MB + `superdoc.min.css` 104 KB）；
   - Host 新增只读路由 `GET /api/review-desk/engine?name=superdoc.js|superdoc.css`，只按白名单发这两个文件；
   - 文稿栏新增「文档」视图：导入 `.docx` 后自动进入，只读展示，可滚动。
2. **划线照常显示**：用浏览器 Custom Highlight API 在引擎渲染出的文字上画下划线，颜色与主分支一致（来源蓝、逻辑紫、扩写绿、关联橙、要求灰）。
3. **选择 → 分析**：在「文档」视图里选中文字，会按文本内容映射回文稿偏移，四项分析照常可用。
4. **修改替换一律先确认**：这个分支关掉了「自动替换原文」，任何改写结果或「替换原文」都会先弹出局部对话框（原文 / 替换为），确认后才改动文稿，取消则什么都不变。这就是「不涉及修改的直接看，涉及修改的走弹窗，改完回到正常状态」的策略。

## 怎么试

插件现在是本地目录安装，直接切分支即可（Host 会热重载，前端刷新一次页面）：

```sh
cd <插件目录>            # 例如 …/default-workspace/plugins/review-desk
git checkout engine
# 刷新 DSH 页面，导入一份 .docx，应自动进入「文档」视图
```

试完回主分支：

```sh
git checkout main
```

## 已知限制

- 引擎只做**只读展示**（`documentMode: 'viewing'`）。文字修改仍写回 Markdown 文本模型；引擎画布显示的是**导入时的原始文件**，所以替换后画布不会跟着变，需要「源码/编辑」视图确认文本。这是当前分支最大的缺口。
- 划线高亮只在**单个文本节点内**匹配；跨节点/跨页的长段落可能漏画。
- PDF 还没接引擎（仍用原有的原稿逐页渲染）；`pdf.js` 是下一步。
- **许可**：SuperDoc 社区版是 **AGPL-3.0**。仓库整体是 MIT，合并前需要决定：整仓改 AGPL、或把引擎改为用户自行安装的运行时依赖、或换 MIT 的 `@forevka/wordcanvas`。不解决许可之前不要把 `vendor/` 合并回 `main`。
- 引擎体积 1.5 MB，且只能由 Host 路由发出（客户端插件不能直接加载 npm 资源）。

## 下一步建议

1. 把「文档」视图的替换真正写回 docx：用 `superdoc.activeEditor.doc.query.match` + `doc.replace`（引擎已提供 Document API），再 `superdoc.export()` 导出；
2. 或者反向：以 docx 为唯一事实源，文本模型改为从 docx 提取，评审偏移全部走引擎的 Document API——工作量大，但这是「真所见即所得」的唯一终点；
3. PDF 接 `pdfjs-dist`（Apache-2.0，无许可问题）。
