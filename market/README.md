# 插件索引 / Plugin index

这个目录是**自托管的市场索引**：`index.json` 是 DSH 插件市场（Plugins → Discover）读的格式 1 索引，随仓库一起发布。

在 DSH 里添加本索引：

1. Plugins → **Discover** → **Add index source**；
2. 填入
   `https://raw.githubusercontent.com/lzyxczw-cyber/dsh-review-desk/main/market/index.json`；
3. 找到「文稿审核台 / Review desk」，**Install**。

## 更新索引（维护者）

索引里的 `source.commit` 必须指向包含该版本代码的提交。发版流程：

```sh
# 1. 改完代码、升 package.json 的 version
git add -A && git commit -m "feat: ..."

# 2. 生成索引（读取当前 HEAD 与 version）
node scripts/market-index.mjs

# 3. 提交并推送
git add market/index.json && git commit -m "chore: refresh market index" && git push
```

> 索引指向的是「上一条提交」，因为索引文件本身在随后的提交里；两个提交之间没有代码差异，安装结果一致。
