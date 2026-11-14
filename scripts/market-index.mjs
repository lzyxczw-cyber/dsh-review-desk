/**
 * Rewrite market/index.json for the current commit.
 *
 * The index references the commit that contains the plugin code, so run this
 * before committing the index and pushing:
 *
 *   node scripts/market-index.mjs && git add market/index.json && git commit && git push
 */
import { execFileSync } from 'node:child_process'
import fs from 'node:fs'
import path from 'node:path'
import { fileURLToPath } from 'node:url'

const root = path.dirname(path.dirname(fileURLToPath(import.meta.url)))
const pkg = JSON.parse(fs.readFileSync(path.join(root, 'package.json'), 'utf8'))
const commit = execFileSync('git', ['-C', root, 'rev-parse', 'HEAD'], { encoding: 'utf8' }).trim()
const repository = pkg.repository.url.replace(/^git\+/, '').replace(/\.git$/, '')

const index = {
  format: 1,
  title: { zh: '本机 DSH 插件', en: 'Local DSH plugins' },
  description: {
    zh: '自托管的 DSH 插件索引，随仓库发布。',
    en: 'A self-hosted DSH plugin index published with its repository.',
  },
  metadata: {
    categories: { review: { title: { zh: '文稿处理', en: 'Manuscripts' } } },
    labels: {
      docx: { title: { zh: 'Word', en: 'Word' } },
      pdf: { title: { zh: 'PDF', en: 'PDF' } },
      zh: { title: { zh: '中文优先', en: 'Chinese first' } },
    },
  },
  plugins: [
    {
      name: pkg.name,
      title: { zh: '审悦', en: 'ShenYue' },
      description: {
        zh: '审悦：三栏审稿工作台 + WPS 文字加载项。来源查证、逻辑分析、扩写、关联分析与自由改写；编辑区按块渲染表格与图片，点哪段改哪段；导入 docx/pdf 保留结构，审核包可携带全部过程，并可在 DSH 里一键安装 WPS 版。',
        en: 'A three-column review workspace: source checks, logic review, expansion, connection analysis, and rewrites, with block-level editing that keeps tables and images rendered, structure-preserving docx/pdf import, and a portable review package.',
      },
      icon: '../icon.svg',
      author: 'lzyxczw-cyber',
      homepage: repository,
      repository: { url: repository },
      metadata: { category: 'review', keywords: ['review', 'manuscript', 'docx', 'pdf', '审核'], labels: ['docx', 'pdf', 'zh'] },
      versions: [
        {
          version: pkg.version,
          publishTimestamp: Math.floor(Date.now() / 1000),
          source: { type: 'git', url: repository, commit },
        },
      ],
    },
  ],
}

const target = path.join(root, 'market', 'index.json')
fs.mkdirSync(path.dirname(target), { recursive: true })
fs.writeFileSync(target, `${JSON.stringify(index, null, 2)}\n`)
console.log(`market/index.json → ${pkg.name}@${pkg.version} at ${commit.slice(0, 7)}`)
