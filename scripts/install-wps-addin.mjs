#!/usr/bin/env node
/**
 * Install / inspect / uninstall the bundled 审悦 (ShenYue) WPS add-in.
 *
 * Usage:
 *   node scripts/install-wps-addin.mjs status
 *   node scripts/install-wps-addin.mjs install
 *   node scripts/install-wps-addin.mjs uninstall
 *
 * The installer copies `wps-addin/` into the local WPS add-in directory and
 * registers it in `publish.xml`. Restart WPS afterwards; the first load asks
 * for trust inside WPS.
 *
 * Env overrides: REVIEW_DESK_WPS_JSADDONS_DIR (add-in directory).
 */
import { installWpsAddin, uninstallWpsAddin, wpsAddinStatus } from '../index.js'

const action = process.argv[2] ?? 'status'
if (!['status', 'install', 'uninstall'].includes(action)) {
  console.error(`未知操作：${action}（可用：status / install / uninstall）`)
  process.exit(2)
}
try {
  const result = action === 'install'
    ? await installWpsAddin()
    : action === 'uninstall'
      ? await uninstallWpsAddin()
      : await wpsAddinStatus()
  console.log(JSON.stringify(result, null, 2))
  if (action === 'install') {
    console.log('\n已安装。请完全退出 WPS 后重新打开，在「审悦」功能区打开审稿台；首次加载需在 WPS 里点「确定」信任。')
    if (result.trusted === false) console.log('提示：WPS 尚未信任该加载项，第一次启动时会弹出确认框。')
  } else if (action === 'uninstall') {
    console.log('\n已卸载。重启 WPS 后生效。')
  } else if (!result.installed) {
    console.log(`\n尚未安装。可运行：node scripts/install-wps-addin.mjs install`)
  }
} catch (error) {
  console.error(`执行失败：${error instanceof Error ? error.message : String(error)}`)
  process.exitCode = 1
}
