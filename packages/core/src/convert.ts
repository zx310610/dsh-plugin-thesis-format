import { execFile } from 'node:child_process'
import { existsSync } from 'node:fs'
import { mkdtemp } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { basename, extname, join, resolve } from 'node:path'
import { promisify } from 'node:util'

const run = promisify(execFile)

const SOFFICE_CANDIDATES = [
  'soffice',
  'C:\\Program Files\\LibreOffice\\program\\soffice.exe',
  'C:\\Program Files (x86)\\LibreOffice\\program\\soffice.exe',
  '/Applications/LibreOffice.app/Contents/MacOS/soffice',
]

/**
 * 把旧版 .doc 转成 .docx，返回新文件路径（放在系统临时目录，不改动原文件）。
 * Windows 上优先用 Word（版式最忠实），否则用 LibreOffice。
 */
export async function ensureDocx(path: string): Promise<string> {
  const ext = extname(path).toLowerCase()
  if (ext === '.docx') return path
  if (ext !== '.doc') throw new Error(`只支持 .doc / .docx 文件：${path}`)

  const outDir = await mkdtemp(join(tmpdir(), 'thesis-check-'))
  const target = join(outDir, basename(path, extname(path)) + '.docx')
  const src = resolve(path)
  const errors: string[] = []

  if (process.platform === 'win32') {
    try {
      await convertWithWord(src, target)
      if (existsSync(target)) return target
    } catch (e) {
      errors.push(`Word: ${(e as Error).message}`)
    }
  }
  for (const bin of SOFFICE_CANDIDATES) {
    try {
      await run(bin, ['--headless', '--convert-to', 'docx', '--outdir', outDir, src], { timeout: 180_000 })
      if (existsSync(target)) return target
    } catch (e) {
      errors.push(`${bin}: ${(e as Error).message.split('\n')[0]}`)
    }
  }
  throw new Error(`无法把 .doc 转换为 .docx，请在 Word 中“另存为 .docx”后重试。\n${errors.join('\n')}`)
}

async function convertWithWord(src: string, target: string) {
  const q = (s: string) => `'${s.replace(/'/g, "''")}'`
  const script = [
    '$ErrorActionPreference = "Stop"',
    '$w = New-Object -ComObject Word.Application',
    '$w.Visible = $false; $w.DisplayAlerts = 0',
    'try {',
    `  $d = $w.Documents.Open(${q(src)}, $false, $true)`,
    `  [string]$t = ${q(target)}`,
    '  $d.SaveAs2([ref]$t, [ref]16); $d.Close($false)',
    '} finally { $w.Quit() }',
  ].join('\n')
  await run('powershell.exe', ['-NoProfile', '-NonInteractive', '-Command', script], { timeout: 180_000 })
}
