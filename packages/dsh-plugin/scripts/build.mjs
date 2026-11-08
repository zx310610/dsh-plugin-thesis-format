// 把核心库及其依赖打包进单个 lib/index.js，使插件包不依赖未发布的 @thesis-format/core，
// 可以直接用 `dsh plugin add` 或 Plugins 页面从本地 tarball / npm 安装。
import { build } from 'esbuild'
import { execFileSync } from 'node:child_process'
import { cpSync, existsSync, readdirSync, readFileSync, rmSync, writeFileSync } from 'node:fs'
import { createRequire } from 'node:module'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'

const root = join(dirname(fileURLToPath(import.meta.url)), '..')
const require = createRequire(import.meta.url)
rmSync(join(root, 'lib'), { recursive: true, force: true })

const result = await build({
  entryPoints: [join(root, 'src/index.ts')],
  outfile: join(root, 'lib/index.js'),
  bundle: true,
  platform: 'node',
  format: 'esm',
  target: 'node24',
  // 只用到 cordis 的类型，运行时不导入；保险起见仍标为外部
  external: ['@deepseek-ai/*'],
  // jszip 是 CommonJS，内部 require('stream') 等需要 ESM 下的 require
  banner: { js: "import { createRequire as __cr } from 'node:module'; const require = __cr(import.meta.url);" },
  // 第三方许可证不内嵌在代码注释里，而是汇总到 THIRD_PARTY_LICENSES.txt（见下文）
  legalComments: 'none',
  metafile: true,
  logLevel: 'warning',
})

// 类型声明
execFileSync(process.execPath, [require.resolve('typescript/bin/tsc'), '-p', join(root, 'tsconfig.json'), '--emitDeclarationOnly'], { stdio: 'inherit' })

// 核心库按 import.meta.url 定位 ../specs/*.yaml，打包后即本包的 specs/
const coreDir = dirname(require.resolve('@thesis-format/core/package.json'))
cpSync(join(coreDir, 'specs'), join(root, 'specs'), { recursive: true })
// 汇总打包进 lib/index.js 的第三方包的许可证（MIT 等要求随附版权与许可声明）
// 标准 MIT 许可正文：取本仓库 LICENSE 中版权行之后的部分
const MIT_TEXT = readFileSync(join(root, 'LICENSE'), 'utf8').split(/\r?\n/).slice(4).join('\n').trim()
const pkgDirs = new Set()
for (const input of Object.keys(result.metafile.inputs)) {
  const m = /^(.*node_modules[\\/](?:@[^\\/]+[\\/])?[^\\/]+)[\\/]/.exec(input)
  if (m) pkgDirs.add(join(root, m[1]))
}
const sections = []
for (const dir of [...pkgDirs].sort()) {
  const pkg = JSON.parse(readFileSync(join(dir, 'package.json'), 'utf8'))
  const file = readdirSync(dir).find((f) => /^(licen[cs]e|copying)(\.|$)/i.test(f))
  const license = typeof pkg.license === 'string' ? pkg.license : JSON.stringify(pkg.license)
  let text
  if (file) text = readFileSync(join(dir, file), 'utf8').trim()
  else if (license === 'MIT' && pkg.author) {
    // 个别包只在 package.json 中声明 MIT 而未附 LICENSE 文件，按声明的作者补全标准 MIT 文本
    const author = typeof pkg.author === 'string' ? pkg.author : pkg.author.name
    text = `(package.json declares "license": "MIT"; no LICENSE file is shipped. Standard MIT text:)\n\nCopyright (c) ${author}\n\n${MIT_TEXT}`
  } else throw new Error(`${pkg.name} 没有 LICENSE 文件，需手动补充许可声明`)
  const rule = '-'.repeat(60)
  sections.push(`${pkg.name}@${pkg.version} (${license})\n${rule}\n${text}\n`)
}
const header = 'lib/index.js 中打包了以下第三方软件，其版权与许可声明如下。\nThe following third-party packages are bundled into lib/index.js.\n\n'
writeFileSync(join(root, 'THIRD_PARTY_LICENSES.txt'), header + sections.join('\n\n'))
console.log(`built lib/index.js + specs/ + THIRD_PARTY_LICENSES.txt (${sections.length} packages)`)
