#!/usr/bin/env node
import { writeFile } from 'node:fs/promises'
import { parseArgs } from 'node:util'
import { checkThesis, openThesis } from './check.js'
import { formatReport, summarize, type ReportFormat } from './report.js'
import { DEFAULT_SPEC } from './spec.js'

const HELP = `用法: thesis-check <论文.docx|论文.doc> [选项]

选项:
  -s, --spec <id|path>    格式规范：内置 id 或 YAML 文件路径（默认 ${DEFAULT_SPEC}）
  -f, --format <fmt>      报告格式：text | markdown | json（默认 text）
  -o, --out <file>        把报告写入文件
      --only <ids>        只运行指定规则，逗号分隔，如 headings,references
      --dump              输出每个段落的区域/类型识别结果（调试用）
  -h, --help              显示帮助

退出码: 0 = 无错误，1 = 存在格式错误，2 = 运行失败`

async function main() {
  const { values, positionals } = parseArgs({
    allowPositionals: true,
    options: {
      spec: { type: 'string', short: 's' },
      format: { type: 'string', short: 'f', default: 'text' },
      out: { type: 'string', short: 'o' },
      only: { type: 'string' },
      dump: { type: 'boolean' },
      help: { type: 'boolean', short: 'h' },
    },
  })
  const file = positionals[0]
  if (values.help || !file) {
    console.log(HELP)
    process.exit(values.help ? 0 : 2)
  }

  if (values.dump) {
    const doc = await openThesis(file)
    for (const p of doc.paragraphs) {
      const t = p.text.replace(/\s+/g, ' ').slice(0, 50)
      console.log(`${String(p.index + 1).padStart(4)}  ${p.region.padEnd(15)} ${p.kind.padEnd(14)} ${t}`)
    }
    return
  }

  const result = await checkThesis(file, { spec: values.spec, only: values.only?.split(',').map((s) => s.trim()) })
  const report = formatReport(result, values.format as ReportFormat)
  if (values.out) {
    await writeFile(values.out, report, 'utf8')
    const s = summarize(result.issues)
    console.log(`报告已写入 ${values.out}：${s.error} 个错误，${s.warning} 个警告`)
  } else {
    console.log(report)
  }
  process.exitCode = summarize(result.issues).error > 0 ? 1 : 0
}

main().catch((e) => {
  console.error(`检查失败：${(e as Error).message}`)
  process.exit(2)
})
