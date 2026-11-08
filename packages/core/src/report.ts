import type { CheckResult } from './check.js'
import type { Issue, Severity } from './rules/index.js'

export type ReportFormat = 'text' | 'markdown' | 'json'

const SEVERITY_LABEL: Record<Severity, string> = { error: '错误', warning: '警告', info: '提示' }
const SEVERITY_ORDER: Severity[] = ['error', 'warning', 'info']

interface Group {
  key: string
  issue: Issue
  where: Issue[]
}

/** 把相同问题（key 相同）合并，保留出现位置 */
export function groupIssues(issues: Issue[]): Group[] {
  const map = new Map<string, Group>()
  for (const i of issues) {
    const g = map.get(i.key)
    if (g) g.where.push(i)
    else map.set(i.key, { key: i.key, issue: i, where: [i] })
  }
  return [...map.values()].sort(
    (a, b) => SEVERITY_ORDER.indexOf(a.issue.severity) - SEVERITY_ORDER.indexOf(b.issue.severity),
  )
}

export function summarize(issues: Issue[]): Record<Severity, number> {
  const s = { error: 0, warning: 0, info: 0 }
  for (const i of issues) s[i.severity]++
  return s
}

export function formatReport(result: CheckResult, format: ReportFormat = 'text', maxExamples = 5): string {
  if (format === 'json') return JSON.stringify({ ...result, summary: summarize(result.issues) }, null, 2)
  const md = format === 'markdown'
  const s = summarize(result.issues)
  const lines: string[] = []
  lines.push(md ? `# 论文格式检查报告` : '论文格式检查报告')
  lines.push('')
  lines.push(`${md ? '- ' : ''}文件：${result.file}`)
  lines.push(`${md ? '- ' : ''}规范：${result.spec.name} (${result.spec.id})`)
  lines.push(`${md ? '- ' : ''}结果：${s.error} 个错误，${s.warning} 个警告，${s.info} 个提示（相同问题已合并）`)

  const byRule = new Map<string, Issue[]>()
  for (const i of result.issues) byRule.set(i.ruleId, [...(byRule.get(i.ruleId) ?? []), i])

  for (const rule of result.rules) {
    const issues = byRule.get(rule.id)
    if (!issues?.length) continue
    lines.push('')
    lines.push(md ? `## ${rule.title}` : `■ ${rule.title}`)
    for (const g of groupIssues(issues)) {
      const count = g.where.length > 1 ? `（共 ${g.where.length} 处）` : ''
      lines.push(`${md ? '- ' : '  '}[${SEVERITY_LABEL[g.issue.severity]}] ${g.issue.message}${count}`)
      const located = g.where.filter((w) => w.paragraphIndex !== undefined)
      for (const w of located.slice(0, maxExamples)) {
        lines.push(`${md ? '  - ' : '      · '}第 ${w.paragraphIndex! + 1} 段：${w.snippet ?? ''}`)
      }
      if (located.length > maxExamples) lines.push(`${md ? '  - ' : '      · '}…其余 ${located.length - maxExamples} 处略`)
    }
  }
  if (!result.issues.length) lines.push('', '未发现格式问题。')
  lines.push('')
  lines.push(
    '说明：段落序号按 Word 中的段落顺序计数（含表格内段落）。页码、页眉页脚、脚注、表格边框等需要排版信息的项目暂未检查。',
  )
  return lines.join('\n')
}
