import { basename } from 'node:path'
import { ensureDocx } from './convert.js'
import { loadDocx, type DocModel, type Region } from './docx/document.js'
import { ALL_RULES, type Issue, type Rule } from './rules/index.js'
import { loadSpec, type Spec } from './spec.js'
import { analyzeStructure } from './structure.js'

export interface CheckOptions {
  /** 内置规范 id 或 YAML 路径，默认 bbctb-cs */
  spec?: string | Spec
  /** 只运行这些规则 */
  only?: string[]
  rules?: Rule[]
}

export interface CheckResult {
  file: string
  spec: { id: string; name: string }
  issues: Issue[]
  /** 实际运行的规则 */
  rules: { id: string; title: string }[]
  stats: { paragraphs: number; tables: number; sections: number; regions: Partial<Record<Region, number>> }
}

export async function openThesis(path: string): Promise<DocModel> {
  return analyzeStructure(await loadDocx(await ensureDocx(path)))
}

export async function checkThesis(path: string, opts: CheckOptions = {}): Promise<CheckResult> {
  const spec = typeof opts.spec === 'object' ? opts.spec : await loadSpec(opts.spec)
  const doc = await openThesis(path)
  return checkDocument(doc, spec, opts, basename(path))
}

export function checkDocument(doc: DocModel, spec: Spec, opts: CheckOptions = {}, file = ''): CheckResult {
  const disabled = new Set(spec.disabledRules ?? [])
  const rules = (opts.rules ?? ALL_RULES).filter((r) => !disabled.has(r.id) && (!opts.only || opts.only.includes(r.id)))
  const issues: Issue[] = []
  for (const rule of rules) {
    try {
      issues.push(...rule.check({ doc, spec }))
    } catch (e) {
      issues.push({ ruleId: rule.id, severity: 'error', message: `规则执行出错：${(e as Error).message}`, key: `${rule.id}|crash` })
    }
  }
  const regions: Partial<Record<Region, number>> = {}
  for (const p of doc.paragraphs) regions[p.region] = (regions[p.region] ?? 0) + 1
  return {
    file,
    spec: { id: spec.id, name: spec.name },
    issues,
    rules: rules.map((r) => ({ id: r.id, title: r.title })),
    stats: { paragraphs: doc.paragraphs.length, tables: doc.tableCount, sections: doc.sections.length, regions },
  }
}

/** 取某个区域的纯文本，供大模型做语义层面的审阅 */
export function regionText(doc: DocModel, region: Region): string {
  return doc.paragraphs
    .filter((p) => p.region === region && p.kind !== 'empty')
    .map((p) => p.text)
    .join('\n')
}
