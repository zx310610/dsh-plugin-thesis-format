import type { DocModel, Paragraph, Region } from '../docx/document.js'
import type { Spec } from '../spec.js'

export type Severity = 'error' | 'warning' | 'info'

export interface Issue {
  ruleId: string
  severity: Severity
  message: string
  /** 同类问题的聚合键，报告中按此合并展示 */
  key: string
  paragraphIndex?: number
  region?: Region
  /** 段落开头文字，用于在 Word 中定位 */
  snippet?: string
  expected?: string
  actual?: string
}

export interface RuleContext {
  doc: DocModel
  spec: Spec
}

export interface Rule {
  id: string
  title: string
  check(ctx: RuleContext): Issue[]
}

export function snippet(p: Paragraph, len = 24): string {
  const t = p.text.replace(/\s+/g, ' ').trim()
  return t.length > len ? `${t.slice(0, len)}…` : t
}

export function issueAt(
  p: Paragraph | undefined,
  init: Omit<Issue, 'key' | 'paragraphIndex' | 'region' | 'snippet'> & { key?: string },
): Issue {
  return {
    key: init.key ?? `${init.ruleId}|${init.message}`,
    paragraphIndex: p?.index,
    region: p?.region,
    snippet: p ? snippet(p) : undefined,
    ...init,
  }
}
