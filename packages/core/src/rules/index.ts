import type { Region } from '../docx/document.js'
import { captionsRule, bodyTextRule, headingsRule, textHygieneRule } from './body.js'
import { abstractEnRule, abstractZhRule, regionTitleRule, titleBlockRule, tocRule } from './front.js'
import { pageSetupRule } from './page.js'
import { referencesRule } from './references.js'
import type { Issue, Rule } from './types.js'

const REQUIRED: [Region, string][] = [
  ['abstract-zh', '中文摘要'],
  ['abstract-en', '英文摘要'],
  ['toc', '目录'],
  ['body', '正文'],
  ['acknowledgement', '致谢'],
  ['references', '参考文献'],
]

/** 论文整体结构：必需部分是否齐全、顺序是否正确 */
export const structureRule: Rule = {
  id: 'structure',
  title: '论文结构',
  check({ doc }) {
    const out: Issue[] = []
    const firstIndex = new Map<Region, number>()
    for (const p of doc.paragraphs) if (!firstIndex.has(p.region)) firstIndex.set(p.region, p.index)
    for (const [region, name] of REQUIRED) {
      if (!firstIndex.has(region)) {
        out.push({ ruleId: this.id, severity: 'error', message: `未识别到“${name}”部分`, key: `${this.id}|missing-${region}` })
      }
    }
    const order = REQUIRED.slice(0, 4).map(([r]) => r).filter((r) => firstIndex.has(r))
    for (let i = 1; i < order.length; i++) {
      if (firstIndex.get(order[i]!)! < firstIndex.get(order[i - 1]!)!) {
        out.push({ ruleId: this.id, severity: 'error', message: '论文各部分顺序应为：中文摘要 → 英文摘要 → 目录 → 正文', key: `${this.id}|order` })
        break
      }
    }
    return out
  },
}

export const ALL_RULES: Rule[] = [
  structureRule,
  pageSetupRule,
  titleBlockRule,
  abstractZhRule,
  abstractEnRule,
  regionTitleRule,
  tocRule,
  headingsRule,
  bodyTextRule,
  captionsRule,
  referencesRule,
  textHygieneRule,
]

export type { Issue, Rule, RuleContext, Severity } from './types.js'
