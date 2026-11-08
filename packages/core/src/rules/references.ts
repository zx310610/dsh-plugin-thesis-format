import { CJK } from '../units.js'
import { trimStart } from '../structure.js'
import { checkFormat } from './format.js'
import { issueAt, type Issue, type Rule } from './types.js'

/** 正文中的引用标注：[1] [1,2] [1-3] [1，3] */
const CITE = /[[［](\d+(?:\s*[-–—~,，、]\s*\d+)*)[\]］]/g

function expandCitation(s: string): number[] {
  const out: number[] = []
  for (const part of s.split(/[,，、]/)) {
    const [a, b] = part.split(/[-–—~]/).map((x) => Number(x.trim()))
    if (a === undefined || Number.isNaN(a)) continue
    if (b !== undefined && !Number.isNaN(b) && b >= a && b - a < 50) for (let i = a; i <= b; i++) out.push(i)
    else out.push(a)
  }
  return out
}

export const referencesRule: Rule = {
  id: 'references',
  title: '参考文献',
  check({ doc, spec }) {
    const out: Issue[] = []
    const rs = spec.references
    const entries = doc.paragraphs.filter((p) => p.kind === 'ref-entry')
    if (!entries.length) {
      out.push({ ruleId: this.id, severity: 'error', message: '未找到参考文献条目（应以 [1] 形式编号）', key: `${this.id}|missing` })
      return out
    }
    if (entries.length < rs.minCount) {
      out.push(issueAt(entries[0], { ruleId: this.id, severity: 'error', message: `参考文献应不少于 ${rs.minCount} 篇，实际 ${entries.length} 篇`, key: `${this.id}|count` }))
    }
    const english = entries.filter((p) => !CJK.test(p.text)).length
    if (english < rs.minEnglish) {
      out.push(issueAt(entries[0], { ruleId: this.id, severity: 'error', message: `英文文献应不少于 ${rs.minEnglish} 篇，实际 ${english} 篇`, key: `${this.id}|english` }))
    }

    const markerRe = new RegExp(`[[［](${rs.typeMarkers.map((m) => m.replace('/', '\\/')).join('|')})[\\]］]`)
    entries.forEach((p, i) => {
      const t = trimStart(p.text).trimEnd()
      const n = Number(/^[[［](\d+)/.exec(t)?.[1])
      if (n !== i + 1) {
        out.push(issueAt(p, { ruleId: this.id, severity: 'error', message: `参考文献序号应为 [${i + 1}]，实际为 [${n}]` }))
      }
      if (/^[［]|[］]/.test(t.slice(0, 6))) {
        out.push(issueAt(p, { ruleId: this.id, severity: 'warning', message: '参考文献序号应使用半角方括号 [ ]', key: `${this.id}|fullwidth-bracket` }))
      }
      if (!markerRe.test(t)) {
        out.push(issueAt(p, { ruleId: this.id, severity: 'error', message: `缺少文献类型标识（${rs.typeMarkers.map((m) => `[${m}]`).join(' ')}）` }))
      }
      if (!rs.endWith.some((e) => t.endsWith(e))) {
        out.push(issueAt(p, { ruleId: this.id, severity: 'error', message: `每条参考文献应以“${rs.endWith.join('”或“')}”结尾`, key: `${this.id}|end` }))
      }
      if (/\[J\]/.test(t) && !/\d+\s*[-–—~]\s*\d+(\+\d+)?\s*[.。]?\s*$/.test(t)) {
        // “:103.106.” 是把起止页码之间的“-”误写成了圆点
        const dotted = /[:：]\s*(\d+)\.(\d+)((?:\+\d+)?)\s*[.。]?\s*$/.exec(t)
        out.push(issueAt(p, dotted
          ? {
              ruleId: this.id, severity: 'error',
              message: `期刊文献起止页码之间应使用“-”，不应使用圆点（如 ${dotted[1]}.${dotted[2]} 应为 ${dotted[1]}-${dotted[2]}）`,
              key: `${this.id}|journal-pages-dot`,
            }
          : { ruleId: this.id, severity: 'warning', message: '期刊文献 [J] 末尾应有起止页码', key: `${this.id}|journal-pages` }))
      }
      out.push(...checkFormat(p, rs.entry, { ruleId: this.id, label: '参考文献条目' }))
    })

    for (const p of doc.paragraphs.filter((q) => q.region === 'references' && q.kind === 'other')) {
      out.push(issueAt(p, {
        ruleId: this.id, severity: 'error',
        message: '参考文献中出现没有序号的段落：可能漏写了序号，或一条文献被回车拆成了两段',
        key: `${this.id}|unnumbered`,
      }))
    }

    // ---- 正文引用 ----
    const order: number[] = []
    for (const p of doc.paragraphs) {
      if (p.region !== 'body' || p.kind !== 'body') continue
      for (const m of p.text.matchAll(CITE)) {
        for (const n of expandCitation(m[1]!)) {
          if (n > entries.length) {
            out.push(issueAt(p, { ruleId: this.id, severity: 'error', message: `正文引用 [${n}] 在参考文献列表中不存在` }))
          }
          if (!order.includes(n)) order.push(n)
        }
      }
    }
    const uncited = entries.map((_, i) => i + 1).filter((n) => !order.includes(n))
    if (uncited.length) {
      out.push({
        ruleId: this.id, severity: 'warning',
        message: `以下参考文献未在正文中引用：${uncited.map((n) => `[${n}]`).join(' ')}`,
        key: `${this.id}|uncited`,
      })
    }
    const outOfOrder = order.findIndex((n, i) => i > 0 && n < order[i - 1]!)
    if (outOfOrder > 0) {
      out.push({
        ruleId: this.id, severity: 'warning',
        message: `参考文献应按正文中首次出现的先后排序；正文首次引用顺序为 ${order.slice(0, 12).map((n) => `[${n}]`).join(' ')}${order.length > 12 ? ' …' : ''}`,
        key: `${this.id}|order`,
      })
    }
    return out
  },
}
