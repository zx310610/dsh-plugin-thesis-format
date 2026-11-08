import type { Paragraph } from '../docx/document.js'
import { chineseToInt, squash, trimStart } from '../structure.js'
import { checkFormat } from './format.js'
import { issueAt, type Issue, type Rule } from './types.js'

const LEVEL_LABEL = ['一级标题', '二级标题', '三级标题', '四级标题'] as const

function headingNumber(p: Paragraph, level: number): number | undefined {
  const t = trimStart(p.text)
  if (level === 1) return chineseToInt(/^([一二三四五六七八九十]+)、/.exec(t)?.[1] ?? '')
  if (level === 2) return chineseToInt(/^[（(]([一二三四五六七八九十]+)[）)]/.exec(t)?.[1] ?? '')
  if (level === 3) return Number(/^(\d+)/.exec(t)?.[1])
  return Number(/^[（(](\d+)[）)]/.exec(t)?.[1])
}

export const headingsRule: Rule = {
  id: 'headings',
  title: '正文各级标题',
  check({ doc, spec }) {
    const out: Issue[] = []
    const specs = [spec.headings.level1, spec.headings.level2, spec.headings.level3, spec.headings.level4]
    /** 每一级当前的序号；上级变化时下级重置 */
    const counters = [0, 0, 0, 0]
    for (const p of doc.paragraphs) {
      const m = /^heading([1-4])$/.exec(p.kind)
      if (!m) continue
      const level = Number(m[1])
      const label = LEVEL_LABEL[level - 1]!
      const t = trimStart(p.text).trimEnd()
      out.push(...checkFormat(p, specs[level - 1]!, { ruleId: this.id, label }))

      if (level <= 3 && /[。，、；：.,;:！？!?]$/.test(t)) {
        out.push(issueAt(p, { ruleId: this.id, severity: 'error', message: `${label}末尾不加标点` }))
      }
      if (/^([一二三四五六七八九十]+、|[（(][一二三四五六七八九十\d]+[）)]|\d+[．.])\s/.test(t)) {
        out.push(issueAt(p, { ruleId: this.id, severity: 'error', message: `${label}序号后面不空格` }))
      }
      if ((level === 2 || level === 4) && /^\(/.test(t)) {
        out.push(issueAt(p, { ruleId: this.id, severity: 'error', message: `${label}序号应使用全角括号“（）”` }))
      }

      const n = headingNumber(p, level)
      const expected = counters[level - 1]! + 1
      // 四级标题常作为列表使用，同一节里可以有几个各自从（1）开始的列表
      const newList = level === 4 && n === 1
      if (n !== undefined && !Number.isNaN(n) && n !== expected && !newList) {
        out.push(issueAt(p, {
          ruleId: this.id, severity: 'warning',
          message: `${label}编号不连续：此处为 ${n}，按顺序应为 ${expected}`,
        }))
      }
      if (n !== undefined && !Number.isNaN(n)) counters[level - 1] = n
      for (let i = level; i < 4; i++) counters[i] = 0
    }
    return out
  },
}

export const bodyTextRule: Rule = {
  id: 'body-text',
  title: '正文段落',
  check({ doc, spec }) {
    const out: Issue[] = []
    const { minChars, ...fmt } = spec.body
    let chars = 0
    for (const p of doc.paragraphs) {
      if (p.region !== 'body') continue
      if (p.kind === 'body' || /^heading/.test(p.kind)) chars += squash(p.text).length
      if (p.kind !== 'body') continue
      out.push(...checkFormat(p, fmt, { ruleId: this.id, label: '正文' }))
    }
    if (chars < minChars) {
      out.push({ ruleId: this.id, severity: 'warning', message: `正文约 ${chars} 字，要求不少于 ${minChars} 字`, key: `${this.id}|length` })
    }
    for (const p of doc.paragraphs.filter((q) => q.kind === 'ack-text')) {
      out.push(...checkFormat(p, spec.acknowledgement.text, { ruleId: this.id, label: '致谢正文' }))
    }
    return out
  },
}

function captionNumber(text: string): string {
  return /^[图表]\s*(\d+(?:[.．\-－]\d+)?)/.exec(trimStart(text))?.[1] ?? ''
}

export const captionsRule: Rule = {
  id: 'captions',
  title: '图表题注',
  check({ doc, spec }) {
    const out: Issue[] = []
    const ps = doc.paragraphs
    const bodyText = ps.filter((p) => p.kind === 'body').map((p) => p.text).join('\n')
    const nonEmpty = (from: number, step: 1 | -1) => {
      for (let i = from + step; i >= 0 && i < ps.length; i += step) {
        const q = ps[i]!
        if (q.kind !== 'empty') return q
      }
      return undefined
    }

    for (const [kind, name, fs] of [
      ['caption-table', '表', spec.captions.table],
      ['caption-figure', '图', spec.captions.figure],
    ] as const) {
      const caps = ps.filter((p) => p.kind === kind && p.region === 'body')
      let prev = 0
      const chapterStyle = caps.some((p) => /[.．\-－]/.test(captionNumber(p.text)))
      for (const p of caps) {
        const label = `${name}题注`
        out.push(...checkFormat(p, fs, { ruleId: this.id, label }))
        const t = trimStart(p.text)
        if (!new RegExp(`^${name}\\s*\\d+(?:[.．\\-－]\\d+)?[ \\u3000]\\S`).test(t)) {
          out.push(issueAt(p, { ruleId: this.id, severity: 'error', message: `${name}序号与${name}名之间应空一格，如“${name}1 名称”` }))
        }
        const num = captionNumber(t)
        if (!chapterStyle) {
          const n = Number(num)
          if (n !== prev + 1) {
            out.push(issueAt(p, { ruleId: this.id, severity: 'warning', message: `${name}编号不连续：此处为${name}${n}，按顺序应为${name}${prev + 1}` }))
          }
          prev = n
        }
        if (kind === 'caption-table') {
          const next = nonEmpty(p.index, 1)
          if (next?.tableId === undefined) {
            out.push(issueAt(p, { ruleId: this.id, severity: 'warning', message: '表题应位于表格正上方，但其后紧接的不是表格' }))
          }
        } else {
          const before = nonEmpty(p.index, -1)
          if (!p.hasDrawing && !before?.hasDrawing) {
            out.push(issueAt(p, { ruleId: this.id, severity: 'warning', message: '图题应位于图的正下方，但其前面不是图片' }))
          }
        }
        if (spec.captions.requireTextReference && !new RegExp(`见${name}\\s*${num.replace(/[.\-]/g, '\\$&')}(?!\\d)`).test(bodyText)) {
          out.push(issueAt(p, {
            ruleId: this.id, severity: 'warning',
            message: `正文中未找到“（见${name}${num}）”字样`,
            key: `${this.id}|no-ref-${name}`,
          }))
        }
      }
    }
    for (const p of ps.filter((q) => q.kind === 'source-note')) {
      out.push(...checkFormat(p, spec.captions.sourceNote, { ruleId: this.id, label: '资料来源' }))
    }

    // 正文中每张表格上方都应有表题。直接紧贴另一张表的（嵌套表）除外；
    // 两张表之间只隔着空行时，Word 中是两张独立的表，后一张同样需要表题。
    const seen = new Set<number>()
    for (const p of ps) {
      if (p.region !== 'body' || p.tableId === undefined || seen.has(p.tableId)) continue
      seen.add(p.tableId)
      if (ps[p.index - 1]?.tableId !== undefined) continue
      // 用单元格表格排版的代码清单，上方是“代码n 名称”，不需要表题
      if (['caption-table', 'caption-code'].includes(nonEmpty(p.index, -1)?.kind ?? '')) continue
      out.push(issueAt(p, {
        ruleId: this.id, severity: 'error',
        message: '表格上方缺少表题（应为“表n 表名”，居中，黑体五号）；下列为该表第一个单元格',
        key: `${this.id}|table-no-caption`,
      }))
    }
    return out
  },
}

export const textHygieneRule: Rule = {
  id: 'text-hygiene',
  title: '数字与标点用法',
  check({ doc }) {
    const out: Issue[] = []
    const checks: [RegExp, string, 'error' | 'warning'][] = [
      [/[０-９]/, '数字不得使用全角字符（如１２３），应使用半角（123）', 'error'],
      [/[一-鿿]\s*-{2,}|-{2,}\s*[一-鿿]/, '破折号应为中文“——”，不得用多个减号“--”代替', 'error'],
      [/(?<![\d.])\d{1,3}(?:,\d{3})+(?![\d,])/, '数字千分位应使用空格（123 456），不得使用逗号（123,456）', 'warning'],
      [/[A-Za-z]、/, '英文字母后不得使用顿号，应使用逗号或圆点', 'warning'],
    ]
    for (const p of doc.paragraphs) {
      if (!['body', 'abstract-zh', 'ack-text', 'heading1', 'heading2', 'heading3', 'heading4'].includes(p.kind)) continue
      for (const [re, msg, severity] of checks) {
        if (re.test(p.text)) out.push(issueAt(p, { ruleId: this.id, severity, message: msg, key: `${this.id}|${msg}` }))
      }
    }
    return out
  },
}

