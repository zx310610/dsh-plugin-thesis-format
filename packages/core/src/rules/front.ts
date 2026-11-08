import type { Paragraph } from '../docx/document.js'
import { squash, trimStart } from '../structure.js'
import { countCjk, TWIPS_PER_LINE } from '../units.js'
import { checkKeywordCase } from './english.js'
import { checkFormat, splitRuns } from './format.js'
import { issueAt, type Issue, type Rule } from './types.js'

const ofKind = (ps: Paragraph[], kind: Paragraph['kind']) => ps.filter((p) => p.kind === kind)

/** 中文题目、署名行、指导教师行 */
export const titleBlockRule: Rule = {
  id: 'title-block',
  title: '中文题目与署名',
  check({ doc, spec }) {
    const ps = doc.paragraphs
    const out: Issue[] = []
    const title = ofKind(ps, 'title-zh')[0]
    if (!title) {
      out.push({ ruleId: this.id, severity: 'warning', message: '未识别到摘要页的中文题目（应位于署名行和“指导教师：”之前）', key: `${this.id}|missing` })
      return out
    }
    out.push(...checkFormat(title, spec.front.titleZh, { ruleId: this.id, label: '中文题目' }))
    for (const p of ofKind(ps, 'author')) {
      out.push(...checkFormat(p, spec.front.author, { ruleId: this.id, label: '署名行' }))
      if (!/^\S+ \S+专业 \d{4}级\S*班 \d+$/.test(p.text.trim())) {
        out.push(issueAt(p, {
          ruleId: this.id, severity: 'warning',
          message: '署名行格式应为“姓名 专业 年级班级 学号”，各部分之间空一格，如“张三 金融学专业 2020级1班 20203019001”',
        }))
      }
    }
    for (const p of ofKind(ps, 'advisor')) {
      out.push(...checkFormat(p, spec.front.advisor, { ruleId: this.id, label: '指导教师行' }))
      if (!/^指导教师：\S+ \S+$/.test(p.text.trim())) {
        out.push(issueAt(p, {
          ruleId: this.id, severity: 'warning',
          message: '指导教师行应写作“指导教师：姓名 职称”（全角冒号，姓名与职称之间空一格），如“指导教师：**** 副教授”',
        }))
      }
    }
    return out
  },
}

export const abstractZhRule: Rule = {
  id: 'abstract-zh',
  title: '中文摘要与关键词',
  check({ doc, spec }) {
    const ps = doc.paragraphs
    const out: Issue[] = []
    const paras = ofKind(ps, 'abstract-zh')
    if (!paras.length) {
      out.push({ ruleId: this.id, severity: 'error', message: '未找到中文摘要（应以“摘  要：”开头）', key: `${this.id}|missing` })
      return out
    }
    const first = paras[0]!
    const labelRe = new RegExp(spec.abstractZh.labelPattern)
    const m = labelRe.exec(trimStart(first.text))
    if (!m) {
      out.push(issueAt(first, { ruleId: this.id, severity: 'error', message: '摘要标签应写作“摘  要：”（两字之间空两格，全角冒号）' }))
    }
    const labelEnd = first.text.search(/[：:]/) + 1
    if (labelEnd > 0) {
      const [label, rest] = splitRuns(first, labelEnd)
      if (spec.abstractZh.labelBold) out.push(...checkFormat(first, { bold: true }, { ruleId: this.id, label: '“摘  要：”标签', runs: label }))
      out.push(...checkFormat(first, spec.abstractZh.text, { ruleId: this.id, label: '摘要正文', runs: rest }))
    }
    for (const p of paras.slice(1)) out.push(...checkFormat(p, spec.abstractZh.text, { ruleId: this.id, label: '摘要正文' }))

    const chars = paras.reduce((n, p) => n + countCjk(p.text), 0) - 2
    const { minChars, maxChars } = spec.abstractZh
    if (chars < minChars || chars > maxChars) {
      out.push(issueAt(first, {
        ruleId: this.id, severity: 'warning',
        message: `中文摘要约 ${chars} 个汉字，要求 ${minChars}–${maxChars} 字`,
        key: `${this.id}|length`,
      }))
    }

    const kw = ofKind(ps, 'keywords-zh')[0]
    if (!kw) {
      out.push({ ruleId: this.id, severity: 'error', message: '未找到中文关键词（应以“关键词：”开头）', key: `${this.id}|kw-missing` })
      return out
    }
    const kwSpec = spec.keywordsZh
    if (!new RegExp(kwSpec.labelPattern).test(trimStart(kw.text))) {
      out.push(issueAt(kw, { ruleId: this.id, severity: 'error', message: '关键词标签应写作“关键词：”（全角冒号）' }))
    }
    const kwEnd = kw.text.search(/[：:]/) + 1
    const [kwLabel, kwRest] = splitRuns(kw, kwEnd)
    out.push(...checkFormat(kw, { bold: true }, { ruleId: this.id, label: '“关键词：”标签', runs: kwLabel }))
    out.push(...checkFormat(kw, kwSpec.text, { ruleId: this.id, label: '关键词', runs: kwRest }))

    const body = kw.text.slice(kwEnd).trim()
    if (/[，,;、]/.test(body)) {
      out.push(issueAt(kw, { ruleId: this.id, severity: 'error', message: `关键词之间应使用全角分号“${kwSpec.separator}”分隔` }))
    }
    if (/[。；.;，,]$/.test(body)) {
      out.push(issueAt(kw, { ruleId: this.id, severity: 'error', message: '最后一个关键词后不加标点符号' }))
    }
    const n = body.split(kwSpec.separator).filter((s) => s.trim()).length
    if (n < kwSpec.min || n > kwSpec.max) {
      out.push(issueAt(kw, { ruleId: this.id, severity: 'error', message: `关键词应为 ${kwSpec.min}–${kwSpec.max} 个，实际 ${n} 个` }))
    }
    return out
  },
}

const SMALL_WORDS = new Set([
  'a', 'an', 'the', 'and', 'or', 'but', 'nor', 'for', 'so', 'yet', 'of', 'in', 'on', 'at', 'to', 'by', 'with',
  'from', 'into', 'onto', 'under', 'over', 'as', 'via', 'about', 'between', 'among', 'through', 'based',
])

export const abstractEnRule: Rule = {
  id: 'abstract-en',
  title: '英文题目、摘要与关键词',
  check({ doc, spec }) {
    const ps = doc.paragraphs
    const out: Issue[] = []
    const title = ofKind(ps, 'title-en')[0]
    if (title) {
      out.push(...checkFormat(title, spec.titleEn, { ruleId: this.id, label: '英文题目' }))
      const words = title.text.trim().split(/\s+/)
      const bad = words.filter((w, i) => {
        const bare = w.replace(/^[^A-Za-z]+/, '')
        if (!/^[a-z]/.test(bare)) return false
        return i === 0 || !SMALL_WORDS.has(bare.toLowerCase())
      })
      if (bad.length) {
        out.push(issueAt(title, {
          ruleId: this.id, severity: 'warning',
          message: `英文题目中除介词、连词、冠词外的实词首字母应大写：${bad.join(', ')}`,
        }))
      }
    } else {
      out.push({ ruleId: this.id, severity: 'warning', message: '未识别到英文题目（应位于中文关键词之后、Abstract 之前）', key: `${this.id}|title-missing` })
    }

    const paras = ofKind(ps, 'abstract-en')
    if (!paras.length) {
      out.push({ ruleId: this.id, severity: 'error', message: '未找到英文摘要（应以“Abstract:”开头）', key: `${this.id}|missing` })
      return out
    }
    if (!new RegExp(spec.abstractEn.labelPattern).test(trimStart(paras[0]!.text))) {
      out.push(issueAt(paras[0], { ruleId: this.id, severity: 'error', message: '英文摘要标签应写作“Abstract:”（英文冒号）' }))
    }
    for (const p of paras) out.push(...checkFormat(p, spec.abstractEn.text, { ruleId: this.id, label: '英文摘要' }))

    const kw = ofKind(ps, 'keywords-en')[0]
    if (!kw) {
      out.push({ ruleId: this.id, severity: 'error', message: '未找到英文关键词（应以“Keywords:”开头）', key: `${this.id}|kw-missing` })
      return out
    }
    const kwSpec = spec.keywordsEn
    if (!new RegExp(kwSpec.labelPattern).test(trimStart(kw.text))) {
      out.push(issueAt(kw, { ruleId: this.id, severity: 'error', message: '英文关键词标签应写作“Keywords:”' }))
    }
    out.push(...checkFormat(kw, spec.abstractEn.text, { ruleId: this.id, label: '英文关键词' }))
    const body = kw.text.replace(/^[^:：]*[:：]/, '').trim()
    if (/[；，、]/.test(body)) {
      out.push(issueAt(kw, { ruleId: this.id, severity: 'error', message: '英文关键词之间应使用英文分号“;”分隔' }))
    }
    const items = body.split(/[;；]/).map((s) => s.trim()).filter(Boolean)
    if (kwSpec.lowercase) {
      const bad = checkKeywordCase(items, kwSpec.properNouns)
      if (bad.length) {
        out.push(issueAt(kw, {
          ruleId: this.id, severity: 'warning',
          message: `英文关键词首字母不用大写（专有名词、缩写保留原样），应改为：${bad.map((b) => `${b.keyword} → ${b.suggestion}`).join('；')}`,
          key: `${this.id}|kw-case`,
        }))
      }
    }
    const zh = ofKind(ps, 'keywords-zh')[0]
    if (zh) {
      const nZh = zh.text.replace(/^[^：:]*[：:]/, '').split('；').filter((s) => s.trim()).length
      if (nZh !== items.length) {
        out.push(issueAt(kw, { ruleId: this.id, severity: 'warning', message: `英文关键词 ${items.length} 个，与中文关键词 ${nZh} 个不一致` }))
      }
    }
    return out
  },
}

export const regionTitleRule: Rule = {
  id: 'region-title',
  title: '目录/致谢/参考文献/附录 标题',
  check({ doc, spec }) {
    const out: Issue[] = []
    const names = { toc: '目录', acknowledgement: '致谢', references: '参考文献', appendix: '附录' } as const
    for (const p of doc.paragraphs.filter((q) => q.kind === 'region-title')) {
      const key = p.region as keyof typeof names
      const rs = spec.regionTitles[key]
      if (!rs) continue
      const actual = p.text.trim().replace(/　/g, '  ')
      if (actual !== rs.text) {
        out.push(issueAt(p, {
          ruleId: this.id, severity: 'warning',
          message: `“${names[key]}”标题应写作“${rs.text}”，实际为“${actual}”`,
        }))
      }
      const { spaceBeforeLines, spaceAfterLines, text: _text, ...fmt } = rs
      out.push(...checkFormat(p, fmt, { ruleId: this.id, label: `“${names[key]}”标题` }))
      // “段前段后各空一行”：设置了约 1 行的段间距，或用一个空段落隔开，都视为符合
      const ps = doc.paragraphs
      const gap = (lines: number | undefined, twips: number | undefined) =>
        lines ? lines / 100 : (twips ?? 0) / TWIPS_PER_LINE
      const sides = [
        ['段前', spaceBeforeLines, gap(p.props.beforeLines, p.props.before), ps[p.index - 1]],
        ['段后', spaceAfterLines, gap(p.props.afterLines, p.props.after), ps[p.index + 1]],
      ] as const
      for (const [side, want, got, neighbor] of sides) {
        if (want === undefined) continue
        const ok = Math.abs(got - want) <= 0.15 || (want === 1 && got >= 0.85 && got <= 1.5) || (want > 0 && neighbor?.kind === 'empty')
        if (!ok) {
          out.push(issueAt(p, {
            ruleId: this.id, severity: 'error',
            message: `“${names[key]}”标题${side}应空 ${want} 行（段间距或空一行均可），实际${side}间距 ${+got.toFixed(2)} 行且无空行`,
          }))
        }
      }
    }
    return out
  },
}

export const tocRule: Rule = {
  id: 'toc',
  title: '目录',
  check({ doc, spec }) {
    const out: Issue[] = []
    const entries = doc.paragraphs.filter((p) => p.kind === 'toc-entry')
    if (!entries.length) {
      out.push({ ruleId: this.id, severity: 'error', message: '未找到目录（应以“目  录”为标题）', key: `${this.id}|missing` })
      return out
    }
    if (spec.toc.requireAuto && !entries.some((p) => p.inTocField)) {
      out.push(issueAt(entries[0], { ruleId: this.id, severity: 'error', message: '目录应使用 Word“引用→目录”自动生成，当前为手工输入' }))
    }
    const text = entries.map((p) => squash(p.text)).join('\n')
    for (const must of spec.toc.mustInclude) {
      if (!text.includes(must)) {
        out.push(issueAt(entries[0], { ruleId: this.id, severity: 'error', message: `目录中缺少“${must}”`, key: `${this.id}|missing-${must}` }))
      }
    }
    for (const p of entries) {
      const t = trimStart(p.text)
      const level1 = /^[一二三四五六七八九十]+、/.test(t)
      if (spec.toc.maxLevel <= 2 && /^(\d+[．.]|[（(]\d+[）)])/.test(t)) {
        out.push(issueAt(p, { ruleId: this.id, severity: 'error', message: '目录只需列到二级标题', key: `${this.id}|too-deep` }))
      }
      out.push(...checkFormat(p, { ...spec.toc.entry, bold: level1 ? true : undefined }, { ruleId: this.id, label: '目录条目' }))
    }
    return out
  },
}
