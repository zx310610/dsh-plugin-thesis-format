import type { Paragraph, Run } from '../docx/document.js'
import type { FormatSpec } from '../spec.js'
import { EAST_ASIAN, normalizeFont, ptLabel, sizeToPt, TWIPS_PER_LINE, WESTERN } from '../units.js'
import { issueAt, type Issue, type Severity } from './types.js'

const ALIGN_LABEL: Record<string, string> = {
  left: '左对齐', center: '居中', right: '右对齐', justify: '两端对齐', both: '两端对齐', distribute: '分散对齐',
}

const hasText = (r: Run) => /\S/.test(r.text.replace(/　/g, ''))

/** 段落主要字号：按字符数加权取最多的那个，缺省 12pt */
export function dominantSize(p: Paragraph): number {
  const tally = new Map<number, number>()
  for (const r of p.runs) if (r.props.sizePt) tally.set(r.props.sizePt, (tally.get(r.props.sizePt) ?? 0) + r.text.length)
  return [...tally.entries()].sort((a, b) => b[1] - a[1])[0]?.[0] ?? 12
}

/** 首行缩进（字符数），已计入行首手工空格；返回 [总缩进, 其中空格贡献] */
export function firstLineIndent(p: Paragraph): [number, number] {
  const pr = p.props
  let chars = 0
  // Word 中字符单位为 0 时会改用 twips 值（实测：firstLineChars=0 + firstLine=482 显示为 2 字符缩进）
  if (pr.firstLineChars || pr.hangingChars) {
    chars = (pr.firstLineChars ?? 0) / 100 - (pr.hangingChars ?? 0) / 100
  } else if (pr.firstLine !== undefined || pr.hanging !== undefined) {
    chars = ((pr.firstLine ?? 0) - (pr.hanging ?? 0)) / (dominantSize(p) * 20)
  }
  const lead = /^[ 　 ]*/.exec(p.text)?.[0] ?? ''
  let spaces = 0
  for (const ch of lead) spaces += ch === '　' ? 1 : 0.5
  return [chars + spaces, spaces]
}

function spacingLines(lines: number | undefined, twips: number | undefined): number {
  if (lines) return lines / 100 // 行单位为 0 时 Word 改用 twips 值
  return (twips ?? 0) / TWIPS_PER_LINE
}

function describeLineSpacing(p: Paragraph): string {
  const { line, lineRule } = p.props
  if (line === undefined) return '单倍行距'
  if (lineRule === 'exact') return `固定值 ${line / 20} 磅`
  if (lineRule === 'atLeast') return `最小值 ${line / 20} 磅`
  return `${+(line / 240).toFixed(2)} 倍行距`
}

export interface CheckOptions {
  ruleId: string
  /** 用于消息中的段落称呼，如“一级标题” */
  label: string
  severity?: Severity
  /** 只检查这些 run（例如摘要中冒号之后的部分） */
  runs?: Run[]
}

/**
 * 按 FormatSpec 检查一个段落，每个属性至多产生一条问题，避免逐 run 刷屏。
 */
export function checkFormat(p: Paragraph, spec: FormatSpec, opt: CheckOptions): Issue[] {
  const out: Issue[] = []
  const severity = opt.severity ?? 'error'
  const runs = (opt.runs ?? p.runs).filter(hasText)
  const add = (prop: string, expected: string, actual: string) =>
    out.push(
      issueAt(p, {
        ruleId: opt.ruleId,
        severity,
        message: `${opt.label}${prop}应为${expected}，实际为${actual}`,
        key: `${opt.ruleId}|${opt.label}|${prop}|${expected}|${actual}`,
        expected,
        actual,
      }),
    )

  // ---- 字符格式 ----
  if (spec.font) {
    const want = normalizeFont(spec.font)
    const bad = distinct(runs.filter((r) => EAST_ASIAN.test(r.text)).map((r) => normalizeFont(r.props.eastAsia) ?? '未设置'))
      .filter((f) => f !== want)
    if (bad.length) add('中文字体', want!, bad.join('、'))
  }
  if (spec.westernFont) {
    const want = normalizeFont(spec.westernFont)
    const bad = distinct(runs.filter((r) => WESTERN.test(r.text)).map((r) => normalizeFont(r.props.ascii) ?? '未设置'))
      .filter((f) => f !== want)
    if (bad.length) add('西文/数字字体', want!, bad.join('、'))
  }
  if (spec.size !== undefined) {
    const want = sizeToPt(spec.size)
    const bad = distinct(runs.map((r) => r.props.sizePt ?? 10.5)).filter((s) => Math.abs(s - want) > 0.01)
    if (bad.length) add('字号', ptLabel(want), bad.map(ptLabel).join('、'))
  }
  if (spec.bold !== undefined) {
    const wrong = runs.filter((r) => (r.props.bold ?? false) !== spec.bold)
    if (wrong.length) {
      const partial = wrong.length < runs.length ? '（部分文字）' : ''
      add('加粗', spec.bold ? '加粗' : '不加粗', `${spec.bold ? '未加粗' : '加粗'}${partial}`)
    }
  }

  // ---- 段落格式 ----
  if (spec.align) {
    const actual = p.props.align === 'both' ? 'justify' : (p.props.align ?? 'left')
    const ok = actual === spec.align || (spec.align === 'left' && actual === 'justify')
    if (!ok) add('对齐方式', ALIGN_LABEL[spec.align]!, ALIGN_LABEL[actual] ?? actual)
  }
  if (spec.firstLineIndentChars !== undefined) {
    const [chars, spaces] = firstLineIndent(p)
    if (Math.abs(chars - spec.firstLineIndentChars) > 0.3) {
      add('首行缩进', `${spec.firstLineIndentChars} 字符`, `${+chars.toFixed(2)} 字符`)
    } else if (spaces > 0) {
      out.push(
        issueAt(p, {
          ruleId: opt.ruleId,
          severity: 'warning',
          message: `${opt.label}用空格代替了首行缩进，应在“段落”设置中使用首行缩进`,
          key: `${opt.ruleId}|${opt.label}|indent-by-spaces`,
        }),
      )
    }
  }
  if (spec.lineSpacing) {
    const { line, lineRule } = p.props
    if (spec.lineSpacing.multiple !== undefined) {
      const actual = line === undefined ? 1 : lineRule === 'auto' ? line / 240 : NaN
      if (!(Math.abs(actual - spec.lineSpacing.multiple) < 0.05)) {
        add('行距', `${spec.lineSpacing.multiple} 倍行距`, describeLineSpacing(p))
      }
    } else if (spec.lineSpacing.exactPt !== undefined) {
      const ok = (lineRule === 'exact' || lineRule === 'atLeast') && Math.abs((line ?? 0) / 20 - spec.lineSpacing.exactPt) < 0.5
      if (!ok) add('行距', `固定值 ${spec.lineSpacing.exactPt} 磅`, describeLineSpacing(p))
    }
  }
  if (spec.spaceBeforeLines !== undefined) {
    const v = spacingLines(p.props.beforeLines, p.props.before)
    if (Math.abs(v - spec.spaceBeforeLines) > 0.15) add('段前间距', `${spec.spaceBeforeLines} 行`, `${+v.toFixed(2)} 行`)
  }
  if (spec.spaceAfterLines !== undefined) {
    const v = spacingLines(p.props.afterLines, p.props.after)
    if (Math.abs(v - spec.spaceAfterLines) > 0.15) add('段后间距', `${spec.spaceAfterLines} 行`, `${+v.toFixed(2)} 行`)
  }
  return out
}

/**
 * 把段落的 run 在字符偏移 offset 处切成前后两组。跨界的 run 被裁剪后分别归入两侧，
 * 裁剪后只含标点/空白的部分会被丢弃，因此“摘  要”标签后紧跟的“：正文…”不会被当成标签。
 */
export function splitRuns(p: Paragraph, textOffset: number): [Run[], Run[]] {
  // offset 以 p.text 为坐标，而 p.text 开头可能有不属于任何 run 的自动编号
  const offset = textOffset - (p.listLabel?.length ?? 0)
  const head: Run[] = []
  const tail: Run[] = []
  const meaningful = (s: string) => /[\p{L}\p{N}]/u.test(s)
  let pos = 0
  for (const r of p.runs) {
    const end = pos + r.text.length
    const a = r.text.slice(0, Math.max(0, offset - pos))
    const b = r.text.slice(Math.max(0, offset - pos))
    if (meaningful(a)) head.push({ ...r, text: a })
    if (meaningful(b)) tail.push({ ...r, text: b })
    pos = end
  }
  return [head, tail]
}

function distinct<T>(xs: T[]): T[] {
  return [...new Set(xs)]
}
