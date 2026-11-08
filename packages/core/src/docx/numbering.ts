import { child, kids, type El } from './xml.js'
import { parsePPr, type ParaProps } from './styles.js'

interface Level {
  start: number
  numFmt: string
  lvlText: string
  /** 编号与正文之间的分隔：tab | space | nothing */
  suff: string
  pPr: ParaProps
}

export interface ListLabel {
  /** 渲染出的编号文字，如 "[1]"、"一、"、"(2)"；项目符号为 "" */
  text: string
  isBullet: boolean
  pPr: ParaProps
}

const CN = ['', '一', '二', '三', '四', '五', '六', '七', '八', '九']

function chinese(n: number): string {
  if (n <= 0 || n >= 100) return String(n)
  if (n < 10) return CN[n]!
  const tens = Math.floor(n / 10)
  return `${tens === 1 ? '' : CN[tens]}十${CN[n % 10]}`
}

function roman(n: number): string {
  const table: [number, string][] = [[1000, 'M'], [900, 'CM'], [500, 'D'], [400, 'CD'], [100, 'C'], [90, 'XC'], [50, 'L'], [40, 'XL'], [10, 'X'], [9, 'IX'], [5, 'V'], [4, 'IV'], [1, 'I']]
  let out = ''
  for (const [v, s] of table) while (n >= v) { out += s; n -= v }
  return out
}

function letter(n: number): string {
  return String.fromCharCode(65 + ((n - 1) % 26)).repeat(Math.floor((n - 1) / 26) + 1)
}

export function formatNumber(n: number, fmt: string): string {
  switch (fmt) {
    case 'decimalZero': return n < 10 ? `0${n}` : String(n)
    case 'upperLetter': return letter(n)
    case 'lowerLetter': return letter(n).toLowerCase()
    case 'upperRoman': return roman(n)
    case 'lowerRoman': return roman(n).toLowerCase()
    case 'chineseCounting':
    case 'chineseCountingThousand':
    case 'chineseLegalSimplified':
    case 'ideographDigital':
    case 'japaneseCounting':
    case 'taiwaneseCounting':
      return chinese(n)
    case 'decimalFullWidth':
    case 'decimalFullWidth2':
      return String(n).replace(/\d/g, (d) => String.fromCharCode(0xff10 + Number(d)))
    case 'decimalEnclosedCircle':
    case 'decimalEnclosedCircleChinese':
      return n <= 20 ? String.fromCharCode(0x2460 + n - 1) : String(n)
    case 'none':
    case 'bullet':
      return ''
    default:
      return String(n)
  }
}

/**
 * Word 自动编号（numbering.xml）。按文档顺序调用 next()，返回该段的编号文字。
 * 计数器按 numId 维护；上级编号递增时下级重新开始，与 Word 默认行为一致。
 */
export class Numbering {
  private readonly abstracts = new Map<string, Map<number, Level>>()
  private readonly nums = new Map<string, { abstractId: string; startOverride: Map<number, number> }>()
  private readonly counters = new Map<string, (number | undefined)[]>()

  constructor(xml: El | undefined) {
    const root = child(xml, 'w:numbering')
    for (const a of kids(root, 'w:abstractNum')) {
      const levels = new Map<number, Level>()
      for (const lvl of kids(a, 'w:lvl')) {
        levels.set(Number(lvl.attrs['w:ilvl'] ?? 0), {
          start: Number(child(lvl, 'w:start')?.attrs['w:val'] ?? 1),
          numFmt: child(lvl, 'w:numFmt')?.attrs['w:val'] ?? 'decimal',
          lvlText: child(lvl, 'w:lvlText')?.attrs['w:val'] ?? '',
          suff: child(lvl, 'w:suff')?.attrs['w:val'] ?? 'tab',
          pPr: parsePPr(child(lvl, 'w:pPr')),
        })
      }
      this.abstracts.set(a.attrs['w:abstractNumId'] ?? '', levels)
    }
    for (const n of kids(root, 'w:num')) {
      const startOverride = new Map<number, number>()
      for (const o of kids(n, 'w:lvlOverride')) {
        const v = child(o, 'w:startOverride')?.attrs['w:val']
        if (v !== undefined) startOverride.set(Number(o.attrs['w:ilvl'] ?? 0), Number(v))
      }
      this.nums.set(n.attrs['w:numId'] ?? '', {
        abstractId: child(n, 'w:abstractNumId')?.attrs['w:val'] ?? '',
        startOverride,
      })
    }
  }

  next(numId: string, ilvl: number): ListLabel | undefined {
    const num = this.nums.get(numId)
    const levels = num && this.abstracts.get(num.abstractId)
    const level = levels?.get(ilvl)
    if (!num || !levels || !level) return undefined

    const startOf = (i: number) => num.startOverride.get(i) ?? levels.get(i)?.start ?? 1
    const c = this.counters.get(numId) ?? []
    c[ilvl] = c[ilvl] === undefined ? startOf(ilvl) : c[ilvl]! + 1
    for (let i = ilvl + 1; i < 9; i++) c[i] = undefined
    this.counters.set(numId, c)

    const isBullet = level.numFmt === 'bullet'
    let text = isBullet
      ? ''
      : level.lvlText.replace(/%([1-9])/g, (_, k: string) => {
          const i = Number(k) - 1
          const v = c[i] ?? startOf(i)
          return formatNumber(v, levels.get(i)?.numFmt ?? 'decimal')
        })
    if (text && level.suff === 'space') text += ' '
    return { text, isBullet, pPr: level.pPr }
  }
}
