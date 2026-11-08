import { child, find, kids, type El } from './xml.js'

/** 字符级格式（已解析为易比较的形式） */
export interface RunProps {
  /** 西文字体 (w:ascii) */
  ascii?: string
  hAnsi?: string
  /** 中文字体 (w:eastAsia) */
  eastAsia?: string
  sizePt?: number
  bold?: boolean
  italic?: boolean
}

/** 段落级格式。缩进/间距保留 Word 原始单位，判断时再换算。 */
export interface ParaProps {
  /** left | center | right | both | distribute */
  align?: string
  /** 首行缩进，单位：百分之一字符 */
  firstLineChars?: number
  /** 首行缩进，单位：twips */
  firstLine?: number
  hangingChars?: number
  hanging?: number
  /** 段前，单位：百分之一行 */
  beforeLines?: number
  afterLines?: number
  /** 段前，单位：twips */
  before?: number
  after?: number
  /** lineRule=auto 时为 240 分之几倍；exact/atLeast 时为 twips */
  line?: number
  lineRule?: string
  pageBreakBefore?: boolean
  /** 自动编号；"0" 表示显式取消编号 */
  numId?: string
  ilvl?: number
}

export interface ThemeFonts {
  majorLatin?: string
  majorEa?: string
  minorLatin?: string
  minorEa?: string
}

interface StyleDef {
  id: string
  name: string
  type: string
  basedOn?: string
  pPr: ParaProps
  rPr: RunProps
}

export function merge<T extends object>(...layers: T[]): T {
  const out = {} as T
  for (const layer of layers) {
    for (const [k, v] of Object.entries(layer)) {
      if (v !== undefined) (out as Record<string, unknown>)[k] = v
    }
  }
  return out
}

function onOff(el: El | undefined): boolean | undefined {
  if (!el) return undefined
  const v = el.attrs['w:val']
  return v === undefined || !['false', '0', 'off', 'none'].includes(v)
}

function num(v: string | undefined): number | undefined {
  if (v === undefined || v === '') return undefined
  const n = Number(v)
  return Number.isFinite(n) ? n : undefined
}

function themeFont(token: string | undefined, theme: ThemeFonts): string | undefined {
  if (!token) return undefined
  const major = token.startsWith('major')
  const ea = token.includes('EastAsia')
  if (major) return ea ? theme.majorEa : theme.majorLatin
  return ea ? theme.minorEa : theme.minorLatin
}

export function parseRPr(rPr: El | undefined, theme: ThemeFonts): RunProps {
  if (!rPr) return {}
  const p: RunProps = {}
  const f = child(rPr, 'w:rFonts')
  if (f) {
    // 主题字体优先于显式字体名（与 Word 行为一致）
    p.ascii = themeFont(f.attrs['w:asciiTheme'], theme) ?? f.attrs['w:ascii']
    p.hAnsi = themeFont(f.attrs['w:hAnsiTheme'], theme) ?? f.attrs['w:hAnsi']
    p.eastAsia = themeFont(f.attrs['w:eastAsiaTheme'], theme) ?? f.attrs['w:eastAsia']
  }
  const sz = num(child(rPr, 'w:sz')?.attrs['w:val'])
  if (sz !== undefined) p.sizePt = sz / 2
  p.bold = onOff(child(rPr, 'w:b'))
  p.italic = onOff(child(rPr, 'w:i'))
  return merge(p)
}

export function parsePPr(pPr: El | undefined): ParaProps {
  if (!pPr) return {}
  const p: ParaProps = {}
  const jc = child(pPr, 'w:jc')?.attrs['w:val']
  if (jc) p.align = jc === 'start' ? 'left' : jc === 'end' ? 'right' : jc
  const ind = child(pPr, 'w:ind')
  if (ind) {
    p.firstLineChars = num(ind.attrs['w:firstLineChars'])
    p.firstLine = num(ind.attrs['w:firstLine'])
    p.hangingChars = num(ind.attrs['w:hangingChars'])
    p.hanging = num(ind.attrs['w:hanging'])
  }
  const sp = child(pPr, 'w:spacing')
  if (sp) {
    p.before = num(sp.attrs['w:before'])
    p.after = num(sp.attrs['w:after'])
    p.beforeLines = num(sp.attrs['w:beforeLines'])
    p.afterLines = num(sp.attrs['w:afterLines'])
    p.line = num(sp.attrs['w:line'])
    p.lineRule = sp.attrs['w:lineRule']
    // 显式写了 line 但没写 lineRule 时，默认为 auto
    if (p.line !== undefined && p.lineRule === undefined) p.lineRule = 'auto'
  }
  p.pageBreakBefore = onOff(child(pPr, 'w:pageBreakBefore'))
  const numPr = child(pPr, 'w:numPr')
  if (numPr) {
    p.numId = child(numPr, 'w:numId')?.attrs['w:val']
    p.ilvl = num(child(numPr, 'w:ilvl')?.attrs['w:val'])
  }
  return merge(p)
}

export function parseTheme(theme: El | undefined): ThemeFonts {
  const scheme = find(theme, 'a:fontScheme')
  const pick = (font: El | undefined) => ({
    latin: child(font, 'a:latin')?.attrs.typeface || undefined,
    ea:
      child(font, 'a:ea')?.attrs.typeface ||
      kids(font, 'a:font').find((f) => f.attrs.script === 'Hans')?.attrs.typeface ||
      undefined,
  })
  const major = pick(child(scheme, 'a:majorFont'))
  const minor = pick(child(scheme, 'a:minorFont'))
  return { majorLatin: major.latin, majorEa: major.ea, minorLatin: minor.latin, minorEa: minor.ea }
}

/** 样式表：按 docDefaults → basedOn 链 → 样式本身 的顺序解析继承 */
export class StyleSheet {
  readonly defaults: { pPr: ParaProps; rPr: RunProps }
  private readonly byId = new Map<string, StyleDef>()
  private defaultParaStyle?: string

  constructor(stylesXml: El | undefined, readonly theme: ThemeFonts) {
    const root = child(stylesXml, 'w:styles')
    const dd = child(root, 'w:docDefaults')
    this.defaults = {
      pPr: parsePPr(child(child(dd, 'w:pPrDefault'), 'w:pPr')),
      rPr: parseRPr(child(child(dd, 'w:rPrDefault'), 'w:rPr'), theme),
    }
    for (const s of kids(root, 'w:style')) {
      const id = s.attrs['w:styleId']
      if (!id) continue
      const def: StyleDef = {
        id,
        name: child(s, 'w:name')?.attrs['w:val'] ?? id,
        type: s.attrs['w:type'] ?? 'paragraph',
        basedOn: child(s, 'w:basedOn')?.attrs['w:val'],
        pPr: parsePPr(child(s, 'w:pPr')),
        rPr: parseRPr(child(s, 'w:rPr'), theme),
      }
      this.byId.set(id, def)
      if (def.type === 'paragraph' && ['1', 'true', 'on'].includes(s.attrs['w:default'] ?? '')) {
        this.defaultParaStyle = id
      }
    }
  }

  styleName(id: string | undefined): string | undefined {
    const sid = id ?? this.defaultParaStyle
    return sid ? this.byId.get(sid)?.name : undefined
  }

  private chain(id: string | undefined): StyleDef[] {
    const out: StyleDef[] = []
    const seen = new Set<string>()
    let cur = id ? this.byId.get(id) : undefined
    while (cur && !seen.has(cur.id)) {
      seen.add(cur.id)
      out.unshift(cur)
      cur = cur.basedOn ? this.byId.get(cur.basedOn) : undefined
    }
    return out
  }

  /** 段落样式（含文档默认值）解析后的段落格式与字符格式 */
  resolveParagraphStyle(styleId: string | undefined): { pPr: ParaProps; rPr: RunProps } {
    const chain = this.chain(styleId ?? this.defaultParaStyle)
    return {
      pPr: merge(this.defaults.pPr, ...chain.map((s) => s.pPr)),
      rPr: merge(this.defaults.rPr, ...chain.map((s) => s.rPr)),
    }
  }

  /** 字符样式链（不含文档默认值） */
  resolveCharStyle(styleId: string | undefined): RunProps {
    return merge({}, ...this.chain(styleId).map((s) => s.rPr))
  }
}
