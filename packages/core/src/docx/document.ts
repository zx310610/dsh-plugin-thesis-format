import JSZip from 'jszip'
import { readFile } from 'node:fs/promises'
import { child, find, kids, parseXml, textContent, type El } from './xml.js'
import { Numbering } from './numbering.js'
import { merge, parsePPr, parseRPr, parseTheme, StyleSheet, type ParaProps, type RunProps } from './styles.js'

export interface Run {
  text: string
  /** 已合并文档默认值、段落样式、字符样式和直接格式后的最终字符格式 */
  props: RunProps
}

/** 文档结构分析后为段落打上的区域标签 */
export type Region =
  | 'front' // 封面、声明、中文题目/署名
  | 'abstract-zh'
  | 'abstract-en'
  | 'toc'
  | 'body'
  | 'acknowledgement'
  | 'references'
  | 'appendix'

export type ParagraphKind =
  | 'empty'
  | 'title-zh' | 'author' | 'advisor'
  | 'abstract-zh' | 'keywords-zh'
  | 'title-en' | 'abstract-en' | 'keywords-en'
  | 'region-title' | 'toc-entry'
  | 'heading1' | 'heading2' | 'heading3' | 'heading4'
  | 'body' | 'figure' | 'caption-figure' | 'caption-table' | 'source-note' | 'table-cell'
  | 'code' | 'caption-code'
  | 'ack-text' | 'ref-entry' | 'appendix-text'
  | 'other'

export interface Paragraph {
  /** 在全文段落序列中的序号（从 0 开始，含表格内段落） */
  index: number
  /** 段落可见文字；若使用 Word 自动编号，开头包含渲染出的编号（见 listLabel） */
  text: string
  /** Word 自动编号生成的编号文字，如 "[1]"；它不属于任何 run */
  listLabel?: string
  runs: Run[]
  styleId?: string
  styleName?: string
  props: ParaProps
  /** 所在表格编号；不在表格中为 undefined */
  tableId?: number
  hasDrawing: boolean
  /** 属于自动目录（TOC 域 / 目录内容控件 / toc 样式） */
  inTocField: boolean
  region: Region
  kind: ParagraphKind
}

export interface Section {
  /** 本节最后一个段落的 index */
  lastParagraphIndex: number
  pageWidth?: number
  pageHeight?: number
  orient?: string
  margins: { top?: number; bottom?: number; left?: number; right?: number; gutter?: number; header?: number; footer?: number }
}

export interface DocModel {
  paragraphs: Paragraph[]
  sections: Section[]
  tableCount: number
  styles: StyleSheet
}

const RUN_CONTAINERS = new Set([
  'w:hyperlink', 'w:ins', 'w:smartTag', 'w:fldSimple', 'w:customXml', 'w:sdt', 'w:sdtContent', 'w:moveTo',
])
const DRAWING = new Set(['w:drawing', 'w:pict', 'w:object'])

export async function loadDocx(input: string | Uint8Array): Promise<DocModel> {
  const data = typeof input === 'string' ? await readFile(input) : input
  const zip = await JSZip.loadAsync(data)
  const read = async (path: string) => {
    const f = zip.file(path)
    return f ? parseXml(await f.async('string')) : undefined
  }
  const docXml = await read('word/document.xml')
  if (!docXml) throw new Error('不是有效的 .docx：缺少 word/document.xml')
  const theme = parseTheme(await read('word/theme/theme1.xml'))
  const styles = new StyleSheet(await read('word/styles.xml'), theme)
  const numbering = new Numbering(await read('word/numbering.xml'))
  return new Builder(styles, numbering).build(docXml)
}

class Builder {
  private paragraphs: Paragraph[] = []
  private sections: Section[] = []
  private tableCount = 0
  /** 跨段落的域栈：TOC 域通常跨越整个目录 */
  private fieldStack: { instr: string }[] = []

  constructor(
    private readonly styles: StyleSheet,
    private readonly numbering: Numbering,
  ) {}

  build(docXml: El): DocModel {
    const body = child(child(docXml, 'w:document'), 'w:body')
    if (!body) throw new Error('document.xml 中缺少 w:body')
    this.walk(body.children, undefined, false)
    return { paragraphs: this.paragraphs, sections: this.sections, tableCount: this.tableCount, styles: this.styles }
  }

  private walk(nodes: El[], tableId: number | undefined, inTocSdt: boolean) {
    for (const n of nodes) {
      switch (n.name) {
        case 'w:p':
          this.paragraph(n, tableId, inTocSdt)
          break
        case 'w:tbl': {
          const id = this.tableCount++
          for (const tr of kids(n, 'w:tr')) for (const tc of kids(tr, 'w:tc')) this.walk(tc.children, id, inTocSdt)
          break
        }
        case 'w:sdt': {
          const gallery = find(child(n, 'w:sdtPr'), 'w:docPartGallery')?.attrs['w:val'] ?? ''
          this.walk(child(n, 'w:sdtContent')?.children ?? [], tableId, inTocSdt || /table of contents/i.test(gallery))
          break
        }
        case 'w:customXml':
        case 'w:ins':
          this.walk(n.children, tableId, inTocSdt)
          break
        case 'w:sectPr':
          this.section(n)
          break
      }
    }
  }

  private paragraph(p: El, tableId: number | undefined, inTocSdt: boolean) {
    const pPrEl = child(p, 'w:pPr')
    const styleId = child(pPrEl, 'w:pStyle')?.attrs['w:val']
    const base = this.styles.resolveParagraphStyle(styleId)
    const direct = parsePPr(pPrEl)
    // 编号：直接格式中的 numPr 优先于样式中的；numId 为 0 表示取消编号
    const numId = direct.numId ?? base.pPr.numId
    const ilvl = direct.ilvl ?? base.pPr.ilvl ?? 0
    const list = numId && numId !== '0' ? this.numbering.next(numId, ilvl) : undefined
    // 优先级：样式 < 编号级别的缩进 < 段落直接格式
    const props = merge(base.pPr, list?.pPr ?? {}, direct)
    const styleName = this.styles.styleName(styleId)

    const runEls: El[] = []
    collectRuns(p, runEls)
    const runs: Run[] = []
    let hasDrawing = false
    let inToc = inTocSdt || /^toc/i.test(styleName ?? '')
    for (const r of runEls) {
      const rPrEl = child(r, 'w:rPr')
      const rStyle = child(rPrEl, 'w:rStyle')?.attrs['w:val']
      const runProps = merge(base.rPr, this.styles.resolveCharStyle(rStyle), parseRPr(rPrEl, this.styles.theme))
      let text = ''
      for (const c of r.children) {
        if (c.name === 'w:t') text += textContent(c)
        else if (c.name === 'w:tab') text += '\t'
        else if (c.name === 'w:br' || c.name === 'w:cr') text += '\n'
        else if (c.name === 'w:noBreakHyphen') text += '-'
        else if (DRAWING.has(c.name)) hasDrawing = true
        else if (c.name === 'w:fldChar') {
          const type = c.attrs['w:fldCharType']
          if (type === 'begin') this.fieldStack.push({ instr: '' })
          else if (type === 'end') this.fieldStack.pop()
        } else if (c.name === 'w:instrText') {
          const top = this.fieldStack.at(-1)
          if (top) top.instr += textContent(c)
        }
      }
      if (this.fieldStack.some((f) => /^\s*TOC\b/i.test(f.instr))) inToc = true
      if (text) runs.push({ text, props: runProps })
    }
    if (find(p, 'w:drawing') || find(p, 'w:pict') || find(p, 'w:object')) hasDrawing = true

    this.paragraphs.push({
      index: this.paragraphs.length,
      text: (list?.text ?? '') + runs.map((r) => r.text).join(''),
      listLabel: list?.text || undefined,
      runs,
      styleId,
      styleName,
      props,
      tableId,
      hasDrawing,
      inTocField: inToc,
      region: 'front',
      kind: 'other',
    })

    const sect = child(pPrEl, 'w:sectPr')
    if (sect) this.section(sect)
  }

  private section(s: El) {
    const sz = child(s, 'w:pgSz')
    const mar = child(s, 'w:pgMar')
    const n = (el: El | undefined, a: string) => {
      const v = el?.attrs[a]
      return v === undefined ? undefined : Math.abs(Number(v))
    }
    this.sections.push({
      lastParagraphIndex: this.paragraphs.length - 1,
      pageWidth: n(sz, 'w:w'),
      pageHeight: n(sz, 'w:h'),
      orient: sz?.attrs['w:orient'],
      margins: {
        top: n(mar, 'w:top'), bottom: n(mar, 'w:bottom'), left: n(mar, 'w:left'), right: n(mar, 'w:right'),
        gutter: n(mar, 'w:gutter'), header: n(mar, 'w:header'), footer: n(mar, 'w:footer'),
      },
    })
  }
}

function collectRuns(el: El, out: El[]) {
  for (const c of el.children) {
    if (c.name === 'w:r') out.push(c)
    else if (RUN_CONTAINERS.has(c.name)) collectRuns(c, out)
  }
}
