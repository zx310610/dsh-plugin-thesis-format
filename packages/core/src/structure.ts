import type { DocModel, Paragraph, ParagraphKind, Region } from './docx/document.js'
import { CJK, countCjk } from './units.js'

const CN_NUM = '一二三四五六七八九十百'

export const PATTERNS = {
  heading1: new RegExp(`^[${CN_NUM}]+、`),
  heading2: new RegExp(`^[（(][${CN_NUM}]+[）)]`),
  heading3: /^\d+[．.](?!\d)/,
  heading4: /^[（(]\d+[）)]/,
  captionFigure: /^图\s*\d+(?:[.．\-－]\d+)?(?!\d)/,
  captionTable: /^表\s*\d+(?:[.．\-－]\d+)?(?!\d)/,
  /** 代码清单标题，如“代码1 xxx”“程序清单2 xxx” */
  captionCode: /^(代码清单|程序清单|代码|程序|清单)\s*\d+(?:[.．\-－]\d+)?(?!\d)/,
  sourceNote: /^(资料|数据)来源[:：]/,
  refEntry: /^[[［]\d+[\]］]/,
  /** 手工输入的目录行：……… + 页码 */
  tocLine: /[…·.．]{3,}\s*\d+\s*$/,
}

/** 去掉所有空白（含全角空格），用于匹配“目  录”“致  谢”这类标题 */
export const squash = (s: string) => s.replace(/[\s　 ]/g, '')
/** 去掉行首空白 */
export const trimStart = (s: string) => s.replace(/^[\s　 ]+/, '')

const REGION_TITLES: Record<string, Region> = {
  目录: 'toc',
  致谢: 'acknowledgement',
  参考文献: 'references',
  附录: 'appendix',
}

/** 按中文标点结尾或长度判断是否像一个标题（而不是以编号开头的正文句子） */
function headingLike(t: string, maxLen: number) {
  return squash(t).length <= maxLen && !/[。；;]$/.test(t.trim())
}

function bodyKind(p: Paragraph): ParagraphKind {
  const t = trimStart(p.text)
  if (p.tableId !== undefined) return 'table-cell'
  if (squash(t) === '') return p.hasDrawing ? 'figure' : 'empty'
  // “一、”“（一）”开头的特征足够明显：短段落即使误加了句号也当作标题，交给标题规则报错
  if (PATTERNS.heading1.test(t) && squash(t).length <= 30) return 'heading1'
  if (PATTERNS.heading2.test(t) && squash(t).length <= 30) return 'heading2'
  if (PATTERNS.heading4.test(t)) return 'heading4'
  // “1.系统管理员的功能用例，如图1所示：”是引出图表的句子；标题中一般不含逗号
  if (PATTERNS.heading3.test(t) && headingLike(t, 40) && !/[，,]/.test(t)) return 'heading3'
  if (PATTERNS.captionFigure.test(t) && headingLike(t, 50)) return 'caption-figure'
  if (PATTERNS.captionTable.test(t) && headingLike(t, 50)) return 'caption-table'
  if (PATTERNS.sourceNote.test(t)) return 'source-note'
  if (PATTERNS.captionCode.test(t) && headingLike(t, 50)) return 'caption-code'
  // 中文论文正文中完全没有汉字的段落，几乎都是代码、SQL 或配置
  if (!CJK.test(t)) return 'code'
  return 'body'
}

/**
 * 代码块中夹带中文注释的行（如 "# 用户外键"、"name = models.CharField('名称')"）没有被“无汉字”规则覆盖。
 * 紧跟在代码行之后、以注释符开头或汉字占比很低的行，也视为代码。
 */
function extendCodeBlocks(ps: Paragraph[]) {
  let prev: Paragraph | undefined
  for (const p of ps) {
    if (p.kind === 'empty') continue
    if (p.kind === 'body' && (prev?.kind === 'code' || prev?.kind === 'caption-code')) {
      const t = trimStart(p.text)
      const ratio = countCjk(t) / Math.max(1, squash(t).length)
      if (/^(#|\/\/|\/\*|\*|--|<!--)/.test(t) || ratio < 0.3) p.kind = 'code'
    }
    prev = p
  }
}

/**
 * 为每个段落标注所属区域 (region) 与段落类型 (kind)。
 * 识别依据是论文模板中的固定标志：“摘  要：”“关键词：”“Abstract:”“目  录”、一级标题“一、”、
 * “致  谢”“参考文献”“附  录”。
 */
export function analyzeStructure(doc: DocModel): DocModel {
  let region: Region = 'front'
  let afterKeywordsZh = false

  for (const p of doc.paragraphs) {
    const t = trimStart(p.text)
    const s = squash(t)

    // ---- 区域切换 ----
    if (p.tableId === undefined) {
      if (region === 'front' && /^摘要[:：]?/.test(s) && s.length >= 2) {
        region = 'abstract-zh'
      } else if (region === 'abstract-zh' && afterKeywordsZh && s !== '') {
        region = 'abstract-en'
      } else if (s in REGION_TITLES && (!p.inTocField || s === "目录")) {
        // “目  录”标题常被包在自动目录的内容控件里，不能因 inTocField 而跳过
        region = REGION_TITLES[s]!
        p.region = region
        p.kind = 'region-title'
        continue
      } else if (
        (region === 'toc' || region === 'abstract-en' || region === 'abstract-zh') &&
        !p.inTocField &&
        !PATTERNS.tocLine.test(t) &&
        PATTERNS.heading1.test(t)
      ) {
        region = 'body'
      }
    }

    p.region = region
    // 无文字的段落：正文/附录中带图片的算作“图”，其余（含仅锚定文本框的段落）一律视为空段
    if (s === '' && (!p.hasDrawing || (region !== 'body' && region !== 'appendix'))) {
      p.kind = 'empty'
      continue
    }

    switch (region) {
      case 'front':
        p.kind = 'other'
        break
      case 'abstract-zh':
        // “关键字：”是常见误写，同样识别为关键词段，由规则报错
        if (/^关键[词字][:：]/.test(s)) {
          p.kind = 'keywords-zh'
          afterKeywordsZh = true
        } else p.kind = 'abstract-zh'
        break
      case 'abstract-en':
        if (/^key\s*words?[:：]/i.test(t)) p.kind = 'keywords-en'
        else if (/^abstract[:：]?/i.test(t)) p.kind = 'abstract-en'
        else {
          // 第一个英文段落是英文题目；Abstract 之后的续段仍算摘要正文
          const prevKinds = doc.paragraphs.slice(0, p.index).filter((q) => q.region === 'abstract-en').map((q) => q.kind)
          p.kind = prevKinds.includes('abstract-en') ? 'abstract-en' : 'title-en'
        }
        break
      case 'toc':
        p.kind = 'toc-entry'
        break
      case 'body':
        p.kind = bodyKind(p)
        break
      case 'acknowledgement':
        p.kind = p.tableId === undefined ? 'ack-text' : 'table-cell'
        break
      case 'references':
        p.kind = PATTERNS.refEntry.test(t) ? 'ref-entry' : 'other'
        break
      case 'appendix':
        p.kind = bodyKind(p) === 'empty' ? 'empty' : 'appendix-text'
        break
    }
  }

  markTitleBlock(doc.paragraphs)
  demoteListContinuations(doc.paragraphs)
  extendCodeBlocks(doc.paragraphs)
  return doc
}

/**
 * “1. xxx：很长的说明…”“2. …”之后紧跟一个短的“3. 核心功能模块实现：”时，它是正文编号列表的一项，
 * 不是三级标题。判断依据：同一二级标题下，编号 n-1 的段落被识别为正文。
 */
function demoteListContinuations(ps: Paragraph[]) {
  for (const p of ps) {
    if (p.kind !== 'heading3') continue
    const n = Number(/^(\d+)/.exec(trimStart(p.text))?.[1])
    if (!(n > 1)) continue
    for (let i = p.index - 1; i >= 0; i--) {
      const q = ps[i]!
      if (/^heading[12]$/.test(q.kind) || q.region !== 'body') break
      if (new RegExp(`^${n - 1}[．.](?!\\d)`).test(trimStart(q.text))) {
        if (q.kind === 'body') p.kind = 'body'
        break
      }
    }
  }
}

/** 中文摘要页的题目与署名：位于“指导教师：”之前的两个非空段落 */
function markTitleBlock(ps: Paragraph[]) {
  // 封面上是“指导教师及职称”，摘要页上是“指导教师：姓名 职称”（漏写冒号也要识别出来，交给规则报错）
  const advisor = ps.find((p) => p.region === 'front' && /^指导教师(?!及)/.test(squash(p.text)))
  if (!advisor) return
  advisor.kind = 'advisor'
  const before = ps.slice(0, advisor.index).filter((p) => squash(p.text) !== '' && p.tableId === undefined)
  const author = before.at(-1)
  const title = before.at(-2)
  if (author) author.kind = 'author'
  if (title) title.kind = 'title-zh'
}

/** 中文数字 → 整数（支持 1–99） */
export function chineseToInt(s: string): number | undefined {
  const map: Record<string, number> = { 一: 1, 二: 2, 三: 3, 四: 4, 五: 5, 六: 6, 七: 7, 八: 8, 九: 9 }
  if (!s) return undefined
  if (s === '十') return 10
  if (s.length === 1) return map[s]
  if (s.startsWith('十')) return 10 + (map[s[1]!] ?? NaN)
  if (s.length === 2 && s[1] === '十') return (map[s[0]!] ?? NaN) * 10
  if (s.length === 3 && s[1] === '十') return (map[s[0]!] ?? NaN) * 10 + (map[s[2]!] ?? NaN)
  return undefined
}
