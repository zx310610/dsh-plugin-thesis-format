import { describe, expect, it } from 'vitest'
import { checkDocument, loadSpec } from '../src/index.js'
import { chineseToInt } from '../src/structure.js'
import { normalizeFont, ptLabel, sizeToPt } from '../src/units.js'
import { docFrom, goodThesis, para, NUMBERING, TOC_STYLES, A4_SECT } from './helpers.js'

async function check(body: string, only?: string[]) {
  return checkDocument(await docFrom(body, TOC_STYLES), await loadSpec(), { only })
}

describe('units', () => {
  it('converts Chinese font sizes', () => {
    expect(sizeToPt('小四')).toBe(12)
    expect(sizeToPt('小四号')).toBe(12)
    expect(sizeToPt('五号')).toBe(10.5)
    expect(sizeToPt(14)).toBe(14)
    expect(ptLabel(15)).toBe('小三(15pt)')
    expect(ptLabel(11)).toBe('11pt')
  })
  it('normalizes font aliases', () => {
    expect(normalizeFont('SimSun')).toBe('宋体')
    expect(normalizeFont('SimHei')).toBe('黑体')
    expect(normalizeFont('times new roman')).toBe('Times New Roman')
  })
  it('parses Chinese numerals', () => {
    expect(chineseToInt('三')).toBe(3)
    expect(chineseToInt('十')).toBe(10)
    expect(chineseToInt('十二')).toBe(12)
    expect(chineseToInt('二十一')).toBe(21)
  })
})

describe('style inheritance', () => {
  it('merges docDefaults → basedOn chain → direct formatting', async () => {
    const styles = `<w:styles>
      <w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:eastAsia="宋体" w:ascii="Times New Roman"/><w:sz w:val="21"/></w:rPr></w:rPrDefault></w:docDefaults>
      <w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:pPr><w:spacing w:line="360" w:lineRule="auto"/></w:pPr></w:style>
      <w:style w:type="paragraph" w:styleId="H"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:rPr><w:b/><w:sz w:val="24"/></w:rPr></w:style>
    </w:styles>`
    const doc = await docFrom(
      `<w:p><w:pPr><w:pStyle w:val="H"/></w:pPr><w:r><w:rPr><w:rFonts w:eastAsia="黑体"/></w:rPr><w:t>标题</w:t></w:r></w:p>` +
        `<w:p><w:r><w:t>正文</w:t></w:r></w:p>`,
      styles,
    )
    const [h, b] = doc.paragraphs
    expect(h!.runs[0]!.props).toMatchObject({ eastAsia: '黑体', ascii: 'Times New Roman', sizePt: 12, bold: true })
    expect(h!.props.line).toBe(360)
    expect(b!.runs[0]!.props).toMatchObject({ eastAsia: '宋体', sizePt: 10.5 })
    expect(b!.runs[0]!.props.bold).toBeUndefined()
  })
})

describe('structure', () => {
  it('recognizes all regions of a well-formed thesis', async () => {
    const doc = await docFrom(goodThesis(), TOC_STYLES)
    const kinds = (k: string) => doc.paragraphs.filter((p) => p.kind === k).length
    expect(kinds('title-zh')).toBe(1)
    expect(kinds('abstract-zh')).toBe(1)
    expect(kinds('keywords-zh')).toBe(1)
    expect(kinds('title-en')).toBe(1)
    expect(kinds('keywords-en')).toBe(1)
    expect(kinds('toc-entry')).toBe(3)
    expect(kinds('heading1')).toBe(1)
    expect(kinds('caption-table')).toBe(1)
    expect(kinds('ref-entry')).toBe(15)
  })
})

describe('rules on a well-formed thesis', () => {
  it('reports no errors', async () => {
    const r = await check(goodThesis())
    const errors = r.issues.filter((i) => i.severity === 'error')
    expect(errors.map((e) => e.message)).toEqual([])
  })
})

describe('rules catch single defects', () => {
  it('heading not bold', async () => {
    const r = await check(goodThesis({ h1: { text: '一、引言', font: '宋体', size: 12, indentChars: 2 } }), ['headings'])
    expect(r.issues.map((i) => i.message)).toContain('一级标题加粗应为加粗，实际为未加粗')
  })
  it('heading with trailing punctuation', async () => {
    const r = await check(goodThesis({ h1: { text: '一、引言。', font: '宋体', size: 12, bold: true, indentChars: 2 } }), ['headings'])
    expect(r.issues.some((i) => i.message === '一级标题末尾不加标点')).toBe(true)
  })
  it('body text in wrong size and western font', async () => {
    const body = { text: '正文 Spring Boot（见表1）[1]。', font: '宋体', western: '宋体', size: 10.5, indentChars: 2, line: [360, 'auto'] as [number, string] }
    const msgs = (await check(goodThesis({ body }), ['body-text'])).issues.map((i) => i.message)
    expect(msgs).toContain('正文字号应为小四(12pt)，实际为五号(10.5pt)')
    expect(msgs).toContain('正文西文/数字字体应为Times New Roman，实际为宋体')
  })
  it('body indented with spaces instead of first-line indent', async () => {
    const body = { text: '　　正文（见表1）[1]。', font: '宋体', western: 'Times New Roman', size: 12, line: [360, 'auto'] as [number, string] }
    const r = await check(goodThesis({ body }), ['body-text'])
    expect(r.issues.some((i) => i.message.includes('用空格代替了首行缩进'))).toBe(true)
  })
  it('firstLineChars="0" falls back to the twips value, as Word does', async () => {
    const h1 = para({ text: '一、引言', font: '宋体', size: 12, bold: true }).replace('<w:pPr>', '<w:pPr><w:ind w:firstLineChars="0" w:firstLine="482"/>')
    const thesis = goodThesis().replace(para({ text: '一、引言', font: '宋体', size: 12, bold: true, indentChars: 2 }), h1)
    const r = await check(thesis, ['headings'])
    expect(r.issues.filter((i) => i.message.includes('首行缩进'))).toEqual([])
  })
  it('keywords with wrong separator', async () => {
    const r = await check(goodThesis({ kw: { text: '关键词：甲，乙，丙', font: '宋体', size: 12 } }), ['abstract-zh'])
    expect(r.issues.some((i) => i.message.includes('全角分号'))).toBe(true)
  })
  it('wrong margins', async () => {
    const r = await check(goodThesis().replace(A4_SECT, A4_SECT.replace('w:left="1984"', 'w:left="1800"')), ['page-setup'])
    expect(r.issues.map((i) => i.message)).toContain('全文左边距应为 3.5cm，实际为 3.17cm')
  })
  it('reference numbering gap', async () => {
    const thesis = goodThesis().replace('[5]张三', '[6]张三')
    const r = await check(thesis, ['references'])
    expect(r.issues.some((i) => i.message === '参考文献序号应为 [5]，实际为 [6]')).toBe(true)
  })
  it('manual TOC is flagged', async () => {
    const doc = await docFrom(goodThesis()) // 不提供 toc 样式 → 视为手工目录
    const r = checkDocument(doc, await loadSpec(), { only: ['toc'] })
    expect(r.issues.some((i) => i.message.includes('自动生成'))).toBe(true)
  })
  it('full-width digits are flagged', async () => {
    const body = { text: '共有１２３个车位（见表1）[1]。', font: '宋体', western: 'Times New Roman', size: 12, indentChars: 2, line: [360, 'auto'] as [number, string] }
    const r = await check(goodThesis({ body }), ['text-hygiene'])
    expect(r.issues.some((i) => i.message.includes('全角'))).toBe(true)
  })
})

describe('Word auto-numbering', () => {
  it('renders list labels into text and inherits the level indent', async () => {
    const doc = await docFrom(goodThesis({}, { autoNumber: true }), TOC_STYLES, NUMBERING)
    const refs = doc.paragraphs.filter((p) => p.kind === 'ref-entry')
    expect(refs).toHaveLength(15)
    expect(refs[0]!.listLabel).toBe('[1]')
    expect(refs[14]!.text.startsWith('[15]')).toBe(true)
    expect(doc.paragraphs.find((p) => p.kind === 'heading1')!.text).toBe('一、引言')
    const r = checkDocument(doc, await loadSpec(), { only: ['references', 'headings'] })
    expect(r.issues.filter((i) => i.severity === 'error').map((i) => i.message)).toEqual([])
  })
})

describe('recognition is lenient, rules are strict', () => {
  it('“关键字：” is still recognized as keywords and flagged', async () => {
    const doc = await docFrom(goodThesis().replace('关键词：', '关键字：'), TOC_STYLES)
    expect(doc.paragraphs.filter((p) => p.kind === 'keywords-zh')).toHaveLength(1)
    expect(doc.paragraphs.some((p) => p.region === 'toc')).toBe(true)
    const r = checkDocument(doc, await loadSpec(), { only: ['abstract-zh'] })
    expect(r.issues.some((i) => i.message.startsWith('关键词标签应写作'))).toBe(true)
  })
  it('“目  录” inside the TOC content control still starts the TOC region', async () => {
    const title = para({ text: '目  录', font: '宋体', size: 15, bold: true, align: 'center' })
    const doc = await docFrom(goodThesis().replace(title, title.replace('<w:pPr>', '<w:pPr><w:pStyle w:val="TOC1"/>')), TOC_STYLES)
    expect(doc.paragraphs.find((p) => p.region === 'toc')?.kind).toBe('region-title')
    expect(doc.paragraphs.filter((p) => p.kind === 'toc-entry')).toHaveLength(3)
  })
})

describe('missing pieces', () => {
  it('table without a caption above it', async () => {
    const thesis = goodThesis().replace(para({ text: '表1 测试表', font: '黑体', size: 10.5, align: 'center' }), '')
    const r = await check(thesis, ['captions'])
    expect(r.issues.some((i) => i.message.startsWith('表格上方缺少表题'))).toBe(true)
  })
  it('reference paragraph without a number', async () => {
    const thesis = goodThesis().replace('[5]张三', '张三')
    const r = await check(thesis, ['references'])
    expect(r.issues.some((i) => i.message.startsWith('参考文献中出现没有序号的段落'))).toBe(true)
  })
})

describe('references and numbering details', () => {
  it('page range written with a dot gets a specific message', async () => {
    const thesis = goodThesis().replace('(1):1-2.', '(1):103.106.')
    const msgs = (await check(thesis, ['references'])).issues.map((i) => i.message)
    expect(msgs).toContain('期刊文献起止页码之间应使用“-”，不应使用圆点（如 103.106 应为 103-106）')
    expect(msgs).not.toContain('期刊文献 [J] 末尾应有起止页码')
  })
  it('page range with a continuation page is accepted', async () => {
    const thesis = goodThesis().replace('(1):1-2.', '(1):72-74+77.')
    const msgs = (await check(thesis, ['references'])).issues.map((i) => i.message)
    expect(msgs.filter((m) => m.includes('页码'))).toEqual([])
  })
  it('a level-4 list may restart at (1) within one section', async () => {
    const h4 = (t: string) => para({ text: t, font: '宋体', size: 12, indentChars: 2 })
    const lists = ['（1）甲', '（2）乙', '（1）丙', '（3）丁'].map(h4).join('')
    const body = para({ text: '正文内容，使用 Spring Boot 实现（见表1）[1]。', font: '宋体', western: 'Times New Roman', size: 12, indentChars: 2, line: [360, 'auto'] })
    const r = await check(goodThesis().replace(body, body + lists), ['headings'])
    expect(r.issues.map((i) => i.message)).toEqual(['四级标题编号不连续：此处为 3，按顺序应为 2'])
  })
})

describe('code listings', () => {
  it('are not checked as body text', async () => {
    const code = [
      para({ text: '代码1 模型定义', font: '黑体', size: 10.5, align: 'center' }),
      para({ text: 'class User(models.Model):', size: 10.5 }),
      para({ text: "    name = models.CharField('用户名')  # 用户名", size: 10.5 }),
      para({ text: ')', size: 10.5 }),
    ].join('')
    const thesis = goodThesis().replace(para({ text: '表1 测试表', font: '黑体', size: 10.5, align: 'center' }), code + '$&')
    const doc = await docFrom(thesis.replace('$&', para({ text: '表1 测试表', font: '黑体', size: 10.5, align: 'center' })), TOC_STYLES)
    expect(doc.paragraphs.filter((p) => p.kind === 'code').map((p) => p.text.trim())).toEqual([
      'class User(models.Model):', "name = models.CharField('用户名')  # 用户名", ')',
    ])
    expect(doc.paragraphs.filter((p) => p.kind === 'caption-code')).toHaveLength(1)
    const r = checkDocument(doc, await loadSpec(), { only: ['body-text'] })
    expect(r.issues.filter((i) => i.severity === 'error')).toEqual([])
  })
})

describe('title block', () => {
  it('recognizes an advisor line without colon and flags it', async () => {
    const r = await check(goodThesis().replace('指导教师：李四 讲师', '指导教师 李四 讲师'), ['title-block'])
    expect(r.issues.some((i) => i.message.startsWith('指导教师行应写作'))).toBe(true)
    expect(r.issues.some((i) => i.message.includes('未识别到'))).toBe(false)
  })
})

it('para helper escapes text', () => {
  expect(para({ text: 'a<b&c' })).toContain('a&lt;b&amp;c')
})
