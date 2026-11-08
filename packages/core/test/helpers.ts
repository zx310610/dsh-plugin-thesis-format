import JSZip from 'jszip'
import { analyzeStructure, checkDocument, loadDocx, loadSpec, type CheckResult, type DocModel } from '../src/index.js'

const NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'

export interface P {
  text: string
  font?: string
  western?: string
  /** 磅值 */
  size?: number
  bold?: boolean
  align?: string
  indentChars?: number
  /** [line, rule]，如 [360, 'auto'] 表示 1.5 倍 */
  line?: [number, string]
  style?: string
  /** Word 自动编号的 numId（见 NUMBERING） */
  numId?: number
}

const esc = (s: string) => s.replace(/&/g, '&amp;').replace(/</g, '&lt;')

export function para(p: P): string {
  const ppr = [
    p.style ? `<w:pStyle w:val="${p.style}"/>` : '',
    p.numId ? `<w:numPr><w:ilvl w:val="0"/><w:numId w:val="${p.numId}"/></w:numPr>` : '',
    p.line ? `<w:spacing w:line="${p.line[0]}" w:lineRule="${p.line[1]}"/>` : '',
    p.indentChars !== undefined ? `<w:ind w:firstLineChars="${p.indentChars * 100}"/>` : '',
    p.align ? `<w:jc w:val="${p.align}"/>` : '',
  ].join('')
  const rpr = [
    p.font || p.western ? `<w:rFonts w:eastAsia="${p.font ?? ''}" w:ascii="${p.western ?? ''}" w:hAnsi="${p.western ?? ''}"/>` : '',
    p.bold ? '<w:b/>' : '',
    p.size ? `<w:sz w:val="${p.size * 2}"/>` : '',
  ].join('')
  return `<w:p><w:pPr>${ppr}</w:pPr><w:r><w:rPr>${rpr}</w:rPr><w:t xml:space="preserve">${esc(p.text)}</w:t></w:r></w:p>`
}

/** 带加粗标签的段落，如“摘  要：”+正文 */
export function labeled(label: string, text: string, p: Omit<P, 'text'>): string {
  const [l, r] = [para({ ...p, text: label, bold: true }), para({ ...p, text })]
  // 取第二个段落的 run 拼到第一个段落里
  const run = r.slice(r.indexOf('<w:r>'), r.lastIndexOf('</w:p>'))
  return l.replace('</w:p>', run + '</w:p>')
}

export const A4_SECT =
  '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1417" w:right="1417" w:bottom="1417" w:left="1984" w:gutter="0"/></w:sectPr>'

const withNs = (xml: string, root: string) => (xml.includes('xmlns') ? xml : xml.replace(`<${root}`, `<${root} ${NS}`))

export async function makeDocx(bodyXml: string, stylesXml = `<w:styles ${NS}/>`, numberingXml?: string): Promise<Uint8Array> {
  const zip = new JSZip()
  zip.file('word/document.xml', `<?xml version="1.0" encoding="UTF-8"?><w:document ${NS}><w:body>${bodyXml}</w:body></w:document>`)
  zip.file('word/styles.xml', withNs(stylesXml, 'w:styles'))
  if (numberingXml) zip.file('word/numbering.xml', withNs(numberingXml, 'w:numbering'))
  return zip.generateAsync({ type: 'uint8array' })
}

export async function docFrom(bodyXml: string, stylesXml?: string, numberingXml?: string): Promise<DocModel> {
  return analyzeStructure(await loadDocx(await makeDocx(bodyXml, stylesXml, numberingXml)))
}

/** numId 5：参考文献 “[%1]”，自带首行缩进 2 字符；numId 6：一级标题 “%1、”（中文数字） */
export const NUMBERING = `<w:numbering>
  <w:abstractNum w:abstractNumId="1"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="[%1]"/><w:pPr><w:ind w:firstLineChars="200"/></w:pPr></w:lvl></w:abstractNum>
  <w:abstractNum w:abstractNumId="2"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="chineseCounting"/><w:lvlText w:val="%1、"/><w:suff w:val="nothing"/></w:lvl></w:abstractNum>
  <w:num w:numId="5"><w:abstractNumId w:val="1"/></w:num>
  <w:num w:numId="6"><w:abstractNumId w:val="2"/></w:num>
</w:numbering>`

export async function checkBody(bodyXml: string, only?: string[]): Promise<CheckResult> {
  return checkDocument(await docFrom(bodyXml), await loadSpec(), { only })
}

/** 一篇格式正确的最小论文，各测试在此基础上制造单个错误 */
export function goodThesis(
  overrides: Partial<Record<'h1' | 'body' | 'kw', P>> = {},
  /** 参考文献和一级标题改用 Word 自动编号（需配合 NUMBERING） */
  { autoNumber = false } = {},
): string {
  const body: P = { text: '正文内容，使用 Spring Boot 实现（见表1）[1]。', font: '宋体', western: 'Times New Roman', size: 12, indentChars: 2, line: [360, 'auto'] }
  const refs = Array.from({ length: 15 }, (_, i) => {
    const n = i + 1
    const label = autoNumber ? '' : `[${n}]`
    const t = n <= 3 ? `${label}Smith J. Title ${n}[J].Journal,2024,(1):1-${n + 1}.` : `${label}张三.题名${n}[M].北京:出版社,2023.`
    return autoNumber
      ? para({ text: t, font: '宋体', size: 10.5, numId: 5 })
      : para({ text: t, font: '宋体', size: 10.5, indentChars: 2 })
  })
  if (autoNumber && !overrides.h1) overrides = { ...overrides, h1: { text: '引言', font: '宋体', size: 12, bold: true, indentChars: 2, numId: 6 } }
  return [
    para({ text: '基于某技术的系统设计', font: '黑体', size: 18, align: 'center' }),
    para({ text: '张三 计算机科学与技术专业 2022级1班 20220000001', font: '宋体', size: 12, align: 'center' }),
    para({ text: '指导教师：李四 讲师', font: '宋体', size: 12, align: 'center' }),
    labeled('摘  要：', '这是摘要内容'.repeat(60), { font: '宋体', size: 12, line: [400, 'exact'] }),
    overrides.kw ? para(overrides.kw) : labeled('关键词：', '甲；乙；丙', { font: '宋体', size: 12 }),
    para({ text: 'Design of a System Based on Some Technology', western: 'Times New Roman', size: 14 }),
    para({ text: 'Abstract: This is the abstract.', western: 'Times New Roman', size: 10.5, line: [360, 'auto'], align: 'both', indentChars: 2.5 }),
    para({ text: 'Keywords: alpha; beta; gamma', western: 'Times New Roman', size: 10.5, line: [360, 'auto'], align: 'both', indentChars: 2.5 }),
    para({ text: '' }),
    para({ text: '目  录', font: '宋体', size: 15, bold: true, align: 'center' }),
    para({ text: '' }),
    para({ text: '一、引言……………1', font: '宋体', size: 12, bold: true, line: [360, 'auto'], style: 'TOC1' }),
    para({ text: '致谢……………2', font: '宋体', size: 12, line: [360, 'auto'], style: 'TOC1' }),
    para({ text: '参考文献……………3', font: '宋体', size: 12, line: [360, 'auto'], style: 'TOC1' }),
    para({ text: '' }),
    para(overrides.h1 ?? { text: '一、引言', font: '宋体', size: 12, bold: true, indentChars: 2 }),
    para(overrides.body ?? body),
    para({ text: '表1 测试表', font: '黑体', size: 10.5, align: 'center' }),
    '<w:tbl><w:tr><w:tc>' + para({ text: '单元格' }) + '</w:tc></w:tr></w:tbl>',
    para({ text: '' }),
    para({ text: '致  谢', font: '宋体', size: 15, bold: true, align: 'center' }),
    para({ text: '' }),
    para({ text: '感谢老师。', font: '宋体', size: 12, line: [360, 'auto'] }),
    para({ text: '参考文献', font: '黑体', size: 12, indentChars: 2 }),
    ...refs,
    A4_SECT,
  ].join('')
}

export const TOC_STYLES = `<w:styles><w:style w:type="paragraph" w:styleId="TOC1"><w:name w:val="toc 1"/></w:style></w:styles>`
