/** 中文字号 → 磅值 */
export const ZIHAO: Record<string, number> = {
  初号: 42, 小初: 36, 一号: 26, 小一: 24, 二号: 22, 小二: 18,
  三号: 16, 小三: 15, 四号: 14, 小四: 12, 五号: 10.5, 小五: 9,
  六号: 7.5, 小六: 6.5,
}

/** 接受 "小四" / "小四号" / "五号" / 12 / "12pt"，返回磅值 */
export function sizeToPt(size: string | number): number {
  if (typeof size === 'number') return size
  let s = size.trim()
  if (s.startsWith('小') && s.endsWith('号')) s = s.slice(0, -1)
  if (s in ZIHAO) return ZIHAO[s]!
  const n = Number.parseFloat(s)
  if (Number.isFinite(n)) return n
  throw new Error(`无法识别的字号：${size}`)
}

/** 12 → "小四(12pt)"，11 → "11pt" */
export function ptLabel(pt: number | undefined): string {
  if (pt === undefined) return '未设置'
  const name = Object.entries(ZIHAO).find(([, v]) => Math.abs(v - pt) < 0.01)?.[0]
  return name ? `${name}(${pt}pt)` : `${pt}pt`
}

export const TWIPS_PER_CM = 566.93
export const twipsToCm = (t: number) => t / TWIPS_PER_CM
export const cmToTwips = (cm: number) => cm * TWIPS_PER_CM

/** A4 纸张，单位 twips */
export const A4 = { width: 11906, height: 16838 }

/**
 * Word "段前/段后 1 行" 在只存 twips 时的换算基准。
 * 中文 Word 默认网格行高 15.6pt = 312 twips。
 */
export const TWIPS_PER_LINE = 312

const FONT_ALIASES: Record<string, string> = {
  simsun: '宋体', 宋体: '宋体', nsimsun: '新宋体', 新宋体: '新宋体',
  simhei: '黑体', 黑体: '黑体',
  kaiti: '楷体', 楷体: '楷体', kaiti_gb2312: '楷体', 楷体_gb2312: '楷体',
  fangsong: '仿宋', 仿宋: '仿宋', fangsong_gb2312: '仿宋', 仿宋_gb2312: '仿宋',
  'microsoft yahei': '微软雅黑', 微软雅黑: '微软雅黑',
  'times new roman': 'Times New Roman',
}

/** 把 SimSun / 宋体 等别名统一成规范名，便于比较 */
export function normalizeFont(name: string | undefined): string | undefined {
  if (!name) return undefined
  return FONT_ALIASES[name.trim().toLowerCase()] ?? name.trim()
}

export const CJK = /[㐀-鿿豈-﫿]/
/** 由东亚字体渲染的字符：汉字 + 全角标点/符号 */
export const EAST_ASIAN = /[　-〿㐀-鿿豈-﫿＀-￯‘-”…—]/
export const WESTERN = /[A-Za-z0-9]/

export function countCjk(s: string): number {
  let n = 0
  for (const ch of s) if (CJK.test(ch)) n++
  return n
}
