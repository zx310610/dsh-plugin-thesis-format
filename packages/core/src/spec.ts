import { readFile } from 'node:fs/promises'
import { fileURLToPath } from 'node:url'
import { parse } from 'yaml'

/** 一段文字应有的格式。未写的字段不检查。 */
export interface FormatSpec {
  /** 中文字体，如 宋体 / 黑体 */
  font?: string
  /** 西文与数字字体，如 Times New Roman */
  westernFont?: string
  /** 字号：小四 / 五号 / 14 (pt) */
  size?: string | number
  bold?: boolean
  align?: 'left' | 'center' | 'right' | 'justify'
  firstLineIndentChars?: number
  lineSpacing?: { multiple?: number; exactPt?: number }
  spaceBeforeLines?: number
  spaceAfterLines?: number
}

export interface RegionTitleSpec extends FormatSpec {
  /** 标题的标准写法，空格数量有意义，如 “目  录” */
  text: string
}

export interface Spec {
  id: string
  name: string
  appliesTo?: number[]
  sources?: string[]
  disabledRules?: string[]
  page: {
    width: number
    height: number
    margins: { top: number; bottom: number; left: number; right: number; gutter: number }
    toleranceCm: number
  }
  front: { titleZh: FormatSpec; author: FormatSpec; advisor: FormatSpec }
  abstractZh: { labelPattern: string; labelBold: boolean; text: FormatSpec; minChars: number; maxChars: number }
  keywordsZh: { labelPattern: string; separator: string; min: number; max: number; text: FormatSpec }
  titleEn: FormatSpec
  abstractEn: { labelPattern: string; text: FormatSpec }
  keywordsEn: {
    labelPattern: string
    separator: string
    lowercase: boolean
    /** 追加的专有名词（保留大小写），内置列表见 rules/english.ts */
    properNouns?: string[]
  }
  toc: { entry: FormatSpec; maxLevel: number; mustInclude: string[]; requireAuto: boolean }
  regionTitles: Partial<Record<'toc' | 'acknowledgement' | 'references' | 'appendix', RegionTitleSpec>>
  headings: { level1: FormatSpec; level2: FormatSpec; level3: FormatSpec; level4: FormatSpec }
  body: FormatSpec & { minChars: number }
  acknowledgement: { text: FormatSpec }
  captions: { table: FormatSpec; figure: FormatSpec; sourceNote: FormatSpec; requireTextReference: boolean }
  references: {
    minCount: number
    minEnglish: number
    entry: FormatSpec
    endWith: string[]
    typeMarkers: string[]
  }
}

/** 随包附带的规范文件 */
export const BUILTIN_SPECS: Record<string, string> = {
  'bbctb-cs': fileURLToPath(new URL('../specs/bbctb-cs.yaml', import.meta.url)),
}
export const DEFAULT_SPEC = 'bbctb-cs'

/** 传入内置规范 id 或 YAML 文件路径 */
export async function loadSpec(idOrPath: string = DEFAULT_SPEC): Promise<Spec> {
  const path = BUILTIN_SPECS[idOrPath] ?? idOrPath
  const spec = parse(await readFile(path, 'utf8')) as Spec
  for (const key of ['id', 'page', 'headings', 'body', 'references'] as const) {
    if (!spec?.[key]) throw new Error(`规范文件 ${path} 缺少必填字段 ${key}`)
  }
  return spec
}
