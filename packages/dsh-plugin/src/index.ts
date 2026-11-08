/**
 * DeepSeek Harness 插件：毕业论文格式检查。
 *
 * 注册两个工具：
 *   thesis_format_check  —— 确定性的格式检查（字体、字号、行距、缩进、页边距、编号、参考文献…）
 *   thesis_read_section  —— 读取论文某一部分的纯文本，供模型做语义层面的审阅
 *                           （摘要是否写明背景与技术、关键词是否为短语、参考文献时效性等）
 */
import { stat } from 'node:fs/promises'
import { resolve } from 'node:path'
import type { Context } from '@deepseek-ai/cordis'
import {
  checkDocument,
  DEFAULT_SPEC,
  formatReport,
  loadSpec,
  openThesis,
  regionText,
  type DocModel,
  type Region,
} from '@thesis-format/core'
import type { ToolDefinition, ToolRegistry } from './dsh-types.js'

export const name = 'thesis-format-check'
export const inject = ['tools']

export interface Config {
  /** 默认使用的格式规范：内置 id 或 YAML 路径 */
  spec?: string
}

const REGIONS: Region[] = ['front', 'abstract-zh', 'abstract-en', 'toc', 'body', 'acknowledgement', 'references', 'appendix']

/** 同一文件未修改时复用解析结果（.doc 转换和解析都比较慢） */
const cache = new Map<string, { mtime: number; doc: Promise<DocModel> }>()

async function open(path: string): Promise<DocModel> {
  const abs = resolve(path)
  const mtime = (await stat(abs)).mtimeMs
  const hit = cache.get(abs)
  if (hit && hit.mtime === mtime) return hit.doc
  const doc = openThesis(abs)
  cache.set(abs, { mtime, doc })
  doc.catch(() => cache.delete(abs))
  return doc
}

function str(args: unknown, key: string, required: true): string
function str(args: unknown, key: string, required?: false): string | undefined
function str(args: unknown, key: string, required = false): string | undefined {
  const v = (args as Record<string, unknown> | null)?.[key]
  if (v === undefined && !required) return undefined
  if (typeof v !== 'string' || !v) throw new Error(`参数 ${key} 必须是非空字符串`)
  return v
}

const text = (_args: unknown, value: unknown) => [{ type: 'text' as const, text: String(value) }]

export function checkTool(config: Config = {}): ToolDefinition {
  return {
    name: 'thesis_format_check',
    description:
      '检查毕业论文（.docx 或 .doc）的排版格式是否符合学校规范，返回按类别分组的问题清单（含段落序号和原文片段）。' +
      '检查项：纸张页边距、中英文题目摘要关键词、目录、各级标题、正文字体字号行距缩进、图表题注、参考文献编号与格式、数字标点用法。' +
      '这是确定性检查，结果可直接采信；页码、页眉页脚、脚注、表格边框暂不检查。' +
      '如需评估内容层面的问题（摘要是否交代研究背景和技术实现、关键词是否为短语、参考文献时效性），再用 thesis_read_section 读取原文。',
    parameters: {
      type: 'object',
      properties: {
        path: { type: 'string', description: '论文文件的路径（绝对路径或相对于当前工作目录）' },
        spec: { type: 'string', description: `格式规范：内置 id 或 YAML 文件路径，默认 ${config.spec ?? DEFAULT_SPEC}` },
        format: { type: 'string', enum: ['markdown', 'json'], description: '报告格式，默认 markdown' },
        only: { type: 'string', description: '只运行指定规则，逗号分隔，如 "headings,references"' },
      },
      required: ['path'],
    },
    output: { schema: { type: 'string' }, render: text },
    timeoutMs: 300_000,
    async execute(args) {
      const path = str(args, 'path', true)
      const spec = await loadSpec(str(args, 'spec') ?? config.spec)
      const only = str(args, 'only')?.split(',').map((s) => s.trim()).filter(Boolean)
      const format = str(args, 'format') === 'json' ? 'json' : 'markdown'
      const result = checkDocument(await open(path), spec, { only }, path)
      return formatReport(result, format)
    },
  }
}

export function sectionTool(): ToolDefinition {
  return {
    name: 'thesis_read_section',
    description:
      '读取毕业论文某一部分的纯文本（不含格式信息），用于内容层面的审阅。' +
      'region 取值：front（封面与题目署名）、abstract-zh、abstract-en、toc、body（正文）、acknowledgement、references、appendix。',
    parameters: {
      type: 'object',
      properties: {
        path: { type: 'string', description: '论文文件路径' },
        region: { type: 'string', enum: REGIONS, description: '要读取的部分' },
        maxChars: { type: 'integer', description: '最多返回的字符数，默认 20000' },
      },
      required: ['path', 'region'],
    },
    output: { schema: { type: 'string' }, render: text },
    async execute(args) {
      const path = str(args, 'path', true)
      const region = str(args, 'region', true) as Region
      if (!REGIONS.includes(region)) throw new Error(`region 必须是 ${REGIONS.join(' / ')} 之一`)
      const max = Number((args as Record<string, unknown>).maxChars ?? 20_000)
      const body = regionText(await open(path), region)
      if (!body) return `（未识别到 ${region} 部分）`
      return body.length > max ? `${body.slice(0, max)}\n…（已截断，共 ${body.length} 字）` : body
    },
  }
}

export function apply(ctx: Context, config: Config = {}) {
  const tools = (ctx as unknown as { tools: ToolRegistry }).tools
  tools.register(checkTool(config))
  tools.register(sectionTool())
}
