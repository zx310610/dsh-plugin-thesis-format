/**
 * 用真实论文做回归测试。samples/ 含个人信息、不入库，缺失时自动跳过。
 */
import { existsSync, readdirSync } from 'node:fs'
import { join } from 'node:path'
import { describe, expect, it } from 'vitest'
import { checkThesis } from '../src/index.js'

const dir = join(import.meta.dirname, '../../../samples')
const thesis = existsSync(dir) ? readdirSync(dir).find((f) => f.endsWith('.docx') && !f.includes('模板')) : undefined

describe.skipIf(!thesis)('sample thesis', () => {
  it('finds the known issues', async () => {
    const r = await checkThesis(join(dir, thesis!))
    const msgs = r.issues.map((i) => i.message)
    expect(r.stats.regions.body).toBeGreaterThan(100)
    expect(msgs.some((m) => m.includes('纸张应为 A4'))).toBe(true)
    expect(msgs).toContain('目录应使用 Word“引用→目录”自动生成，当前为手工输入')
    expect(msgs).toContain('正文西文/数字字体应为Times New Roman，实际为宋体')
    expect(r.issues.filter((i) => i.paragraphIndex !== undefined && i.ruleId === 'crash')).toEqual([])
  }, 60_000)
})
