import { mkdtemp, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { describe, expect, it } from 'vitest'
import { apply, inject, name } from '../src/index.js'
import type { ToolDefinition } from '../src/dsh-types.js'
import { goodThesis, makeDocx, TOC_STYLES } from '../../core/test/helpers.js'

function fakeCtx() {
  const registered: ToolDefinition[] = []
  return { ctx: { tools: { register: (t: ToolDefinition) => registered.push(t) } }, registered }
}

describe('dsh plugin', () => {
  it('exports cordis plugin shape', () => {
    expect(name).toBe('thesis-format-check')
    expect(inject).toEqual(['tools'])
  })

  it('registers tools whose schemas use only the supported JSON Schema subset', () => {
    const { ctx, registered } = fakeCtx()
    apply(ctx as never)
    expect(registered.map((t) => t.name)).toEqual(['thesis_format_check', 'thesis_read_section'])
    const allowed = new Set(['type', 'properties', 'required', 'items', 'enum', 'description'])
    for (const t of registered) {
      expect(t.parameters.type).toBe('object')
      for (const p of Object.values(t.parameters.properties!)) for (const k of Object.keys(p)) expect(allowed).toContain(k)
    }
  })

  it('runs both tools against a .docx file', async () => {
    const dir = await mkdtemp(join(tmpdir(), 'thesis-plugin-'))
    const file = join(dir, 't.docx')
    await writeFile(file, await makeDocx(goodThesis(), TOC_STYLES))
    const { ctx, registered } = fakeCtx()
    apply(ctx as never)
    const [check, section] = registered
    const exec = { callId: '1', name: 'x', signal: new AbortController().signal }
    const report = String(await check!.execute({ path: file }, exec))
    expect(report).toContain('# 论文格式检查报告')
    expect(report).toContain('0 个错误')
    const refs = String(await section!.execute({ path: file, region: 'references' }, exec))
    expect(refs).toContain('[15]')
    expect(check!.output.render({}, report)).toEqual([{ type: 'text', text: report }])
    await expect(check!.execute({}, exec)).rejects.toThrow('path')
  })
})
