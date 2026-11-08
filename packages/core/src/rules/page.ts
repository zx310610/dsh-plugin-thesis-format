import { twipsToCm } from '../units.js'
import type { Issue, Rule } from './types.js'

const cm = (t: number | undefined) => (t === undefined ? '未设置' : `${twipsToCm(t).toFixed(2)}cm`)

export const pageSetupRule: Rule = {
  id: 'page-setup',
  title: '页面设置（纸张与页边距）',
  check({ doc, spec }) {
    const out: Issue[] = []
    const tol = spec.page.toleranceCm
    doc.sections.forEach((s, i) => {
      const where = doc.sections.length > 1 ? `第 ${i + 1} 节` : '全文'
      const near = (twips: number | undefined, wantCm: number) =>
        twips !== undefined && Math.abs(twipsToCm(twips) - wantCm) <= tol
      const w = s.orient === 'landscape' ? s.pageHeight : s.pageWidth
      const h = s.orient === 'landscape' ? s.pageWidth : s.pageHeight
      if (!near(w, spec.page.width) || !near(h, spec.page.height)) {
        out.push({
          ruleId: this.id,
          severity: 'error',
          message: `${where}纸张应为 A4 (${spec.page.width}×${spec.page.height}cm)，实际为 ${cm(w)}×${cm(h)}`,
          key: `${this.id}|size|${cm(w)}×${cm(h)}`,
        })
      }
      const labels = { top: '上', bottom: '下', left: '左', right: '右', gutter: '装订线' } as const
      for (const k of Object.keys(labels) as (keyof typeof labels)[]) {
        const want = spec.page.margins[k]
        const got = s.margins[k] ?? (k === 'gutter' ? 0 : undefined)
        if (!near(got, want)) {
          out.push({
            ruleId: this.id,
            severity: 'error',
            message: `${where}${labels[k]}边距应为 ${want}cm，实际为 ${cm(got)}`,
            key: `${this.id}|${k}|${cm(got)}`,
            expected: `${want}cm`,
            actual: cm(got),
          })
        }
      }
    })
    return out
  },
}
