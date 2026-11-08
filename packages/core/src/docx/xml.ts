import { XMLParser } from 'fast-xml-parser'

/** 简化的 XML 元素树：保留子元素顺序，方便按文档顺序遍历段落。 */
export interface El {
  name: string
  attrs: Record<string, string>
  children: El[]
  /** 仅 name === '#text' 的文本节点有值 */
  text?: string
}

const parser = new XMLParser({
  preserveOrder: true,
  ignoreAttributes: false,
  attributeNamePrefix: '',
  trimValues: false,
  parseTagValue: false,
  parseAttributeValue: false,
  processEntities: true,
})

export function parseXml(xml: string): El {
  return { name: '#root', attrs: {}, children: convert(parser.parse(xml) as RawNode[]) }
}

type RawNode = Record<string, unknown>

function convert(nodes: RawNode[]): El[] {
  const out: El[] = []
  for (const node of nodes) {
    const attrs = (node[':@'] as Record<string, string> | undefined) ?? {}
    for (const key of Object.keys(node)) {
      if (key === ':@') continue
      if (key === '#text') {
        out.push({ name: '#text', attrs: {}, children: [], text: String(node[key]) })
        continue
      }
      out.push({ name: key, attrs, children: convert((node[key] as RawNode[]) ?? []) })
    }
  }
  return out
}

export function child(el: El | undefined, name: string): El | undefined {
  return el?.children.find((c) => c.name === name)
}

export function kids(el: El | undefined, name: string): El[] {
  return el ? el.children.filter((c) => c.name === name) : []
}

/** 深度优先查找第一个同名后代元素 */
export function find(el: El | undefined, name: string): El | undefined {
  if (!el) return undefined
  for (const c of el.children) {
    if (c.name === name) return c
    const hit = find(c, name)
    if (hit) return hit
  }
  return undefined
}

export function textContent(el: El): string {
  if (el.name === '#text') return el.text ?? ''
  return el.children.map(textContent).join('')
}
