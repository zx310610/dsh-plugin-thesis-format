/**
 * 本插件用到的 dsh-tools 类型子集。
 *
 * 依据 @deepseek-ai/dsh-tools@0.0.1-rc.1 的 lib/types/index.d.ts（ToolDefinition / ToolOutputDefinition）。
 * 不直接依赖该包：它的 rc 版本依赖了尚未发布到 npm 的 @deepseek-ai/dsh-type-meta，装不上；
 * 而 ctx.tools.register() 本身就接受原始 ToolDefinition，运行时由 dsh 宿主提供 ToolRegistry。
 * dsh 处于开发者预览阶段，升级 dsh 后若接口变化，只需改这个文件和 index.ts。
 */

/** 支持的 JSON Schema 关键字：type oneOf properties required additionalProperties items enum const + 注解 */
export interface JsonSchema {
  type?: 'object' | 'array' | 'string' | 'number' | 'integer' | 'boolean' | 'null'
  properties?: Record<string, JsonSchema>
  required?: string[]
  additionalProperties?: boolean | JsonSchema
  items?: JsonSchema
  enum?: (string | number | boolean | null)[]
  description?: string
  title?: string
  default?: unknown
}

export interface TextBlock {
  type: 'text'
  text: string
}

export interface ToolRunContext {
  callId: string
  name: string
  signal: AbortSignal
}

export interface ToolDefinition {
  name: string
  description: string
  /** 根节点必须是 type: 'object' */
  parameters: JsonSchema & { type: 'object' }
  output: {
    schema: JsonSchema
    render(args: unknown, value: unknown): TextBlock[]
  }
  execute(args: unknown, exec: ToolRunContext): Promise<unknown>
  timeoutMs?: number
}

export interface ToolRegistry {
  register(tool: ToolDefinition): unknown
}
