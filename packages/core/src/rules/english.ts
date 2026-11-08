/**
 * 英文关键词大小写检查。
 *
 * 规范原文：“关键词的首字母不用大写”，示例为
 *   Keywords: traditional industry clusters; economy globalization; innovation
 * 但专有名词、缩写在英文中本来就要大写（React、Spring Boot、MySQL），不能算错。
 * 做法：先识别出专有名词/缩写，其余普通单词应为小写；违规时给出改正后的写法。
 */

/** 计算机类论文常见专有名词。大小写以这里为准；规范 YAML 中的 keywordsEn.properNouns 会追加到这里。 */
export const DEFAULT_PROPER_NOUNS = [
  // 前端
  'Vue', 'Vue.js', 'Vue3', 'Vue 3', 'React', 'Angular', 'jQuery', 'Bootstrap', 'Element UI', 'Element Plus',
  'Ant Design', 'Layui', 'ECharts', 'Axios', 'Vite', 'Webpack', 'TypeScript', 'JavaScript', 'HTML', 'HTML5', 'CSS', 'CSS3',
  'Uni-app', 'uni-app', 'Taro', 'Flutter', 'Electron', 'Pinia', 'Vuex',
  // 后端与框架
  'Java', 'Python', 'PHP', 'Go', 'Golang', 'Kotlin', 'Swift', 'C#', 'C++', '.NET', 'ASP.NET',
  'Spring', 'Spring Boot', 'SpringBoot', 'Spring Cloud', 'Spring Security', 'Spring MVC', 'MyBatis', 'MyBatis-Plus',
  'Hibernate', 'Node.js', 'Express', 'Koa', 'NestJS', 'Django', 'Flask', 'FastAPI', 'ThinkPHP', 'Laravel',
  'Tomcat', 'Nginx', 'Maven', 'Gradle', 'Swagger', 'Postman', 'Git', 'Docker', 'Kubernetes', 'Linux', 'Windows',
  'Android', 'iOS', 'HarmonyOS',
  // 数据
  'MySQL', 'Redis', 'MongoDB', 'PostgreSQL', 'Oracle', 'SQL Server', 'SQLite', 'Elasticsearch', 'Kafka', 'RabbitMQ',
  'Hadoop', 'Spark', 'Hive', 'Flink',
  // AI 与其他
  'TensorFlow', 'PyTorch', 'OpenCV', 'YOLO', 'BERT', 'ChatGPT', 'DeepSeek', 'Transformer',
  'WeChat', 'WeChat Mini Program', 'Mini Program', 'Alipay', 'Baidu', 'Alibaba', 'Tencent', 'Huawei', 'Unity', 'WebSocket', 'RESTful', 'JWT',
]

export interface KeywordCaseIssue {
  keyword: string
  suggestion: string
}

/** 单词本身看起来就是专有名词/缩写：含 2 个以上大写字母、数字或 . / # + 等符号 */
function looksSpecial(word: string): boolean {
  return (word.match(/[A-Z]/g)?.length ?? 0) >= 2 || /[\d./#+]/.test(word.replace(/[.,]$/, ''))
}

const escapeRe = (s: string) => s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')

export function checkKeywordCase(keywords: string[], extraProperNouns: string[] = []): KeywordCaseIssue[] {
  // 长的优先，保证 “Spring Boot” 先于 “Spring” 匹配
  const nouns = [...new Set([...DEFAULT_PROPER_NOUNS, ...extraProperNouns])].sort((a, b) => b.length - a.length)
  const nounRe = new RegExp(`(?<![\\w.])(${nouns.map(escapeRe).join('|')})(?![\\w])`, 'gi')
  const issues: KeywordCaseIssue[] = []

  for (const keyword of keywords) {
    // 用占位把专有名词保护起来，剩下的才是普通单词
    const protectedParts: string[] = []
    const masked = keyword.replace(nounRe, (m) => {
      protectedParts.push(m)
      return `\u0000${protectedParts.length - 1}\u0000`
    })
    const fixed = masked
      .split(/(\s+|-)/)
      .map((w) => (w.includes('\u0000') || looksSpecial(w) ? w : w.toLowerCase()))
      .join('')
      .replace(/\u0000(\d+)\u0000/g, (_, i: string) => protectedParts[Number(i)]!)
    if (fixed !== keyword) issues.push({ keyword, suggestion: fixed })
  }
  return issues
}
