# 毕业论文格式检查插件

自动检查毕业论文（.docx / .doc）的排版格式是否符合学校规范，可作为命令行工具使用，也可作为 [DeepSeek Harness](https://github.com/deepseek-ai/deepseek-harness) 插件让模型调用。

当前内置规范：**蚌埠工商学院 计算机科学与技术专业 毕业设计（2025、2026 届通用）**。

## 目录结构

```
packages/
  core/                 格式检查核心库 + CLI（不依赖 dsh，可单独使用）
    specs/bbctb-cs.yaml   格式规范（从学校模板整理，带原文出处注释）
    src/docx/             .docx 解析：XML、样式继承（docDefaults → basedOn → 直接格式）
    src/structure.ts      论文结构识别：摘要/目录/正文/致谢/参考文献、各级标题、图表题注
    src/rules/            检查规则，每个文件一组
    src/convert.ts        .doc → .docx（Windows 用 Word，其他平台用 LibreOffice）
    test/                 单元测试（内存中生成 .docx）+ 真实论文回归测试
  dsh-plugin/           DeepSeek Harness 插件外壳（注册两个工具）
docs/规范歧义.md         模板中自相矛盾或含糊、需要确认的条目
samples/                 测试用论文与模板（含个人信息，已 gitignore）
```

## 使用

需要 Node.js ≥ 24.9 和 pnpm。

```bash
pnpm install
pnpm build
pnpm test

# 检查一篇论文（.doc 会自动转换，不改动原文件）
node packages/core/lib/cli.js 论文.docx
node packages/core/lib/cli.js 论文.doc -f markdown -o reports/报告.md
node packages/core/lib/cli.js 论文.docx --only headings,references
node packages/core/lib/cli.js 论文.docx --dump      # 查看每段被识别成什么，调试规则用
```

退出码：0 无错误，1 有格式错误，2 运行失败。

## 已实现的检查

| 规则 id | 内容 |
|---|---|
| `structure` | 必需部分是否齐全、顺序是否正确 |
| `page-setup` | A4 纸张，上下右 2.5cm、左 3.5cm、装订线 0 |
| `title-block` | 中文题目黑体小二居中；署名行格式 |
| `abstract-zh` | “摘  要：”写法与加粗、正文宋体小四固定 20 磅、300–500 字；关键词 3–5 个、全角分号、末尾无标点 |
| `abstract-en` | 英文题目 TNR 14pt、实词首字母大写；摘要 TNR 五号 1.5 倍、缩进 2.5 字符；Keywords 小写、英文分号、数量与中文一致 |
| `region-title` | “目  录”“致  谢”“参考文献”“附 录”的写法、字体、对齐、段前段后 |
| `toc` | 自动生成、只到二级、包含参考文献与致谢、宋体小四 1.5 倍 |
| `headings` | 各级标题字体字号加粗、缩进 2 字符、段前段后 0、末尾无标点、序号后不空格、编号连续 |
| `body-text` | 正文宋体小四、西文数字 TNR、1.5 倍行距、首行缩进 2 字符（识别用空格假缩进）、不少于 8000 字；致谢正文 |
| `captions` | 图表题注黑体五号居中、序号与名称空一格、编号连续、表题在表上/图题在图下、正文有“见表n/见图n” |
| `references` | ≥15 篇、英文 ≥3 篇、编号连续、类型标识 [J][M]…、以“.”结尾、期刊有页码、宋体五号；正文引用存在且按首次出现排序 |
| `text-hygiene` | 全角数字、“--”代替破折号、千分位逗号、英文字母后用顿号 |

同类问题会合并显示（如“共 36 处”，列出前 5 处的段落序号和原文开头）。

结构识别的原则是“识别从宽、检查从严”：写错的标签（“关键字：”“指导教师 xx”）、Word 自动编号的“[1]”“一、”、放在自动目录控件里的“目  录”都能识别，再由规则报出格式问题。代码清单（无汉字的段落及其中的注释行、“代码1 xxx”标题）不按正文格式检查。

**暂未检查**（需要 Word 排版引擎才能拿到的信息）：页码位置与起始页、页眉页脚、脚注、表格三线表边框、分页。可以后续通过 Word COM 或 LibreOffice 渲染补上。

## 修改或新增规范

规范全部写在 [packages/core/specs/bbctb-cs.yaml](packages/core/specs/bbctb-cs.yaml)，每项都注明了模板原文。

- 某届要求变了：复制一份 YAML 改差异项，用 `--spec 路径` 指定。
- 不想要某条规则：在 `disabledRules` 里列出规则 id。
- 字号可写“小四”“五号”或磅值；字体名中英文别名通用（宋体 = SimSun）。

## 新增检查规则

在 `packages/core/src/rules/` 下实现 `Rule` 接口（`id`、`title`、`check({ doc, spec }) => Issue[]`），加到 `rules/index.ts` 的 `ALL_RULES`。
格式比对用 `checkFormat(paragraph, formatSpec, { ruleId, label })`，它会处理样式继承、字号换算、缩进单位等细节。
每段的 `region` / `kind` 由 `structure.ts` 给出，先用 `--dump` 确认识别结果再写规则。

## 作为 DeepSeek Harness 插件

插件注册两个工具：

- `thesis_format_check(path, spec?, format?, only?)`：运行上面全部确定性检查，返回 Markdown 报告。
- `thesis_read_section(path, region)`：返回某部分纯文本，让模型做内容层面的审阅（摘要是否交代背景与技术实现、关键词是否为短语、参考文献时效性等，这些不适合写死成规则）。

### 打包

插件包是一个 dsh **bundle**（`package.json` 中声明 `dsh.bundle.patch`，由 [cordis.patch.yml](packages/dsh-plugin/cordis.patch.yml) 把插件插入配置树）。
构建时用 esbuild 把核心库及其依赖打进单个 `lib/index.js`，并附带 `specs/`，安装时不需要任何其他包。

```bash
pnpm -r build
pnpm --filter dsh-plugin-thesis-format pack:local   # 生成 dist/dsh-plugin-thesis-format-0.1.0.tgz
```

### 安装到 dsh

任选一种：

- **命令行**（`desktop` 换成你的 profile 名，可在 `~/.dsh/profiles/` 下看到）：

  ```bash
  dsh plugin --profile desktop add "file:D:/2027届毕业设计/论文格式插件/dist/dsh-plugin-thesis-format-0.1.0.tgz"
  ```

  安装后插件自动加入该 profile 的 `dsh.profile.bundles`，重启 dsh 即生效（开启 HMR 时立即生效）。

- **Web / Desktop 界面**：侧边栏 **Plugins** 页面安装同一个 tarball，或安装发布到 npm 后的包名。

在 dsh 中对模型说“检查一下 D:\…\某某论文.doc 的格式”，模型会调用 `thesis_format_check`。
默认规范可在 profile 的 `cordis.patch.yml` 中按 id `thesis-format-check` 覆盖 `config.spec`（内置 id 或 YAML 绝对路径），写法见 [插件 README](packages/dsh-plugin/README.md#配置)。插件页面目前没有该配置的表单。

插件包内附 `LICENSE`（MIT）和构建时自动生成的 `THIRD_PARTY_LICENSES.txt`（打包进 `lib/index.js` 的 22 个第三方包的许可声明；缺少声明时构建失败）。

### 已验证

在隔离的 DSH home 中，对 dsh **0.2.1-rc.3**（当前 Desktop 版本）和 **0.2.2-rc.1** 各做了一遍：
`dsh plugin add` 安装 → `--dump-config` 确认插件进入配置树 → 启动真实运行时，从 dsh 的 ToolRegistry 取出两个工具并逐篇调用。
8 篇样本论文在 dsh 内生成的报告与 CLI 报告逐字一致。

### 注意

- dsh 仍处于开发者预览阶段，接口可能变化。插件只依赖 `ctx.tools.register()` 和原始 `ToolDefinition` 结构（见 [packages/dsh-plugin/src/dsh-types.ts](packages/dsh-plugin/src/dsh-types.ts)），升级时只需改这一层。`engines.dsh` 声明为 `>=0.2.1-0 <0.3.0`，dsh 0.3 发布后需重新验证再放宽。
- 不依赖 `@deepseek-ai/dsh-tools`：插件按原始 `ToolDefinition` 结构注册，运行时由 dsh 提供工具注册表。
- 工具在 dsh 进程内运行，可读取 dsh 进程有权限访问的任意文件路径；`.doc` 转换会调用本机 Word 或 LibreOffice。
- 发布到 npm 后，按 dsh 的约定给仓库加上 `dsh-plugin` topic 便于他人发现。

## 作者

周翔，蚌埠工商学院计算机与人工智能学院

- 邮箱：120230021@bctb.edu.cn
- GitHub：[zx310610](https://github.com/zx310610)
