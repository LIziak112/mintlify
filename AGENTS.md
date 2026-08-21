# Documentation project instructions

## About this project

- 这是基于 [Mintlify](https://mintlify.com) 的中文技术站点
- 页面为带 YAML frontmatter 的 MDX
- 配置在 `docs.json`
- 核心内容：`word-format` 与 `mathtype-re-render` 两个 skill，以及相关技术博客
- Skill 源码在 `skills/`，下载包在 `downloads/`
- Mintlify 产品知识可用：`npx skills add https://mintlify.com/docs`

## Terminology

- 使用「skill」指 Agent 可安装的能力包（含 `SKILL.md` 与脚本）
- MathType 可编辑公式写作 `Equation.DSMT4` / OLE，不要与 OMML 混淆
- 版式迁移（word-format）与公式配置重渲染（mathtype-re-render）是两条流水线

## Style preferences

- 站点文案默认**中文**
- 使用主动语态与第二人称（「你」）
- 句子尽量短——一句一个意思
- 标题用句子式大小写习惯：中文标题不加英文 Title Case
- UI 元素加粗：点击 **设置**
- 文件名、命令、路径、代码引用用行内代码
- 优先使用 Mintlify 组件组织内容：`Card` / `CardGroup`、`Steps`、`Tip` / `Note` / `Warning`、`Tabs`、`Accordion`

## Content boundaries

- 不保留英文随笔模板、假浏览器工具页
- 不把整仓 skill 源码贴进博客正文；博客讲原理与用法，安装指向 `/tools/skills`
- 不鼓励用 HTML 重建含可编辑 MathType 的整篇文档
