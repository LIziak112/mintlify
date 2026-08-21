---
name: word-format
description: 将参考 Word (.docx) 的字体、字号、缩进、行距、对齐、页面设置、样式表、主题和页眉页脚复刻到新内容；支持 Windows Word COM、双栏习题册模板的 OOXML 修复，以及 Equation.DSMT4 MathType/OLE 保留。用户要求按 Word 模板排版、在已有文档上续写、修复文档后半部分版式或统一 Word 格式时使用。
---

# word-format

使用 Word 自己导出的 Filtered HTML 作为排版模板，只修改文字节点，再通过浏览器渲染、Windows 剪贴板和 Word 粘贴完成输出。

## 工作原理

```
reference.docx
    │  docx_to_html.ps1（Word COM，Filtered HTML）
    ▼
reference.html / append.html
    │  只修改文字，保留原有标签和 inline style
    ▼
Chrome 或 Edge 渲染 → DevTools 选中页面并复制 → 系统剪贴板
    │
    ▼
Word COM 打开副本/新文档 → 粘贴 → 保存为 .docx
```

优先使用续写模式。它先复制参考 `.docx`，再把新内容粘到副本末尾，因此页面设置、样式表、主题、页眉页脚和字体表仍由参考文档承载。

## 习题册模板化排版（首选 OOXML）

对已有的双栏习题册，先读取已排版模板，再使用容器级脚本；不要把整篇文档转成 HTML 后重新粘贴。该路径直接复制源文档的 ZIP 部件，只改 `word/document.xml`、`word/styles.xml` 和模板规定的栏断点，因此可以保留 MathType/OLE、二维码、图片、关系和嵌入对象。

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\scripts\inspect_template.ps1 `
  -InputDocx 'C:\path\M1-template.docx'

powershell -NoProfile -ExecutionPolicy Bypass -File .\scripts\apply_exercise_template_ooxml.ps1 `
  -TemplateDocx 'C:\path\M1-template.docx' `
  -SourceDocx 'C:\path\M4-source.docx' `
  -OutputDocx 'C:\path\M4-formatted.docx'

powershell -NoProfile -ExecutionPolicy Bypass -File .\scripts\audit_exercise_template_ooxml.ps1 `
  -Docx 'C:\path\M4-formatted.docx' `
  -ReferenceDocx 'C:\path\M4-source.docx'
```

默认习题册规则是：题型一保持首个内容位置；其他题型、检测题、检测题答案和反馈通道从下一栏顶部开始；母题之间只保留模板段落间距；反馈通道标题、两个二维码和题册调整必须同栏；`Step n` 的冒号及后续文字/公式清除直接加粗。普通正文套用模板样式时清除残留直接缩进，图题/图示等特殊几何段落保留自身定位；`MTDisplayEquation` 独立公式段保留公式 OLE，但删除开头定位 Tab，取消段前/段后间距并显式左对齐。完整规则和验收清单见 [`references/exercise_book_layout.md`](references/exercise_book_layout.md)。需要覆盖已有输出时显式传入 `-Force`。

完成 OOXML 修改后必须运行审计器，再使用 skill 自带的渲染入口生成 PDF、逐页 PNG、检查页和 JSON 报告：

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\scripts\render_docx_for_review.ps1 `
  -InputDocx 'C:\path\formatted.docx' `
  -OutputDirectory 'C:\path\formatted-render'
```

脚本使用 Word COM 导出 `document.pdf`，再使用 Windows 自带的 PDF 渲染 API生成 `page-0001.png` 等页面图，同时写出 `review.html` 和 `render-report.json`。`Status=PASS` 只表示 PDF 页数、PNG 数量和非空文件检查通过；仍须打开 `review.html` 或逐张查看 PNG，检查栏起始、二维码、公式完整性、重叠、截断和空白页。不要只抽查第一页。

## 环境要求

- Windows 10/11
- Microsoft Word 桌面版（Microsoft 365 或 Office 2016 及以上，必须支持 COM；网页版 Word 不支持）
- Google Chrome 或 Microsoft Edge
- Windows PowerShell 5.1。内置 PDF→PNG 渲染器使用 Windows Runtime，不依赖 Python、Poppler 或 LibreOffice。
- 脚本运行时必须有可交互的桌面会话。复制由 DevTools 完成，不模拟键盘；运行期间不要主动覆盖系统剪贴板。

首次使用先检查环境：

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\scripts\verify_env.ps1
```

## 硬约束：编辑 HTML 时遵守

1. 所有新样式写在标签的 `style` 属性中；不要加入 `<style>` 块或外部 CSS。
2. 字号使用 `pt`，不要使用 `px`。
3. 表格直接使用 `<table align="center" ...>` 居中，不使用 `margin:auto`。
4. `<body>` 保持 `margin:0;padding:0;`，避免粘贴后整体偏移。
5. 大表格使用 `style="width:440pt"`，小表格使用 `style="width:auto"`。
6. 从原 HTML 复制同类型段落作为模板，只替换文字；不要重建标签层级或引入原文没有的字体名。
7. HTML 文本节点中不要写 Markdown（例如 `**粗体**`、`# 标题`、`- 列表`）。

Word 导出的 Filtered HTML 可能同时使用 `<style>` 中的类定义和 `<p class=数字>` 等 class 属性。这里的 `<style>` 是原文档的样式表，不要删除、扁平化或重命名；“不要加入 `<style>`”只针对新建样式。

## 自定义样式与 MathType/OLE

习题册模板通常含有“章节名称”“题型”“母题”“子题1”“解析加粗”等自定义样式。先检查模板，确认实际样式名称和公式对象数量：

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\scripts\inspect_template.ps1 `
  -InputDocx 'C:\path\reference.docx'
```

`Equation.DSMT4` 是 MathType OLE 对象，不是 OMML。Word 导出的 HTML 会把它们变成 `*.files/*.png` 图片，浏览器剪贴板也只能传图片，不能凭空恢复可编辑 OLE。因此：

- 包含 `Equation.DSMT4` 的模板必须使用 `-AppendTo`。脚本会把原 `.docx` 复制为输出承载文档，原有公式和自定义样式不经过 HTML 管线。
- `docx_to_html.ps1` 会在 HTML 末尾写入不可见的公式数量标记；`render_and_paste.ps1` 在新建模式发现该标记时会停止，防止静默把公式变成图片。
- 续写时，新粘贴的公式如果来自浏览器 HTML 仍然是图片。需要新增可编辑 MathType 公式时，应在 Word/MathType 中插入或复制已有 OLE 对象，不能用普通 HTML 公式替代。
- 编辑导出的 HTML 时保留原有 class、`<style>`、`<img>` 和锚点；只改题干、解析等文字节点。

## 完整流程

### 1. 导出参考文档

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\scripts\docx_to_html.ps1 `
  -InputDocx 'C:\path\reference.docx'
# 默认输出到 $env:LOCALAPPDATA\word-format-skill\reference.html
# 输出还会报告 Equation.DSMT4 数量，并在 HTML 中写入不可见标记
```

也可以显式指定输出：

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\scripts\docx_to_html.ps1 `
  -InputDocx 'C:\path\reference.docx' `
  -OutputHtml 'C:\path\reference.html'
```

### 2. 编辑 HTML

续写时只生成要追加的片段，例如从参考 HTML 复制一个正文 `<p>`，保留完整 `style`，只替换文本内容并保存为 `append.html`。不需要完整的 `<html>` 外壳；浏览器可以直接渲染零散段落。

整文重写时复制原 HTML 后编辑：

```powershell
Copy-Item -LiteralPath 'C:\path\reference.html' -Destination 'C:\path\reference.edited.html'
```

编辑过程中保留所有标签、`style` 和嵌套结构。

### 3. 渲染、粘贴和保存

续写模式（推荐，保留完整模板）：

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\scripts\render_and_paste.ps1 `
  -AppendTo 'C:\path\reference.docx' `
  -InputHtml 'C:\path\append.html' `
  -OutputDocx 'C:\path\final.docx'
```

新建模式（只继承字符和段落级格式）：

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\scripts\render_and_paste.ps1 `
  -InputHtml 'C:\path\reference.edited.html' `
  -OutputDocx 'C:\path\final.docx'
```

若该 HTML 来自含 MathType/OLE 的模板，脚本会拒绝新建模式；改用上面的 `-AppendTo` 形式。这样可以保留原文档已有的 OLE 公式。

`-Browser Chrome` 或 `-Browser Edge` 可强制选择浏览器；默认自动选择已安装的 Chrome，否则使用 Edge。`-RenderDelaySeconds 5` 可在字体或复杂页面加载较慢时增加等待时间。

脚本默认保存后关闭本次创建的 Word 实例和隔离浏览器，并恢复脚本运行前的前台窗口；不会关闭用户原有的 Word/Chrome。调试时传 `-KeepOpen` 可保留本次实例。

## 故障排查

| 现象 | 处理 |
|---|---|
| 找不到 Word COM | 安装/修复 Microsoft Word 桌面版；不要使用网页版 Word。 |
| 找不到浏览器窗口 | 安装 Chrome 或 Edge，确认脚本运行在有桌面的用户会话中。 |
| 粘贴为空或内容不全 | 增大 `-RenderDelaySeconds`；运行期间不要切换窗口或操作键鼠。 |
| 模式 B 页面设置变成默认 | 确认传入了 `-AppendTo`，且输出路径与参考文档不同。 |
| 字体或字号变化 | 检查是否引入了原文没有的字体名或 `px` 单位。 |
| 页面整体左移 | 检查 `<body>` 是否包含 `margin:0;padding:0;`。 |
| 表格没有居中 | 在 `<table>` 标签上补 `align="center"`。 |
| 新建模式提示 Equation.DSMT4 | 这是为了避免把 MathType/OLE 公式变成 PNG；复制参考 `.docx` 并使用 `-AppendTo`。 |
| 新增公式无法在 MathType 中编辑 | 浏览器剪贴板只能传图片；在 Word/MathType 中插入或复制 OLE 对象。 |
| DOCX 导出 PDF 时出现 `80070520` | 保存并关闭本轮自动化遗留的 Word 实例后重试；不要擅自结束用户正在使用的 Word 进程。 |
| 渲染目录已存在 | 换用新的输出目录；确认仅替换渲染产物时才传 `-Force`。 |
| `render-report.json` 为 PASS 但页面仍有问题 | PASS 只检查文件和页数；必须继续打开 `review.html` 逐页目检。 |

`docx_to_html.py` 在 Windows 上也会自动转调同目录的 `docx_to_html.ps1`；直接使用 PowerShell 入口可以减少 Python 环境依赖。`export_docx_pdf.ps1` 和 `pdf_to_page_images.ps1` 可单独调用，通常直接使用 `render_docx_for_review.ps1`。目录中的 `.sh` 文件仅供原 macOS 流程使用。
