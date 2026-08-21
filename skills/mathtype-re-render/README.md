# MathType Re-render 使用说明

## 这个技能是做什么的

一句话：把 MathType 的格式配置文件（`.eqp`）应用到 DOCX 文档中所有**可编辑公式对象**的二进制数据里，并重新渲染公式预览图，使 Word 里显示的公式与二进制配置真正一致。

完整链路：

```text
EQP 配置 -> MTEF 二进制审计/修补 -> 二进制完整性证明
         -> 唯一 WMF 映射 -> x86 OLE 渲染
         -> WMF 替换 + Word 几何同步 -> 最终审计/渲染 QA
```

关键认知：**MTEF 二进制被修改 ≠ Word 显示已更新**。这两层是分开的，各有各的验收门槛，本技能同时处理两层。

## 适用范围：可以处理哪些问题

### 适用对象

- 格式：`.docx` 文档
- 公式类型：可编辑的 `Equation.DSMT4` OLE 对象（OLE2 复合文件中的 `Equation Native` 流）
- 配置类型：标准 `.eqp` 文件（8 项 `[Sizes]` + 30 项 `[Spacing]`，共 38 个数值）

### 能解决的问题

| 问题 | 处理方式 |
| --- | --- |
| 公式二进制配置（字号/间距）与目标 EQP 不一致，如不同来源文档公式格式漂移、批量规范格式 | 保守修补 `Equation Native` 流中的 `EQN_PREFS` 记录（38 值块），支持同步字体记录 |
| 修改配置后 Word 中显示的预览图仍是旧的 | 用 32 位 STA 工作进程重新渲染修补后的 OLE，生成可放置 WMF 替换预览 |
| 多个公式共享同一张预览图，重渲染会互相覆盖 | 拆分共享预览关系，保证一个公式对象对应一张预览 |
| 公式显示尺寸与内容不匹配（拉伸、压缩、裁剪、重叠） | 依据 WMF 物理尺寸同步 `w:dxaOrig/dyaOrig` 与 VML 宽高，保留基线位置 |
| 需要证明修补安全（改了哪、没动哪） | 证明非目标字节、非 `Equation Native` 的 OLE 流、所有非目标 DOCX 部件逐字节不变，ZIP CRC 通过 |
| 批量套用预设格式（教辅、正文、标题等） | 使用 `assets/presets/` 下现成配置一键跑全流程 |

### 不适用 / 处理不了

- 静态图片公式（PNG/EMF 图片），不可编辑、没有 OLE `Equation Native` 流
- Office 原生公式（OMML，MathML）
- **修改数学公式内容本身**——本技能只改配置和显示，不动公式内容
- 非 Windows 环境（依赖 Windows + 已注册的 MathType 32 位 OLE server）
- 未知的新字体映射——需先做对照方程字节比较、扩展补丁器后才能支持（会停下，不猜测）

## 环境要求

- Windows 系统，已安装并注册 MathType 的 32 位 OLE server
- Python（运行脚本）
- 输入：含可编辑公式的 `.docx` + `.eqp` 配置
- 输出：**总是写入新的 DOCX**，禁止源文件与输出同文件，覆盖需明确授权

## 使用方法

### 一键流程（推荐）

```powershell
python .\run_full_pipeline.py target.eqp source.docx refreshed.docx `
  --work-dir .\work\job-name `
  --cache .\cache `
  --batch-size 25
```

任一门槛失败即停止。成功后再用 documents 技能渲染整份文档并逐页检查（结构成功 ≠ 视觉合格）。

### 分步流程（7 个门槛）

1. **预检**：`find_deviating.py target.eqp source.docx --out .\work\source-mtef-audit.json`（退出码 `2` = 存在偏差，需要修复时属预期）
2. **修补**：`patch_mathtype_mtef.py target.eqp source.docx patched.docx --report .\work\mtef-patch-report.json`
3. **证明**：`verify_mtef_patch.py` + 再跑一次 `find_deviating.py`，要求偏差为 0、非目标部件逐字节不变
4. **映射**：`ole-preview-bridge\map_equations.py`；若公式数多于唯一预览数，先 `split_shared_previews.py` 再映射
5. **渲染**：先 `build_renderer32.ps1` 构建 x86 渲染器（每台机器/源码变更后一次），再 `refresh_previews.py`（默认 25 次缓存未命中后重启，避免 MathType 内存泄漏）
6. **几何审计**：`ole-preview-bridge\audit_geometry.py`，要求 `failures=0`
7. **全页渲染 QA**：渲染全部页面，检查漏公式、旧预览框、变形、裁剪、重叠、基线漂移、页面流变化

### 最短安全变体

- 配置错且预览旧 → 跑全流程
- 配置已修补并验证 → 从关系映射 + OLE 重渲染开始，**不要再次修补**
- 仅预览陈旧、EQP 配置已合规 → 跳过修补，但仍要映射、拆分、刷新、审计
- 未知/新字体映射 → 审计后停下，先做对照字节比较，不硬改
- 渲染器不可用 → 诊断 x86 构建与 MathType 注册；**绝不把可编辑公式静默替换成 PNG**

## 内置资源

- `scripts/run_full_pipeline.py`：全流程入口
- `scripts/find_deviating.py`：EQP 对照 MTEF 的审计
- `scripts/patch_mathtype_mtef.py`：保守的 nibble/字体/mini-stream 补丁器
- `scripts/verify_mtef_patch.py`：证明只有允许的 `Equation Native` 流被改动
- `scripts/ole-preview-bridge/`：映射、校验、共享预览拆分、x86 渲染器、刷新引擎、几何审计
- `references/mtef-binary-patching.md`：二进制编码、分配、校验与扩展规则（改补丁器前必读）
- `assets/presets/`：预设配置——教辅格式11.7pm、正文11.9pt、一级标题16pm、二级标题13pm

## 交付物与报告

返回最终 `refreshed.docx`（非中间文件，除非要求）。报告中说明：所选 EQP、公式数量、MTEF 已改/未改数量、共享预览拆分数量、渲染/缓存命中数、几何结果、OLE 保留结果、全页 QA 状态。JSON 报告保留在工作目录中以便复现。
