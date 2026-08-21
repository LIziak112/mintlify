# MX Quan — Word / MathType Skills 站点

基于 [Mintlify](https://mintlify.com) 的中文技术站点，内容围绕：

- **word-format**：双栏习题册版式迁移（保留 MathType OLE）
- **mathtype-re-render**：`.eqp` → MTEF 修补 → WMF 重渲染 → 几何同步

## 安装 Skills

```powershell
npx skills add LIziak112/mintlify -s word-format
npx skills add LIziak112/mintlify -s mathtype-re-render
```

或下载 `downloads/*.zip`。说明见站点内 [Skills：安装与使用](./tools/skills.mdx)。

## 本地预览

```powershell
npm i -g mint
mint dev
```

## 仓库结构

```text
blog/           中文技术博客
tools/          Skills 总览与安装页
skills/         完整 skill 源码（给 Agent / 脚本）
downloads/      可下载 zip 包
docs.json       导航与站点配置
```

写文档时可安装 Mintlify 官方 skill：

```powershell
npx skills add https://mintlify.com/docs
```
