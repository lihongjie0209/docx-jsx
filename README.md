# docx-jsx：用 JSX 编写 Word 文档

使用 JSX/TSX 描述 Word 文档，编译为 DOCX；也可将支持的 DOCX 逆向为可重新编译的
JSX。项目使用 Rust CLI 和内嵌 V8 运行前端，以 [`docx-rs`](https://github.com/bokuweb/docx-rs)
生成 DOCX，无需安装 Node.js。

[最新版本](https://github.com/lihongjie0209/docx-jsx/releases/latest) ·
[DSL 规范](docs/spec.md) · [docx-rs 覆盖审计](docs/docx-rs-coverage.md)

## 快速开始

```tsx
import { Document, Section, Paragraph, Run } from "docx-jsx";

export default ({ name = "World" } = {}) => (
  <Document defaultFont="Arial">
    <Section pageSize="A4">
      <Paragraph align="center">
        <Run bold size={18}>Hello {name}</Run>
      </Paragraph>
    </Section>
  </Document>
);
```

```sh
cargo install --path .
docx-jsx report.tsx --data data.json -o report.docx
docx-jsx report.tsx -o report.pdf
```

校验完整 DSL，但不生成 DOCX 文件：

```sh
docx-jsx validate report.tsx --data data.json
```

校验会汇总能够独立检测到的语义错误和警告，并提供 JSX 文件、行列号及组件路径。
语法、模块加载和运行时错误仍会提前终止。校验可以发现与本地样式定义仅大小写不同的引用，
并提示 JSX 文本中的字面量 `\uXXXX`（JSX 文本不会解析 JavaScript 转义）。未在本地声明的
样式引用仍被允许，因为它可能来自外部 Word 模板。

将 DOCX 逆向为确定性 JSX：

```sh
docx-jsx reverse report.docx -o report.jsx
```

当 `-o` 以 `.pdf` 结尾时，编译命令可直接输出 PDF。`docx-jsx pdf` 则转换已有的 DOCX。
两种方式都调用本机已安装的办公软件（不会随程序打包）：

```sh
docx-jsx report.tsx -o report.pdf --engine word
docx-jsx pdf report.docx -o report.pdf
docx-jsx pdf report.docx --engine wps --force
```

`auto` 会依次尝试 Microsoft Word、WPS 和 LibreOffice。若均未安装，命令会说明安装方式，
以及如何设置 `--engine`、`--soffice`、`DOCX_JSX_SOFFICE`、`DOCX_JSX_WORD` 或
`DOCX_JSX_WPS`。

通过 Word 样式继承集中定义格式，避免在每个组件中重复书写：

```tsx
const styles = [
  { id: "BodyBase", name: "Body Base", type: "paragraph", paragraph: { spacingAfter: 6 } },
  { id: "Body", name: "Body", type: "paragraph", basedOn: "BodyBase", run: { font: "Arial" } },
  { id: "Emphasis", name: "Emphasis", type: "character", run: { italic: true } },
];

export default (
  <Document styles={styles}>
    <Section><Paragraph style="Body"><Run style="Emphasis">Reusable formatting</Run></Paragraph></Section>
  </Document>
);
```

输出 Agent 可读的完整规范，或机器可读的 IR Schema：

```sh
docx-jsx spec
docx-jsx spec --format json-schema
```

GitHub Releases 提供 Linux、macOS 和 Windows 的压缩包及安装程序。

## 功能矩阵

| 功能领域 | 状态 | 支持范围 |
| --- | --- | --- |
| JSX/TSX 执行 | 支持 | 内嵌 V8、本地 ESM 导入、异步组件和 JSON 数据注入 |
| 样式与继承 | 支持 | 段落、字符、编号和表格样式；支持 `basedOn`、段落 `next`、关联样式、类型化引用和循环检测。双向保留分页、孤行控制、双向文本、底纹、字符单位缩进、字体槽位、字号、强调、隐藏文本等属性 |
| 文档默认值与元数据 | 支持 | 多语言字体槽位、主题字体、间距、视图缩放、修订设置、核心属性、文档设置、编辑限制、自定义属性/变量、Web Extensions 和 Custom XML |
| 节与页面 | 支持 | 多种分节符、行号及重启/间距、纸张/方向/页边距、等宽或独立栏宽及栏间距、页眉页脚、首页、文字方向、文档网格、页码格式和章节分隔符 |
| 段落、标题与列表项 | 支持 | 完整行距、磅值及字符单位缩进（包括 `firstLineChars`、`hangingChars` 等）、段前后间距、镜像缩进、边框、制表位、双向文本、大纲级别及段落标记修订 |
| 文字运行与语义文本 | 支持 | 字体和主题字体、拉丁文/复杂文字字号与粗斜体、大小写、Word 下划线、隐藏文字、上下标、效果、边框/底纹、符号、换行、文字修订及语义包装；不会暴露会生成非法 Schema 的修订 setter |
| 表格 | 支持 | 网格、宽度、布局、定位、边距、单元格宽度与换行、边框、行高/规则、重复表头、禁止跨页拆行、网格前后留白、行修订和单元格布局。逆向保留可表示属性；遇到无法无损表示的值会报错 |
| 域与索引 | 支持 | 通用/类型化域；目录、图表目录和自定义索引；支持级别/页码范围、样式映射、书签范围、刷新、可逆删除修订、域段落格式、静态条目、题注及交叉引用 |
| 列表与编号 | 支持 | 项目符号/有序列表简写、`styleLink`、多级编号、各级格式/文本/缩进/运行格式、实例级起始值覆盖；逆向可将 `w:numPr` 还原为 `List` |
| 批注与内容控件 | 支持 | 书签、稳定 ID 批注、线程回复和解决状态、批注范围、脚注、行内/文档级数据绑定内容控件及超链接；保留复合子节点关系 |
| 图片 | 支持 | 光栅图嵌入、尺寸、旋转、辅助功能描述/标题、装饰图语义、浮动锚点、对齐/偏移、环绕距离、层级及 `simplePosition`。逆向可提取正文、页眉和页脚图片 |
| DOCX → JSX | 支持 | 生成文档按组件 1:1 逆向；外部文档支持样式继承、段落/表格/节属性、缩进、行距、修订、图片、元数据、书签/批注范围和超链接等。文本框因上游 writer 不支持且扁平化会丢失几何信息而拒绝逆向，并提供修复建议 |
| 校验与 Agent 规范 | 支持 | DSL 语义校验、带修复建议的诊断、样式大小写检查、可疑 Unicode 转义警告、Agent 规范及 .NET Open XML 校验；JSON Schema 覆盖全部 74 种 IR 组件的属性和子节点约束 |
| PDF 导出 | 可选 | 编译到 `.pdf` 或使用 `docx-jsx pdf`，自动选择 Word、WPS 或 LibreOffice；可用 `--engine` 指定。不包含在发布二进制中 |
| 后端受限 API | 不输出 | `docx-rs 0.4.22` 的主题部件和嵌入字体仅支持读取；缺失 setter 和会生成非法 OOXML 的 writer API 见覆盖审计 |

详细的 writer API 覆盖审计见 [docs/docx-rs-coverage.md](docs/docx-rs-coverage.md)。

完整组件约定见 [v1 规范](docs/spec.md)。

## Open XML 一致性校验

集成测试会编译代表性 JSX 夹具，并使用 Microsoft `DocumentFormat.OpenXml` 3.5.1 校验生成的
DOCX。诊断以 JSON Lines 输出，包含包部件、XPath、原因和修复建议。测试需要安装 .NET 8 SDK，
并通过配置的华为云 NuGet 镜像还原依赖。校验器在临时目录运行，因此本机可使用其他兼容的
.NET 8 SDK，不受仓库 SDK 锁定版本影响。

```sh
./scripts/openxml-test.sh
```

这项测试补充 Rust DSL 校验：它检查生成 DOCX 的 OPC/OOXML Schema 和语义约束；Rust 测试则
继续验证组件映射以及 DOCX 到 JSX 的往返转换。
