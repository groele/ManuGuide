# Manuscript Guide / 学术论文排版审查

A Microsoft Word add-in that automatically detects and fixes formatting, typographical, and stylistic issues in scientific manuscripts — with special support for Chinese-English bilingual academic papers.

一款 Microsoft Word 插件，自动检测并修复学术论文中的排版、拼写和格式问题，特别针对中英文混排学术论文进行了优化。

---

## Features / 功能特性

### Diagnostic Modules / 诊断模块

| Module / 模块 | Description (EN) | 说明 (中文) |
|---|---|---|
| **Punctuation & Formatting** | Full-width punctuation misuse, missing spaces after commas, Greek symbol encoding, equation spacing | 标点与格式：全角标点误用、逗号后缺少空格、希腊符号编码、公式间距与末尾标点 |
| **Capitalization** | Abbreviation first-use definitions, acronym case consistency, figure/table/citation capitalization | 大小写：缩写首次定义、缩写大小写一致性、图/表/引用首字母大小写 |
| **Dash & Spacing** | Hyphen vs en-dash vs em-dash, compound modifier spacing, Chinese-English mixed spacing | 横线与间距：连字符/半角破折号/全角破折号用法、复合修饰语间距、中英文混排间距 |
| **Data & Values** | Number-unit spacing, scientific notation, range expressions, math symbol normalization | 数据与数值：数字-单位间距、科学计数法、范围表达、数学符号规范化 |
| **Variables & Italics** | Variable italicization, function upright, academic phrase conventions | 斜体：变量斜体、函数正体、学术短语斜体规范 |
| **Subscripts & Superscripts** | Chemical formula subscripts, unit powers, native Word formatting & Unicode fallback | 上下标：化学式下标、单位幂次、Word 原生格式与 Unicode 回退方案 |
| **Wording & Grammar** | Common academic writing errors, redundant phrasing, English expression issues | 措辞与语病：常见学术写作错误、冗余表达、英文措辞问题 |

### Additional Features / 其他功能

- **Keyword Search & Highlight / 关键词搜索与高亮** — Search for specific terms (e.g., "PL", "strain") and highlight all occurrences in the document. 搜索特定术语并在文档中高亮所有匹配项。
- **Side Panel / 侧边栏** — A 480px-wide WPF task pane docked on the right, showing issue cards with before/after suggestions, per-module stats, and one-click accept/ignore. 右侧停靠的任务面板，展示问题卡片（含修改前后预览）、各模块统计，支持一键接受/忽略。
- **One-Click Fix All / 一键全部修复** — Runs all scanners then batch-applies auto-fixable corrections, leaving light shading as an audit trail. 运行所有扫描器，批量应用可自动修复的更正，保留浅色底纹作为审计记录。
- **Protected Range Handling / 受保护区域处理** — Automatically identifies and protects formula/inline-equation regions from being corrupted. 自动识别并保护公式/行内公式区域不被破坏。
- **Correction Tracking & Undo / 更正追踪与撤销** — Tracks all applied corrections with full undo support. 追踪所有已应用的更正，支持完整撤销。
- **Advanced Settings / 高级设置** — Per-rule toggles, color customization, whitelists, thresholds, and subscript formatting mode. 支持按规则开关、颜色自定义、白名单、阈值设置和下标格式模式选择。

---

## Prerequisites / 环境要求

| Requirement / 要求 | Details / 详情 |
|---|---|
| **OS / 操作系统** | Windows |
| **Microsoft Word** | 2013 or later / 2013 或更高版本 |
| **IDE / 开发工具** | Visual Studio 2022 with VSTO workload / Visual Studio 2022（需安装 VSTO 工作负载） |
| **.NET Framework** | 4.7.2 |
| **VSTO Runtime** | 4.0 |

---

## Getting Started / 快速开始

### Build & Run / 构建与运行

1. Clone this repository / 克隆本仓库：
   ```bash
   git clone https://github.com/<your-username>/Manuscript-Guide.git
   ```

2. Open the solution in Visual Studio 2022 / 在 Visual Studio 2022 中打开解决方案：
   ```
   Manuscript guide/Manuscript guide.sln
   ```

3. Build the solution (Debug or Release) / 构建解决方案（Debug 或 Release）。

4. Press **F5** to launch Word with the add-in loaded / 按 **F5** 启动 Word 并加载插件。

5. Word will display a new Ribbon group with diagnostic buttons, and the task pane sidebar ("学术论文排版审查") will auto-open on the right. Word 将显示新的功能区按钮组，侧边栏任务面板会自动在右侧打开。

---

## Project Structure / 项目结构

```
Manuscript Guide/
├── Guideline/                          # Reference docs for academic writing rules
│   │                                   # 学术论文写作规范参考文档
│   ├── 科技论文中横线的使用规范.md
│   ├── 科技论文中斜体的使用规范.md
│   ├── 科技论文中大小写的使用规范.md
│   ├── 科技论文中数据与数值表达规范.md
│   ├── 科技论文中常见写作语病与措辞规范.md
│   ├── 科技论文中中英文标点与混排规范.md
│   ├── 科技论文中角标字符与排版上标下标的规范.md
│   └── ...
│
└── Manuscript guide/                   # Main VSTO add-in project / 主项目
    ├── ThisAddIn.cs                    # Add-in entry point / 插件入口
    ├── Ribbon1.cs                      # Ribbon UI & button handlers / 功能区与按钮处理
    ├── Models/
    │   └── IssueItem.cs                # Issue data model / 问题数据模型
    ├── Scanners/                       # Diagnostic scanner modules / 诊断扫描模块
    │   ├── ISpecializedScanner.cs      # Scanner interface / 扫描器接口
    │   ├── PunctuationScanner.cs
    │   ├── CapitalizationScanner.cs
    │   ├── DashSpacingScanner.cs
    │   ├── DataValueScanner.cs
    │   ├── ItalicsScanner.cs
    │   ├── SubscriptScanner.cs
    │   └── WordingGrammarScanner.cs
    ├── Services/                       # Core services / 核心服务
    │   ├── EventBus.cs                 # Pub/sub event bus / 事件总线
    │   ├── SettingsManager.cs          # JSON settings persistence / JSON 设置持久化
    │   ├── ShadingManager.cs           # Word highlight/shading / Word 高亮/底纹
    │   ├── ProtectedRangeService.cs    # Formula region protection / 公式区域保护
    │   └── CorrectionTracker.cs        # Undo tracking / 撤销追踪
    └── UI/                             # User interface / 用户界面
        ├── TaskPaneHost.cs             # WinForms host for WPF / WPF 的 WinForms 宿主
        └── TaskPaneWPFControl.xaml     # Sidebar task pane / 侧边栏任务面板
```

---

## Architecture / 架构设计

```
┌──────────────┐     EventBus      ┌──────────────────┐
│   Ribbon UI  │ ──── (pub/sub) ────▶  Task Pane (WPF) │
│  (WinForms)  │                    │  Issue cards,    │
└──────────────┘                    │  settings, fixes │
       │                            └────────┬─────────┘
       │ invoke                              │
       ▼                                     │ scan / fix
┌──────────────────┐                         │
│  Diagnostic      │◀────────────────────────┘
│  Module Registry │
└───────┬──────────┘
        │ dispatch
        ▼
┌──────────────────────────────────────┐
│  ISpecializedScanner implementations │
│  (Punctuation, Capitalization, Dash, │
│   Data, Italics, Subscript, Wording) │
└──────────────────────────────────────┘
```

- **Event Bus / 事件总线**: Decouples Ribbon buttons from task pane logic via static pub/sub events. 通过静态发布/订阅事件将功能区按钮与任务面板逻辑解耦。
- **Scanner Interface / 扫描器接口**: All modules implement `ISpecializedScanner.Scan(Document) → List<IssueItem>` for a clean plugin architecture. 所有模块实现统一的扫描器接口，采用插件式架构。
- **Settings / 设置**: Persisted as JSON via `SettingsManager`, storing per-rule toggles, colors, whitelists, and thresholds. 通过 JSON 持久化设置，包括规则开关、颜色、白名单和阈值。

---

## Contributing / 参与贡献

1. Fork this repository / Fork 本仓库
2. Create a feature branch / 创建功能分支：`git checkout -b feature/my-feature`
3. Commit your changes / 提交更改：`git commit -m "Add my feature"`
4. Push to the branch / 推送到分支：`git push origin feature/my-feature`
5. Open a Pull Request / 创建 Pull Request

---

## License / 许可证

This project is for academic and research purposes. / 本项目仅供学术和研究用途。
