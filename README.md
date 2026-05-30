<div align="center">

# ManuGuide

**面向中英文科技论文的 Word 排版、格式与学术语言审查插件**  
*Microsoft Word add-in for manuscript formatting, style diagnostics, bilingual academic writing checks, and controlled one-click correction.*

![Type](https://img.shields.io/badge/type-Word%20Add--in-blue?style=flat-square)
![Platform](https://img.shields.io/badge/platform-Windows%20%2B%20Word-green?style=flat-square)
![Language](https://img.shields.io/badge/language-C%23%20%2B%20VSTO-blueviolet?style=flat-square)
![Architecture](https://img.shields.io/badge/architecture-modular%20scanner-purple?style=flat-square)
![License](https://img.shields.io/badge/license-MIT-yellow?style=flat-square)

Part of **ResearchFlow Lab** — a local-first research productivity ecosystem for literature, manuscripts, data, and scientific visualization.

</div>

---

## 01. Overview

**ManuGuide** is a Microsoft Word add-in for scientific manuscript checking. It detects formatting, punctuation, capitalization, spacing, data-value expression, italicization, subscript/superscript, and academic wording issues, with special attention to Chinese-English mixed scientific writing.

**ManuGuide** 是一个面向科技论文写作的 Word 插件，目标是把论文排版规范、学术表达规范和可审计的一键修复流程整合进 Word 侧边栏。它特别适合中英文混排论文、材料/物理/工程类论文、含大量单位、变量、公式、化学式和图表引用的手稿。

---

## 02. Why this project exists

Scientific manuscripts often contain small but repeated formatting issues: wrong dash types, inconsistent figure references, missing spaces between numbers and units, non-italic variables, malformed chemical subscripts, Chinese-English punctuation conflicts, and redundant academic expressions. These issues are tedious to find manually and risky to fix with blind global replacement.

ManuGuide provides a controlled scanner-and-fix workflow inside Word.

核心目标：

- Detect manuscript issues using specialized rule modules.
- Present issues as inspectable cards instead of uncontrolled global replacement.
- Protect equations and formula-like regions from destructive edits.
- Support Chinese-English mixed academic writing conventions.
- Leave an auditable correction trail with undo support.

---

## 03. Key features

| Module | What it does | 中文说明 |
|---|---|---|
| Punctuation & Formatting | Detects full-width punctuation misuse, comma spacing, Greek symbol encoding, and equation punctuation | 检查全角标点误用、逗号后空格、希腊字符编码、公式标点等 |
| Capitalization | Checks abbreviation definitions, acronym consistency, and figure/table/citation capitalization | 检查缩写首次定义、大小写一致性、图表引用大小写 |
| Dash & Spacing | Distinguishes hyphen, en-dash, em-dash, compound modifiers, and mixed Chinese-English spacing | 检查连字符、短横线、长破折号、复合修饰语和中英文间距 |
| Data & Values | Normalizes number-unit spacing, ranges, scientific notation, and math symbols | 规范数字-单位间距、范围表达、科学计数法和数学符号 |
| Variables & Italics | Detects variable italicization and upright function conventions | 检查变量斜体、函数正体和学术符号规范 |
| Subscripts & Superscripts | Handles chemical formulas, unit powers, Unicode fallback, and native Word formatting | 检查化学式下标、单位幂次、Unicode/Word 原生上下标 |
| Wording & Grammar | Flags redundant academic phrases and common English expression issues | 标记冗余学术表达和常见英文措辞问题 |
| Protected Ranges | Identifies equation and formula regions to avoid accidental corruption | 自动保护公式和行内公式区域 |
| Task Pane UI | Shows issue cards, before/after suggestions, module stats, accept/ignore actions | 右侧任务栏展示问题卡、修改建议、模块统计和接受/忽略操作 |
| Correction Tracker | Tracks applied fixes and supports undo | 记录已应用修复并支持撤销 |

---

## 04. Product philosophy

ManuGuide follows four design principles:

1. **Audit before correction** — every fix should be visible before it changes the manuscript.
2. **Rule modularity** — punctuation, data values, italics, subscripts, and wording should be separate scanners.
3. **Document safety** — formulas, symbols, citations, and protected ranges must not be damaged by automated fixes.
4. **Academic specificity** — rules should target real scientific-writing problems rather than generic grammar correction only.

---

## 05. Architecture

```text
Word Document
    ↓
Ribbon Command / Task Pane Trigger
    ↓
Diagnostic Module Registry
    ↓
Specialized Scanner Modules
├── PunctuationScanner
├── CapitalizationScanner
├── DashSpacingScanner
├── DataValueScanner
├── ItalicsScanner
├── SubscriptScanner
└── WordingGrammarScanner
    ↓
IssueItem Model
    ↓
WPF Task Pane Issue Cards
    ↓
Accept / Ignore / Fix All
    ↓
Correction Tracker + Protected Range Service
```

---

## 06. Quick start

Requirements:

| Requirement | Version |
|---|---|
| Windows | Required |
| Microsoft Word | 2013 or later recommended |
| Visual Studio | 2022 with VSTO workload |
| .NET Framework | 4.7.2 |
| VSTO Runtime | 4.0 |

Development run:

```bash
git clone https://github.com/groele/ManuGuide.git
cd ManuGuide
```

Then:

1. Open the solution file in Visual Studio 2022.
2. Restore dependencies if required.
3. Build the project in Debug or Release mode.
4. Press **F5** to launch Word with the add-in loaded.
5. Open a manuscript and run diagnostics from the Ribbon or task pane.

---

## 07. Recommended workflow

```text
Open manuscript → Run scanners → Review issue cards
                → Accept / ignore individual fixes
                → Run controlled Fix All for safe rules
                → Inspect highlighted audit trail
                → Export / save revised manuscript
```

Typical use cases:

- Final formatting check before journal submission.
- Chinese-English mixed manuscript cleanup.
- Unit, range, and scientific notation normalization.
- Equation-safe search, highlight, and correction.
- Repeated manuscript style enforcement across projects.

---

## 08. Project structure

```text
ManuGuide
├── Guideline/
│   ├── 科技论文中横线的使用规范.md
│   ├── 科技论文中斜体的使用规范.md
│   ├── 科技论文中大小写的使用规范.md
│   ├── 科技论文中数据与数值表达规范.md
│   ├── 科技论文中常见写作语病与措辞规范.md
│   ├── 科技论文中中英文标点与混排规范.md
│   └── 科技论文中角标字符与排版上标下标的规范.md
└── Manuscript guide/
    ├── ThisAddIn.cs
    ├── Ribbon1.cs
    ├── Models/
    ├── Scanners/
    ├── Services/
    └── UI/
```

---

## 09. Roadmap

- [ ] More robust long-document performance optimization
- [ ] Regex-driven custom rule extension layer
- [ ] User-defined journal style profiles
- [ ] Batch checking across multiple Word documents
- [ ] Exportable diagnostic report
- [ ] Safer auto-fix preview for complex ranges
- [ ] Rule severity levels and per-project whitelist

---

## 10. Privacy and data ownership

ManuGuide runs inside Microsoft Word on the user's local machine. Manuscript content is not uploaded by default. Any future AI-assisted module should be optional, user-configured, and clearly separated from local rule-based diagnostics.

---

## 11. Related projects

- **ResearchFlow Companion** — research workflow operating system
- **PaperPilot Pro** — academic search and publisher-page enhancement
- **ClipNote** — browser-native quick notes and Markdown capture
- **Witec-Matlab** — spectroscopy data analysis workflow
- **Scientific Color Lab** — scientific color and visualization workspace

---

## 12. License

MIT License.

Developed by **Shikun Hou / groele**.
