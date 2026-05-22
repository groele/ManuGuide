using System.Collections.Generic;

namespace Manuscript_guide.Services
{
    public sealed class DiagnosticRuleDefinition
    {
        public DiagnosticRuleDefinition(string moduleType, string ruleId, string title, string description, bool defaultEnabled = true)
        {
            ModuleType = moduleType;
            RuleId = ruleId;
            Title = title;
            Description = description;
            DefaultEnabled = defaultEnabled;
        }

        public string ModuleType { get; private set; }
        public string RuleId { get; private set; }
        public string Title { get; private set; }
        public string Description { get; private set; }
        public bool DefaultEnabled { get; private set; }

        public string Key
        {
            get { return ModuleType + "." + RuleId; }
        }
    }

    public static class DiagnosticRuleRegistry
    {
        private static readonly List<DiagnosticRuleDefinition> Rules = new List<DiagnosticRuleDefinition>
        {
            new DiagnosticRuleDefinition("global", "filter_references_and_citations", "过滤引用与参考文献区域", "优先保护 Zotero、EndNote、Mendeley 等引用软件生成的 Word Field，再用 References/Bibliography/参考文献标题兜底。开启后，检测、标记、替换、修复、清理和格式规范化都不会作用于这些区域。"),
            new DiagnosticRuleDefinition("global", "skip_latex_formula_regions", "跳过 LaTeX 公式模块", "在所有检测、标记、替换和改写功能中跳过 LaTeX 公式、Word 原生公式对象和 MathType 公式对象，避免误改公式内容并提升长文档处理速度。"),
            new DiagnosticRuleDefinition("global", "mark_skipped_formula_regions", "标记已跳过的 LaTeX 公式区域", "使用插件专属高亮标记被保护的公式区域，便于检查；清除插件标记时可一键移除，不影响用户手动高亮。", false),
            new DiagnosticRuleDefinition("global", "preserve_user_highlights", "清除时保护用户手动标记", "执行清除插件痕迹时，只清除带有 Manuscript Guide 私有标识的书签、批注和底纹，不清理用户手动添加的高亮或批注。"),
            new DiagnosticRuleDefinition("global", "auto_backup_before_bulk_fix", "一键修复前自动备份", "执行一键修复前，在原文档目录生成时间戳备份，降低批量修改带来的误改风险。"),

            new DiagnosticRuleDefinition("punc", "fullwidth_punctuation", "全角标点识别", "识别英文论文中误用的中文全角逗号、句号、分号和冒号。"),
            new DiagnosticRuleDefinition("punc", "punctuation_spacing", "半角标点后空格", "识别英文半角逗号、分号、冒号、问号、感叹号后缺失空格的情况。"),
            new DiagnosticRuleDefinition("punc", "greek_mu_encoding", "微符号编码规范化", "识别 Micro Sign (U+00B5)，建议统一为 Greek Mu (U+03BC)。"),
            new DiagnosticRuleDefinition("punc", "equation_spacing", "公式域前后空格", "检查 Word 公式域与前后英文/数字之间是否缺失半角空格。"),
            new DiagnosticRuleDefinition("punc", "equation_terminal_punctuation", "公式结尾标点", "检查公式作为句子结尾时是否遗漏句号、逗号等标点。"),

            new DiagnosticRuleDefinition("cap", "acronym_definition", "缩写首次定义", "检查缩写是否首次定义、是否重复定义、是否先使用后定义。"),
            new DiagnosticRuleDefinition("cap", "casing_consistency", "大小写一致性", "根据全文多数派拼写判断术语大小写是否前后一致。"),
            new DiagnosticRuleDefinition("cap", "crossref_capitalization", "图表/文献引用首字母", "检查 figure、table、fig.、eq.、ref. 等交叉引用是否应首字母大写。"),
            new DiagnosticRuleDefinition("cap", "physical_variable_casing_lock", "物理变量大小写保护", "在大小写一致性判断中保护 F、f、E、e 等物理/数学变量，避免被全局拼写多数派误改。"),

            new DiagnosticRuleDefinition("dash", "cjk_latin_spacing", "中文后接英文/数字空格", "检查中文字符后紧跟英文或数字时是否缺失半角空格。"),
            new DiagnosticRuleDefinition("dash", "latin_cjk_spacing", "英文/数字后接中文空格", "检查英文或数字后紧跟中文字符时是否缺失半角空格。"),
            new DiagnosticRuleDefinition("dash", "academic_hyphen_endash", "学术复合词横线", "检查特定学术复合短语是否应使用 en dash。"),

            new DiagnosticRuleDefinition("data", "value_unit_spacing", "数值与单位空格", "检查 15nm、300K 等数值与物理单位之间是否缺失空格。"),
            new DiagnosticRuleDefinition("data", "scientific_notation", "科学计数法乘号", "检查 5.2x10^12、5.2*10^12 等是否应使用标准乘号 ×。"),
            new DiagnosticRuleDefinition("data", "from_range_expression", "from...to 范围表达", "检查 from 10-300 这类范围是否应改为 from 10 to 300。"),
            new DiagnosticRuleDefinition("data", "range_endash", "数值范围 en dash", "检查 10 - 300 或 10-300 是否应统一为紧凑 en dash。"),
            new DiagnosticRuleDefinition("data", "plus_minus_sign", "正负号 ±", "检查 +/- 是否应替换为标准正负号 ±。"),
            new DiagnosticRuleDefinition("data", "negative_minus_sign", "负号 −", "检查负数前的普通连字符是否应替换为排版减号 −。"),

            new DiagnosticRuleDefinition("ital", "latin_phrase_upright", "拉丁短语正体", "检查 in situ、in vitro、et al. 等拉丁短语是否被误设为斜体。"),
            new DiagnosticRuleDefinition("ital", "math_function_upright", "数学函数正体", "检查 sin、cos、exp、log 等函数名是否被误设为斜体。"),
            new DiagnosticRuleDefinition("ital", "existing_italics_review", "已存在斜体复核", "额外列出正文中所有已存在斜体文本，供人工复核。", false),
            new DiagnosticRuleDefinition("ital", "variable_italic", "单字母变量斜体", "检查独立单字母物理/数学变量是否应设为斜体。"),

            new DiagnosticRuleDefinition("sub", "unicode_subsup_to_native", "Unicode 角标转规范", "在 Word 原生角标规范下，识别 Unicode 角标并建议转为 ASCII + Word 原生格式。"),
            new DiagnosticRuleDefinition("sub", "element_formula_subscript", "元素化学式数字下标", "识别 WSe2、MoS2、Bi2O2Se 等化学式并标注未下标化数字。"),
            new DiagnosticRuleDefinition("sub", "descriptive_subscript", "描述性物理角标", "检查 Eg、EF、kB、Vg 等描述性物理角标是否规范。"),
            new DiagnosticRuleDefinition("sub", "latex_inline_subsup", "LaTeX 行内上下标", "识别 E_g、T_{c}、x^2 等 LaTeX 风格行内上下标并转换为 Word 格式。"),

            new DiagnosticRuleDefinition("word", "data_plural_agreement", "data 复数谓语", "检查 data is/has/was 等主谓一致风险。"),
            new DiagnosticRuleDefinition("word", "redundant_in_order_to", "in order to 精简", "建议将冗余的 in order to 精简为 to。"),
            new DiagnosticRuleDefinition("word", "redundant_as_is_shown", "as is shown in 精简", "建议将 as is shown in 精简为 as shown in。"),
            new DiagnosticRuleDefinition("word", "chinglish_companion_syntax", "中式伴随状语", "识别 with the increase of 等中式表达并给出替代表达。")
        };

        public static IEnumerable<DiagnosticRuleDefinition> AllRules
        {
            get { return Rules; }
        }

        public static IEnumerable<DiagnosticRuleDefinition> GetRulesForModule(string moduleType)
        {
            foreach (DiagnosticRuleDefinition rule in Rules)
            {
                if (rule.ModuleType == moduleType)
                {
                    yield return rule;
                }
            }
        }

        public static DiagnosticRuleDefinition Find(string moduleType, string ruleId)
        {
            foreach (DiagnosticRuleDefinition rule in Rules)
            {
                if (rule.ModuleType == moduleType && rule.RuleId == ruleId)
                {
                    return rule;
                }
            }

            return null;
        }

        public static DiagnosticRuleDefinition FindByKey(string key)
        {
            if (string.IsNullOrEmpty(key))
            {
                return null;
            }

            foreach (DiagnosticRuleDefinition rule in Rules)
            {
                if (rule.Key == key)
                {
                    return rule;
                }
            }

            return null;
        }

        public static string ResolveRuleId(string moduleType, string subtype)
        {
            if (string.IsNullOrEmpty(moduleType))
            {
                return subtype;
            }

            switch (moduleType)
            {
                case "punc":
                    if (subtype == "MissingSpaceAfterPunctuation") return "punc.punctuation_spacing";
                    if (subtype == "MicroSymbolEncoding") return "punc.greek_mu_encoding";
                    if (subtype == "EquationSpacing") return "punc.equation_spacing";
                    if (subtype == "EquationPunctuation") return "punc.equation_terminal_punctuation";
                    return "punc.fullwidth_punctuation";
                case "cap":
                    if (subtype == "CasingInconsistency") return "cap.casing_consistency";
                    if (subtype == "CrossRefCapitalization") return "cap.crossref_capitalization";
                    return "cap.acronym_definition";
                case "dash":
                    if (subtype == "CJKLatinSpacing") return "dash.cjk_latin_spacing";
                    if (subtype == "LatinCJKSpacing") return "dash.latin_cjk_spacing";
                    return "dash.academic_hyphen_endash";
                case "data":
                    if (subtype == "ValueUnitSpacing") return "data.value_unit_spacing";
                    if (subtype == "ScientificNotation") return "data.scientific_notation";
                    if (subtype == "RangeExpression") return "data.from_range_expression";
                    if (subtype == "RangeEnDash") return "data.range_endash";
                    if (subtype == "PlusMinusSign") return "data.plus_minus_sign";
                    return "data.negative_minus_sign";
                case "ital":
                    if (subtype == "LatinUpright") return "ital.latin_phrase_upright";
                    if (subtype == "MathFunctionUpright") return "ital.math_function_upright";
                    if (subtype == "ExistingItalicReview") return "ital.existing_italics_review";
                    return "ital.variable_italic";
                case "sub":
                    if (subtype == "UnicodeSubscript") return "sub.unicode_subsup_to_native";
                    if (subtype == "ChemicalSubscriptMissing") return "sub.element_formula_subscript";
                    if (subtype == "DescriptiveSubscript") return "sub.descriptive_subscript";
                    return "sub.latex_inline_subsup";
                case "word":
                    if (subtype == "DataPluralAgreement") return "word.data_plural_agreement";
                    if (subtype == "RedundantInOrderTo") return "word.redundant_in_order_to";
                    if (subtype == "RedundantAsIsShown") return "word.redundant_as_is_shown";
                    return "word.chinglish_companion_syntax";
                default:
                    return moduleType + "." + subtype;
            }
        }
    }
}
