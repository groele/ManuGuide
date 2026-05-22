using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using Microsoft.Office.Interop.Word;
using Manuscript_guide.Models;
using Manuscript_guide.Services;

namespace Manuscript_guide.Scanners
{
    public class SubscriptScanner : ISpecializedScanner
    {
        public string ModuleType => "sub";

        private static readonly HashSet<string> ElementSymbols = new HashSet<string>(StringComparer.Ordinal)
        {
            "H", "He", "Li", "Be", "B", "C", "N", "O", "F", "Ne",
            "Na", "Mg", "Al", "Si", "P", "S", "Cl", "Ar", "K", "Ca",
            "Sc", "Ti", "V", "Cr", "Mn", "Fe", "Co", "Ni", "Cu", "Zn",
            "Ga", "Ge", "As", "Se", "Br", "Kr", "Rb", "Sr", "Y", "Zr",
            "Nb", "Mo", "Tc", "Ru", "Rh", "Pd", "Ag", "Cd", "In", "Sn",
            "Sb", "Te", "I", "Xe", "Cs", "Ba", "La", "Ce", "Pr", "Nd",
            "Pm", "Sm", "Eu", "Gd", "Tb", "Dy", "Ho", "Er", "Tm", "Yb",
            "Lu", "Hf", "Ta", "W", "Re", "Os", "Ir", "Pt", "Au", "Hg",
            "Tl", "Pb", "Bi", "Po", "At", "Rn", "Fr", "Ra", "Ac", "Th",
            "Pa", "U", "Np", "Pu", "Am", "Cm", "Bk", "Cf", "Es", "Fm",
            "Md", "No", "Lr", "Rf", "Db", "Sg", "Bh", "Hs", "Mt", "Ds",
            "Rg", "Cn", "Nh", "Fl", "Mc", "Lv", "Ts", "Og"
        };

        private static readonly HashSet<string> CommonSingleElementFormulaTokens = new HashSet<string>(StringComparer.Ordinal)
        {
            "H2", "O2", "O3", "N2", "F2", "Cl2", "Br2", "I2", "P4", "S8", "C60", "C70"
        };

        private static readonly Regex UnicodeSubSupRegex = new Regex(
            @"(?<![A-Za-z0-9_])(?:[A-Za-z]+[₀₁₂₃₄₅₆₇₈₉ₓ]+|[A-Za-z]+[⁰¹²³⁴⁵⁶⁷⁸⁹⁺⁻⁼]+)(?![A-Za-z0-9_])",
            RegexOptions.Compiled);

        private static readonly Regex ElementFormulaRegex = new Regex(
            @"(?<![A-Za-z0-9_])(?:[A-Z][a-z]?\d{0,3})*\d(?:[A-Z][a-z]?\d{0,3})*(?![A-Za-z0-9_])",
            RegexOptions.Compiled);

        private static readonly Regex DescriptiveSubscriptRegex = new Regex(
            @"(?<![A-Za-z0-9_])(?:E[gF]|kB|V(?:g|th)|T(?:c|m)|R(?:s|sh)|I(?:d|sd|ph|PL)|PL)(?![A-Za-z0-9_])",
            RegexOptions.Compiled);

        private static readonly Regex LatexInlineSubSupRegex = new Regex(
            @"(?<![A-Za-z0-9_])(?:[A-Za-z][A-Za-z0-9]{0,4})_\{?[A-Za-z0-9\-]{1,8}\}?|(?<![A-Za-z0-9_])(?:[A-Za-z][A-Za-z0-9]{0,4})\^\{?[A-Za-z0-9+\-=]{1,8}\}?",
            RegexOptions.Compiled);

        public List<IssueItem> Scan(Document doc)
        {
            List<IssueItem> issues = new List<IssueItem>();
            string text = DocumentScanContext.GetText(doc);
            if (string.IsNullOrEmpty(text)) return issues;
            PluginSettings settings = SettingsManager.Current;

            // --- 1. Unicode Subscript/Superscript Detection ---
            // Subscripts: ₀₁₂₃₄₅₆₇₈₉ₓ
            // Superscripts: ⁰¹²³⁴⁵⁶⁷⁸⁹⁺⁻⁼
            if (SettingsManager.IsRuleEnabled(ModuleType, "unicode_subsup_to_native"))
            {
                foreach (Match match in UnicodeSubSupRegex.Matches(text))
                {
                    string origText = match.Value;
                    // Determine sub-type
                    string formatType = "chemical";
                    if (Regex.IsMatch(origText, @"[⁰¹²³⁴⁵⁶⁷⁸⁹⁺⁻⁼]"))
                    {
                        formatType = "unit";
                    }

                    // Predict clean translation recommendation
                    string recommend = settings.UseNativeSubscript
                        ? SubscriptFormatter.ConvertToAsciiText(origText, formatType)
                        : origText;

                    string description = settings.UseNativeSubscript
                        ? $"当前正文角标规范为 Word 原生角标。检测到 Unicode 编码的上下标字符“{origText}”，建议转换为标准 ASCII 字符并应用 Word 原生上下标格式。"
                        : $"当前正文角标规范为 Unicode 角标字符。检测到 Unicode 编码的上下标字符“{origText}”，此项作为全文角标复核命中；如需改用 Word 原生角标，可在高级设置中切换输出方式。";

                    AddIssue(doc, text, issues, "UnicodeSubscript", match.Index, match.Length, origText, recommend,
                        description);
                }
            }

            // --- 2. Element-driven Chemical Formula Digits Missing Subscripts ---
            // Based on Element subscript conversion.js: only accept tokens that can be
            // fully segmented into valid chemical element symbols. This catches WSe2,
            // MoS2, Bi2O2Se, CuInP2S6, and avoids ordinary prose words.
            if (SettingsManager.IsRuleEnabled(ModuleType, "element_formula_subscript"))
            {
                foreach (Match match in ElementFormulaRegex.Matches(text))
                {
                    string formula = match.Value;
                    if (LooksLikeSupplementalReferenceToken(text, match.Index, formula))
                    {
                        DocumentScanContext.RecordCandidate(ModuleType, "ChemicalSubscriptMissing");
                        DocumentScanContext.RecordSkip(ModuleType, "ChemicalSubscriptMissing", ScannerSkipReason.RuleFilter);
                        continue;
                    }

                    if (!IsLikelyElementFormulaToken(formula))
                    {
                        DocumentScanContext.RecordCandidate(ModuleType, "ChemicalSubscriptMissing");
                        DocumentScanContext.RecordSkip(ModuleType, "ChemicalSubscriptMissing", ScannerSkipReason.RuleFilter);
                        continue;
                    }

                    Range r = DocumentScanContext.CreateRangeFromTextSpan(doc, match.Index, match.Length, formula);
                    if (r == null)
                    {
                        DocumentScanContext.RecordCandidate(ModuleType, "ChemicalSubscriptMissing");
                        DocumentScanContext.RecordSkip(ModuleType, "ChemicalSubscriptMissing", ScannerSkipReason.RangeMismatch);
                        continue;
                    }

                    bool hasUnformattedDigit = HasUnformattedFormulaDigit(r);

                    if (hasUnformattedDigit)
                    {
                        string recommendation = settings.UseNativeSubscript ? formula : SubscriptFormatter.ConvertToUnicodeSubSup(formula, "chemical");
                        string description = settings.UseNativeSubscript
                            ? $"当前正文角标规范为 Word 原生角标。检测到由合法元素符号组成的化学式“{formula}”，其中数字未进行下标格式化；建议保留 ASCII 文本并应用 Word 原生下标。"
                            : $"当前正文角标规范为 Unicode 角标字符。检测到由合法元素符号组成的化学式“{formula}”，建议直接改写为“{SubscriptFormatter.ConvertToUnicodeSubSup(formula, "chemical")}”。";
                        AddIssue(doc, text, issues, "ChemicalSubscriptMissing", match.Index, match.Length, formula, recommendation, description);
                    }
                    else
                    {
                        DocumentScanContext.RecordCandidate(ModuleType, "ChemicalSubscriptMissing");
                        DocumentScanContext.RecordSkip(ModuleType, "ChemicalSubscriptMissing", ScannerSkipReason.AlreadyCorrect);
                    }
                }
            }

            // --- 3. Academic Descriptive Subscripts Uprighting ---
            // Eg -> Eg, Eg (gap), Fermi level EF, Boltzmann constant kB, gate voltage Vg, Rs, IPL
            // Often typed as "Eg", "EF", "kB", "Vg", "Rs", "IPL" in raw manuscripts
            // We search for these patterns
            if (SettingsManager.IsRuleEnabled(ModuleType, "descriptive_subscript"))
            {
                foreach (Match match in DescriptiveSubscriptRegex.Matches(text))
                {
                    string word = match.Value;
                    if (!LooksLikeDescriptiveSubscriptContext(text, match.Index, word))
                    {
                        DocumentScanContext.RecordCandidate(ModuleType, "DescriptiveSubscript");
                        DocumentScanContext.RecordSkip(ModuleType, "DescriptiveSubscript", ScannerSkipReason.RuleFilter);
                        continue;
                    }

                    Range r = DocumentScanContext.CreateRangeFromTextSpan(doc, match.Index, match.Length, word);
                    if (r == null)
                    {
                        DocumentScanContext.RecordCandidate(ModuleType, "DescriptiveSubscript");
                        DocumentScanContext.RecordSkip(ModuleType, "DescriptiveSubscript", ScannerSkipReason.RangeMismatch);
                        continue;
                    }

                    // Check if it's already formatted correctly (first char italic, rest subscript + upright)
                    bool needFormatting = true;
                    if (r.Characters.Count > 1)
                    {
                        Range subPart = doc.Range(r.Characters[2].Start, r.End);
                        // If subpart is already Subscript and NOT Italic, then it's correct!
                        if (subPart.Font.Subscript == 1 && subPart.Font.Italic == 0)
                        {
                            needFormatting = false;
                        }
                    }

                    if (needFormatting)
                    {
                        AddIssue(doc, text, issues, "DescriptiveSubscript", match.Index, match.Length, word, word,
                            $"学术描述性角标“{word}”格式不规范。按照学术标准，表示物理意义的角标（如 gap 缩写 g，Bohr 缩写 B）必须设为【正体下标】（如 $E_\\mathrm{{g}}$、$k_\\mathrm{{B}}$），而非斜体或普通字符。");
                    }
                    else
                    {
                        DocumentScanContext.RecordCandidate(ModuleType, "DescriptiveSubscript");
                        DocumentScanContext.RecordSkip(ModuleType, "DescriptiveSubscript", ScannerSkipReason.AlreadyCorrect);
                    }
                }
            }

            // --- 4. LaTeX Style Inline上下标 Cleaner ---
            // Matches constructs like E_g, T_{c}, or x^2
            if (SettingsManager.IsRuleEnabled(ModuleType, "latex_inline_subsup"))
            {
                foreach (Match match in LatexInlineSubSupRegex.Matches(text))
                {
                    string origText = match.Value;
                    if (!LooksLikeInlineSubSup(origText, text, match.Index))
                    {
                        DocumentScanContext.RecordCandidate(ModuleType, "LaTeXStyleClean");
                        DocumentScanContext.RecordSkip(ModuleType, "LaTeXStyleClean", ScannerSkipReason.RuleFilter);
                        continue;
                    }

                    // Extract recommendation
                    string cleanText = origText.Replace("_", "").Replace("^", "").Replace("{", "").Replace("}", "");
                    AddIssue(doc, text, issues, "LaTeXStyleClean", match.Index, match.Length, origText, cleanText,
                        $"检测到 LaTeX 风格的行内上下标表达式“{origText}”。建议去除下划线或上标符，并在 Word 中自动应用原生的上下标排版属性。");
                }
            }

            return issues;
        }

        private static bool IsLikelyElementFormulaToken(string token)
        {
            if (string.IsNullOrEmpty(token) || !Regex.IsMatch(token, @"\d"))
            {
                return false;
            }

            int index = 0;
            int elementCount = 0;

            while (index < token.Length)
            {
                if (!char.IsUpper(token[index]))
                {
                    return false;
                }

                string symbol = null;
                if (index + 1 < token.Length && char.IsLower(token[index + 1]))
                {
                    string twoChar = token.Substring(index, 2);
                    if (ElementSymbols.Contains(twoChar))
                    {
                        symbol = twoChar;
                        index += 2;
                    }
                }

                if (symbol == null)
                {
                    string oneChar = token.Substring(index, 1);
                    if (!ElementSymbols.Contains(oneChar))
                    {
                        return false;
                    }

                    index += 1;
                }

                elementCount++;

                int digitCount = 0;
                while (index < token.Length && char.IsDigit(token[index]))
                {
                    digitCount++;
                    if (digitCount > 3)
                    {
                        return false;
                    }

                    index++;
                }
            }

            if (elementCount >= 2)
            {
                return true;
            }

            return CommonSingleElementFormulaTokens.Contains(token);
        }

        private static bool LooksLikeSupplementalReferenceToken(string text, int index, string token)
        {
            if (string.IsNullOrEmpty(text) || string.IsNullOrEmpty(token))
            {
                return false;
            }

            if (!Regex.IsMatch(token, @"^[A-Z]\d+[A-Za-z]?$"))
            {
                return false;
            }

            int start = Math.Max(0, index - 28);
            string before = text.Substring(start, index - start);
            return Regex.IsMatch(before, @"(?:Fig\.?|Figs\.?|Figure|Figures|Table|Tables|Scheme|Schemes|Supplementary|Movie)\s*$", RegexOptions.IgnoreCase);
        }

        private static bool HasUnformattedFormulaDigit(Range range)
        {
            for (int i = 1; i <= range.Characters.Count; i++)
            {
                Range c = range.Characters[i];
                string ct = c.Text;
                if (ct.Length == 1 && char.IsDigit(ct[0]) && c.Font.Subscript != 1)
                {
                    return true;
                }
            }

            return false;
        }

        private static bool LooksLikeDescriptiveSubscriptContext(string text, int index, string token)
        {
            if (string.IsNullOrEmpty(token))
            {
                return false;
            }

            if (token == "PL")
            {
                return LooksLikePhysicalContext(text, index) &&
                       Regex.IsMatch(GetWindow(text, index, 40), @"\b(intensity|peak|spectrum|spectra|mapping|signal|emission)\b", RegexOptions.IgnoreCase);
            }

            return LooksLikePhysicalContext(text, index);
        }

        private static bool LooksLikeInlineSubSup(string expression, string text, int index)
        {
            if (string.IsNullOrEmpty(expression))
            {
                return false;
            }

            if (expression.Contains("_"))
            {
                return Regex.IsMatch(expression, @"^[A-Za-z][A-Za-z0-9]{0,4}_\{?[A-Za-z0-9\-]{1,8}\}?$");
            }

            if (expression.Contains("^"))
            {
                if (Regex.IsMatch(expression, @"^(cm|m|s|K|eV|Hz|W|A|V|mol|J|Pa)\^\{?[-+]?\d{1,2}\}?$", RegexOptions.IgnoreCase))
                {
                    return true;
                }

                return LooksLikePhysicalContext(text, index);
            }

            return false;
        }

        private static bool LooksLikePhysicalContext(string text, int index)
        {
            string window = GetWindow(text, index, 52);
            return Regex.IsMatch(window, @"\b(gap|energy|Fermi|level|Boltzmann|gate|threshold|temperature|critical|resistance|sheet|current|drain|source|photocurrent|photoluminescence|PL|voltage|field|carrier|mobility|peak|spectrum|spectra|intensity)\b", RegexOptions.IgnoreCase);
        }

        private static string GetWindow(string text, int index, int radius)
        {
            int start = Math.Max(0, index - radius);
            int end = Math.Min(text.Length, index + radius + 1);
            return text.Substring(start, end - start);
        }

        private void AddIssue(Document doc, string text, List<IssueItem> issues, string subtype, int start, int length, string originalText, string recommendFix, string description)
        {
            IssueItem issue = IssueMatchFactory.Create(
                doc,
                text,
                ModuleType,
                subtype,
                start,
                length,
                originalText,
                recommendFix,
                description);
            if (issue != null)
            {
                issues.Add(issue);
            }
        }
    }
}

