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
            @"(?<![A-Za-z0-9_])(?:[A-Za-z0-9_]|[₀₁₂₃₄₅₆₇₈₉ₓ]|[⁰¹²³⁴⁵⁶⁷⁸⁹⁺⁻⁼])*(?:[₀₁₂₃₄₅₆₇₈₉ₓ]|[⁰¹²³⁴⁵⁶⁷⁸⁹⁺⁻⁼])(?:[A-Za-z0-9_]|[₀₁₂₃₄₅₆₇₈₉ₓ]|[⁰¹²³⁴⁵⁶⁷⁸⁹⁺⁻⁼])*(?![A-Za-z0-9_])",
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

            var context = DocumentScanContext.Current;
            var snapshot = context?.Snapshot;

            // --- 1. Unicode Subscript/Superscript Detection ---
            if (SettingsManager.IsRuleEnabled(ModuleType, "unicode_subsup_to_native"))
            {
                foreach (Match match in UnicodeSubSupRegex.Matches(text))
                {
                    string origText = match.Value;
                    string formatType = "chemical";
                    if (Regex.IsMatch(origText, @"[⁰¹²³⁴⁵⁶⁷⁸⁹⁺⁻⁼]"))
                    {
                        formatType = "unit";
                    }

                    string recommend = settings.UseNativeSubscript
                        ? SubscriptFormatter.ConvertToAsciiText(origText, formatType)
                        : origText;

                    string description = $"当前正文角标规范为 Word 原生角标。检测到 Unicode 编码的上下标字符“{origText}”，建议转换为标准 ASCII 字符并应用 Word 原生上下标格式。";

                    AddIssue(doc, text, "UnicodeSubscript", match.Index, match.Length, origText, recommend,
                        description, issues);
                }
            }

            // --- 2. Element-driven Chemical Formula Digits Missing Subscripts ---
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

                    bool hasUnformattedDigit = false;
                    if (!settings.UseNativeSubscript)
                    {
                        // In Unicode subscript mode, any ASCII digits in the formula are considered incorrect and need conversion
                        hasUnformattedDigit = true;
                    }
                    else if (snapshot != null)
                    {
                        hasUnformattedDigit = HasUnformattedFormulaDigit(snapshot, match.Index, match.Length);
                    }
                    else
                    {
                        Range r = DocumentScanContext.CreateRangeFromTextSpan(doc, match.Index, match.Length, formula);
                        if (r == null)
                        {
                            DocumentScanContext.RecordCandidate(ModuleType, "ChemicalSubscriptMissing");
                            DocumentScanContext.RecordSkip(ModuleType, "ChemicalSubscriptMissing", ScannerSkipReason.RangeMismatch);
                            continue;
                        }
                        hasUnformattedDigit = HasUnformattedFormulaDigitFallback(r);
                    }

                    if (hasUnformattedDigit)
                    {
                        string recommendation = settings.UseNativeSubscript ? formula : SubscriptFormatter.ConvertToUnicodeSubSup(formula, "chemical");
                        string description = settings.UseNativeSubscript
                            ? $"当前正文角标规范为 Word 原生角标。检测到由合法元素符号组成的化学式“{formula}”，其中数字未进行下标格式化；建议保留 ASCII 文本并应用 Word 原生下标。"
                            : $"当前正文角标规范为 Unicode 角标字符。检测到由合法元素符号组成的化学式“{formula}”，建议直接改写为“{SubscriptFormatter.ConvertToUnicodeSubSup(formula, "chemical")}”。";
                        AddIssue(doc, text, "ChemicalSubscriptMissing", match.Index, match.Length, formula, recommendation, description, issues);
                    }
                    else
                    {
                        DocumentScanContext.RecordCandidate(ModuleType, "ChemicalSubscriptMissing");
                        DocumentScanContext.RecordSkip(ModuleType, "ChemicalSubscriptMissing", ScannerSkipReason.AlreadyCorrect);
                    }
                }
            }

            // --- 3. Academic Descriptive Subscripts Uprighting ---
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

                    bool needFormatting = true;
                    if (snapshot != null)
                    {
                        int start = match.Index;
                        int length = match.Length;
                        if (length > 1)
                        {
                            bool isSubscript = snapshot.Subscripts.IsSpanFullyCovered(start + 1, start + length);
                            bool isItalic = snapshot.Italics.HasAnyFormatting(start + 1, start + length);
                            if (isSubscript && !isItalic)
                            {
                                needFormatting = false;
                            }
                        }
                    }
                    else
                    {
                        Range r = DocumentScanContext.CreateRangeFromTextSpan(doc, match.Index, match.Length, word);
                        if (r == null)
                        {
                            DocumentScanContext.RecordCandidate(ModuleType, "DescriptiveSubscript");
                            DocumentScanContext.RecordSkip(ModuleType, "DescriptiveSubscript", ScannerSkipReason.RangeMismatch);
                            continue;
                        }

                        if (word.Length > 1)
                        {
                            Range subPart = doc.Range(r.Start + 1, r.End);
                            if (subPart.Font.Subscript == 1 && subPart.Font.Italic == 0)
                            {
                                needFormatting = false;
                            }
                        }
                    }

                    if (needFormatting)
                    {
                        AddIssue(doc, text, "DescriptiveSubscript", match.Index, match.Length, word, word,
                            $"学术描述性角标“{word}”格式不规范。按照学术标准，表示物理意义的角标（如 gap 缩写 g，Bohr 缩写 B）必须设为【正体下标】（如 $E_\\mathrm{{g}}$、$k_\\mathrm{{B}}$），而非斜体或普通字符。", issues);
                    }
                    else
                    {
                        DocumentScanContext.RecordCandidate(ModuleType, "DescriptiveSubscript");
                        DocumentScanContext.RecordSkip(ModuleType, "DescriptiveSubscript", ScannerSkipReason.AlreadyCorrect);
                    }
                }
            }

            // --- 4. LaTeX Style Inline上下标 Cleaner ---
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

                    string cleanText = origText.Replace("_", "").Replace("^", "").Replace("{", "").Replace("}", "");
                    AddIssue(doc, text, "LaTeXStyleClean", match.Index, match.Length, origText, cleanText,
                        $"检测到 LaTeX 风格的行内上下标表达式“{origText}”。建议去除下划线或上标符，并在 Word 中自动应用原生的上下标排版属性。", issues);
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

        private static bool HasUnformattedFormulaDigit(DocumentSnapshot snapshot, int start, int length)
        {
            for (int i = 0; i < length; i++)
            {
                char c = snapshot.FullText[start + i];
                if (char.IsDigit(c))
                {
                    if (!snapshot.Subscripts.IsSpanFullyCovered(start + i, start + i + 1))
                    {
                        return true;
                    }
                }
            }
            return false;
        }

        private static bool HasUnformattedFormulaDigitFallback(Range range)
        {
            try
            {
                Range findRange = range.Duplicate;
                Find find = findRange.Find;
                find.ClearFormatting();
                find.Text = "^#"; // Any digit
                find.Forward = true;
                find.Wrap = WdFindWrap.wdFindStop;

                while (find.Execute())
                {
                    if (findRange.Start >= range.End) break;
                    if (findRange.Font.Subscript != 1)
                    {
                        return true;
                    }
                    findRange.Collapse(WdCollapseDirection.wdCollapseEnd);
                    findRange.End = range.End;
                }
            }
            catch
            {
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

        private void AddIssue(Document doc, string text, string subtype, int start, int length, string originalText, string recommendFix, string description, List<IssueItem> issues)
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
