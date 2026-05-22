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

        public List<IssueItem> Scan(Document doc)
        {
            List<IssueItem> issues = new List<IssueItem>();
            string text = DocumentScanContext.GetText(doc);
            if (string.IsNullOrEmpty(text)) return issues;
            PluginSettings settings = SettingsManager.Current;

            // --- 1. Unicode Subscript/Superscript Detection ---
            // Subscripts: ₀₁₂₃₄₅₆₇₈₉ₓ
            // Superscripts: ⁰¹²³⁴⁵⁶⁷⁸⁹⁺⁻⁼
            if (settings.UseNativeSubscript && SettingsManager.IsRuleEnabled(ModuleType, "unicode_subsup_to_native"))
            {
                Regex unicodeSubSupRegex = new Regex(@"([a-zA-Z]+[₀₁₂₃₄₅₆₇₈₉ₓ]+)+|([a-zA-Z]+[⁰¹²³⁴⁵⁶⁷⁸⁹⁺⁻⁼]+)+");
                foreach (Match match in unicodeSubSupRegex.Matches(text))
                {
                    string origText = match.Value;
                    Range r = doc.Range(match.Index, match.Index + match.Length);
                    string issueId = Guid.NewGuid().ToString();

                    CorrectionTracker.Instance.CreateBookmark(doc, issueId, r, ModuleType);
                    ShadingManager.ApplyActiveShading(r, ModuleType);

                    // Determine sub-type
                    string formatType = "chemical";
                    if (Regex.IsMatch(origText, @"[⁰¹²³⁴⁵⁶⁷⁸⁹⁺⁻⁼]"))
                    {
                        formatType = "unit";
                    }

                    // Predict clean translation recommendation
                    string recommend = SubscriptFormatter.ConvertToAsciiText(origText, formatType);

                    issues.Add(new IssueItem
                    {
                        IssueId = issueId,
                        Type = ModuleType,
                        Subtype = "UnicodeSubscript",
                        Start = match.Index,
                        End = match.Index + match.Length,
                        OriginalText = origText,
                        RecommendFix = recommend, // Under the hood, custom correction applies SubscriptFormatter
                        Desc = $"当前正文角标规范为 Word 原生角标。检测到 Unicode 编码的上下标字符“{origText}”，建议转换为标准 ASCII 字符并应用 Word 原生上下标格式。",
                        Context = PunctuationScanner.GetContextSnippet(text, match.Index, match.Length)
                    });
                }
            }

            // --- 2. Element-driven Chemical Formula Digits Missing Subscripts ---
            // Based on Element subscript conversion.js: only accept tokens that can be
            // fully segmented into valid chemical element symbols. This catches WSe2,
            // MoS2, Bi2O2Se, CuInP2S6, and avoids ordinary prose words.
            if (SettingsManager.IsRuleEnabled(ModuleType, "element_formula_subscript"))
            {
                Regex elementFormulaRegex = new Regex(@"\b(?:[A-Z][a-z]?\d{0,3}){2,}\b");
                foreach (Match match in elementFormulaRegex.Matches(text))
                {
                    string formula = match.Value;
                    if (!IsLikelyElementFormulaToken(formula)) continue;

                    Range r = doc.Range(match.Index, match.Index + match.Length);
                    bool hasUnformattedDigit = HasUnformattedFormulaDigit(r);

                    if (hasUnformattedDigit)
                    {
                        string issueId = Guid.NewGuid().ToString();
                        CorrectionTracker.Instance.CreateBookmark(doc, issueId, r, ModuleType);
                        ShadingManager.ApplyActiveShading(r, ModuleType);

                        issues.Add(new IssueItem
                        {
                            IssueId = issueId,
                            Type = ModuleType,
                            Subtype = "ChemicalSubscriptMissing",
                            Start = match.Index,
                            End = match.Index + match.Length,
                            OriginalText = formula,
                            RecommendFix = settings.UseNativeSubscript ? formula : SubscriptFormatter.ConvertToUnicodeSubSup(formula, "chemical"),
                            Desc = settings.UseNativeSubscript
                                ? $"当前正文角标规范为 Word 原生角标。检测到由合法元素符号组成的化学式“{formula}”，其中数字未进行下标格式化；建议保留 ASCII 文本并应用 Word 原生下标。"
                                : $"当前正文角标规范为 Unicode 角标字符。检测到由合法元素符号组成的化学式“{formula}”，建议直接改写为“{SubscriptFormatter.ConvertToUnicodeSubSup(formula, "chemical")}”。",
                            Context = PunctuationScanner.GetContextSnippet(text, match.Index, match.Length)
                        });
                    }
                }
            }

            // --- 3. Academic Descriptive Subscripts Uprighting ---
            // Eg -> Eg, Eg (gap), Fermi level EF, Boltzmann constant kB, gate voltage Vg, Rs, IPL
            // Often typed as "Eg", "EF", "kB", "Vg", "Rs", "IPL" in raw manuscripts
            // We search for these patterns
            if (SettingsManager.IsRuleEnabled(ModuleType, "descriptive_subscript"))
            {
                Regex descSubscriptRegex = new Regex(@"\b(E[gF]|k[B]|V[g]|R[s]|I[P][L]|T[c])\b");
                foreach (Match match in descSubscriptRegex.Matches(text))
                {
                    string word = match.Value;
                    Range r = doc.Range(match.Index, match.Index + match.Length);

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
                        string issueId = Guid.NewGuid().ToString();
                        CorrectionTracker.Instance.CreateBookmark(doc, issueId, r, ModuleType);
                        ShadingManager.ApplyActiveShading(r, ModuleType);

                        string mainVar = word.Substring(0, 1);
                        string subLabel = word.Substring(1);

                        issues.Add(new IssueItem
                        {
                            IssueId = issueId,
                            Type = ModuleType,
                            Subtype = "DescriptiveSubscript",
                            Start = match.Index,
                            End = match.Index + match.Length,
                            OriginalText = word,
                            RecommendFix = word, // Corrects via SubscriptFormatter "desc" (mainVar italic, subLabel subscript+upright)
                            Desc = $"学术描述性角标“{word}”格式不规范。按照学术标准，表示物理意义的角标（如 gap 缩写 g，Bohr 缩写 B）必须设为【正体下标】（如 $E_\\mathrm{{g}}$、$k_\\mathrm{{B}}$），而非斜体或普通字符。",
                            Context = PunctuationScanner.GetContextSnippet(text, match.Index, match.Length)
                        });
                    }
                }
            }

            // --- 4. LaTeX Style Inline上下标 Cleaner ---
            // Matches constructs like E_g, T_{c}, or x^2
            if (SettingsManager.IsRuleEnabled(ModuleType, "latex_inline_subsup"))
            {
                Regex latexRegex = new Regex(@"([a-zA-Z0-9]+)_\{?([a-zA-Z0-9\-]+)\}?|([a-zA-Z0-9]+)\^\{?([a-zA-Z0-9\-]+)\}?");
                foreach (Match match in latexRegex.Matches(text))
                {
                    string origText = match.Value;
                    Range r = doc.Range(match.Index, match.Index + match.Length);
                    string issueId = Guid.NewGuid().ToString();

                    CorrectionTracker.Instance.CreateBookmark(doc, issueId, r, ModuleType);
                    ShadingManager.ApplyActiveShading(r, ModuleType);

                    // Extract recommendation
                    string cleanText = origText.Replace("_", "").Replace("^", "").Replace("{", "").Replace("}", "");
                    string formatType = origText.Contains("_") ? "chemical" : "unit";

                    // If it is a descriptive subscript, treat as desc
                    if (origText.StartsWith("E_") || origText.StartsWith("k_") || origText.StartsWith("V_"))
                    {
                        formatType = "desc";
                    }

                    issues.Add(new IssueItem
                    {
                        IssueId = issueId,
                        Type = ModuleType,
                        Subtype = "LaTeXStyleClean",
                        Start = match.Index,
                        End = match.Index + match.Length,
                        OriginalText = origText,
                        RecommendFix = cleanText,
                        Desc = $"检测到 LaTeX 风格的行内上下标表达式“{origText}”。建议去除下划线或上标符，并在 Word 中自动应用原生的上下标排版属性。",
                        Context = PunctuationScanner.GetContextSnippet(text, match.Index, match.Length)
                    });
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

            return elementCount >= 2;
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
    }
}

