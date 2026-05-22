using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using Microsoft.Office.Interop.Word;
using Manuscript_guide.Models;
using Manuscript_guide.Services;

namespace Manuscript_guide.Scanners
{
    public class DataValueScanner : ISpecializedScanner
    {
        public string ModuleType => "data";

        public List<IssueItem> Scan(Document doc)
        {
            List<IssueItem> issues = new List<IssueItem>();
            string text = DocumentScanContext.GetText(doc);
            if (string.IsNullOrEmpty(text)) return issues;

            // 1. Value-Unit spacing: e.g. 300K -> 300 K, 15nm -> 15 nm
            // Match digits followed immediately by unit (K, nm, mA, GHz, cm, eV, V, Hz, Pa, s, min, h, K, °C, % is excluded)
            if (SettingsManager.IsRuleEnabled(ModuleType, "value_unit_spacing"))
            {
                Regex unitSpacingRegex = new Regex(@"\b(\d+(\.\d+)?)\s*(K|nm|nm|mA|GHz|cm|eV|V|Hz|Pa|m|s|min|kg|mW|MW|dB)\b");
                foreach (Match match in unitSpacingRegex.Matches(text))
                {
                    string origText = match.Value;
                    string value = match.Groups[1].Value;
                    string unit = match.Groups[3].Value;

                    // If there's no space between value and unit, flag it
                    if (!origText.Contains(" ") && !origText.Contains("\u00A0")) // check space or non-breaking space
                    {
                        Range r = doc.Range(match.Index, match.Index + match.Length);
                        string issueId = Guid.NewGuid().ToString();

                        CorrectionTracker.Instance.CreateBookmark(doc, issueId, r, ModuleType);
                        ShadingManager.ApplyActiveShading(r, ModuleType);

                        issues.Add(new IssueItem
                        {
                            IssueId = issueId,
                            Type = ModuleType,
                            Subtype = "ValueUnitSpacing",
                            Start = match.Index,
                            End = match.Index + match.Length,
                            OriginalText = origText,
                            RecommendFix = value + " " + unit,
                            Desc = $"数值与物理单位“{origText}”之间缺失空格，建议补充一个半角空格。",
                            Context = PunctuationScanner.GetContextSnippet(text, match.Index, match.Length)
                        });
                    }
                }
            }

            // 2. Scientific Notation multiplier: e.g. 5.2x10^12 -> 5.2 × 10^12
            if (SettingsManager.IsRuleEnabled(ModuleType, "scientific_notation"))
            {
                Regex scientificRegex = new Regex(@"(\d+(\.\d+)?)\s*[x*]\s*10\s*([\^⁺⁻⁰¹²³⁴⁵⁶⁷⁸⁹]+)");
                foreach (Match match in scientificRegex.Matches(text))
                {
                    string origText = match.Value;
                    string baseVal = match.Groups[1].Value;
                    string expMarker = match.Groups[3].Value;

                    Range r = doc.Range(match.Index, match.Index + match.Length);
                    string issueId = Guid.NewGuid().ToString();

                    CorrectionTracker.Instance.CreateBookmark(doc, issueId, r, ModuleType);
                    ShadingManager.ApplyActiveShading(r, ModuleType);

                    issues.Add(new IssueItem
                    {
                        IssueId = issueId,
                        Type = ModuleType,
                        Subtype = "ScientificNotation",
                        Start = match.Index,
                        End = match.Index + match.Length,
                        OriginalText = origText,
                        RecommendFix = $"{baseVal} × 10{expMarker}", // Standard multiplication sign (×, U+00D7)
                        Desc = $"科学计数法乘号误用了字符“x”或“*”，建议替换为标准学术乘号“×”。",
                        Context = PunctuationScanner.GetContextSnippet(text, match.Index, match.Length)
                    });
                }
            }

            // 3. Range representations: "from 10-300 K" -> "from 10 to 300 K"
            if (SettingsManager.IsRuleEnabled(ModuleType, "from_range_expression"))
            {
                Regex fromRangeRegex = new Regex(@"\bfrom\s+(\d+)\s*[-–—]\s*(\d+)\b", RegexOptions.IgnoreCase);
                foreach (Match match in fromRangeRegex.Matches(text))
                {
                    string origText = match.Value;
                    string val1 = match.Groups[1].Value;
                    string val2 = match.Groups[2].Value;

                    Range r = doc.Range(match.Index, match.Index + match.Length);
                    string issueId = Guid.NewGuid().ToString();

                    CorrectionTracker.Instance.CreateBookmark(doc, issueId, r, ModuleType);
                    ShadingManager.ApplyActiveShading(r, ModuleType);

                    issues.Add(new IssueItem
                    {
                        IssueId = issueId,
                        Type = ModuleType,
                        Subtype = "RangeExpression",
                        Start = match.Index,
                        End = match.Index + match.Length,
                        OriginalText = origText,
                        RecommendFix = $"from {val1} to {val2}",
                        Desc = $"在英文学术论文中，使用“from”引导范围时应使用“to”连接数值，不能使用横线（如 from A to B）。",
                        Context = PunctuationScanner.GetContextSnippet(text, match.Index, match.Length)
                    });
                }
            }

            // 4. Standalone numerical range without "from": e.g. 10 - 300 K or 10-300K -> 10–300 K using En-dash without spaces
            if (SettingsManager.IsRuleEnabled(ModuleType, "range_endash"))
            {
                Regex rangeHyphenRegex = new Regex(@"\b(\d+)\s*[-—]\s*(\d+)\b");
                foreach (Match match in rangeHyphenRegex.Matches(text))
                {
                    // Skip if preceded by "from" (handled by previous regex)
                    int startIndex = match.Index;
                    if (startIndex >= 5)
                    {
                        string prev = text.Substring(startIndex - 5, 5).ToLower();
                        if (prev.Contains("from")) continue;
                    }

                    string origText = match.Value;
                    string val1 = match.Groups[1].Value;
                    string val2 = match.Groups[2].Value;

                    // If there's spaces around the hyphen, or it is a simple hyphen instead of En-dash, flag it
                    if (origText.Contains(" ") || origText.Contains("-"))
                    {
                        Range r = doc.Range(match.Index, match.Index + match.Length);
                        string issueId = Guid.NewGuid().ToString();

                        CorrectionTracker.Instance.CreateBookmark(doc, issueId, r, ModuleType);
                        ShadingManager.ApplyActiveShading(r, ModuleType);

                        issues.Add(new IssueItem
                        {
                            IssueId = issueId,
                            Type = ModuleType,
                            Subtype = "RangeEnDash",
                            Start = match.Index,
                            End = match.Index + match.Length,
                            OriginalText = origText,
                            RecommendFix = $"{val1}–{val2}", // En-dash, no spaces
                            Desc = $"纯数值范围连接应使用紧凑的 En-dash（–，U+2013）连接线，且两侧不留空格。",
                            Context = PunctuationScanner.GetContextSnippet(text, match.Index, match.Length)
                        });
                    }
                }
            }

            // 5. Mathematical sign corrections:
            // +/- to ± (U+00B1)
            if (SettingsManager.IsRuleEnabled(ModuleType, "plus_minus_sign"))
            {
                int index = text.IndexOf("+/-");
                while (index != -1)
                {
                    Range r = doc.Range(index, index + 3);
                    string issueId = Guid.NewGuid().ToString();

                    CorrectionTracker.Instance.CreateBookmark(doc, issueId, r, ModuleType);
                    ShadingManager.ApplyActiveShading(r, ModuleType);

                    issues.Add(new IssueItem
                    {
                        IssueId = issueId,
                        Type = ModuleType,
                        Subtype = "PlusMinusSign",
                        Start = index,
                        End = index + 3,
                        OriginalText = "+/-",
                        RecommendFix = "±",
                        Desc = $"检测到 ASCII 拼写的“+/-”符号，建议规范替换为标准的物理正负号“±”（U+00B1）。",
                        Context = PunctuationScanner.GetContextSnippet(text, index, 3)
                    });

                    index = text.IndexOf("+/-", index + 1);
                }
            }

            // Standard minus sign U+2212 inside negative number values: e.g. -30 K -> −30 K
            if (SettingsManager.IsRuleEnabled(ModuleType, "negative_minus_sign"))
            {
                Regex negativeNumberRegex = new Regex(@"(?<=\s|^)-(\d+(\.\d+)?)\b");
                foreach (Match match in negativeNumberRegex.Matches(text))
                {
                    string origText = match.Value;
                    string val = match.Groups[1].Value;

                    Range r = doc.Range(match.Index, match.Index + match.Length);
                    string issueId = Guid.NewGuid().ToString();

                    CorrectionTracker.Instance.CreateBookmark(doc, issueId, r, ModuleType);
                    ShadingManager.ApplyActiveShading(r, ModuleType);

                    issues.Add(new IssueItem
                    {
                        IssueId = issueId,
                        Type = ModuleType,
                        Subtype = "NegativeMinusSign",
                        Start = match.Index,
                        End = match.Index + match.Length,
                        OriginalText = origText,
                        RecommendFix = "−" + val, // typographic minus (U+2212)
                        Desc = $"负号连接符误用了普通连字符（-），建议规范替换为排版专用的减号“−”（U+2212）。",
                        Context = PunctuationScanner.GetContextSnippet(text, match.Index, match.Length)
                    });
                }
            }

            return issues;
        }
    }
}

