using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using Microsoft.Office.Interop.Word;
using Manuscript_guide.Models;
using Manuscript_guide.Services;

namespace Manuscript_guide.Scanners
{
    public class ItalicsScanner : ISpecializedScanner
    {
        public string ModuleType => "ital";

        private const int MaxVariableIssues = 300;
        private const int MaxExistingItalicIssues = 150;

        private static readonly Regex LatinPhraseRegex = new Regex(
            @"\b(et\s+al\.|in\s+situ|in\s+vivo|in\s+vitro|vs\.|e\.g\.|i\.e\.|ca\.)\b",
            RegexOptions.IgnoreCase | RegexOptions.Compiled);

        private static readonly Regex MathFunctionRegex = new Regex(
            @"\b(sin|cos|tan|exp|log|ln|lim|max|min|det)\b(?=\s*[\(\[])",
            RegexOptions.IgnoreCase | RegexOptions.Compiled);

        private static readonly Regex SingleLetterRegex = new Regex(
            @"\b([A-Za-z])\b",
            RegexOptions.Compiled);

        private static readonly HashSet<string> EnglishSingleWords = new HashSet<string>(StringComparer.Ordinal)
        {
            "a", "A", "I"
        };

        private static readonly HashSet<char> CandidateVariables = new HashSet<char>
        {
            'a', 'b', 'c', 'd', 'e', 'f', 'g', 'h', 'i', 'j', 'k', 'l', 'm', 'n',
            'p', 'q', 'r', 's', 't', 'u', 'v', 'w', 'x', 'y', 'z',
            'A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N',
            'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z'
        };

        public List<IssueItem> Scan(Document doc)
        {
            List<IssueItem> issues = new List<IssueItem>();
            string documentText = DocumentScanContext.GetText(doc);
            if (string.IsNullOrEmpty(documentText))
            {
                return issues;
            }

            if (SettingsManager.IsRuleEnabled(ModuleType, "latin_phrase_upright"))
            {
                AddUprightIssues(doc, documentText, LatinPhraseRegex, "LatinUpright",
                    "拉丁短语“{0}”在现代学术写作中推荐使用标准的常规正体（Upright/Roman）。", issues);
            }

            if (SettingsManager.IsRuleEnabled(ModuleType, "math_function_upright"))
            {
                AddUprightIssues(doc, documentText, MathFunctionRegex, "MathFunctionUpright",
                    "标准数学函数名称“{0}”必须使用正体排版，不能设为斜体。", issues);
            }

            if (SettingsManager.IsRuleEnabled(ModuleType, "existing_italics_review"))
            {
                AddExistingItalicReviewIssues(doc, documentText, issues);
            }

            if (SettingsManager.IsRuleEnabled(ModuleType, "variable_italic"))
            {
                AddVariableItalicIssues(doc, documentText, issues);
            }

            return issues;
        }

        private void AddUprightIssues(Document doc, string documentText, Regex regex, string subtype, string description, List<IssueItem> issues)
        {
            foreach (Match match in regex.Matches(documentText))
            {
                Range targetRange = doc.Range(match.Index, match.Index + match.Length);
                if (ProtectedRangeService.IsRangeProtected(targetRange) || targetRange.Font.Italic != 1)
                {
                    continue;
                }

                AddIssue(doc, documentText, targetRange, subtype, match.Value, match.Value,
                    string.Format(description, match.Value), issues);
            }
        }

        private void AddExistingItalicReviewIssues(Document doc, string documentText, List<IssueItem> issues)
        {
            Range italicRange = doc.Content;
            italicRange.Find.ClearFormatting();
            italicRange.Find.Font.Italic = 1;
            italicRange.Find.Text = "";
            italicRange.Find.Forward = true;
            italicRange.Find.Format = true;
            italicRange.Find.Wrap = WdFindWrap.wdFindStop;

            int guard = 0;
            while (italicRange.Find.Execute())
            {
                if (++guard > MaxExistingItalicIssues)
                {
                    break;
                }

                string text = italicRange.Text == null ? string.Empty : italicRange.Text.Trim();
                if (!string.IsNullOrEmpty(text) && !ProtectedRangeService.IsRangeProtected(italicRange))
                {
                    AddIssue(doc, documentText, doc.Range(italicRange.Start, italicRange.End), "ExistingItalicReview",
                        text, text,
                        $"检测到已存在的斜体文本“{text}”。此项仅用于复核正文中所有斜体位置，确认其是否确实应作为变量、物种名或期刊要求的斜体保留。",
                        issues);
                }

                int nextStart = Math.Max(italicRange.End, italicRange.Start + 1);
                if (nextStart >= doc.Content.End)
                {
                    break;
                }

                italicRange.SetRange(nextStart, doc.Content.End);
                italicRange.Find.ClearFormatting();
                italicRange.Find.Font.Italic = 1;
                italicRange.Find.Text = "";
                italicRange.Find.Forward = true;
                italicRange.Find.Format = true;
                italicRange.Find.Wrap = WdFindWrap.wdFindStop;
            }
        }

        private void AddVariableItalicIssues(Document doc, string documentText, List<IssueItem> issues)
        {
            int count = 0;
            foreach (Match match in SingleLetterRegex.Matches(documentText))
            {
                string text = match.Groups[1].Value;
                if (EnglishSingleWords.Contains(text) || !CandidateVariables.Contains(text[0]))
                {
                    continue;
                }

                if (!LooksLikeVariableContext(documentText, match.Index, text[0]))
                {
                    continue;
                }

                Range targetRange = doc.Range(match.Index, match.Index + match.Length);
                if (ProtectedRangeService.IsRangeProtected(targetRange) || targetRange.Font.Italic == 1)
                {
                    continue;
                }

                AddIssue(doc, documentText, targetRange, "VariableItalic", text, text,
                    $"作为物理或数学公式变量的单字母“{text}”在学术规范中应设为斜体（Italic）排版。",
                    issues);

                count++;
                if (count >= MaxVariableIssues)
                {
                    break;
                }
            }
        }

        private static bool LooksLikeVariableContext(string text, int index, char letter)
        {
            int start = Math.Max(0, index - 36);
            int end = Math.Min(text.Length, index + 37);
            string window = text.Substring(start, end - start);

            if (Regex.IsMatch(window, @"(=|≈|~|∝|<|>|≤|≥|\+|-|×|/|\(|\)|\[|\]|\{|\})"))
            {
                return true;
            }

            if (Regex.IsMatch(window, @"\b(where|variable|parameter|constant|field|temperature|voltage|current|energy|time|axis|coordinate|denoted|defined|plotted|fitted|measured|calculated|value|values)\b", RegexOptions.IgnoreCase))
            {
                return true;
            }

            if ((letter == 'x' || letter == 'y' || letter == 'z') &&
                Regex.IsMatch(window, @"\b(axis|coordinate|direction|component|plot|versus|vs\.?)\b", RegexOptions.IgnoreCase))
            {
                return true;
            }

            return false;
        }

        private void AddIssue(Document doc, string documentText, Range targetRange, string subtype, string originalText, string recommendFix, string desc, List<IssueItem> issues)
        {
            string issueId = Guid.NewGuid().ToString();
            string bookmarkName = CorrectionTracker.Instance.CreateBookmark(doc, issueId, targetRange, ModuleType);
            if (string.IsNullOrEmpty(bookmarkName))
            {
                return;
            }

            ShadingManager.ApplyActiveShading(targetRange, ModuleType);
            issues.Add(new IssueItem
            {
                IssueId = issueId,
                Type = ModuleType,
                Subtype = subtype,
                Start = targetRange.Start,
                End = targetRange.End,
                OriginalText = originalText,
                RecommendFix = recommendFix,
                Desc = desc,
                Context = PunctuationScanner.GetContextSnippet(documentText, targetRange.Start, Math.Max(1, targetRange.End - targetRange.Start))
            });
        }
    }
}
