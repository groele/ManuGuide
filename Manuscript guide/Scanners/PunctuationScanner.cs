using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using Microsoft.Office.Interop.Word;
using Manuscript_guide.Models;
using Manuscript_guide.Services;

namespace Manuscript_guide.Scanners
{
    public class PunctuationScanner : ISpecializedScanner
    {
        public string ModuleType => "punc";

        public List<IssueItem> Scan(Document doc)
        {
            List<IssueItem> issues = new List<IssueItem>();
            string text = DocumentScanContext.GetText(doc);
            if (string.IsNullOrEmpty(text)) return issues;

            // 1. Full-width punctuation checks inside English context
            // Search for full-width comma ，, full-width period 。 , full-width semicolon ；
            if (SettingsManager.IsRuleEnabled(ModuleType, "fullwidth_punctuation"))
            {
                MatchFullWidth(text, "，", ", ", "全角逗号", "在英文手稿中误用了中文全角逗号“，”", doc, issues);
                MatchFullWidth(text, "。", ". ", "全角句号", "在英文手稿中误用了中文全角句号“。”", doc, issues);
                MatchFullWidth(text, "；", "; ", "全角分号", "在英文手稿中误用了中文全角分号“；”", doc, issues);
                MatchFullWidth(text, "：", ": ", "全角冒号", "在英文手稿中误用了中文全角冒号“：”", doc, issues);
            }

            // 2. Missing half-width space after comma/semicolon/colon
            // Match any letter/digit followed by a half-width comma/semicolon/colon and immediately another letter/digit (without space)
            if (SettingsManager.IsRuleEnabled(ModuleType, "punctuation_spacing"))
            {
                Regex missingSpaceRegex = new Regex(@"([a-zA-Z0-9])([,;:?!])([a-zA-Z])");
                foreach (Match match in missingSpaceRegex.Matches(text))
                {
                    if (LooksLikeUrlOrEmailContext(text, match.Index))
                    {
                        continue;
                    }

                    string orig = match.Value;
                    string replacement = match.Groups[1].Value + match.Groups[2].Value + " " + match.Groups[3].Value;

                    IssueItem issue = IssueMatchFactory.Create(
                        doc,
                        text,
                        ModuleType,
                        "MissingSpaceAfterPunctuation",
                        match.Index,
                        match.Length,
                        orig,
                        replacement,
                        $"半角标点“{match.Groups[2].Value}”后缺失英文空格。");
                    if (issue != null)
                    {
                        issues.Add(issue);
                    }
                }
            }

            // 3. Greek Mu / Micro Symbol normalization
            if (SettingsManager.IsRuleEnabled(ModuleType, "greek_mu_encoding"))
            {
                // Micro Sign (U+00B5) to standard Greek Mu (U+03BC)
                // U+00B5 character: 'µ'
                // U+03BC character: 'μ'
                for (int i = 0; i < text.Length; i++)
                {
                    if (text[i] == '\u00B5')
                    {
                        IssueItem issue = IssueMatchFactory.Create(
                            doc,
                            text,
                            ModuleType,
                            "MicroSymbolEncoding",
                            i,
                            1,
                            "\u00B5",
                            "\u03BC",
                            "微符号编码不规范 (U+00B5)，建议自动规范化为学术出版标准的希腊字母 Mu (U+03BC) 以避免 XML 解析错误。");
                        if (issue != null)
                        {
                            issues.Add(issue);
                        }
                    }
                }
            }

            // 4. Word Inline Equation punctuation and spacing audit
            bool checkEquationSpacing = SettingsManager.IsRuleEnabled(ModuleType, "equation_spacing");
            bool checkEquationTerminalPunctuation = SettingsManager.IsRuleEnabled(ModuleType, "equation_terminal_punctuation");
            if (ProtectedRangeService.IsFormulaProtectionEnabled())
            {
                checkEquationSpacing = false;
                checkEquationTerminalPunctuation = false;
            }

            if (checkEquationSpacing || checkEquationTerminalPunctuation)
            {
                try
                {
                    foreach (OMath omath in doc.OMaths)
                    {
                        Range omathRange = omath.Range;
                        int start = omathRange.Start;
                        int end = omathRange.End;

                        // Check equation spacing before
                        if (checkEquationSpacing && start > 0)
                        {
                            Range preRange = doc.Range(start - 1, start);
                            string preText = preRange.Text;
                            if (!string.IsNullOrEmpty(preText) && Regex.IsMatch(preText, @"[a-zA-Z0-9]"))
                            {
                                IssueItem issue = IssueMatchFactory.Create(
                                    doc,
                                    text,
                                    ModuleType,
                                    "EquationSpacing",
                                    start - 1,
                                    1,
                                    preText,
                                    preText + " ",
                                    "公式域与前文英文单词之间缺失空格，建议添加半角空格以满足国际主流期刊排版规范。");
                                if (issue != null)
                                {
                                    issues.Add(issue);
                                }
                            }
                        }

                        // Check equation spacing & punctuation after
                        if (end < text.Length)
                        {
                            Range postRange = doc.Range(end, end + 1);
                            string postText = postRange.Text;

                            // Missing space after formula when followed by an alphanumeric character
                            if (checkEquationSpacing && !string.IsNullOrEmpty(postText) && Regex.IsMatch(postText, @"[a-zA-Z0-9]"))
                            {
                                IssueItem issue = IssueMatchFactory.Create(
                                    doc,
                                    text,
                                    ModuleType,
                                    "EquationSpacing",
                                    end,
                                    1,
                                    postText,
                                    " " + postText,
                                    "公式域与后文英文单词之间缺失空格，建议添加半角空格以满足国际主流期刊排版规范。");
                                if (issue != null)
                                {
                                    issues.Add(issue);
                                }
                            }

                            // Missing punctuation when equation paragraph ends
                            if (checkEquationTerminalPunctuation && !string.IsNullOrEmpty(postText) && (postText == "\r" || postText == "\n"))
                            {
                                Range lastCharRange = doc.Range(end - 1, end);
                                string lastChar = lastCharRange.Text;
                                if (!string.IsNullOrEmpty(lastChar) && !Regex.IsMatch(lastChar, @"[,.;:?!，。；：？！]"))
                                {
                                    IssueItem issue = IssueMatchFactory.Create(
                                        doc,
                                        text,
                                        ModuleType,
                                        "EquationPunctuation",
                                        end - 1,
                                        1,
                                        lastChar,
                                        lastChar + ".",
                                        "行内或独立公式作为句子结尾，末尾遗漏了标点符号（如句点“.”或逗号“,”），请检查并规范。");
                                    if (issue != null)
                                    {
                                        issues.Add(issue);
                                    }
                                }
                            }
                        }
                    }
                }
                catch (Exception)
                {
                    // Fail-safe for document OMath traversal
                }
            }

            return issues;
        }

        private void MatchFullWidth(string text, string target, string replacement, string subtype, string desc, Document doc, List<IssueItem> issues)
        {
            int index = text.IndexOf(target);
            while (index != -1)
            {
                if (!IsLikelyEnglishContext(text, index))
                {
                    index = text.IndexOf(target, index + 1);
                    continue;
                }

                IssueItem issue = IssueMatchFactory.Create(
                    doc,
                    text,
                    ModuleType,
                    subtype,
                    index,
                    target.Length,
                    target,
                    replacement,
                    desc);
                if (issue != null)
                {
                    issues.Add(issue);
                }

                index = text.IndexOf(target, index + 1);
            }
        }

        private static bool IsLikelyEnglishContext(string text, int index)
        {
            int start = Math.Max(0, index - 18);
            int end = Math.Min(text.Length, index + 19);
            for (int i = start; i < end; i++)
            {
                if ((text[i] >= 'A' && text[i] <= 'Z') || (text[i] >= 'a' && text[i] <= 'z'))
                {
                    return true;
                }
            }

            return false;
        }

        private static bool LooksLikeUrlOrEmailContext(string text, int index)
        {
            int start = Math.Max(0, index - 30);
            int end = Math.Min(text.Length, index + 31);
            string window = text.Substring(start, end - start);
            return Regex.IsMatch(window, @"(https?://|www\.|@\w|[\w.-]+\.[A-Za-z]{2,})", RegexOptions.IgnoreCase);
        }

        public static string GetContextSnippet(string text, int index, int length)
        {
            int start = Math.Max(0, index - 20);
            int end = Math.Min(text.Length, index + length + 20);
            string snippet = text.Substring(start, end - start);
            snippet = snippet.Replace("\r", " ").Replace("\n", " ");
            
            // Use plain text markers because the WPF TextBlock does not render HTML tags.
            int matchOffset = index - start;
            if (matchOffset >= 0 && matchOffset < snippet.Length)
            {
                string before = snippet.Substring(0, matchOffset);
                string matchText = snippet.Substring(matchOffset, Math.Min(length, snippet.Length - matchOffset));
                string after = snippet.Substring(Math.Min(matchOffset + length, snippet.Length));
                return $"{before}【{matchText}】{after}";
            }
            return snippet;
        }
    }
}

