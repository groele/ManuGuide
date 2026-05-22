using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using Microsoft.Office.Interop.Word;
using Manuscript_guide.Models;
using Manuscript_guide.Services;

namespace Manuscript_guide.Scanners
{
    public class DashSpacingScanner : ISpecializedScanner
    {
        public string ModuleType => "dash";

        public List<IssueItem> Scan(Document doc)
        {
            List<IssueItem> issues = new List<IssueItem>();
            string text = DocumentScanContext.GetText(doc);
            if (string.IsNullOrEmpty(text)) return issues;

            // 1. Spacing between Chinese characters and English/Numbers
            // CJK range: [\u4e00-\u9fa5]
            
            // Rule 1A: Chinese character followed immediately by English/Number without space
            if (SettingsManager.IsRuleEnabled(ModuleType, "cjk_latin_spacing"))
            {
                Regex cjkToLatin = new Regex(@"([\u4e00-\u9fa5])([a-zA-Z0-9])");
                foreach (Match match in cjkToLatin.Matches(text))
                {
                    string matched = match.Value;
                    if (ShouldSkipMixedSpacing(text, match.Index, match.Length))
                    {
                        DocumentScanContext.RecordCandidate(ModuleType, "CJKLatinSpacing");
                        DocumentScanContext.RecordSkip(ModuleType, "CJKLatinSpacing", ScannerSkipReason.RuleFilter);
                        continue;
                    }

                    IssueItem issue = IssueMatchFactory.Create(
                        doc,
                        text,
                        ModuleType,
                        "CJKLatinSpacing",
                        match.Index,
                        match.Length,
                        matched,
                        match.Groups[1].Value + " " + match.Groups[2].Value,
                        $"中文与英文/数字“{matched}”之间应留有一个半角空格以优化排版视觉体验。");
                    if (issue != null)
                    {
                        issues.Add(issue);
                    }
                }
            }

            // Rule 1B: English/Number followed immediately by Chinese character without space
            if (SettingsManager.IsRuleEnabled(ModuleType, "latin_cjk_spacing"))
            {
                Regex latinToCjk = new Regex(@"([a-zA-Z0-9])([\u4e00-\u9fa5])");
                foreach (Match match in latinToCjk.Matches(text))
                {
                    string matched = match.Value;
                    if (ShouldSkipMixedSpacing(text, match.Index, match.Length))
                    {
                        DocumentScanContext.RecordCandidate(ModuleType, "LatinCJKSpacing");
                        DocumentScanContext.RecordSkip(ModuleType, "LatinCJKSpacing", ScannerSkipReason.RuleFilter);
                        continue;
                    }

                    IssueItem issue = IssueMatchFactory.Create(
                        doc,
                        text,
                        ModuleType,
                        "LatinCJKSpacing",
                        match.Index,
                        match.Length,
                        matched,
                        match.Groups[1].Value + " " + match.Groups[2].Value,
                        $"英文/数字与中文“{matched}”之间应留有一个半角空格以优化排版视觉体验。");
                    if (issue != null)
                    {
                        issues.Add(issue);
                    }
                }
            }

            // 2. Hyphen to En-dash for Academic Compound Words
            // Scan for common academic noun pairs connected by hyphens
            string[] compoundPatterns = new[]
            {
                "exciton-polariton", "spin-orbit", "metal-semiconductor", "gate-doping",
                "electron-phonon", "density-functional", "metal-insulator", "structure-property",
                "donor-acceptor", "hole-electron", "vapor-liquid-solid", "temperature-dependent"
            };

            if (SettingsManager.IsRuleEnabled(ModuleType, "academic_hyphen_endash"))
            {
                foreach (var term in compoundPatterns)
                {
                    Regex termRegex = new Regex(@"\b" + term + @"\b", RegexOptions.IgnoreCase);
                    foreach (Match match in termRegex.Matches(text))
                    {
                        string matched = match.Value;
                        string replacement = matched.Replace("-", "–"); // En-dash (U+2013)

                        IssueItem issue = IssueMatchFactory.Create(
                            doc,
                            text,
                            ModuleType,
                            "HyphenToEnDash",
                            match.Index,
                            match.Length,
                            matched,
                            replacement,
                            $"表示并列或关联特征的学术复合词组“{matched}”，中间应使用标准的 En-dash（–，U+2013）而非普通连字符（-）。");
                        if (issue != null)
                        {
                            issues.Add(issue);
                        }
                    }
                }
            }

            return issues;
        }

        private static bool ShouldSkipMixedSpacing(string text, int index, int length)
        {
            int start = Math.Max(0, index - 20);
            int end = Math.Min(text.Length, index + length + 20);
            string window = text.Substring(start, end - start);
            return Regex.IsMatch(window, @"(https?://|www\.|doi\s*:|10\.\d{4,9}/|[A-Z][a-z]?\d+[A-Za-z]*|[A-Za-z]*\d+[A-Z][a-z]?|sample\s*[A-Z]?\d+|Fig\.?\s*\d+|Table\s*\d+)", RegexOptions.IgnoreCase);
        }
    }
}

