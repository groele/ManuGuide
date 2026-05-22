using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using Microsoft.Office.Interop.Word;
using Manuscript_guide.Models;
using Manuscript_guide.Services;

namespace Manuscript_guide.Scanners
{
    public sealed class ScannerMatch
    {
        public ScannerMatch(Match match)
        {
            Match = match;
            Index = match.Index;
            Length = match.Length;
            Value = match.Value;
        }

        public Match Match { get; private set; }
        public int Index { get; private set; }
        public int Length { get; private set; }
        public string Value { get; private set; }
    }

    public sealed class ScannerIssueSpec
    {
        public string Subtype { get; set; }
        public int Start { get; set; }
        public int Length { get; set; }
        public string OriginalText { get; set; }
        public string RecommendFix { get; set; }
        public string Description { get; set; }
    }

    public static class ScannerRuleRunner
    {
        public static void AddRegexIssues(
            Document doc,
            string text,
            string moduleType,
            string subtype,
            Regex regex,
            List<IssueItem> issues,
            Func<ScannerMatch, ScannerIssueSpec> createSpec)
        {
            if (regex == null || issues == null || createSpec == null || string.IsNullOrEmpty(text))
            {
                return;
            }

            foreach (Match match in regex.Matches(text))
            {
                ScannerIssueSpec spec = createSpec(new ScannerMatch(match));
                if (spec == null)
                {
                    DocumentScanContext.RecordCandidate(moduleType, subtype);
                    DocumentScanContext.RecordSkip(moduleType, subtype, ScannerSkipReason.RuleFilter);
                    continue;
                }

                string effectiveSubtype = string.IsNullOrEmpty(spec.Subtype) ? subtype : spec.Subtype;
                IssueItem issue = IssueMatchFactory.Create(
                    doc,
                    text,
                    moduleType,
                    effectiveSubtype,
                    spec.Start,
                    spec.Length,
                    spec.OriginalText,
                    spec.RecommendFix,
                    spec.Description);
                if (issue != null)
                {
                    issues.Add(issue);
                }
            }
        }
    }
}
