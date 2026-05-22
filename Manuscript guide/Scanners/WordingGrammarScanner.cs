using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using Microsoft.Office.Interop.Word;
using Manuscript_guide.Models;
using Manuscript_guide.Services;

namespace Manuscript_guide.Scanners
{
    public class WordingGrammarScanner : ISpecializedScanner
    {
        public string ModuleType => "word";

        private sealed class PhraseRule
        {
            public PhraseRule(string settingKey, string subtype, string pattern, string replacement, string description)
            {
                SettingKey = settingKey;
                Subtype = subtype;
                Pattern = pattern;
                Replacement = replacement;
                Description = description;
            }

            public string SettingKey { get; private set; }
            public string Subtype { get; private set; }
            public string Pattern { get; private set; }
            public string Replacement { get; private set; }
            public string Description { get; private set; }
        }

        private static readonly List<PhraseRule> PhraseRules = new List<PhraseRule>
        {
            new PhraseRule(
                "redundant_in_order_to",
                "RedundantInOrderTo",
                @"\bin\s+order\s+to\b",
                "to",
                "冗余的学术连词。建议精简为简洁明了的“to”以提高句子的可读性。"),
            new PhraseRule(
                "redundant_as_is_shown",
                "RedundantAsIsShown",
                @"\bas\s+is\s+shown\s+in\b",
                "as shown in",
                "啰嗦的学术句式。在学术写作中推荐使用更精炼的“as shown in Fig.”等结构。"),
            new PhraseRule(
                "chinglish_companion_syntax",
                "ChinglishCompanionSyntax",
                @"\bwith\s+the\s+increase\s+of\b",
                "with increasing (or: as temperature increases)",
                "典型的中式伴随状语。在学术英语中，推荐使用“with increasing [variable]”或“as [variable] increases”等动态连贯表达。")
        };

        public List<IssueItem> Scan(Document doc)
        {
            List<IssueItem> issues = new List<IssueItem>();
            string text = DocumentScanContext.GetText(doc);
            if (string.IsNullOrEmpty(text)) return issues;

            // 1. Data Plural subject-verb agreement
            // Match "the data is" or "data is" (case insensitive)
            if (SettingsManager.IsRuleEnabled(ModuleType, "data_plural_agreement"))
            {
                Regex dataIsRegex = new Regex(@"\b(data)\s+(is|has|was)\b", RegexOptions.IgnoreCase);
                foreach (Match match in dataIsRegex.Matches(text))
                {
                    string origText = match.Value;
                    string verb = match.Groups[2].Value.ToLower();
                    string recommendVerb = "are";
                    if (verb == "has") recommendVerb = "have";
                    if (verb == "was") recommendVerb = "were";

                    string recommend = match.Groups[1].Value + " " + recommendVerb;

                    IssueItem issue = IssueMatchFactory.Create(
                        doc,
                        text,
                        ModuleType,
                        "DataPluralAgreement",
                        match.Index,
                        match.Length,
                        origText,
                        recommend,
                        $"名词“data”在科技论文中属于复数名词，其搭配的单数谓语动词“{match.Groups[2].Value}”应规范改写为复数形式“{recommendVerb}”。");
                    if (issue != null)
                    {
                        issues.Add(issue);
                    }
                }
            }

            foreach (PhraseRule rule in PhraseRules)
            {
                if (SettingsManager.IsRuleEnabled(ModuleType, rule.SettingKey))
                {
                    MatchWordyPhrase(text, rule.Pattern, rule.Replacement, rule.Subtype, rule.Description, doc, issues);
                }
            }

            return issues;
        }

        private void MatchWordyPhrase(string text, string pattern, string replacement, string subtype, string desc, Document doc, List<IssueItem> issues)
        {
            Regex r = new Regex(pattern, RegexOptions.IgnoreCase);
            foreach (Match match in r.Matches(text))
            {
                string origText = match.Value;
                IssueItem issue = IssueMatchFactory.Create(
                    doc,
                    text,
                    ModuleType,
                    subtype,
                    match.Index,
                    match.Length,
                    origText,
                    replacement,
                    desc);
                if (issue != null)
                {
                    issues.Add(issue);
                }
            }
        }
    }
}

