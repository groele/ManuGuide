using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using Microsoft.Office.Interop.Word;
using Manuscript_guide.Models;
using Manuscript_guide.Services;

namespace Manuscript_guide.Scanners
{
    public class CapitalizationScanner : ISpecializedScanner
    {
        public string ModuleType => "cap";

        public List<IssueItem> Scan(Document doc)
        {
            List<IssueItem> issues = new List<IssueItem>();
            string text = DocumentScanContext.GetText(doc);
            if (string.IsNullOrEmpty(text)) return issues;

            PluginSettings settings = SettingsManager.Current;

            // Load Whitelists
            HashSet<string> casingWhitelist = ParseWhitelist(settings.WhitelistCasing);
            HashSet<string> acronymWhitelist = ParseWhitelist(settings.WhitelistAcronyms);

            // --- 1. Acronym First-Definition Analysis ---
            if (SettingsManager.IsRuleEnabled(ModuleType, "acronym_definition"))
            {
                AnalyzeAbbreviationsFlow(doc, text, acronymWhitelist, issues);
            }

            // --- 2. Casing Consistency Analysis ---
            if (SettingsManager.IsRuleEnabled(ModuleType, "casing_consistency"))
            {
                AnalyzeCasingConsistencyFlow(doc, text, casingWhitelist, settings, issues);
            }

            // --- 3. Cross-Reference Capitalization ---
            if (SettingsManager.IsRuleEnabled(ModuleType, "crossref_capitalization"))
            {
                AnalyzeCrossRefCapitalizationFlow(doc, text, issues);
            }

            return issues;
        }

        private HashSet<string> ParseWhitelist(string csv)
        {
            var set = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            if (string.IsNullOrEmpty(csv)) return set;
            foreach (var item in csv.Split(','))
            {
                string trimmed = item.Trim();
                if (trimmed.Length > 0)
                {
                    set.Add(trimmed);
                }
            }
            return set;
        }

        private void AnalyzeAbbreviationsFlow(Document doc, string text, HashSet<string> acronymWhitelist, List<IssueItem> issues)
        {
            // Acronyms definitions: e.g. "photoluminescence (PL)" or "PL (photoluminescence)"
            Dictionary<string, int> definitionPositions = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            Dictionary<string, string> definitionTexts = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

            // 1. Scan for definitions: Full name (Acronym)
            // Group 1: Full name, Group 2: Acronym
            Regex defRegex = new Regex(@"\b([a-zA-Z\s\-]{3,50}) \(([A-Z]{2,6})\)\b");
            foreach (Match match in defRegex.Matches(text))
            {
                string full = match.Groups[1].Value.Trim();
                string acronym = match.Groups[2].Value;

                // Simple validation to ensure the full text words match acronym letters roughly
                if (IsLikelyAcronymDefinition(full, acronym))
                {
                    if (!definitionPositions.ContainsKey(acronym))
                    {
                        definitionPositions[acronym] = match.Index;
                        definitionTexts[acronym] = full;
                    }
                    else
                    {
                        // Redundant definition: second definition occurrence in text
                        int idx = match.Index;
                        AddIssue(doc, text, issues, "RedundantDefinition", idx, match.Length, match.Value, acronym,
                            $"缩写词 '{acronym}' 已在前文定义过，此处重复定义。");
                    }
                }
            }

            // 2. Scan all acronyms: \b[A-Z]{2,6}\b
            Regex acronymRegex = new Regex(@"\b[A-Z]{2,6}\b");
            HashSet<string> commonShortWords = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
            {
                "THE", "AND", "FOR", "BUT", "NOT", "YES", "WHO", "YOU", "ITS", "HAS", "HAD", "WAS", "ARE"
            };

            // Keep track of the first usage of each acronym
            Dictionary<string, int> firstUsagePositions = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            
            foreach (Match match in acronymRegex.Matches(text))
            {
                string acronym = match.Value;
                if (commonShortWords.Contains(acronym) || acronymWhitelist.Contains(acronym)) continue;

                // Filter out standard Roman numerals (I, II, III, IV, VI, VII, VIII, IX, X, etc.)
                if (IsRomanNumeral(acronym)) continue;

                if (!firstUsagePositions.ContainsKey(acronym))
                {
                    firstUsagePositions[acronym] = match.Index;
                }
            }

            // Check acronyms for missing definition or definition lag
            foreach (var kvp in firstUsagePositions)
            {
                string acronym = kvp.Key;
                int firstUseIndex = kvp.Value;

                if (!definitionPositions.ContainsKey(acronym))
                {
                    // Acronym has no definition anywhere in the document
                    AddIssue(doc, text, issues, "UndefinedAcronym", firstUseIndex, acronym.Length, acronym, acronym + " (Supplement Definition)",
                        $"缩写词 '{acronym}' 首次出现，但全文中均未检测到对应的全称定义。");
                }
                else
                {
                    int defIndex = definitionPositions[acronym];
                    // Check if definition is lagging significantly (e.g. acronym used before definition, or defined >1000 chars late)
                    if (firstUseIndex < defIndex)
                    {
                        // Definition lag! Acronym used before the definition
                        string fullDefText = definitionTexts[acronym];

                        AddIssue(doc, text, issues, "DefinitionLag", firstUseIndex, acronym.Length, acronym, $"{fullDefText} ({acronym})",
                            $"缩写词 '{acronym}' 出现滞后定义。在第 {firstUseIndex} 字符处首次使用，但全称定义 '{fullDefText}' 却延迟在第 {defIndex} 字符处才给出。建议将全称定义迁移至首次出现处。");
                    }
                }
            }
        }

        private void AnalyzeCasingConsistencyFlow(Document doc, string text, HashSet<string> casingWhitelist, PluginSettings settings, List<IssueItem> issues)
        {
            // Sentence and paragraph start check
            // We find paragraph starts and sentence starts.
            Regex sentenceStartRegex = new Regex(@"(^|\r|\n|[.!?]\s+)([a-zA-Z]+)");
            HashSet<int> sentenceStartIndices = new HashSet<int>();
            foreach (Match match in sentenceStartRegex.Matches(text))
            {
                Group wordGroup = match.Groups[2];
                sentenceStartIndices.Add(wordGroup.Index);
            }

            // Word frequency maps
            // LowercaseKey -> Dictionary of Spelling -> List of Indices
            var casingMap = new Dictionary<string, Dictionary<string, List<int>>>(StringComparer.OrdinalIgnoreCase);

            Regex wordRegex = new Regex(@"\b[a-zA-Z]{3,20}\b");
            foreach (Match match in wordRegex.Matches(text))
            {
                string word = match.Value;

                // Skip if this index belongs to a sentence or paragraph start (casing is syntactically uppercase)
                if (sentenceStartIndices.Contains(match.Index)) continue;

                // Skip if whitelisted
                if (casingWhitelist.Contains(word)) continue;

                // Skip if variable lock is active and the word is in the locked physical variables/units set
                if (SettingsManager.IsRuleEnabled(ModuleType, "physical_variable_casing_lock") && LockedVariables.Contains(word)) continue;

                string lower = word.ToLower();

                if (!casingMap.ContainsKey(lower))
                {
                    casingMap[lower] = new Dictionary<string, List<int>>();
                }

                if (!casingMap[lower].ContainsKey(word))
                {
                    casingMap[lower][word] = new List<int>();
                }

                casingMap[lower][word].Add(match.Index);
            }

            // Majority-decision check
            foreach (var kvp in casingMap)
            {
                var variants = kvp.Value;
                if (variants.Count > 1)
                {
                    // Find the majority variation
                    string majorityVariant = null;
                    int maxCount = -1;
                    int totalCount = 0;

                    foreach (var varKvp in variants)
                    {
                        int c = varKvp.Value.Count;
                        totalCount += c;
                        if (c > maxCount)
                        {
                            maxCount = c;
                            majorityVariant = varKvp.Key;
                        }
                    }

                    // Check majority ratio threshold (e.g. 70%)
                    double ratio = (double)maxCount / totalCount * 100;
                    if (ratio >= settings.CasingRatioThreshold)
                    {
                        foreach (var varKvp in variants)
                        {
                            string spelling = varKvp.Key;
                            if (spelling == majorityVariant) continue;

                            // All minority occurrences are flagged as inconsistent
                            foreach (int charIndex in varKvp.Value)
                            {
                                AddIssue(doc, text, issues, "CasingInconsistency", charIndex, spelling.Length, spelling, majorityVariant,
                                    $"大小写全局不一致。全篇主要使用 '{majorityVariant}' （{maxCount}次，占比 {ratio:F1}%），而此处拼写为 '{spelling}'。");
                            }
                        }
                    }
                }
            }
        }

        private bool IsLikelyAcronymDefinition(string fullName, string acronym)
        {
            string cleanFull = Regex.Replace(fullName, @"[^a-zA-Z\s\-]", "");
            string[] words = cleanFull.Split(new[] { ' ', '-' }, StringSplitOptions.RemoveEmptyEntries);

            if (words.Length < acronym.Length) return false;

            // Check if the letters of acronym roughly match the first letters of words
            int acronymIndex = 0;
            for (int i = 0; i < words.Length; i++)
            {
                if (acronymIndex < acronym.Length && 
                    words[i].Length > 0 && 
                    char.ToLower(words[i][0]) == char.ToLower(acronym[acronymIndex]))
                {
                    acronymIndex++;
                }
            }

            return acronymIndex == acronym.Length;
        }

        private bool IsRomanNumeral(string text)
        {
            return Regex.IsMatch(text, @"^(I|II|III|IV|V|VI|VII|VIII|IX|X|XI|XII|XIII|XIV|XV)$", RegexOptions.IgnoreCase);
        }

        private static readonly HashSet<string> LockedVariables = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "e", "E", "f", "F", "p", "P", "v", "V", "t", "T", "c", "C", "q", "Q", "k", "K", "a", "A", "d", "D", "h", "H", "m", "M", "n", "N", "r", "R", "s", "S", "x", "X", "y", "Y", "z", "Z",
            "eV", "keV", "MeV", "GeV", "pH", "kB", "kb", "Hz", "kHz", "MHz", "GHz", "THZ", "THz", "dB", "Raman", "afm", "AFM", "raman"
        };

        private void AnalyzeCrossRefCapitalizationFlow(Document doc, string text, List<IssueItem> issues)
        {
            // Match pattern like "figure 1", "fig. 2", "table 3", "eq 4", "equation 5", "ref 6", "reference 7", etc.
            Regex crossRefRegex = new Regex(@"\b(figure|table|fig|eq|equation|ref|reference)(\.?\s+)(\d+|\b[A-Z]\b)", RegexOptions.IgnoreCase);
            foreach (Match match in crossRefRegex.Matches(text))
            {
                string word = match.Groups[1].Value;
                string spacing = match.Groups[2].Value;
                string number = match.Groups[3].Value;

                // If first letter is lowercase
                if (char.IsLower(word[0]))
                {
                    string capitalizedWord = char.ToUpper(word[0]) + word.Substring(1);
                    string orig = match.Value;
                    string replacement = capitalizedWord + spacing + number;

                    AddIssue(doc, text, issues, "CrossRefCapitalization", match.Index, match.Length, orig, replacement,
                        $"学术交叉引用首字母应大写。此处将 '{orig}' 识别为对特定图表/文献的引用，建议规范化为大写开头的 '{replacement}'。");
                }
            }
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

