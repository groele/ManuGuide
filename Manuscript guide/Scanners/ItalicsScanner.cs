using System;
using System.Collections.Generic;
using Microsoft.Office.Interop.Word;
using Manuscript_guide.Models;
using Manuscript_guide.Services;

namespace Manuscript_guide.Scanners
{
    public class ItalicsScanner : ISpecializedScanner
    {
        public string ModuleType => "ital";

        public List<IssueItem> Scan(Document doc)
        {
            List<IssueItem> issues = new List<IssueItem>();
            string documentText = DocumentScanContext.GetText(doc);

            // --- 1. Uprighting Italicized Math Functions and Latin Phrases ---
            // We search for italicized text in the document
            Range italicRange = doc.Content;
            italicRange.Find.ClearFormatting();
            italicRange.Find.Font.Italic = 1; // Search for Italicized text
            italicRange.Find.Text = "";
            italicRange.Find.Forward = true;
            italicRange.Find.Format = true;

            HashSet<string> latinPhrases = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
            {
                "et al.", "in situ", "in vivo", "in vitro", "vs.", "e.g.", "i.e.", "ca."
            };

            HashSet<string> mathFunctions = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
            {
                "sin", "cos", "tan", "exp", "log", "ln", "lim", "max", "min", "det"
            };

            bool checkLatinPhrases = SettingsManager.IsRuleEnabled(ModuleType, "latin_phrase_upright");
            bool checkMathFunctions = SettingsManager.IsRuleEnabled(ModuleType, "math_function_upright");
            bool checkExistingItalics = SettingsManager.IsRuleEnabled(ModuleType, "existing_italics_review");

            while (italicRange.Find.Execute())
            {
                string text = italicRange.Text?.Trim();
                if (string.IsNullOrEmpty(text)) continue;

                bool handledBySpecificRule = false;

                // Rule 1A: Italicized Latin expression
                if (checkLatinPhrases && latinPhrases.Contains(text))
                {
                    handledBySpecificRule = true;
                    string issueId = Guid.NewGuid().ToString();
                    Range targetRange = doc.Range(italicRange.Start, italicRange.End);
                    
                    CorrectionTracker.Instance.CreateBookmark(doc, issueId, targetRange, ModuleType);
                    ShadingManager.ApplyActiveShading(targetRange, ModuleType);

                    issues.Add(new IssueItem
                    {
                        IssueId = issueId,
                        Type = ModuleType,
                        Subtype = "LatinUpright",
                        Start = targetRange.Start,
                        End = targetRange.End,
                        OriginalText = text,
                        RecommendFix = text, // C# can accept correction and set Font.Italic = 0
                        Desc = $"拉丁短语“{text}”在现代学术写作中推荐使用标准的常规正体（Upright/Roman）。",
                        Context = $"... {text} ..."
                    });
                }
                // Rule 1B: Italicized standard math function name
                else if (checkMathFunctions && mathFunctions.Contains(text))
                {
                    handledBySpecificRule = true;
                    string issueId = Guid.NewGuid().ToString();
                    Range targetRange = doc.Range(italicRange.Start, italicRange.End);

                    CorrectionTracker.Instance.CreateBookmark(doc, issueId, targetRange, ModuleType);
                    ShadingManager.ApplyActiveShading(targetRange, ModuleType);

                    issues.Add(new IssueItem
                    {
                        IssueId = issueId,
                        Type = ModuleType,
                        Subtype = "MathFunctionUpright",
                        Start = targetRange.Start,
                        End = targetRange.End,
                        OriginalText = text,
                        RecommendFix = text, // Accept correction sets Font.Italic = 0
                        Desc = $"标准数学函数名称“{text}”必须使用正体排版，不能设为斜体。",
                        Context = $"... {text}(x) ..."
                    });
                }

                if (checkExistingItalics && !handledBySpecificRule)
                {
                    string issueId = Guid.NewGuid().ToString();
                    Range targetRange = doc.Range(italicRange.Start, italicRange.End);

                    CorrectionTracker.Instance.CreateBookmark(doc, issueId, targetRange, ModuleType);
                    ShadingManager.ApplyActiveShading(targetRange, ModuleType);

                    issues.Add(new IssueItem
                    {
                        IssueId = issueId,
                        Type = ModuleType,
                        Subtype = "ExistingItalicReview",
                        Start = targetRange.Start,
                        End = targetRange.End,
                        OriginalText = text,
                        RecommendFix = text,
                        Desc = $"检测到已存在的斜体文本“{text}”。此项仅用于复核正文中所有斜体位置，确认其是否确实应作为变量、物种名或期刊要求的斜体保留。",
                        Context = PunctuationScanner.GetContextSnippet(documentText, targetRange.Start, Math.Max(1, targetRange.End - targetRange.Start))
                    });
                }
            }

            // --- 2. Italicizing Single-Letter Physical Variables ---
            if (!SettingsManager.IsRuleEnabled(ModuleType, "variable_italic"))
            {
                return issues;
            }

            // We search for non-italicized single English characters that are stand-alone
            Range nonItalicRange = doc.Content;
            nonItalicRange.Find.ClearFormatting();
            nonItalicRange.Find.Font.Italic = 0; // Search for Non-Italic text
            nonItalicRange.Find.Text = "<[A-Za-z]>"; // Wildcard search for a single character word
            nonItalicRange.Find.MatchWildcards = true;
            nonItalicRange.Find.Forward = true;
            nonItalicRange.Find.Format = true;

            HashSet<string> EnglishSingleWords = new HashSet<string>(StringComparer.Ordinal)
            {
                "a", "A", "I" // Exclude "a", "A", and pronoun "I"
            };

            while (nonItalicRange.Find.Execute())
            {
                string text = nonItalicRange.Text?.Trim();
                if (string.IsNullOrEmpty(text) || EnglishSingleWords.Contains(text)) continue;

                // Check if it's likely a physical variable (e.g. x, y, T, E, P, V, k, L, t, etc.)
                // Let's exclude numbers, punctuation, or spaces
                if (text.Length == 1 && char.IsLetter(text[0]))
                {
                    // Double check surroundings to make sure it's stand-alone and not inside standard text
                    string contextText = documentText;
                    int startIdx = nonItalicRange.Start;
                    
                    // Simple contextual check: often variables are flanked by operators or descriptions like "where x is" or "var x"
                    bool isVariable = true;
                    if (startIdx > 0 && startIdx < contextText.Length)
                    {
                        // Check if it is capitalized I which is more likely a pronoun in English
                        if (text == "I")
                        {
                            // If it's "I", check if there are standard mathematical operations around it, or if it represents electric current/intensity
                            isVariable = false;
                            int scanStart = Math.Max(0, startIdx - 15);
                            int scanEnd = Math.Min(contextText.Length, startIdx + 15);
                            string contextWindow = contextText.Substring(scanStart, scanEnd - scanStart);
                            if (contextWindow.Contains("intensity") || contextWindow.Contains("current") || contextWindow.Contains("=") || contextWindow.Contains("+"))
                            {
                                isVariable = true;
                            }
                        }
                    }

                    if (isVariable)
                    {
                        string issueId = Guid.NewGuid().ToString();
                        Range targetRange = doc.Range(nonItalicRange.Start, nonItalicRange.End);

                        CorrectionTracker.Instance.CreateBookmark(doc, issueId, targetRange, ModuleType);
                        ShadingManager.ApplyActiveShading(targetRange, ModuleType);

                        issues.Add(new IssueItem
                        {
                            IssueId = issueId,
                            Type = ModuleType,
                            Subtype = "VariableItalic",
                            Start = targetRange.Start,
                            End = targetRange.End,
                            OriginalText = text,
                            RecommendFix = text, // Accept correction sets Font.Italic = 1
                            Desc = $"作为物理或数学公式变量的单字母“{text}”在学术规范中应设为斜体（Italic）排版。",
                            Context = PunctuationScanner.GetContextSnippet(contextText, startIdx, 1)
                        });
                    }
                }
            }

            return issues;
        }
    }
}

