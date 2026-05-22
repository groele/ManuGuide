using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using Word = Microsoft.Office.Interop.Word;
using Manuscript_guide.Models;

namespace Manuscript_guide.Services
{
    public interface ICorrectionActionProvider
    {
        bool CanHandle(IssueItem issue);
        Action<Word.Range> CreateAction(IssueItem issue, Word.Document doc);
    }

    public static class CorrectionActionResolver
    {
        private static readonly List<ICorrectionActionProvider> Providers = new List<ICorrectionActionProvider>
        {
            new SubscriptCorrectionActionProvider(),
            new ItalicsCorrectionActionProvider()
        };

        public static void Register(ICorrectionActionProvider provider)
        {
            if (provider != null)
            {
                Providers.Add(provider);
            }
        }

        public static Action<Word.Range> Resolve(IssueItem issue, Word.Document doc)
        {
            if (issue == null)
            {
                return null;
            }

            foreach (ICorrectionActionProvider provider in Providers)
            {
                if (provider.CanHandle(issue))
                {
                    return provider.CreateAction(issue, doc);
                }
            }

            return null;
        }
    }

    internal sealed class SubscriptCorrectionActionProvider : ICorrectionActionProvider
    {
        public bool CanHandle(IssueItem issue)
        {
            return issue != null && issue.Type == "sub";
        }

        public Action<Word.Range> CreateAction(IssueItem issue, Word.Document doc)
        {
            string formatType = ResolveFormatType(issue);
            return (newRange) =>
            {
                SubscriptFormatter.ApplyCorrection(
                    doc,
                    newRange,
                    issue.RecommendFix,
                    formatType,
                    SettingsManager.Current.UseNativeSubscript
                );
            };
        }

        private string ResolveFormatType(IssueItem issue)
        {
            string originalText = issue.OriginalText ?? string.Empty;
            string subtype = issue.Subtype ?? string.Empty;

            if (subtype == "UnicodeSubscript" && Regex.IsMatch(originalText, @"[⁰¹²³⁴⁵⁶⁷⁸⁹⁺⁻⁼]"))
            {
                return "unit";
            }

            if (subtype == "LaTeXStyleClean" && originalText.Contains("^"))
            {
                return "unit";
            }

            if (subtype == "DescriptiveSubscript" ||
                originalText.StartsWith("E_") ||
                originalText.StartsWith("k_") ||
                originalText.StartsWith("V_"))
            {
                return "desc";
            }

            return "chemical";
        }
    }

    internal sealed class ItalicsCorrectionActionProvider : ICorrectionActionProvider
    {
        public bool CanHandle(IssueItem issue)
        {
            return issue != null && issue.Type == "ital";
        }

        public Action<Word.Range> CreateAction(IssueItem issue, Word.Document doc)
        {
            if (issue.Subtype == "LatinUpright" || issue.Subtype == "MathFunctionUpright")
            {
                return (newRange) => { newRange.Font.Italic = 0; };
            }

            if (issue.Subtype == "VariableItalic")
            {
                return (newRange) => { newRange.Font.Italic = 1; };
            }

            return null;
        }
    }
}
