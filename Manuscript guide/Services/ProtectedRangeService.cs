using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using Microsoft.Office.Interop.Word;
using Manuscript_guide.Models;

namespace Manuscript_guide.Services
{
    public sealed class ProtectedTextRange
    {
        public int Start { get; set; }
        public int End { get; set; }
        public string Source { get; set; }
    }

    public static class ProtectedRangeService
    {
        private static readonly object CacheLock = new object();
        private static readonly Dictionary<string, ProtectedRangeCache> Cache = new Dictionary<string, ProtectedRangeCache>();

        private static readonly string[] CitationFieldMarkers =
        {
            "ADDIN ZOTERO_ITEM",
            "CSL_CITATION",
            "CSL_BIBLIOGRAPHY",
            "ADDIN EN.CITE",
            "ADDIN EN.REFLIST",
            "ADDIN MENDELEY CITATION",
            "ADDIN MENDELEY BIBLIOGRAPHY",
            "MENDELEY CITATION",
            "MENDELEY BIBLIOGRAPHY",
            "REFWORKS",
            "CITAVI",
            "NOTEEXPRESS"
        };

        private static readonly string[] FormulaFieldMarkers =
        {
            "EQUATION.DSMT",
            "MATHTYPE",
            "MTEF",
            "OMATH"
        };

        private static readonly Regex ReferencesHeadingRegex = new Regex(
            @"^\s*(references|reference|bibliography|works\s+cited|literature\s+cited|参考文献|参考资料)\s*[:：]?\s*$",
            RegexOptions.IgnoreCase | RegexOptions.Compiled);

        private static readonly Regex LatexBlockRegex = new Regex(
            @"\$\$[\s\S]*?\$\$|\\\[[\s\S]*?\\\]|\\begin\{(?:equation|equation\*|align|align\*|gather|gather\*|multline|multline\*)\}[\s\S]*?\\end\{(?:equation|equation\*|align|align\*|gather|gather\*|multline|multline\*)\}",
            RegexOptions.Compiled);

        private static readonly Regex LatexInlineRegex = new Regex(
            @"(?<!\\)\$[^\r\n$]{1,500}?(?<!\\)\$|\\\([^\r\n]{1,500}?\\\)",
            RegexOptions.Compiled);

        public static bool IsProtectionEnabled()
        {
            return IsReferenceProtectionEnabled() || IsFormulaProtectionEnabled();
        }

        public static bool IsReferenceProtectionEnabled()
        {
            return SettingsManager.IsRuleEnabled("global", "filter_references_and_citations");
        }

        public static bool IsFormulaProtectionEnabled()
        {
            return SettingsManager.IsRuleEnabled("global", "skip_latex_formula_regions");
        }

        public static bool ShouldMarkSkippedFormulaRegions()
        {
            return IsFormulaProtectionEnabled() &&
                   SettingsManager.IsRuleEnabled("global", "mark_skipped_formula_regions");
        }

        public static bool IsRangeProtected(Range range)
        {
            if (range == null)
            {
                return false;
            }

            try
            {
                return IsRangeProtected(range.Document, range.Start, range.End);
            }
            catch
            {
                return false;
            }
        }

        public static bool IsRangeProtected(Document doc, int start, int end)
        {
            if (!IsProtectionEnabled() || doc == null)
            {
                return false;
            }

            List<ProtectedTextRange> ranges = DocumentScanContext.GetProtectedRanges(doc) ?? GetProtectedRanges(doc);
            return IntersectsAny(ranges, start, end);
        }

        public static List<IssueItem> FilterIssues(Document doc, List<IssueItem> issues)
        {
            if (!IsProtectionEnabled() || issues == null || issues.Count == 0)
            {
                return issues ?? new List<IssueItem>();
            }

            List<ProtectedTextRange> protectedRanges = DocumentScanContext.GetProtectedRanges(doc) ?? GetProtectedRanges(doc);
            if (protectedRanges.Count == 0)
            {
                return issues;
            }

            List<IssueItem> filtered = new List<IssueItem>();
            foreach (IssueItem issue in issues)
            {
                if (issue == null)
                {
                    continue;
                }

                int start = DocumentScanContext.TextOffsetToDocumentPosition(doc, issue.Start);
                int end = DocumentScanContext.TextOffsetToDocumentPosition(doc, issue.End);
                if (!IntersectsAny(protectedRanges, start, end))
                {
                    filtered.Add(issue);
                }
            }

            return filtered;
        }

        public static List<ProtectedTextRange> GetProtectedRanges(Document doc)
        {
            if (doc == null)
            {
                return new List<ProtectedTextRange>();
            }

            List<ProtectedTextRange> activeRanges = DocumentScanContext.GetProtectedRanges(doc);
            if (activeRanges != null)
            {
                return activeRanges;
            }

            string documentText = doc.Content.Text ?? string.Empty;
            return GetProtectedRanges(doc, documentText);
        }

        public static List<ProtectedTextRange> GetProtectedRanges(Document doc, string documentText)
        {
            if (doc == null)
            {
                return new List<ProtectedTextRange>();
            }

            string key = GetCacheKey(doc);
            int contentEnd = SafeContentEnd(doc);
            int fieldCount = SafeFieldCount(doc);
            int omathCount = SafeOMathCount(doc);
            int inlineShapeCount = SafeInlineShapeCount(doc);
            int textHash = string.IsNullOrEmpty(documentText) ? 0 : documentText.GetHashCode();

            lock (CacheLock)
            {
                if (Cache.ContainsKey(key))
                {
                    ProtectedRangeCache cached = Cache[key];
                    if (cached.ContentEnd == contentEnd &&
                        cached.FieldCount == fieldCount &&
                        cached.OMathCount == omathCount &&
                        cached.InlineShapeCount == inlineShapeCount &&
                        cached.TextHash == textHash)
                    {
                        return cached.Ranges;
                    }
                }
            }

            List<ProtectedTextRange> ranges = BuildProtectedRanges(doc, documentText);
            MergeRanges(ranges);

            lock (CacheLock)
            {
                Cache[key] = new ProtectedRangeCache
                {
                    ContentEnd = contentEnd,
                    FieldCount = fieldCount,
                    OMathCount = omathCount,
                    InlineShapeCount = inlineShapeCount,
                    TextHash = textHash,
                    Ranges = ranges
                };
            }

            return ranges;
        }

        public static string MaskProtectedText(string text, List<ProtectedTextRange> ranges)
        {
            return MaskProtectedText(text, ranges, 0);
        }

        public static string MaskProtectedText(string text, List<ProtectedTextRange> ranges, int contentStart)
        {
            if (string.IsNullOrEmpty(text) || ranges == null || ranges.Count == 0)
            {
                return text ?? string.Empty;
            }

            char[] chars = text.ToCharArray();
            foreach (ProtectedTextRange range in ranges)
            {
                int start = Math.Max(0, Math.Min(chars.Length, range.Start - contentStart));
                int end = Math.Max(start, Math.Min(chars.Length, range.End - contentStart));
                for (int i = start; i < end; i++)
                {
                    if (chars[i] != '\r' && chars[i] != '\n')
                    {
                        chars[i] = ' ';
                    }
                }
            }

            return new string(chars);
        }

        public static void RefreshProtectedMarkers(Document doc)
        {
            if (doc == null)
            {
                return;
            }

            ShadingManager.ClearModuleShading(doc, "protected");
            if (!ShouldMarkSkippedFormulaRegions())
            {
                return;
            }

            List<ProtectedTextRange> ranges = DocumentScanContext.GetProtectedRanges(doc) ?? GetProtectedRanges(doc);
            foreach (ProtectedTextRange protectedRange in ranges)
            {
                if (!IsFormulaSource(protectedRange.Source))
                {
                    continue;
                }

                try
                {
                    Range range = doc.Range(protectedRange.Start, protectedRange.End);
                    string issueId = Guid.NewGuid().ToString();
                    CorrectionTracker.Instance.CreateBookmark(doc, issueId, range, "protected", true);
                    ShadingManager.ApplyActiveShading(range, "protected", true);
                }
                catch
                {
                }
            }
        }

        private static List<ProtectedTextRange> BuildProtectedRanges(Document doc, string documentText)
        {
            List<ProtectedTextRange> ranges = new List<ProtectedTextRange>();
            if (IsReferenceProtectionEnabled())
            {
                AddCitationFieldRanges(doc, ranges);
            }
            if (IsFormulaProtectionEnabled())
            {
                AddFormulaFieldRanges(doc, ranges);
                AddWordEquationRanges(doc, ranges);
                AddMathTypeRanges(doc, ranges);
                AddLatexTextRanges(doc, documentText, ranges);
            }
            if (IsReferenceProtectionEnabled())
            {
                AddReferencesSectionRange(doc, ranges);
            }
            return ranges;
        }

        private static void AddCitationFieldRanges(Document doc, List<ProtectedTextRange> ranges)
        {
            try
            {
                foreach (Field field in doc.Fields)
                {
                    string codeText = GetFieldCodeText(field);
                    if (!ContainsAnyMarker(codeText, CitationFieldMarkers))
                    {
                        continue;
                    }

                    AddFieldCombinedRange(ranges, field, "citation-field");
                    AddRange(ranges, field.Code, "citation-code");
                    AddRange(ranges, field.Result, "citation-result");
                }
            }
            catch
            {
            }
        }

        private static void AddFormulaFieldRanges(Document doc, List<ProtectedTextRange> ranges)
        {
            try
            {
                foreach (Field field in doc.Fields)
                {
                    string codeText = GetFieldCodeText(field);
                    if (!ContainsAnyMarker(codeText, FormulaFieldMarkers))
                    {
                        continue;
                    }

                    AddFieldCombinedRange(ranges, field, "formula-field");
                    AddRange(ranges, field.Code, "formula-code");
                    AddRange(ranges, field.Result, "formula-result");
                }
            }
            catch
            {
            }
        }

        private static void AddWordEquationRanges(Document doc, List<ProtectedTextRange> ranges)
        {
            try
            {
                foreach (OMath omath in doc.OMaths)
                {
                    AddRange(ranges, omath.Range, "word-omath");
                }
            }
            catch
            {
            }
        }

        private static void AddMathTypeRanges(Document doc, List<ProtectedTextRange> ranges)
        {
            try
            {
                foreach (InlineShape shape in doc.InlineShapes)
                {
                    if (IsMathTypeInlineShape(shape))
                    {
                        AddRange(ranges, shape.Range, "mathtype-object");
                    }
                }
            }
            catch
            {
            }
        }

        private static void AddLatexTextRanges(Document doc, string documentText, List<ProtectedTextRange> ranges)
        {
            if (string.IsNullOrEmpty(documentText))
            {
                return;
            }

            int contentStart = DocumentScanContext.GetContentStart(doc);
            foreach (Match match in LatexBlockRegex.Matches(documentText))
            {
                ranges.Add(new ProtectedTextRange
                {
                    Start = contentStart + match.Index,
                    End = contentStart + match.Index + match.Length,
                    Source = "latex-block"
                });
            }

            foreach (Match match in LatexInlineRegex.Matches(documentText))
            {
                ranges.Add(new ProtectedTextRange
                {
                    Start = contentStart + match.Index,
                    End = contentStart + match.Index + match.Length,
                    Source = "latex-inline"
                });
            }
        }

        private static bool IsMathTypeInlineShape(InlineShape shape)
        {
            try
            {
                string progId = shape.OLEFormat == null ? string.Empty : shape.OLEFormat.ProgID ?? string.Empty;
                string normalized = progId.ToUpperInvariant();
                return normalized.Contains("MATHTYPE") ||
                       normalized.Contains("EQUATION.DSMT") ||
                       normalized.Contains("MTEF") ||
                       normalized.Contains("MT");
            }
            catch
            {
                return false;
            }
        }

        private static void AddReferencesSectionRange(Document doc, List<ProtectedTextRange> ranges)
        {
            try
            {
                foreach (Paragraph paragraph in doc.Paragraphs)
                {
                    string text = paragraph.Range.Text ?? string.Empty;
                    string normalized = text.Replace("\r", string.Empty).Replace("\n", string.Empty).Trim();
                    if (ReferencesHeadingRegex.IsMatch(normalized))
                    {
                        ranges.Add(new ProtectedTextRange
                        {
                            Start = paragraph.Range.Start,
                            End = doc.Content.End,
                            Source = "references-heading"
                        });
                        return;
                    }
                }
            }
            catch
            {
            }
        }

        private static string GetFieldCodeText(Field field)
        {
            try
            {
                return field.Code == null ? string.Empty : field.Code.Text ?? string.Empty;
            }
            catch
            {
                return string.Empty;
            }
        }

        private static bool ContainsAnyMarker(string fieldCode, string[] markers)
        {
            if (string.IsNullOrEmpty(fieldCode))
            {
                return false;
            }

            string normalized = fieldCode.ToUpperInvariant();
            foreach (string marker in markers)
            {
                if (normalized.Contains(marker))
                {
                    return true;
                }
            }

            return false;
        }

        private static void AddRange(List<ProtectedTextRange> ranges, Range range, string source)
        {
            if (range == null)
            {
                return;
            }

            try
            {
                if (range.End > range.Start)
                {
                    ranges.Add(new ProtectedTextRange
                    {
                        Start = range.Start,
                        End = range.End,
                        Source = source
                    });
                }
            }
            catch
            {
            }
        }

        private static void AddFieldCombinedRange(List<ProtectedTextRange> ranges, Field field, string source)
        {
            try
            {
                int start = Math.Min(field.Code.Start, field.Result.Start);
                int end = Math.Max(field.Code.End, field.Result.End);
                if (end > start)
                {
                    ranges.Add(new ProtectedTextRange
                    {
                        Start = start,
                        End = end,
                        Source = source
                    });
                }
            }
            catch
            {
            }
        }

        private static bool IntersectsAny(List<ProtectedTextRange> ranges, int start, int end)
        {
            if (ranges == null || ranges.Count == 0)
            {
                return false;
            }

            int normalizedEnd = Math.Max(start + 1, end);
            foreach (ProtectedTextRange range in ranges)
            {
                if (start < range.End && normalizedEnd > range.Start)
                {
                    return true;
                }
            }

            return false;
        }

        private static bool IsFormulaSource(string source)
        {
            if (string.IsNullOrEmpty(source))
            {
                return false;
            }

            return source.Contains("latex") ||
                   source.Contains("omath") ||
                   source.Contains("mathtype") ||
                   source.Contains("formula");
        }

        private static void MergeRanges(List<ProtectedTextRange> ranges)
        {
            ranges.Sort((a, b) => a.Start.CompareTo(b.Start));
            for (int i = ranges.Count - 2; i >= 0; i--)
            {
                ProtectedTextRange current = ranges[i];
                ProtectedTextRange next = ranges[i + 1];
                if (current.End >= next.Start)
                {
                    current.End = Math.Max(current.End, next.End);
                    current.Source = current.Source + "+" + next.Source;
                    ranges.RemoveAt(i + 1);
                }
            }
        }

        private static string GetCacheKey(Document doc)
        {
            try
            {
                if (!string.IsNullOrEmpty(doc.FullName))
                {
                    return doc.FullName;
                }
            }
            catch
            {
            }

            return doc.GetHashCode().ToString();
        }

        private static int SafeContentEnd(Document doc)
        {
            try { return doc.Content.End; }
            catch { return -1; }
        }

        private static int SafeFieldCount(Document doc)
        {
            try { return doc.Fields.Count; }
            catch { return -1; }
        }

        private static int SafeOMathCount(Document doc)
        {
            try { return doc.OMaths.Count; }
            catch { return -1; }
        }

        private static int SafeInlineShapeCount(Document doc)
        {
            try { return doc.InlineShapes.Count; }
            catch { return -1; }
        }

        private sealed class ProtectedRangeCache
        {
            public int ContentEnd { get; set; }
            public int FieldCount { get; set; }
            public int OMathCount { get; set; }
            public int InlineShapeCount { get; set; }
            public int TextHash { get; set; }
            public List<ProtectedTextRange> Ranges { get; set; }
        }
    }
}
