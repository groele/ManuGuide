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

        private static readonly Regex ReferencesHeadingRegex = new Regex(
            @"^\s*(references|reference|bibliography|works\s+cited|literature\s+cited|参考文献|参考资料)\s*[:：]?\s*$",
            RegexOptions.IgnoreCase | RegexOptions.Compiled);

        public static bool IsProtectionEnabled()
        {
            return SettingsManager.IsRuleEnabled("global", "filter_references_and_citations");
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

                if (!IntersectsAny(protectedRanges, issue.Start, issue.End))
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

            string key = GetCacheKey(doc);
            int contentEnd = SafeContentEnd(doc);
            int fieldCount = SafeFieldCount(doc);

            lock (CacheLock)
            {
                if (Cache.ContainsKey(key))
                {
                    ProtectedRangeCache cached = Cache[key];
                    if (cached.ContentEnd == contentEnd && cached.FieldCount == fieldCount)
                    {
                        return cached.Ranges;
                    }
                }
            }

            List<ProtectedTextRange> ranges = BuildProtectedRanges(doc);
            MergeRanges(ranges);

            lock (CacheLock)
            {
                Cache[key] = new ProtectedRangeCache
                {
                    ContentEnd = contentEnd,
                    FieldCount = fieldCount,
                    Ranges = ranges
                };
            }

            return ranges;
        }

        private static List<ProtectedTextRange> BuildProtectedRanges(Document doc)
        {
            List<ProtectedTextRange> ranges = new List<ProtectedTextRange>();
            AddCitationFieldRanges(doc, ranges);
            AddReferencesSectionRange(doc, ranges);
            return ranges;
        }

        private static void AddCitationFieldRanges(Document doc, List<ProtectedTextRange> ranges)
        {
            try
            {
                foreach (Field field in doc.Fields)
                {
                    string codeText = string.Empty;
                    try
                    {
                        codeText = field.Code == null ? string.Empty : field.Code.Text ?? string.Empty;
                    }
                    catch
                    {
                    }

                    if (!ContainsCitationMarker(codeText))
                    {
                        continue;
                    }

                    AddFieldCombinedRange(ranges, field);
                    AddRange(ranges, field.Code, "citation-code");
                    AddRange(ranges, field.Result, "citation-result");
                }
            }
            catch
            {
            }
        }

        private static bool ContainsCitationMarker(string fieldCode)
        {
            if (string.IsNullOrEmpty(fieldCode))
            {
                return false;
            }

            string normalized = fieldCode.ToUpperInvariant();
            foreach (string marker in CitationFieldMarkers)
            {
                if (normalized.Contains(marker))
                {
                    return true;
                }
            }

            return false;
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

        private static void AddFieldCombinedRange(List<ProtectedTextRange> ranges, Field field)
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
                        Source = "citation-field"
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
            try
            {
                return doc.Content.End;
            }
            catch
            {
                return -1;
            }
        }

        private static int SafeFieldCount(Document doc)
        {
            try
            {
                return doc.Fields.Count;
            }
            catch
            {
                return -1;
            }
        }

        private sealed class ProtectedRangeCache
        {
            public int ContentEnd { get; set; }
            public int FieldCount { get; set; }
            public List<ProtectedTextRange> Ranges { get; set; }
        }
    }
}
