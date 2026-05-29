using System;
using System.Collections.Generic;
using Microsoft.Office.Interop.Word;
using Manuscript_guide.Models;

namespace Manuscript_guide.Services
{
    public sealed class DocumentScanContext : IDisposable
    {
        [ThreadStatic]
        private static DocumentScanContext current;

        private readonly DocumentScanContext previous;

        private DocumentScanContext(Document doc)
        {
            Document = doc;
            ContentStart = GetContentStart(doc);
            RawText = doc == null ? string.Empty : doc.Content.Text;
            ProtectedRanges = ProtectedRangeService.GetProtectedRanges(doc, RawText);
            Text = ProtectedRangeService.MaskProtectedText(RawText, ProtectedRanges, ContentStart);
            RuleStats = new Dictionary<string, ScannerRuleStats>();
            TextSegments = BuildTextSegments(doc, RawText);
            previous = current;
            current = this;
        }

        private DocumentScanContext(Document doc, DocumentSnapshot snapshot)
        {
            Document = doc;
            Snapshot = snapshot;
            ContentStart = snapshot.ContentStart;
            RawText = snapshot.FullText;
            ProtectedRanges = new List<ProtectedTextRange>(snapshot.ProtectedRanges.Ranges);
            Text = ProtectedRangeService.MaskProtectedText(RawText, ProtectedRanges, ContentStart);
            RuleStats = new Dictionary<string, ScannerRuleStats>();
            
            TextSegments = new List<TextSegment>();
            foreach (var p in snapshot.Paragraphs)
            {
                TextSegments.Add(new TextSegment
                {
                    TextStart = p.Start,
                    TextEnd = p.End,
                    DocumentStart = p.Start + ContentStart,
                    DocumentEnd = p.End + ContentStart,
                    Text = p.Text
                });
            }

            previous = current;
            current = this;
        }

        public Document Document { get; private set; }
        public DocumentSnapshot Snapshot { get; private set; }
        public int ContentStart { get; private set; }
        public string RawText { get; private set; }
        public string Text { get; private set; }
        public List<ProtectedTextRange> ProtectedRanges { get; private set; }
        public Dictionary<string, ScannerRuleStats> RuleStats { get; private set; }
        private List<TextSegment> TextSegments { get; set; }

        public static DocumentScanContext Current
        {
            get { return current; }
        }

        public static DocumentScanContext Begin(Document doc)
        {
            return new DocumentScanContext(doc);
        }

        public static DocumentScanContext Begin(Document doc, DocumentSnapshot snapshot)
        {
            return new DocumentScanContext(doc, snapshot);
        }

        public static string GetText(Document doc)
        {
            // When doc is null (background scan path), use the active context's cached text
            if (current != null && (doc == null || ReferenceEquals(current.Document, doc)))
            {
                return current.Text;
            }

            return doc == null ? string.Empty : doc.Content.Text;
        }

        public static List<ProtectedTextRange> GetProtectedRanges(Document doc)
        {
            if (current != null && (doc == null || ReferenceEquals(current.Document, doc)))
            {
                return current.ProtectedRanges;
            }

            return null;
        }

        public static int GetContentStart(Document doc)
        {
            try
            {
                return doc == null ? 0 : doc.Content.Start;
            }
            catch
            {
                return 0;
            }
        }

        public static int TextOffsetToDocumentPosition(Document doc, int textOffset)
        {
            TextSegment segment = FindTextSegment(doc, textOffset);
            if (segment != null)
            {
                return segment.DocumentStart + (textOffset - segment.TextStart);
            }

            int contentStart = current != null && (doc == null || ReferenceEquals(current.Document, doc))
                ? current.ContentStart
                : GetContentStart(doc);
            return contentStart + textOffset;
        }

        public static int DocumentPositionToTextOffset(Document doc, int documentPosition)
        {
            if (current != null && (doc == null || ReferenceEquals(current.Document, doc)) && current.TextSegments != null)
            {
                foreach (TextSegment segment in current.TextSegments)
                {
                    if (documentPosition >= segment.DocumentStart && documentPosition <= segment.DocumentEnd)
                    {
                        return segment.TextStart + Math.Max(0, documentPosition - segment.DocumentStart);
                    }
                }
            }

            int contentStart = current != null && (doc == null || ReferenceEquals(current.Document, doc))
                ? current.ContentStart
                : GetContentStart(doc);
            return Math.Max(0, documentPosition - contentStart);
        }

        public static Range CreateRangeFromTextSpan(Document doc, int textStart, int length, string expectedText = null)
        {
            if (doc == null || textStart < 0 || length <= 0)
            {
                return null;
            }

            try
            {
                int start = TextOffsetToDocumentPosition(doc, textStart);
                int end = start + length;
                if (start < doc.Content.Start || end > doc.Content.End || end <= start)
                {
                    return null;
                }

                Range range = doc.Range(start, end);
                if (!string.IsNullOrEmpty(expectedText) && !RangeTextMatches(range, expectedText))
                {
                    return FindRangeByTextOccurrence(doc, textStart, length, expectedText);
                }

                return range;
            }
            catch
            {
                return null;
            }
        }

        public static bool RangeTextMatches(Range range, string expectedText)
        {
            if (range == null)
            {
                return false;
            }

            string actual = range.Text ?? string.Empty;
            return string.Equals(NormalizeRangeText(actual), NormalizeRangeText(expectedText), StringComparison.Ordinal);
        }

        private static string NormalizeRangeText(string text)
        {
            return (text ?? string.Empty).Replace("\r", "\n");
        }

        private static TextSegment FindTextSegment(Document doc, int textOffset)
        {
            if (current == null || (doc != null && !ReferenceEquals(current.Document, doc)) || current.TextSegments == null)
            {
                return null;
            }

            foreach (TextSegment segment in current.TextSegments)
            {
                if (textOffset >= segment.TextStart && textOffset < segment.TextEnd)
                {
                    return segment;
                }
            }

            return null;
        }

        private static Range FindRangeByTextOccurrence(Document doc, int textStart, int length, string expectedText)
        {
            TextSegment segment = FindTextSegment(doc, textStart);
            if (segment == null || string.IsNullOrEmpty(expectedText))
            {
                return null;
            }

            int occurrenceIndex = CountPreviousOccurrences(segment.Text, textStart - segment.TextStart, expectedText);
            if (occurrenceIndex < 0)
            {
                return null;
            }

            try
            {
                Range searchRange = doc.Range(segment.DocumentStart, segment.DocumentEnd);
                searchRange.Find.ClearFormatting();
                searchRange.Find.Text = expectedText;
                searchRange.Find.Forward = true;
                searchRange.Find.Wrap = WdFindWrap.wdFindStop;
                searchRange.Find.MatchWildcards = false;

                int found = 0;
                while (searchRange.Find.Execute())
                {
                    if (found == occurrenceIndex && RangeTextMatches(searchRange, expectedText))
                    {
                        return doc.Range(searchRange.Start, searchRange.End);
                    }

                    found++;
                    int nextStart = Math.Max(searchRange.End, searchRange.Start + 1);
                    if (nextStart >= segment.DocumentEnd)
                    {
                        break;
                    }

                    searchRange.SetRange(nextStart, segment.DocumentEnd);
                    searchRange.Find.ClearFormatting();
                    searchRange.Find.Text = expectedText;
                    searchRange.Find.Forward = true;
                    searchRange.Find.Wrap = WdFindWrap.wdFindStop;
                    searchRange.Find.MatchWildcards = false;
                }
            }
            catch
            {
            }

            return null;
        }

        private static int CountPreviousOccurrences(string text, int beforeOffset, string value)
        {
            if (string.IsNullOrEmpty(text) || string.IsNullOrEmpty(value) || beforeOffset < 0)
            {
                return -1;
            }

            int count = 0;
            int index = 0;
            while (index >= 0 && index < beforeOffset)
            {
                index = text.IndexOf(value, index, StringComparison.Ordinal);
                if (index < 0 || index >= beforeOffset)
                {
                    break;
                }

                count++;
                index += Math.Max(1, value.Length);
            }

            return count;
        }

        private static List<TextSegment> BuildTextSegments(Document doc, string rawText)
        {
            List<TextSegment> segments = new List<TextSegment>();
            if (doc == null || string.IsNullOrEmpty(rawText))
            {
                return segments;
            }

            int cursor = 0;
            try
            {
                foreach (Paragraph paragraph in doc.Paragraphs)
                {
                    string paragraphText = paragraph.Range.Text ?? string.Empty;
                    if (paragraphText.Length == 0)
                    {
                        continue;
                    }

                    int textStart = rawText.IndexOf(paragraphText, cursor, StringComparison.Ordinal);
                    if (textStart < 0)
                    {
                        textStart = cursor;
                    }

                    segments.Add(new TextSegment
                    {
                        TextStart = textStart,
                        TextEnd = textStart + paragraphText.Length,
                        DocumentStart = paragraph.Range.Start,
                        DocumentEnd = paragraph.Range.End,
                        Text = paragraphText
                    });

                    cursor = textStart + paragraphText.Length;
                }
            }
            catch
            {
            }

            return segments;
        }

        public static void RecordCandidate(string moduleType, string subtype)
        {
            ScannerRuleStats stats = GetOrCreateStats(moduleType, subtype);
            if (stats != null)
            {
                stats.CandidateCount++;
            }
        }

        public static void RecordIssue(string moduleType, string subtype)
        {
            ScannerRuleStats stats = GetOrCreateStats(moduleType, subtype);
            if (stats != null)
            {
                stats.IssueCount++;
            }
        }

        public static void RecordSkip(string moduleType, string subtype, ScannerSkipReason reason)
        {
            ScannerRuleStats stats = GetOrCreateStats(moduleType, subtype);
            if (stats == null)
            {
                return;
            }

            switch (reason)
            {
                case ScannerSkipReason.ProtectedRange:
                    stats.ProtectedSkipCount++;
                    break;
                case ScannerSkipReason.RangeMismatch:
                    stats.RangeMismatchSkipCount++;
                    break;
                case ScannerSkipReason.AlreadyCorrect:
                    stats.AlreadyCorrectSkipCount++;
                    break;
                case ScannerSkipReason.RuleFilter:
                    stats.RuleFilterSkipCount++;
                    break;
                case ScannerSkipReason.InvalidRange:
                    stats.InvalidRangeSkipCount++;
                    break;
                case ScannerSkipReason.BookmarkFailed:
                    stats.BookmarkFailedSkipCount++;
                    break;
            }
        }

        public static ScannerStatsSnapshot GetStatsSnapshot(string moduleType)
        {
            ScannerStatsSnapshot snapshot = new ScannerStatsSnapshot();
            if (current == null || current.RuleStats == null)
            {
                return snapshot;
            }

            foreach (ScannerRuleStats stats in current.RuleStats.Values)
            {
                if (!string.IsNullOrEmpty(moduleType) && moduleType != "all" && stats.ModuleType != moduleType)
                {
                    continue;
                }

                snapshot.CandidateCount += stats.CandidateCount;
                snapshot.IssueCount += stats.IssueCount;
                snapshot.ProtectedSkipCount += stats.ProtectedSkipCount;
                snapshot.RangeMismatchSkipCount += stats.RangeMismatchSkipCount;
                snapshot.AlreadyCorrectSkipCount += stats.AlreadyCorrectSkipCount;
                snapshot.RuleFilterSkipCount += stats.RuleFilterSkipCount;
                snapshot.InvalidRangeSkipCount += stats.InvalidRangeSkipCount;
                snapshot.BookmarkFailedSkipCount += stats.BookmarkFailedSkipCount;
            }

            return snapshot;
        }

        private static ScannerRuleStats GetOrCreateStats(string moduleType, string subtype)
        {
            if (current == null || current.RuleStats == null)
            {
                return null;
            }

            string ruleId = DiagnosticRuleRegistry.ResolveRuleId(moduleType, subtype);
            string key = string.IsNullOrEmpty(ruleId) ? moduleType + "." + subtype : ruleId;
            if (!current.RuleStats.ContainsKey(key))
            {
                current.RuleStats[key] = new ScannerRuleStats
                {
                    ModuleType = moduleType,
                    RuleId = ruleId
                };
            }

            return current.RuleStats[key];
        }

        public void Dispose()
        {
            current = previous;
        }

        private sealed class TextSegment
        {
            public int TextStart { get; set; }
            public int TextEnd { get; set; }
            public int DocumentStart { get; set; }
            public int DocumentEnd { get; set; }
            public string Text { get; set; }
        }
    }
}
