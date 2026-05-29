using System;
using System.Collections.Generic;

namespace Manuscript_guide.Models
{
    public sealed class FormattingIndex
    {
        public static readonly FormattingIndex Empty = new FormattingIndex();

        private readonly List<(int Start, int End)> _ranges = new List<(int, int)>();
        private (int Start, int End)[] _sortedRanges = Array.Empty<(int, int)>();

        public IReadOnlyList<(int Start, int End)> Ranges => _sortedRanges;

        public void AddRange(int start, int end)
        {
            if (end > start)
            {
                _ranges.Add((start, end));
            }
        }

        public void SortAndMerge()
        {
            _ranges.Sort((a, b) => a.Start.CompareTo(b.Start));
            for (int i = _ranges.Count - 2; i >= 0; i--)
            {
                var current = _ranges[i];
                var next = _ranges[i + 1];
                if (current.End >= next.Start)
                {
                    _ranges[i] = (current.Start, Math.Max(current.End, next.End));
                    _ranges.RemoveAt(i + 1);
                }
            }
            _sortedRanges = _ranges.ToArray();
        }

        public bool IsSpanFullyCovered(int start, int end)
        {
            if (_sortedRanges.Length == 0) return false;

            int low = 0;
            int high = _sortedRanges.Length - 1;
            while (low <= high)
            {
                int mid = low + (high - low) / 2;
                var r = _sortedRanges[mid];
                if (r.Start <= start && r.End >= end)
                {
                    return true;
                }
                if (r.Start > start)
                {
                    high = mid - 1;
                }
                else
                {
                    low = mid + 1;
                }
            }
            return false;
        }

        public bool HasAnyFormatting(int start, int end)
        {
            if (_sortedRanges.Length == 0) return false;

            int low = 0;
            int high = _sortedRanges.Length - 1;
            int normalizedEnd = Math.Max(start + 1, end);

            while (low <= high)
            {
                int mid = low + (high - low) / 2;
                var r = _sortedRanges[mid];
                if (r.End > start && r.Start < normalizedEnd)
                {
                    return true;
                }
                if (r.Start >= normalizedEnd)
                {
                    high = mid - 1;
                }
                else
                {
                    low = mid + 1;
                }
            }
            return false;
        }
    }

    public sealed class ParagraphRange
    {
        public int Start { get; set; }
        public int End { get; set; }
        public string Text { get; set; } = string.Empty;
    }

    public sealed class DocumentSnapshot
    {
        public string SnapshotId { get; set; } = Guid.NewGuid().ToString("N");
        public DateTime CreatedAt { get; set; } = DateTime.UtcNow;

        public int ContentStart { get; set; }
        public int ContentEnd { get; set; }

        public string FullText { get; set; } = string.Empty;

        public IReadOnlyList<ParagraphRange> Paragraphs { get; set; } = Array.Empty<ParagraphRange>();
        public ProtectedRangeIndex ProtectedRanges { get; set; } = ProtectedRangeIndex.Empty;
        public FormattingIndex Italics { get; set; } = FormattingIndex.Empty;
        public FormattingIndex Subscripts { get; set; } = FormattingIndex.Empty;
        public FormattingIndex Superscripts { get; set; } = FormattingIndex.Empty;
        public FormattingIndex OMaths { get; set; } = FormattingIndex.Empty;
    }
}
