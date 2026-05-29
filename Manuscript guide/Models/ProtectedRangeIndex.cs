using System;
using System.Collections.Generic;
using Manuscript_guide.Services;

namespace Manuscript_guide.Models
{
    public sealed class ProtectedRangeIndex
    {
        public static readonly ProtectedRangeIndex Empty = new ProtectedRangeIndex(new List<ProtectedTextRange>());

        private readonly ProtectedTextRange[] _ranges;

        public ProtectedRangeIndex(List<ProtectedTextRange> ranges)
        {
            var list = ranges != null ? new List<ProtectedTextRange>(ranges) : new List<ProtectedTextRange>();
            list.Sort((a, b) => a.Start.CompareTo(b.Start));
            _ranges = list.ToArray();
        }

        public IReadOnlyList<ProtectedTextRange> Ranges => _ranges;

        public bool Intersects(int start, int end)
        {
            if (_ranges.Length == 0) return false;

            int low = 0;
            int high = _ranges.Length - 1;
            int normalizedEnd = Math.Max(start + 1, end);

            while (low <= high)
            {
                int mid = low + (high - low) / 2;
                var r = _ranges[mid];

                // If overlap
                if (r.End > start && r.Start < normalizedEnd)
                    return true;

                if (r.Start >= normalizedEnd)
                    high = mid - 1;
                else
                    low = mid + 1;
            }

            return false;
        }
    }
}
