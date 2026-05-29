using System.Collections.Generic;
using Manuscript_guide.Models;

namespace Manuscript_guide.Services
{
    public static class ScannableSegmentPlanner
    {
        public static List<ScannableSegment> Build(string fullText, List<ProtectedTextRange> protectedRanges)
        {
            var segments = new List<ScannableSegment>();
            if (string.IsNullOrEmpty(fullText)) return segments;

            int cursor = 0;
            int id = 0;

            if (protectedRanges != null)
            {
                foreach (var range in protectedRanges)
                {
                    if (cursor < range.Start)
                    {
                        int segStart = cursor;
                        int segEnd = range.Start;
                        segments.Add(new ScannableSegment
                        {
                            SegmentId = id++,
                            TextStart = segStart,
                            TextEnd = segEnd,
                            Text = fullText.Substring(segStart, segEnd - segStart)
                        });
                    }
                    cursor = System.Math.Max(cursor, range.End);
                }
            }

            if (cursor < fullText.Length)
            {
                segments.Add(new ScannableSegment
                {
                    SegmentId = id,
                    TextStart = cursor,
                    TextEnd = fullText.Length,
                    Text = fullText.Substring(cursor)
                });
            }

            return segments;
        }
    }
}
