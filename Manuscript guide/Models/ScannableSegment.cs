namespace Manuscript_guide.Models
{
    public sealed class ScannableSegment
    {
        public int SegmentId { get; set; }
        public int TextStart { get; set; }
        public int TextEnd { get; set; }
        public string Text { get; set; } = string.Empty;
    }
}
