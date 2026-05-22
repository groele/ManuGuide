using System.Collections.Generic;
using System.Text;

namespace Manuscript_guide.Services
{
    public enum ScannerSkipReason
    {
        ProtectedRange,
        RangeMismatch,
        AlreadyCorrect,
        RuleFilter,
        InvalidRange,
        BookmarkFailed
    }

    public sealed class ScannerRuleStats
    {
        public string ModuleType { get; set; }
        public string RuleId { get; set; }
        public int CandidateCount { get; set; }
        public int IssueCount { get; set; }
        public int ProtectedSkipCount { get; set; }
        public int RangeMismatchSkipCount { get; set; }
        public int AlreadyCorrectSkipCount { get; set; }
        public int RuleFilterSkipCount { get; set; }
        public int InvalidRangeSkipCount { get; set; }
        public int BookmarkFailedSkipCount { get; set; }

        public int SkippedCount
        {
            get
            {
                return ProtectedSkipCount + RangeMismatchSkipCount + AlreadyCorrectSkipCount +
                       RuleFilterSkipCount + InvalidRangeSkipCount + BookmarkFailedSkipCount;
            }
        }
    }

    public sealed class ScannerStatsSnapshot
    {
        public int CandidateCount { get; set; }
        public int IssueCount { get; set; }
        public int ProtectedSkipCount { get; set; }
        public int RangeMismatchSkipCount { get; set; }
        public int AlreadyCorrectSkipCount { get; set; }
        public int RuleFilterSkipCount { get; set; }
        public int InvalidRangeSkipCount { get; set; }
        public int BookmarkFailedSkipCount { get; set; }

        public int SkippedCount
        {
            get
            {
                return ProtectedSkipCount + RangeMismatchSkipCount + AlreadyCorrectSkipCount +
                       RuleFilterSkipCount + InvalidRangeSkipCount + BookmarkFailedSkipCount;
            }
        }

        public string ToStatusText(int visibleIssueCount)
        {
            StringBuilder builder = new StringBuilder();
            builder.Append("发现 ");
            builder.Append(visibleIssueCount);
            builder.Append(" 处排版建议");

            if (CandidateCount > 0)
            {
                builder.Append("；候选 ");
                builder.Append(CandidateCount);
                builder.Append("，有效 ");
                builder.Append(IssueCount);
            }

            if (SkippedCount > 0)
            {
                builder.Append("，跳过 ");
                builder.Append(SkippedCount);
                builder.Append(" 处");
            }

            List<string> details = new List<string>();
            if (ProtectedSkipCount > 0) details.Add("受保护区 " + ProtectedSkipCount);
            if (AlreadyCorrectSkipCount > 0) details.Add("格式已正确 " + AlreadyCorrectSkipCount);
            if (RuleFilterSkipCount > 0) details.Add("规则过滤 " + RuleFilterSkipCount);
            if (RangeMismatchSkipCount > 0) details.Add("定位不一致 " + RangeMismatchSkipCount);
            if (InvalidRangeSkipCount > 0) details.Add("无效定位 " + InvalidRangeSkipCount);
            if (BookmarkFailedSkipCount > 0) details.Add("书签失败 " + BookmarkFailedSkipCount);

            if (details.Count > 0)
            {
                builder.Append("（");
                builder.Append(string.Join("，", details));
                builder.Append("）");
            }

            builder.Append("。");
            return builder.ToString();
        }
    }
}
