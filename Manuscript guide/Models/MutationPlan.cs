using System.Collections.Generic;

namespace Manuscript_guide.Models
{
    public static class MutationPlan
    {
        public static List<IssueItem> BuildDescending(List<IssueItem> issues)
        {
            var sorted = new List<IssueItem>(issues);
            sorted.Sort((a, b) => b.Start.CompareTo(a.Start));
            return sorted;
        }
    }
}
