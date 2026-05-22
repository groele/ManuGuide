using System.Collections.Generic;
using Manuscript_guide.Models;

namespace Manuscript_guide.Services
{
    public static class IssueMetadataService
    {
        public static List<IssueItem> EnrichAll(List<IssueItem> issues)
        {
            if (issues == null)
            {
                return new List<IssueItem>();
            }

            foreach (IssueItem issue in issues)
            {
                Enrich(issue);
            }

            return issues;
        }

        public static IssueItem Enrich(IssueItem issue)
        {
            if (issue == null)
            {
                return null;
            }

            if (string.IsNullOrEmpty(issue.RuleId))
            {
                issue.RuleId = DiagnosticRuleRegistry.ResolveRuleId(issue.Type, issue.Subtype);
            }

            DiagnosticRuleDefinition rule = DiagnosticRuleRegistry.FindByKey(issue.RuleId);
            issue.RuleTitle = rule == null ? issue.Subtype : rule.Title;

            if (string.IsNullOrEmpty(issue.MatchReason))
            {
                issue.MatchReason = issue.Desc;
            }

            return issue;
        }
    }
}
