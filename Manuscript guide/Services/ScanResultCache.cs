using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using Manuscript_guide.Models;

namespace Manuscript_guide.Services
{
    public static class ScanResultCache
    {
        private static readonly object CacheLock = new object();
        private static readonly Dictionary<string, CachedScanResult> Cache = new Dictionary<string, CachedScanResult>();

        private static string lastDocumentHash;
        private static string lastRuleConfigHash;

        public class CachedScanResult
        {
            public string DocumentHash { get; set; }
            public string RuleConfigHash { get; set; }
            public List<IssueItem> Issues { get; set; }
            public DateTime Timestamp { get; set; }
        }

        public static void UpdateHashes(string documentHash, string ruleConfigHash)
        {
            lock (CacheLock)
            {
                if (documentHash != lastDocumentHash || ruleConfigHash != lastRuleConfigHash)
                {
                    Cache.Clear();
                    lastDocumentHash = documentHash;
                    lastRuleConfigHash = ruleConfigHash;
                }
            }
        }

        public static List<IssueItem> GetCachedResults(string moduleType)
        {
            lock (CacheLock)
            {
                if (Cache.TryGetValue(moduleType, out CachedScanResult cached))
                {
                    if (cached.DocumentHash == lastDocumentHash && cached.RuleConfigHash == lastRuleConfigHash)
                    {
                        return cached.Issues;
                    }
                }
                return null;
            }
        }

        public static void StoreResults(string moduleType, List<IssueItem> issues, string documentHash, string ruleConfigHash)
        {
            lock (CacheLock)
            {
                Cache[moduleType] = new CachedScanResult
                {
                    DocumentHash = documentHash,
                    RuleConfigHash = ruleConfigHash,
                    Issues = issues,
                    Timestamp = DateTime.Now
                };
            }
        }

        public static void Invalidate()
        {
            lock (CacheLock)
            {
                Cache.Clear();
                lastDocumentHash = null;
                lastRuleConfigHash = null;
            }
        }

        public static string ComputeDocumentHash(string text)
        {
            if (string.IsNullOrEmpty(text)) return string.Empty;
            using (var sha = SHA256.Create())
            {
                byte[] bytes = sha.ComputeHash(Encoding.UTF8.GetBytes(text));
                return Convert.ToBase64String(bytes);
            }
        }

        public static string ComputeRuleConfigHash()
        {
            // Hash every setting that can change scanner output, not only rule toggles.
            var settings = SettingsManager.Current;
            var sb = new StringBuilder();
            sb.Append("UseNativeSubscript=").Append(settings.UseNativeSubscript).Append(';');
            sb.Append("EnableElementSubscriptConversion=").Append(settings.EnableElementSubscriptConversion).Append(';');
            sb.Append("UnifyGreekMu=").Append(settings.UnifyGreekMu).Append(';');
            sb.Append("PreserveUserHighlights=").Append(settings.PreserveUserHighlights).Append(';');
            sb.Append("EquationPunctuation=").Append(settings.EquationPunctuation).Append(';');
            sb.Append("CrossRefCapitalization=").Append(settings.CrossRefCapitalization).Append(';');
            sb.Append("VariableLock=").Append(settings.VariableLock).Append(';');
            sb.Append("DetectExistingItalics=").Append(settings.DetectExistingItalics).Append(';');

            if (settings.EnabledRules != null)
            {
                foreach (var kvp in settings.EnabledRules.OrderBy(kvp => kvp.Key, StringComparer.Ordinal))
                {
                    sb.Append(kvp.Key).Append('=').Append(kvp.Value).Append(';');
                }
            }
            sb.Append("CasingThreshold=").Append(settings.CasingRatioThreshold);
            sb.Append(";MaxAcronymLag=").Append(settings.MaxAcronymLagCharacters);

            using (var sha = SHA256.Create())
            {
                byte[] bytes = sha.ComputeHash(Encoding.UTF8.GetBytes(sb.ToString()));
                return Convert.ToBase64String(bytes);
            }
        }
    }
}
