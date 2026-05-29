using System;
using Microsoft.Office.Interop.Word;
using Manuscript_guide.Models;
using Manuscript_guide.Services;

namespace Manuscript_guide.Scanners
{
    public static class IssueMatchFactory
    {
        public static IssueItem Create(
            Document doc,
            string documentText,
            string moduleType,
            string subtype,
            int start,
            int length,
            string originalText,
            string recommendFix,
            string description)
        {
            DocumentScanContext.RecordCandidate(moduleType, subtype);

            if (string.IsNullOrEmpty(documentText) || start < 0 || length <= 0)
            {
                DocumentScanContext.RecordSkip(moduleType, subtype, ScannerSkipReason.InvalidRange);
                return null;
            }

            // High Performance: Use snapshot for protection and verification if available
            var context = DocumentScanContext.Current;
            if (context != null && context.Snapshot != null)
            {
                var snapshot = context.Snapshot;
                if (start + length > snapshot.FullText.Length)
                {
                    DocumentScanContext.RecordSkip(moduleType, subtype, ScannerSkipReason.InvalidRange);
                    return null;
                }

                string actualText = snapshot.FullText.Substring(start, length);
                if (!string.Equals(actualText, originalText, StringComparison.Ordinal))
                {
                    DocumentScanContext.RecordSkip(moduleType, subtype, ScannerSkipReason.RangeMismatch);
                    return null;
                }

                if (snapshot.ProtectedRanges.Intersects(start, start + length))
                {
                    DocumentScanContext.RecordSkip(moduleType, subtype, ScannerSkipReason.ProtectedRange);
                    return null;
                }
            }
            else
            {
                // Fallback (only used if running outside a snapshot context)
                if (start + length > documentText.Length)
                {
                    DocumentScanContext.RecordSkip(moduleType, subtype, ScannerSkipReason.InvalidRange);
                    return null;
                }

                string actualText = documentText.Substring(start, length);
                if (!string.Equals(actualText, originalText, StringComparison.Ordinal))
                {
                    DocumentScanContext.RecordSkip(moduleType, subtype, ScannerSkipReason.RangeMismatch);
                    return null;
                }

                if (ProtectedRangeService.IsRangeProtected(doc, start, start + length))
                {
                    DocumentScanContext.RecordSkip(moduleType, subtype, ScannerSkipReason.ProtectedRange);
                    return null;
                }
            }

            // v2.3 Core Principle: Separate scanning from Word COM modifications.
            // NO Bookmark creation or Active Shading is performed here.
            DocumentScanContext.RecordIssue(moduleType, subtype);

            return new IssueItem
            {
                IssueId = Guid.NewGuid().ToString(),
                Type = moduleType,
                Subtype = subtype,
                Start = start,
                End = start + length,
                OriginalText = originalText,
                RecommendFix = recommendFix,
                Desc = description,
                Context = PunctuationScanner.GetContextSnippet(documentText, start, length)
            };
        }

        public static IssueItem CreateFromRange(
            Document doc,
            string documentText,
            string moduleType,
            string subtype,
            Range range,
            string originalText,
            string recommendFix,
            string description)
        {
            if (range == null || string.IsNullOrEmpty(documentText))
            {
                return null;
            }

            // Use context position mapping
            int start = DocumentScanContext.DocumentPositionToTextOffset(doc, range.Start);
            int length = Math.Max(1, range.End - range.Start);
            return Create(doc, documentText, moduleType, subtype, start, length, originalText, recommendFix, description);
        }
    }
}
