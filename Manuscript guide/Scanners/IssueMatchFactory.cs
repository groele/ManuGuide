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

            if (doc == null || string.IsNullOrEmpty(documentText) || start < 0 || length <= 0)
            {
                DocumentScanContext.RecordSkip(moduleType, subtype, ScannerSkipReason.InvalidRange);
                return null;
            }

            if (start + length > documentText.Length)
            {
                DocumentScanContext.RecordSkip(moduleType, subtype, ScannerSkipReason.InvalidRange);
                return null;
            }

            string issueId = Guid.NewGuid().ToString();
            Range range = DocumentScanContext.CreateRangeFromTextSpan(doc, start, length, originalText);
            if (range == null)
            {
                DocumentScanContext.RecordSkip(moduleType, subtype, ScannerSkipReason.RangeMismatch);
                return null;
            }

            if (ProtectedRangeService.IsRangeProtected(range))
            {
                DocumentScanContext.RecordSkip(moduleType, subtype, ScannerSkipReason.ProtectedRange);
                return null;
            }

            string bookmarkName = CorrectionTracker.Instance.CreateBookmark(doc, issueId, range, moduleType);
            if (string.IsNullOrEmpty(bookmarkName))
            {
                DocumentScanContext.RecordSkip(moduleType, subtype, ScannerSkipReason.BookmarkFailed);
                return null;
            }

            ShadingManager.ApplyActiveShading(range, moduleType);
            DocumentScanContext.RecordIssue(moduleType, subtype);

            return new IssueItem
            {
                IssueId = issueId,
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
            if (doc == null || range == null || string.IsNullOrEmpty(documentText))
            {
                return null;
            }

            int start = DocumentScanContext.DocumentPositionToTextOffset(doc, range.Start);
            int length = Math.Max(1, range.End - range.Start);
            return Create(doc, documentText, moduleType, subtype, start, length, originalText, recommendFix, description);
        }
    }
}

