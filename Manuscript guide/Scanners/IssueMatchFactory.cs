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
            if (doc == null || string.IsNullOrEmpty(documentText) || start < 0 || length <= 0)
            {
                return null;
            }

            if (start + length > doc.Content.End)
            {
                return null;
            }

            string issueId = Guid.NewGuid().ToString();
            Range range;
            try
            {
                range = doc.Range(start, start + length);
            }
            catch
            {
                return null;
            }

            if (ProtectedRangeService.IsRangeProtected(range))
            {
                return null;
            }

            string bookmarkName = CorrectionTracker.Instance.CreateBookmark(doc, issueId, range, moduleType);
            if (string.IsNullOrEmpty(bookmarkName))
            {
                return null;
            }

            ShadingManager.ApplyActiveShading(range, moduleType);

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
    }
}

