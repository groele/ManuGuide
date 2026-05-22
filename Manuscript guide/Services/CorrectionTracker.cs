using System;
using System.Collections.Generic;
using Microsoft.Office.Interop.Word;

namespace Manuscript_guide.Services
{
    public class IssueBackup
    {
        public string IssueId { get; set; }
        public string BookmarkName { get; set; }
        public string OriginalText { get; set; }
        
        // Font attributes preservation
        public int IsSubscript { get; set; } // 0 = false, 1 = true, 9999999 = mixed
        public int IsSuperscript { get; set; }
        public int IsItalic { get; set; }
    }

    public class CorrectionTracker
    {
        private readonly Dictionary<string, IssueBackup> _backups = new Dictionary<string, IssueBackup>();
        private static readonly object _lock = new object();
        private static CorrectionTracker _instance = null;

        public static CorrectionTracker Instance
        {
            get
            {
                lock (_lock)
                {
                    if (_instance == null)
                    {
                        _instance = new CorrectionTracker();
                    }
                    return _instance;
                }
            }
        }

        private CorrectionTracker() { }

        // Clears all in-memory backups
        public void Clear()
        {
            _backups.Clear();
        }

        // Creates a native bookmark for an issue range
        public string CreateBookmark(Document doc, string issueId, Range range, string moduleType = null)
        {
            if (ProtectedRangeService.IsRangeProtected(range))
            {
                return null;
            }

            string bookmarkName = CreateBookmarkName(issueId, moduleType);
            try
            {
                if (doc.Bookmarks.Exists(bookmarkName))
                {
                    doc.Bookmarks[bookmarkName].Delete();
                }
                doc.Bookmarks.Add(bookmarkName, range);
                return bookmarkName;
            }
            catch
            {
                return null;
            }
        }

        public string FindBookmarkName(Document doc, string issueId)
        {
            string normalizedIssueId = issueId.Replace("-", "_");
            string legacyName = "mg_" + normalizedIssueId;
            if (doc.Bookmarks.Exists(legacyName))
            {
                return legacyName;
            }

            string compactIssueId = issueId.Replace("-", "");
            string legacySuffix = "_" + normalizedIssueId;
            string compactSuffix = "_" + compactIssueId;
            for (int i = 1; i <= doc.Bookmarks.Count; i++)
            {
                try
                {
                    string name = doc.Bookmarks[i].Name;
                    if (name.StartsWith("mg_", StringComparison.OrdinalIgnoreCase) &&
                        (name.EndsWith(legacySuffix, StringComparison.OrdinalIgnoreCase) ||
                         name.EndsWith(compactSuffix, StringComparison.OrdinalIgnoreCase)))
                    {
                        return name;
                    }
                }
                catch
                {
                }
            }

            return legacyName;
        }

        public static string CreateBookmarkName(string issueId, string moduleType)
        {
            string normalizedIssueId = issueId.Replace("-", "");
            if (string.IsNullOrEmpty(moduleType))
            {
                return "mg_" + normalizedIssueId;
            }

            return GetBookmarkPrefixForModule(moduleType) + normalizedIssueId;
        }

        public static string GetBookmarkPrefixForModule(string moduleType)
        {
            if (string.IsNullOrEmpty(moduleType))
            {
                return "mg_";
            }

            string shortModule = moduleType.Length <= 3 ? moduleType : moduleType.Substring(0, 3);
            return "mg_" + shortModule + "_";
        }

        // Executes correction and creates a backup
        public bool ExecuteCorrection(Document doc, string issueId, string moduleType, string newText, Action<Range> customFormatAction = null)
        {
            string bookmarkName = FindBookmarkName(doc, issueId);
            if (!doc.Bookmarks.Exists(bookmarkName)) return false;

            try
            {
                Range range = doc.Bookmarks[bookmarkName].Range;
                if (ProtectedRangeService.IsRangeProtected(range))
                {
                    return false;
                }

                string originalText = range.Text;

                var backup = new IssueBackup
                {
                    IssueId = issueId,
                    BookmarkName = bookmarkName,
                    OriginalText = originalText,
                    IsSubscript = range.Font.Subscript,
                    IsSuperscript = range.Font.Superscript,
                    IsItalic = range.Font.Italic
                };
                _backups[issueId] = backup;

                // Perform replacement
                range.Text = newText;

                // Re-establish range and bookmark on the corrected text
                Range newRange = doc.Range(range.Start, range.Start + newText.Length);
                if (doc.Bookmarks.Exists(bookmarkName))
                {
                    doc.Bookmarks[bookmarkName].Delete();
                }
                doc.Bookmarks.Add(bookmarkName, newRange);

                // Apply formatting (e.g. subscript formatting or variable uprighting)
                customFormatAction?.Invoke(newRange);
                string finalText = newRange.Text;

                // Keep visible review traces: faded shading plus a native Word comment.
                ShadingManager.ApplyFadedShading(newRange, moduleType);
                AddReviewComment(doc, newRange, moduleType, originalText, finalText);
                return true;
            }
            catch (Exception)
            {
                // Safety guard for Word Interop thread exceptions
                return false;
            }
        }

        private void AddReviewComment(Document doc, Range range, string moduleType, string originalText, string newText)
        {
            try
            {
                string commentText =
                    $"Manuscript Guide [{moduleType}]: 已应用建议修改。原文: {originalText} -> 修改: {newText}";
                doc.Comments.Add(range, commentText);
            }
            catch
            {
                // Some protected or compatibility documents reject comments; shading remains the fallback trace.
            }
        }

        // Reverts a previous correction
        public bool ExecuteUndo(Document doc, string issueId, string moduleType)
        {
            if (!_backups.ContainsKey(issueId)) return false;
            var backup = _backups[issueId];

            if (!doc.Bookmarks.Exists(backup.BookmarkName)) return false;

            try
            {
                Range currentRange = doc.Bookmarks[backup.BookmarkName].Range;
                if (ProtectedRangeService.IsRangeProtected(currentRange))
                {
                    return false;
                }
                
                // Revert text
                currentRange.Text = backup.OriginalText;

                // Re-establish range and bookmark on original text
                Range revertedRange = doc.Range(currentRange.Start, currentRange.Start + backup.OriginalText.Length);
                doc.Bookmarks.Add(backup.BookmarkName, revertedRange);

                // Restore font properties
                revertedRange.Font.Subscript = backup.IsSubscript;
                revertedRange.Font.Superscript = backup.IsSuperscript;
                revertedRange.Font.Italic = backup.IsItalic;

                // Restore high-contrast active diagnostic highlight color
                ShadingManager.ApplyActiveShading(revertedRange, moduleType);

                // Clean up backup entry
                _backups.Remove(issueId);
                return true;
            }
            catch
            {
                return false;
            }
        }

        // Checks if an issue has a backup (can be undone)
        public bool CanUndo(string issueId)
        {
            return _backups.ContainsKey(issueId);
        }
    }
}
