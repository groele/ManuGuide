using System;
using Microsoft.Office.Interop.Word;

namespace Manuscript_guide.Services
{
    public static class ShadingManager
    {
        private const string BookmarkPrefix = "mg_";
        private const string CommentPrefix = "Manuscript Guide";

        public static void ApplyActiveShading(Range range, string moduleType)
        {
            ApplyActiveShading(range, moduleType, false);
        }

        public static void ApplyActiveShading(Range range, string moduleType, bool allowProtected)
        {
            if (!allowProtected && ProtectedRangeService.IsRangeProtected(range))
            {
                return;
            }

            PluginSettings settings = SettingsManager.Current;
            if (settings.Colors.ContainsKey(moduleType))
            {
                range.Shading.BackgroundPatternColor = SettingsManager.HexToWdColor(settings.Colors[moduleType]);
            }
        }

        public static void ApplyFadedShading(Range range, string moduleType)
        {
            ApplyFadedShading(range, moduleType, false);
        }

        public static void ApplyFadedShading(Range range, string moduleType, bool allowProtected)
        {
            if (!allowProtected && ProtectedRangeService.IsRangeProtected(range))
            {
                return;
            }

            PluginSettings settings = SettingsManager.Current;
            if (settings.Colors.ContainsKey(moduleType))
            {
                range.Shading.BackgroundPatternColor = SettingsManager.GetFadedColor(settings.Colors[moduleType]);
            }
        }

        public static void ClearModuleShading(Document doc, string moduleType)
        {
            ClearPluginBookmarks(doc, moduleType);
        }

        public static void ClearPluginTraces(Document doc, string moduleType)
        {
            ClearPluginBookmarks(doc, moduleType);

            if (moduleType == "all")
            {
                ClearPluginComments(doc);
                CorrectionTracker.Instance.Clear();
            }
        }

        private static void ClearPluginBookmarks(Document doc, string moduleType)
        {
            string modulePrefix = string.Equals(moduleType, "all", StringComparison.OrdinalIgnoreCase)
                ? BookmarkPrefix
                : CorrectionTracker.GetBookmarkPrefixForModule(moduleType);
            bool allowProtectedClear = string.Equals(moduleType, "protected", StringComparison.OrdinalIgnoreCase);

            for (int i = doc.Bookmarks.Count; i >= 1; i--)
            {
                try
                {
                    Bookmark bookmark = doc.Bookmarks[i];
                    if (bookmark.Name.StartsWith(modulePrefix, StringComparison.OrdinalIgnoreCase))
                    {
                        bool isProtectedMarker = bookmark.Name.StartsWith(CorrectionTracker.GetBookmarkPrefixForModule("protected"), StringComparison.OrdinalIgnoreCase);
                        if (!allowProtectedClear && !isProtectedMarker && ProtectedRangeService.IsRangeProtected(bookmark.Range))
                        {
                            continue;
                        }

                        ClearPluginRangeTrace(bookmark.Range);
                        bookmark.Delete();
                    }
                }
                catch
                {
                }
            }
        }

        private static void ClearPluginComments(Document doc)
        {
            for (int i = doc.Comments.Count; i >= 1; i--)
            {
                try
                {
                    Comment comment = doc.Comments[i];
                    string commentText = comment.Range.Text ?? string.Empty;
                    if (commentText.StartsWith(CommentPrefix, StringComparison.OrdinalIgnoreCase))
                    {
                        if (ProtectedRangeService.IsRangeProtected(comment.Scope))
                        {
                            continue;
                        }

                        ClearPluginRangeTrace(comment.Scope);
                        comment.Delete();
                    }
                }
                catch
                {
                }
            }
        }

        private static void ClearPluginRangeTrace(Range range)
        {
            if (range == null)
            {
                return;
            }

            try
            {
                range.Shading.BackgroundPatternColor = WdColor.wdColorAutomatic;
            }
            catch
            {
            }
        }
    }
}
