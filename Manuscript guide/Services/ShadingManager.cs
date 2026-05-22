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
            if (ProtectedRangeService.IsRangeProtected(range))
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
            if (ProtectedRangeService.IsRangeProtected(range))
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

            for (int i = doc.Bookmarks.Count; i >= 1; i--)
            {
                try
                {
                    Bookmark bookmark = doc.Bookmarks[i];
                    if (bookmark.Name.StartsWith(modulePrefix, StringComparison.OrdinalIgnoreCase))
                    {
                        if (ProtectedRangeService.IsRangeProtected(bookmark.Range))
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

            try
            {
                for (int i = 1; i <= range.Characters.Count; i++)
                {
                    try
                    {
                        Range charRange = range.Characters[i];
                        charRange.Shading.BackgroundPatternColor = WdColor.wdColorAutomatic;
                    }
                    catch
                    {
                    }
                }
            }
            catch
            {
            }
        }
    }
}
