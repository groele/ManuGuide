using System;
using System.Collections.Generic;
using Manuscript_guide.Models;

namespace Manuscript_guide.Services
{
    public static class EventBus
    {
        // Event triggered when a specialized scan is requested from the Ribbon
        public static event Action<string> ScanTriggered;

        // Event triggered when all specialized scanners should run as one workflow
        public static event Action FullScanRequested;

        // Event triggered when all specialized scanners should scan and fix as one workflow
        public static event Action FullFixRequested;
        
        // Event triggered when a specialized scan is completed with a list of issues
        public static event Action<string, List<IssueItem>> ScanCompleted;

        // Event triggered when a specific correction is accepted in the task pane
        public static event Action<string, string> FixTriggered; // (issueId, moduleType)

        // Event triggered when a user clicks "一键修复本模块"
        public static event Action<string> FixAllTriggered; // (moduleType)

        // Event triggered when a user clicks "撤销 (Undo)" in the task pane
        public static event Action<string, string> UndoTriggered; // (issueId, moduleType)

        // Event triggered when settings are saved/updated (e.g., preference switches)
        public static event Action SettingsChanged;

        // Event triggered when user clicks "清除标记"
        public static event Action<string> ClearHighlightsRequested; // (moduleType: "all" or specific)

        // Event triggered to open advanced settings panel in taskpane
        public static event Action OpenSettingsRequested;

        // Event triggered when keyword search is requested
        public static event Action<string, bool, bool, string> KeywordSearchRequested; // (keywords, caseSensitive, wholeWord, colorHex)

        public static void TriggerScan(string moduleType)
        {
            ScanTriggered?.Invoke(moduleType);
        }

        public static void TriggerFullScan()
        {
            FullScanRequested?.Invoke();
        }

        public static void TriggerFullFix()
        {
            FullFixRequested?.Invoke();
        }

        public static void TriggerScanCompleted(string moduleType, List<IssueItem> issues)
        {
            ScanCompleted?.Invoke(moduleType, issues);
        }

        public static void TriggerFix(string issueId, string moduleType)
        {
            FixTriggered?.Invoke(issueId, moduleType);
        }

        public static void TriggerFixAll(string moduleType)
        {
            FixAllTriggered?.Invoke(moduleType);
        }

        public static void TriggerUndo(string issueId, string moduleType)
        {
            UndoTriggered?.Invoke(issueId, moduleType);
        }

        public static void TriggerSettingsChanged()
        {
            SettingsChanged?.Invoke();
        }

        public static void TriggerClearHighlights(string moduleType)
        {
            ClearHighlightsRequested?.Invoke(moduleType);
        }

        public static void TriggerOpenSettings()
        {
            OpenSettingsRequested?.Invoke();
        }

        public static void TriggerKeywordSearch(string keywords, bool caseSensitive, bool wholeWord, string colorHex)
        {
            KeywordSearchRequested?.Invoke(keywords, caseSensitive, wholeWord, colorHex);
        }
    }
}
