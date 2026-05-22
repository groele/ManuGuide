using System;
using System.Collections.Generic;
using System.IO;
using System.Web.Script.Serialization;

namespace Manuscript_guide.Services
{
    public class PluginSettings
    {
        // 🔬 Subscript formatting
        public bool UseNativeSubscript { get; set; } = true;
        public bool EnableElementSubscriptConversion { get; set; } = true;

        // 🛡️ Academic Safety & Advanced Typography Tuning
        public bool UnifyGreekMu { get; set; } = true;
        public bool PreserveUserHighlights { get; set; } = true;
        public bool EquationPunctuation { get; set; } = true;
        public bool CrossRefCapitalization { get; set; } = true;
        public bool AutoBackup { get; set; } = true;
        public bool VariableLock { get; set; } = true;
        public bool DetectExistingItalics { get; set; } = false;

        // ⚙️ Diagnosis Thresholds
        public int CasingRatioThreshold { get; set; } = 70; // 50% - 95%
        public int MaxAcronymLagCharacters { get; set; } = 1000;

        // ✍️ Whitelists
        public string WhitelistCasing { get; set; } = "";
        public string WhitelistAcronyms { get; set; } = "";

        // 🎨 Custom Colors
        public Dictionary<string, string> Colors { get; set; } = new Dictionary<string, string>()
        {
            { "punc", "#00E5E5" },
            { "cap",  "#3399FF" },
            { "dash", "#FF5C5C" },
            { "data", "#FF66FF" },
            { "ital", "#5CFF5C" },
            { "sub",  "#B84DFF" },
            { "word", "#33B2B2" },
            { "keyword", "#FFFF33" },
            { "protected", "#D0D7DE" }
        };

        // Per-rule enable switches. Key format: module.ruleId, e.g. punc.fullwidth_punctuation.
        public Dictionary<string, bool> EnabledRules { get; set; } = new Dictionary<string, bool>();
    }

    public static class SettingsManager
    {
        private static readonly string SettingsDir = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "ManuscriptGuide"
        );
        private static readonly string SettingsFile = Path.Combine(SettingsDir, "settings.json");
        private static PluginSettings _current = null;
        private static readonly object _lock = new object();

        public static PluginSettings Current
        {
            get
            {
                lock (_lock)
                {
                    if (_current == null)
                    {
                        Load();
                    }
                    return _current;
                }
            }
        }

        public static void Load()
        {
            try
            {
                if (File.Exists(SettingsFile))
                {
                    string json = File.ReadAllText(SettingsFile);
                    var serializer = new JavaScriptSerializer();
                    _current = serializer.Deserialize<PluginSettings>(json);
                }
            }
            catch (Exception)
            {
                // Fallback to defaults on error
            }

            if (_current == null)
            {
                _current = new PluginSettings();
            }

            ClearLegacySampleWhitelists();
            EnsureColorDefaults();
            EnsureRuleDefaults();
        }

        private static void ClearLegacySampleWhitelists()
        {
            if (_current.WhitelistCasing == "GaAs, SrTiO3, LaTeX, iPhone, macOS, raman")
            {
                _current.WhitelistCasing = "";
            }

            if (_current.WhitelistAcronyms == "DNA, RNA, MOSFET, STEM, TEM, AFM")
            {
                _current.WhitelistAcronyms = "";
            }
        }

        private static void EnsureColorDefaults()
        {
            var defaults = new PluginSettings().Colors;
            if (_current.Colors == null)
            {
                _current.Colors = new Dictionary<string, string>(defaults);
                return;
            }

            foreach (var color in defaults)
            {
                if (!_current.Colors.ContainsKey(color.Key))
                {
                    _current.Colors[color.Key] = color.Value;
                }
            }
        }

        private static void EnsureRuleDefaults()
        {
            if (_current.EnabledRules == null)
            {
                _current.EnabledRules = new Dictionary<string, bool>();
            }

            foreach (DiagnosticRuleDefinition rule in DiagnosticRuleRegistry.AllRules)
            {
                if (!_current.EnabledRules.ContainsKey(rule.Key))
                {
                    _current.EnabledRules[rule.Key] = GetLegacyRuleDefault(rule);
                }
            }

        }

        private static bool GetLegacyRuleDefault(DiagnosticRuleDefinition rule)
        {
            switch (rule.Key)
            {
                case "global.filter_references_and_citations":
                case "global.skip_latex_formula_regions":
                    return true;
                case "global.mark_skipped_formula_regions":
                    return false;
                case "global.preserve_user_highlights":
                    return _current.PreserveUserHighlights;
                case "global.auto_backup_before_bulk_fix":
                    return _current.AutoBackup;
                case "punc.greek_mu_encoding":
                    return _current.UnifyGreekMu;
                case "punc.equation_spacing":
                case "punc.equation_terminal_punctuation":
                    return _current.EquationPunctuation;
                case "cap.crossref_capitalization":
                    return _current.CrossRefCapitalization;
                case "cap.physical_variable_casing_lock":
                    return _current.VariableLock;
                case "ital.existing_italics_review":
                    return _current.DetectExistingItalics;
                case "sub.element_formula_subscript":
                    return _current.EnableElementSubscriptConversion;
                default:
                    return rule.DefaultEnabled;
            }
        }

        public static bool IsRuleEnabled(string moduleType, string ruleId)
        {
            DiagnosticRuleDefinition rule = DiagnosticRuleRegistry.Find(moduleType, ruleId);
            string key = moduleType + "." + ruleId;

            if (Current.EnabledRules == null)
            {
                Current.EnabledRules = new Dictionary<string, bool>();
            }

            if (!Current.EnabledRules.ContainsKey(key))
            {
                Current.EnabledRules[key] = rule == null || rule.DefaultEnabled;
            }

            return Current.EnabledRules[key];
        }

        public static void Save()
        {
            try
            {
                if (!Directory.Exists(SettingsDir))
                {
                    Directory.CreateDirectory(SettingsDir);
                }

                var serializer = new JavaScriptSerializer();
                string json = serializer.Serialize(Current);
                File.WriteAllText(SettingsFile, json);
                
                // Signal settings change globally
                EventBus.TriggerSettingsChanged();
            }
            catch (Exception)
            {
                // Suppress save failures silently in office thread
            }
        }

        // Helper to convert hex to Word WdColor
        public static Microsoft.Office.Interop.Word.WdColor HexToWdColor(string hex)
        {
            if (string.IsNullOrEmpty(hex)) return Microsoft.Office.Interop.Word.WdColor.wdColorAutomatic;
            hex = hex.TrimStart('#');
            if (hex.Length != 6) return Microsoft.Office.Interop.Word.WdColor.wdColorAutomatic;

            try
            {
                int r = Convert.ToInt32(hex.Substring(0, 2), 16);
                int g = Convert.ToInt32(hex.Substring(2, 2), 16);
                int b = Convert.ToInt32(hex.Substring(4, 2), 16);

                // Word color is 0x00BBGGRR (BGR format)
                int colorRef = r | (g << 8) | (b << 16);
                return (Microsoft.Office.Interop.Word.WdColor)colorRef;
            }
            catch
            {
                return Microsoft.Office.Interop.Word.WdColor.wdColorAutomatic;
            }
        }

        // Helper to generate a faded background color (light pastel/macaron) from standard Hex Color
        public static Microsoft.Office.Interop.Word.WdColor GetFadedColor(string hex)
        {
            if (string.IsNullOrEmpty(hex)) return Microsoft.Office.Interop.Word.WdColor.wdColorAutomatic;
            hex = hex.TrimStart('#');
            if (hex.Length != 6) return Microsoft.Office.Interop.Word.WdColor.wdColorAutomatic;

            try
            {
                int r = Convert.ToInt32(hex.Substring(0, 2), 16);
                int g = Convert.ToInt32(hex.Substring(2, 2), 16);
                int b = Convert.ToInt32(hex.Substring(4, 2), 16);

                // Blend with white (e.g. 90% white, 10% base color) to get faded review pastel
                double factor = 0.90;
                int fadedR = (int)(r + (255 - r) * factor);
                int fadedG = (int)(g + (255 - g) * factor);
                int fadedB = (int)(b + (255 - b) * factor);

                int colorRef = fadedR | (fadedG << 8) | (fadedB << 16);
                return (Microsoft.Office.Interop.Word.WdColor)colorRef;
            }
            catch
            {
                return Microsoft.Office.Interop.Word.WdColor.wdColorAutomatic;
            }
        }
    }
}
