using System;
using System.Collections.Generic;
using Manuscript_guide.Scanners;

namespace Manuscript_guide.Services
{
    public sealed class DiagnosticModule
    {
        public DiagnosticModule(string moduleType, string displayName, string icon, Func<ISpecializedScanner> scannerFactory, bool includeInFullScan = true)
        {
            ModuleType = moduleType;
            DisplayName = displayName;
            Icon = icon;
            ScannerFactory = scannerFactory;
            IncludeInFullScan = includeInFullScan;
        }

        public string ModuleType { get; private set; }
        public string DisplayName { get; private set; }
        public string Icon { get; private set; }
        public bool IncludeInFullScan { get; private set; }

        private Func<ISpecializedScanner> ScannerFactory { get; set; }

        public ISpecializedScanner CreateScanner()
        {
            return ScannerFactory == null ? null : ScannerFactory();
        }
    }

    public static class DiagnosticModuleRegistry
    {
        private static readonly List<DiagnosticModule> Modules = new List<DiagnosticModule>
        {
            new DiagnosticModule("punc", "标点与格式", "✍️", () => new PunctuationScanner()),
            new DiagnosticModule("cap", "大小写规范", "🔠", () => new CapitalizationScanner()),
            new DiagnosticModule("dash", "横线与混排", "➖", () => new DashSpacingScanner()),
            new DiagnosticModule("data", "数据与数值", "🔢", () => new DataValueScanner()),
            new DiagnosticModule("ital", "变量与斜体", "𝘹", () => new ItalicsScanner()),
            new DiagnosticModule("sub", "角标与改写", "₂", () => new SubscriptScanner()),
            new DiagnosticModule("word", "语病与措辞", "📝", () => new WordingGrammarScanner())
        };

        public static IEnumerable<DiagnosticModule> AllModules
        {
            get { return Modules; }
        }

        public static IEnumerable<DiagnosticModule> FullScanModules
        {
            get
            {
                foreach (DiagnosticModule module in Modules)
                {
                    if (module.IncludeInFullScan)
                    {
                        yield return module;
                    }
                }
            }
        }

        public static int FullScanModuleCount
        {
            get
            {
                int count = 0;
                foreach (DiagnosticModule module in Modules)
                {
                    if (module.IncludeInFullScan)
                    {
                        count++;
                    }
                }

                return count;
            }
        }

        public static DiagnosticModule Find(string moduleType)
        {
            foreach (DiagnosticModule module in Modules)
            {
                if (string.Equals(module.ModuleType, moduleType, StringComparison.OrdinalIgnoreCase))
                {
                    return module;
                }
            }

            return null;
        }

        public static ISpecializedScanner CreateScanner(string moduleType)
        {
            DiagnosticModule module = Find(moduleType);
            return module == null ? null : module.CreateScanner();
        }

        public static string GetDisplayName(string moduleType)
        {
            if (string.Equals(moduleType, "all", StringComparison.OrdinalIgnoreCase))
            {
                return "全部专项";
            }

            if (string.Equals(moduleType, "keyword", StringComparison.OrdinalIgnoreCase))
            {
                return "关键词标记";
            }

            DiagnosticModule module = Find(moduleType);
            return module == null ? "未知模块" : module.DisplayName;
        }

        public static string GetIcon(string moduleType)
        {
            DiagnosticModule module = Find(moduleType);
            return module == null ? "⚙️" : module.Icon;
        }

        public static string GetFullScanDescription()
        {
            List<string> names = new List<string>();
            foreach (DiagnosticModule module in FullScanModules)
            {
                names.Add(module.DisplayName);
            }

            return "已启用 " + names.Count + " 个专项模块：" + string.Join("、", names) + "。";
        }
    }
}
