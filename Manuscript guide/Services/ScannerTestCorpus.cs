using System.Collections.Generic;

namespace Manuscript_guide.Services
{
    public sealed class ScannerTestCase
    {
        public string ModuleType { get; set; }
        public string Text { get; set; }
        public string[] ExpectedMatches { get; set; }
        public string[] ExpectedNonMatches { get; set; }
    }

    public static class ScannerTestCorpus
    {
        public static List<ScannerTestCase> CreateDefaultCases()
        {
            return new List<ScannerTestCase>
            {
                new ScannerTestCase
                {
                    ModuleType = "sub",
                    Text = "WSe2, MoS2, Bi2O2Se, CuInP2S6, CO2, O2, MoSe₂, SiO₂, Eg, kB, cm-2, E_g and T_{c} should be reviewed.",
                    ExpectedMatches = new[] { "WSe2", "MoS2", "Bi2O2Se", "CuInP2S6", "CO2", "O2", "MoSe₂", "SiO₂", "Eg", "kB", "E_g", "T_{c}" },
                    ExpectedNonMatches = new[] { "S1", "Fig1" }
                },
                new ScannerTestCase
                {
                    ModuleType = "ital",
                    Text = "The band and doping words must not mark inner letters, but E = hv and where x is the coordinate should be reviewed.",
                    ExpectedMatches = new[] { "E", "x" },
                    ExpectedNonMatches = new[] { "band", "doping", "temperature" }
                },
                new ScannerTestCase
                {
                    ModuleType = "data",
                    Text = "The sample was measured from 10-300 K at 5.2x10^12 cm-2, not on 2025-05-22 or doi:10.1000/abc-123.",
                    ExpectedMatches = new[] { "from 10-300", "5.2x10^12" },
                    ExpectedNonMatches = new[] { "2025-05-22", "10.1000/abc-123" }
                }
            };
        }
    }
}
