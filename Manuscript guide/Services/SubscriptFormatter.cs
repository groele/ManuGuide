using System;
using Microsoft.Office.Interop.Word;

namespace Manuscript_guide.Services
{
    public static class SubscriptFormatter
    {
        // Applies formatting based on settings preference
        public static void ApplyCorrection(Document doc, Range range, string rawText, string formatType, bool useNativeFormat)
        {
            if (!useNativeFormat)
            {
                // Fallback Mode: Physical Unicode character rewrite (Al₂O₃, cm⁻³)
                range.Text = ConvertToUnicodeSubSup(rawText, formatType);
                range.Font.Subscript = 0;
                range.Font.Superscript = 0;
                range.Font.Italic = 0;
                return;
            }

            // Primary Native Mode: Keep standard ASCII and apply Font Subscript/Superscript
            string cleanAscii = ConvertToAsciiText(rawText, formatType);
            range.Text = cleanAscii;

            // Define the range of the newly inserted text
            Range formattedRange = doc.Range(range.Start, range.Start + cleanAscii.Length);
            formattedRange.Font.Subscript = 0;
            formattedRange.Font.Superscript = 0;

            if (formatType == "chemical")
            {
                // Chemical subscripts: set all numbers to Subscript using native Find
                Range findRange = doc.Range(formattedRange.Start, formattedRange.End);
                Find find = findRange.Find;
                find.ClearFormatting();
                find.Text = "^#"; // Any digit
                find.Forward = true;
                find.Wrap = WdFindWrap.wdFindStop;
                while (find.Execute())
                {
                    if (findRange.Start >= formattedRange.End) break;
                    findRange.Font.Subscript = 1;
                    findRange.Font.Superscript = 0;
                    findRange.Collapse(WdCollapseDirection.wdCollapseEnd);
                    findRange.End = formattedRange.End;
                }
            }
            else if (formatType == "unit")
            {
                // Physical units: e.g. cm-3, set negative symbols and numbers to Superscript using native Find
                // 1. Find and superscript digits
                Range findRange = doc.Range(formattedRange.Start, formattedRange.End);
                Find find = findRange.Find;
                find.ClearFormatting();
                find.Text = "^#";
                find.Forward = true;
                find.Wrap = WdFindWrap.wdFindStop;
                while (find.Execute())
                {
                    if (findRange.Start >= formattedRange.End) break;
                    findRange.Font.Superscript = 1;
                    findRange.Font.Subscript = 0;
                    findRange.Collapse(WdCollapseDirection.wdCollapseEnd);
                    findRange.End = formattedRange.End;
                }

                // 2. Find and superscript minus sign '-'
                findRange = doc.Range(formattedRange.Start, formattedRange.End);
                find = findRange.Find;
                find.ClearFormatting();
                find.Text = "-";
                find.Forward = true;
                find.Wrap = WdFindWrap.wdFindStop;
                while (find.Execute())
                {
                    if (findRange.Start >= formattedRange.End) break;
                    findRange.Font.Superscript = 1;
                    findRange.Font.Subscript = 0;
                    findRange.Collapse(WdCollapseDirection.wdCollapseEnd);
                    findRange.End = formattedRange.End;
                }

                // 3. Find and superscript minus sign '−' (Unicode minus)
                findRange = doc.Range(formattedRange.Start, formattedRange.End);
                find = findRange.Find;
                find.ClearFormatting();
                find.Text = "−";
                find.Forward = true;
                find.Wrap = WdFindWrap.wdFindStop;
                while (find.Execute())
                {
                    if (findRange.Start >= formattedRange.End) break;
                    findRange.Font.Superscript = 1;
                    findRange.Font.Subscript = 0;
                    findRange.Collapse(WdCollapseDirection.wdCollapseEnd);
                    findRange.End = formattedRange.End;
                }
            }
            else if (formatType == "desc")
            {
                // Academic descriptions: e.g., Eg -> E_g, where g is standard subscript (upright)
                // Use absolute character offset (Start + 1) instead of the slow Characters[2] COM call
                if (cleanAscii.Length > 1)
                {
                    Range subRange = doc.Range(formattedRange.Start + 1, formattedRange.End);
                    subRange.Font.Subscript = 1;
                    subRange.Font.Superscript = 0;
                    subRange.Font.Italic = 0; // Upright subscripts for physical labels
                }
            }
        }

        // Translates Unicode subscript/superscript back to standard ASCII
        public static string ConvertToAsciiText(string input, string formatType)
        {
            if (string.IsNullOrEmpty(input)) return input;

            // Subscript translations
            string unicodeSubs = "₀₁₂₃₄₅₆₇₈₉ₓ";
            string asciiSubs   = "0123456789x";

            // Superscript translations
            string unicodeSups = "⁰¹²³⁴⁵⁶⁷⁸⁹⁺⁻⁼";
            string asciiSups   = "0123456789+-=";

            char[] chars = input.ToCharArray();
            for (int i = 0; i < chars.Length; i++)
            {
                int subIdx = unicodeSubs.IndexOf(chars[i]);
                if (subIdx >= 0)
                {
                    chars[i] = asciiSubs[subIdx];
                    continue;
                }

                int supIdx = unicodeSups.IndexOf(chars[i]);
                if (supIdx >= 0)
                {
                    chars[i] = asciiSups[supIdx];
                }
            }

            string result = new string(chars);
            if (formatType == "desc" && result.Contains("_"))
            {
                // Remove LaTeX elements
                result = result.Replace("_", "").Replace("{", "").Replace("}", "");
            }
            return result;
        }

        // Converts ASCII to physical Unicode subscript/superscript (Fallback Mode)
        public static string ConvertToUnicodeSubSup(string input, string formatType)
        {
            if (string.IsNullOrEmpty(input)) return input;
            
            if (formatType == "chemical")
            {
                // Al2O3 -> Al₂O₃
                return input.Replace("0", "₀").Replace("1", "₁").Replace("2", "₂")
                            .Replace("3", "₃").Replace("4", "₄").Replace("5", "₅")
                            .Replace("6", "₆").Replace("7", "₇").Replace("8", "₈")
                            .Replace("9", "₉");
            }
            if (formatType == "unit")
            {
                // cm-3 -> cm⁻³
                return input.Replace("-3", "⁻³").Replace("-2", "⁻²").Replace("-1", "⁻¹")
                            .Replace("^-3", "⁻³").Replace("^-2", "⁻²").Replace("^-1", "⁻¹");
            }
            return input;
        }
    }
}
