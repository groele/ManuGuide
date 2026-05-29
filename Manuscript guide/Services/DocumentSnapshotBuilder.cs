using System;
using System.Collections.Generic;
using Microsoft.Office.Interop.Word;
using Manuscript_guide.Models;

namespace Manuscript_guide.Services
{
    public static class DocumentSnapshotBuilder
    {
        public static DocumentSnapshot Build(Document doc)
        {
            if (doc == null)
            {
                return new DocumentSnapshot();
            }

            // Must run on STA thread. This builds the immutable memory representation of the document.
            int contentStart = 0;
            int contentEnd = 0;
            string rawText = string.Empty;

            try
            {
                contentStart = doc.Content.Start;
                contentEnd = doc.Content.End;
                rawText = doc.Content.Text ?? string.Empty;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("Failed to read document basic properties: " + ex.Message);
            }

            var paragraphs = BuildParagraphs(doc, rawText, contentStart);
            var protectedRanges = ProtectedRangeService.GetProtectedRanges(doc, rawText);
            var protectedIndex = new ProtectedRangeIndex(protectedRanges);

            var italics = BuildFormattingIndex(doc, contentStart, FormattingType.Italic);
            var subscripts = BuildFormattingIndex(doc, contentStart, FormattingType.Subscript);
            var superscripts = BuildFormattingIndex(doc, contentStart, FormattingType.Superscript);
            var omaths = BuildOMathIndex(doc, contentStart);

            return new DocumentSnapshot
            {
                ContentStart = contentStart,
                ContentEnd = contentEnd,
                FullText = rawText,
                Paragraphs = paragraphs,
                ProtectedRanges = protectedIndex,
                Italics = italics,
                Subscripts = subscripts,
                Superscripts = superscripts,
                OMaths = omaths
            };
        }

        private static List<ParagraphRange> BuildParagraphs(Document doc, string rawText, int contentStart)
        {
            var paragraphs = new List<ParagraphRange>();
            if (string.IsNullOrEmpty(rawText))
            {
                return paragraphs;
            }

            try
            {
                int start = 0;
                while (start < rawText.Length)
                {
                    int nextBreak = rawText.IndexOf('\r', start);
                    int end = nextBreak >= 0 ? nextBreak + 1 : rawText.Length;
                    string paragraphText = rawText.Substring(start, end - start);
                    
                    if (paragraphText.Length > 0)
                    {
                        paragraphs.Add(new ParagraphRange
                        {
                            Start = start,
                            End = start + paragraphText.Length,
                            Text = paragraphText
                        });
                    }

                    start = end;
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("Failed to build paragraphs: " + ex.Message);
            }

            return paragraphs;
        }

        private enum FormattingType
        {
            Italic,
            Subscript,
            Superscript
        }

        private static FormattingIndex BuildFormattingIndex(Document doc, int contentStart, FormattingType type)
        {
            var index = new FormattingIndex();
            try
            {
                Range r = doc.Content;
                r.Find.ClearFormatting();
                
                switch (type)
                {
                    case FormattingType.Italic:
                        r.Find.Font.Italic = 1;
                        break;
                    case FormattingType.Subscript:
                        r.Find.Font.Subscript = 1;
                        break;
                    case FormattingType.Superscript:
                        r.Find.Font.Superscript = 1;
                        break;
                }

                r.Find.Text = "";
                r.Find.Forward = true;
                r.Find.Format = true;
                r.Find.Wrap = WdFindWrap.wdFindStop;

                while (r.Find.Execute())
                {
                    int start = r.Start - contentStart;
                    int end = r.End - contentStart;
                    index.AddRange(start, end);

                    int nextStart = Math.Max(r.End, r.Start + 1);
                    if (nextStart >= doc.Content.End)
                    {
                        break;
                    }

                    r.SetRange(nextStart, doc.Content.End);
                    r.Find.ClearFormatting();
                    
                    switch (type)
                    {
                        case FormattingType.Italic:
                            r.Find.Font.Italic = 1;
                            break;
                        case FormattingType.Subscript:
                            r.Find.Font.Subscript = 1;
                            break;
                        case FormattingType.Superscript:
                            r.Find.Font.Superscript = 1;
                            break;
                    }

                    r.Find.Text = "";
                    r.Find.Forward = true;
                    r.Find.Format = true;
                    r.Find.Wrap = WdFindWrap.wdFindStop;
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Failed to build formatting index for {type}: " + ex.Message);
            }

            index.SortAndMerge();
            return index;
        }

        private static FormattingIndex BuildOMathIndex(Document doc, int contentStart)
        {
            var index = new FormattingIndex();
            try
            {
                foreach (OMath omath in doc.OMaths)
                {
                    index.AddRange(omath.Range.Start - contentStart, omath.Range.End - contentStart);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("Failed to build OMath index: " + ex.Message);
            }

            index.SortAndMerge();
            return index;
        }
    }
}
