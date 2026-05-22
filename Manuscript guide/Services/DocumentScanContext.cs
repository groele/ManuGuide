using System;
using System.Collections.Generic;
using Microsoft.Office.Interop.Word;

namespace Manuscript_guide.Services
{
    public sealed class DocumentScanContext : IDisposable
    {
        [ThreadStatic]
        private static DocumentScanContext current;

        private readonly DocumentScanContext previous;

        private DocumentScanContext(Document doc)
        {
            Document = doc;
            Text = doc == null ? string.Empty : doc.Content.Text;
            ProtectedRanges = ProtectedRangeService.GetProtectedRanges(doc);
            previous = current;
            current = this;
        }

        public Document Document { get; private set; }
        public string Text { get; private set; }
        public List<ProtectedTextRange> ProtectedRanges { get; private set; }

        public static DocumentScanContext Current
        {
            get { return current; }
        }

        public static DocumentScanContext Begin(Document doc)
        {
            return new DocumentScanContext(doc);
        }

        public static string GetText(Document doc)
        {
            if (current != null && ReferenceEquals(current.Document, doc))
            {
                return current.Text;
            }

            return doc == null ? string.Empty : doc.Content.Text;
        }

        public static List<ProtectedTextRange> GetProtectedRanges(Document doc)
        {
            if (current != null && ReferenceEquals(current.Document, doc))
            {
                return current.ProtectedRanges;
            }

            return null;
        }

        public void Dispose()
        {
            current = previous;
        }
    }
}
