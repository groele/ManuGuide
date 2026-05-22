using System.Collections.Generic;
using Microsoft.Office.Interop.Word;
using Manuscript_guide.Models;

namespace Manuscript_guide.Scanners
{
    public interface ISpecializedScanner
    {
        string ModuleType { get; }
        List<IssueItem> Scan(Document doc);
    }
}

