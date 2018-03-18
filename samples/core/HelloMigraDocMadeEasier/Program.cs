using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using HelloMigraDoc;
using MigraDoc.Rendering;

namespace HelloMigraDocMadeEasier
{
    /// <summary>
    /// This sample shows some features of MigraDoc.
    /// </summary>
    class Program
    {
        static void Main(string[] args)
        {
            // Create a MigraDoc document.
            var document = Documents.CreateDocument();

            document.SaveMDDDL("MigraDoc.mdddl");

            // Save the document and start a viewer.
#if DEBUG
            var filename = Guid.NewGuid().ToString("N").ToUpper() + ".pdf";
#else
            var filename = "HelloMigraDoc.pdf";
#endif
            document.MakePdf(filename, true);
        }
    }
}
