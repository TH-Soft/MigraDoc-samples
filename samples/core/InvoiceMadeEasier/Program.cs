using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using MigraDoc.Rendering;

namespace Invoice
{
    /// <summary>
    /// This sample shows how to create a simple invoice of a fictional book store. The invoice document
    /// is created with the MigraDoc document object model and then rendered to PDF with PDFsharp.
    /// </summary>
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create an invoice form with the sample invoice data.
                var invoice = new InvoiceForm("../../../../assets/xml/invoice.xml");

                // Create the document using MigraDoc.
                var document = invoice.CreateDocument();
                document.Document.UseCmykColor = true;

#if DEBUG
                // For debugging only...
                MigraDoc.DocumentObjectModel.IO.DdlWriter.WriteToFile(document.Document, "MigraDoc.mdddl");
                //var document2 = MigraDoc.DocumentObjectModel.IO.DdlReader.DocumentFromFile("MigraDoc.mdddl");
                //document = document2;
                // With PDFsharp 1.50 beta 3 there is a known problem: the blank before "by" gets lost while persisting as MDDDL.

                MigraDoc.DocumentObjectModel.IO.DdlWriter.WriteToFile(document.Document.Clone(), "MigraDoc2.mdddl");
#endif

#if true
                // Create a renderer for PDF that uses Unicode font encoding.
                var pdfRenderer = new PdfDocumentRenderer(true);

                // Set the MigraDoc document.
                pdfRenderer.Document = document.Document/*.Clone()*/;
                // $THHO TODO Investigate why Clone() leads to an unusable document.

                // Create the PDF document.
                pdfRenderer.RenderDocument();

                // Save the PDF document...
                var filename = "Invoice.pdf";
#if DEBUG
                // I don't want to close the document constantly...
                filename = "Invoice-" + Guid.NewGuid().ToString("N").ToUpper() + ".pdf";
#endif
                pdfRenderer.Save(filename);
                // ...and start a viewer.
                Process.Start(filename);
#else
                // Create a renderer for PDF that uses Unicode font encoding.
                // Save the PDF document...
                var filename = "Invoice.pdf";
#if DEBUG
                // I don't want to close the document constantly...
                filename = "Invoice-" + Guid.NewGuid().ToString("N").ToUpper() + ".pdf";
#endif
                document.MakePdf(filename, true);
#endif
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                Console.ReadLine();
            }
        }
    }
}
