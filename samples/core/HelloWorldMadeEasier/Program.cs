using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Text;
using MigraDoc.DocumentObjectModel;
using MigraDoc.DocumentObjectModel.Fields;
using MigraDoc.Rendering;
using MigraDocMadeEZ;

namespace HelloWorldMadeEasier
{
    /// <summary>
    /// This sample is the obligatory Hello World program for MigraDoc documents.
    /// </summary>
    class Program
    {
        static void Main(string[] args)
        {
            // Instantiate MigraDocMadeEZR.
            var mez = new MigraDocMadeEZR();

            // Create a MigraDoc document.
            CreateDocument(mez);

#if DEBUG
            //MigraDoc.DocumentObjectModel.IO.DdlWriter dw = new MigraDoc.DocumentObjectModel.IO.DdlWriter("HelloWorld.mdddl");
            //dw.WriteDocument(mez.Document);
            //dw.Close();
            mez.SaveMDDDL("HelloWorld.mdddl");
#endif
            
            // Save the document and start a viewer.
            const string filename = "HelloWorld.pdf";
            mez.MakePdf(filename, true, false);
        }

        /// <summary>
        /// Creates an absolutely minimalistic document.
        /// </summary>
        static void CreateDocument(MigraDocMadeEZR mez)
        {
            // Add a paragraph to the section.
            var mezParagraph = mez.AddParagraph();

            // Set font color.
            mezParagraph.Color(Colors.DarkBlue);

            // Add some text to the paragraph.
            mezParagraph.AddFormattedText(new MezFormattedText("Hello, World!").Bold(true));

            // Create the primary footer.
            var footer = mez.Section.Footers.Primary;

            // _THHO TODO Add support for Headers and Footers to MEZ?
            // Add content to footer.
            var paragraph = footer.AddParagraph();
            paragraph.Add(new DateField() { Format = "yyyy/MM/dd HH:mm:ss" });
            paragraph.Format.Alignment = ParagraphAlignment.Center;
        }
    }
}
