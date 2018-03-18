using MigraDoc.DocumentObjectModel;
using MigraDocMadeEZ;

namespace HelloMigraDoc
{
    class Documents
    {
        public static MigraDocMadeEZR CreateDocument()
        {
            // Create a new MigraDoc document.
            var document = new MigraDocMadeEZR();
            document.InfoTitle = "Hello, MigraDoc";
            document.InfoSubject = "Demonstrates an excerpt of the capabilities of MigraDoc.";
            document.InfoAuthor = "Stefan Lange (modifications by Thomas Hövel)";

            Styles.DefineStyles(document);

            Cover.DefineCover(document);
            TableOfContents.DefineTableOfContents(document);

            DefineContentSection(document);

            Paragraphs.DefineParagraphs(document);
            Tables.DefineTables(document);
            Charts.DefineCharts(document);

            return document;
        }

        /// <summary>
        /// Defines page setup, headers, and footers.
        /// </summary>
        static void DefineContentSection(MigraDocMadeEZR document)
        {
            var section = document.AddSection();
            section.PageSetup.OddAndEvenPagesHeaderFooter = true;
            section.PageSetup.StartingNumber = 1;

            var header = section.Headers.Primary;
            header.AddParagraph("\tOdd Page Header");

            header = section.Headers.EvenPage;
            header.AddParagraph("Even Page Header");

            // Create a paragraph with centered page number. See definition of style "Footer".
            var paragraph = new Paragraph();
            paragraph.AddTab();
            paragraph.AddPageField();

            // Add paragraph to footer for odd pages.
            section.Footers.Primary.Add(paragraph);
            // Add clone of paragraph to footer for odd pages. Cloning is necessary because an object must
            // not belong to more than one other object. If you forget cloning an exception is thrown.
            section.Footers.EvenPage.Add(paragraph.Clone());
        }
    }
}
