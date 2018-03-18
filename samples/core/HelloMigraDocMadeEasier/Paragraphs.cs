using MigraDoc.DocumentObjectModel;
using MigraDocMadeEZ;

namespace HelloMigraDoc
{
    /// <summary>
    /// This class demonstrates MigraDoc paragraphs.
    /// </summary>
    public class Paragraphs
    {
        public static void DefineParagraphs(MigraDocMadeEZR document)
        {
            var paragraph = document.AddParagraph("Heading1", "Paragraph Layout Overview");
            paragraph.AddBookmark("Paragraphs");

            DemonstrateAlignment(document);
            DemonstrateIndent(document);
            DemonstrateFormattedText(document);
            DemonstrateBordersAndShading(document);
        }

        static void DemonstrateAlignment(MigraDocMadeEZR document)
        {
            document.AddParagraph("Heading2", "Alignment");

            document.AddParagraph("Heading3", "Left Aligned");

            document.AddParagraph(FillerText.Text).Alignment(ParagraphAlignment.Left);

            document.AddParagraph("Heading3", "Right Aligned");

            document.AddParagraph(FillerText.Text).Alignment(ParagraphAlignment.Right);

            document.AddParagraph("Heading3", "Centered");

            document.AddParagraph(FillerText.Text).Alignment(ParagraphAlignment.Center);

            document.AddParagraph("Heading3", "Justified");

            document.AddParagraph(FillerText.MediumText).Alignment(ParagraphAlignment.Justify);
        }

        static void DemonstrateIndent(MigraDocMadeEZR document)
        {
            document.AddParagraph("Heading2", "Indent");

            document.AddParagraph("Heading3", "Left Indent");

            document.AddParagraph(FillerText.Text).LeftIndent("2cm");

            document.AddParagraph("Heading3", "Right Indent");

            document.AddParagraph(FillerText.Text).RightIndent("1in");

            document.AddParagraph("Heading3", "First Line Indent");

            document.AddParagraph(FillerText.Text).FirstLineIndent("12mm");

            document.AddParagraph("Heading3", "First Line Negative Indent");

            document.AddParagraph(FillerText.Text).LeftIndent("1.5cm").FirstLineIndent("-1.5cm");
        }

        static void DemonstrateFormattedText(MigraDocMadeEZR document)
        {
            document.AddParagraph("Heading2", "Formatted Text");

            var paragraph = document.AddParagraph();
            paragraph.AddText("Text can be formatted ");
            paragraph.AddFormattedText(new MezFormattedText("bold").Bold(true));
            paragraph.AddText(", ");
            paragraph.AddFormattedText(new MezFormattedText("italic").Italic(true));
            paragraph.AddText(", or ");
            paragraph.AddFormattedText(new MezFormattedText("bold & italic").Bold(true).Italic(true));
            paragraph.AddText(".");
            paragraph.AddLineBreak();
            paragraph.AddText("You can set the ");
            paragraph.AddFormattedText(new MezFormattedText("size").Font(15));
            paragraph.AddText(", the ");
            paragraph.AddFormattedText(new MezFormattedText("color").Color(Colors.Firebrick));
            paragraph.AddText(", the ");
            // Times New Roman looks smaller than Segoe UI, so we make it a bit larger.
            paragraph.AddFormattedText(new MezFormattedText("font").Font("Times New Roman", 12));
            paragraph.AddText(".");
            paragraph.AddLineBreak();
            paragraph.AddText("You can set the ");
            paragraph.AddFormattedText(new MezFormattedText("subscript").Subscript(true));
            paragraph.AddText(" or ");
            paragraph.AddFormattedText(new MezFormattedText("superscript").Superscript(true));
            paragraph.AddText(".");
        }

        static void DemonstrateBordersAndShading(MigraDocMadeEZR document)
        {
            document.AddPageBreak();
            document.AddParagraph("Heading2", "Borders and Shading");

            document.AddParagraph("Heading3", "Border around Paragraph");

            document.AddParagraph(FillerText.MediumText).Borders(2.5, Colors.Navy).BorderDistance(3);

            document.AddParagraph("Heading3", "Shading");

            document.AddParagraph(FillerText.Text).ShadingColor(Colors.LightCoral);

            document.AddParagraph("Heading3", "Borders & Shading");

            document.AddParagraph("TextBox", FillerText.MediumText);
        }
    }
}
