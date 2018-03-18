using MigraDoc.DocumentObjectModel;
using MigraDocMadeEZ;

namespace HelloMigraDoc
{
    /// <summary>
    /// Defines the styles used in the document.
    /// </summary>
    public class Styles
    {
        /// <summary>
        /// Defines the styles used in the document.
        /// </summary>
        public static void DefineStyles(MigraDocMadeEZR document)
        {
            // Change the predefined style Normal.
            // Because all styles are derived from Normal, the next line changes the 
            // font of the whole document. Or, more exactly, it changes the font of
            // all styles and paragraphs that do not redefine the font.
            document.Style(StyleNames.Normal).Font("Segoe UI");

            // Heading1 to Heading9 are predefined styles with an outline level. An outline level
            // other than OutlineLevel.BodyText automatically creates the outline (or bookmarks) 
            // in PDF.

            document.Style(StyleNames.Heading1).Font("Segoe UI Light", 16)
                .Color(Colors.DarkBlue).SpaceAfter(6).PageBreakBefore(true).KeepWithNext(true);
            // Set KeepWithNext for all headings to prevent headings from appearing all alone
            // at the bottom of a page. The other headings inherit this from Heading1.

            document.Style(StyleNames.Heading2).Font(14).SpaceBefore(6).SpaceAfter(6).PageBreakBefore(false);

            document.Style(StyleNames.Heading3).Font(12).Italic(true).SpaceBefore(6).SpaceAfter(3);

            document.Style(StyleNames.Header).SetTabStop("16cm", TabAlignment.Right);

            document.Style(StyleNames.Footer).SetTabStop("8cm", TabAlignment.Center);

            // Create a new style called TextBox based on style Normal.
            document.AddStyle("TextBox", StyleNames.Normal).Alignment(ParagraphAlignment.Justify)
                .Borders(2.5).BorderDistance("3pt").ShadingColor(Colors.SkyBlue);

            // Create a new style called TOC based on style Normal.
            document.AddStyle("TOC", StyleNames.Normal)
                .SetTabStop("16cm", TabAlignment.Right, TabLeader.Dots).Color(Colors.Blue);
        }
    }
}
