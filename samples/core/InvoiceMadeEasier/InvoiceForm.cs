using System;
using System.Globalization;
using System.Xml;
using System.Xml.XPath;
using MigraDoc.DocumentObjectModel;
using MigraDoc.DocumentObjectModel.Tables;
using MigraDoc.DocumentObjectModel.Shapes;
using System.Diagnostics;
using MigraDocMadeEZ;

namespace Invoice
{
    /// <summary>
    /// Creates the invoice form.
    /// </summary>
    public class InvoiceForm
    {
        /// <summary>
        /// The MigraDoc document that represents the invoice.
        /// </summary>
        MigraDocMadeEZR _document;

        /// <summary>
        /// The root navigator for the XML document.
        /// </summary>
        readonly XPathNavigator _navigator;

        /// <summary>
        /// The text frame of the MigraDoc document that contains the address.
        /// </summary>
        TextFrame _addressFrame;

        /// <summary>
        /// The table of the MigraDoc document that contains the invoice items.
        /// </summary>
        MezTable _table;

        /// <summary>
        /// Initializes a new instance of the class InvoiceForm and opens the specified XML document.
        /// </summary>
        public InvoiceForm(string filename)
        {
            var invoice = new XmlDocument();
            // An XML invoice based on a sample created with Microsoft InfoPath.
            invoice.Load(filename);
            _navigator = invoice.CreateNavigator();
        }

        /// <summary>
        /// Creates the invoice document.
        /// </summary>
        public MigraDocMadeEZR CreateDocument()
        {
            // Create a new MigraDoc document.
            _document = new MigraDocMadeEZR();
            _document.InfoTitle = "A sample invoice";
            _document.InfoSubject = "Demonstrates how to create an invoice.";
            _document.InfoAuthor = "Stefan Lange";

            DefineStyles();

            CreatePage();

            FillContent();

            return _document;
        }

        /// <summary>
        /// Defines the styles used to format the MigraDoc document.
        /// </summary>
        void DefineStyles()
        {
            // Modify the predefined style Normal.
            // Because all styles are derived from Normal, the next line changes the 
            // font of the whole document. Or, more exactly, it changes the font of
            // all styles and paragraphs that do not redefine the font.
            var style = _document.Style(StyleNames.Normal).Font("Segoe UI");

            _document.Style(StyleNames.Header).SetTabStop("16cm", TabAlignment.Right);

            _document.Style(StyleNames.Footer).SetTabStop("8cm", TabAlignment.Center);

            // Create a new style called Table based on style Normal.
            _document.AddStyle("Table", StyleNames.Normal).Font("Segoe UI Semilight", 9);

            // Create a new style called Title based on style Normal.
            _document.AddStyle("Title", StyleNames.Normal).Font("Segoe UI Semibold", 9);

            // Create a new style called Reference based on style Normal.
            _document.AddStyle("Reference", StyleNames.Normal).SpaceBefore("5mm").SpaceAfter("5mm")
                .SetTabStop("16cm", TabAlignment.Right);
        }

        /// <summary>
        /// Creates the static parts of the invoice.
        /// </summary>
        void CreatePage()
        {
            var section = _document.Section;
            // Define the page setup. We use an image in the header, therefore the
            // default top margin is too small for our invoice.
            // We increase the TopMargin to prevent the document body from overlapping the page header.
            // We have an image of 3.5 cm height in the header.
            // The default position for the header is 1.25 cm.
            // We add 0.5 cm spacing between header image and body and get 5.25 cm.
            // Default value is 2.5 cm.
            section.PageSetup.TopMargin = "5.25cm";

            // Put the logo in the header.
#if true
            var image = new MezImage("../../../../assets/images/MigraDoc.png")
                .Height("3.5cm")
                .LockAspectRatio(true)
                .RelativeVertical(RelativeVertical.Line)
                .RelativeHorizontal(RelativeHorizontal.Margin)
                .Top(ShapePosition.Top)
                .Left(ShapePosition.Right)
                .WrapFormatStyle(WrapStyle.Through);
            section.Headers.Primary.Add(image);
#else
            var image = section.Headers.Primary.AddImage("../../../../assets/images/MigraDoc.png");
            image.Height = "3.5cm";
            image.LockAspectRatio = true;
            image.RelativeVertical = RelativeVertical.Line;
            image.RelativeHorizontal = RelativeHorizontal.Margin;
            image.Top = ShapePosition.Top;
            image.Left = ShapePosition.Right;
            image.WrapFormat.Style = WrapStyle.Through;
#endif

            // Create the footer.
            var paragraph2 = new MezParagraph("PowerBooks Inc ● Sample Street 42 ● 56789 Cologne ● Germany")
                .Alignment(ParagraphAlignment.Center).Font(9);
            section.Footers.Primary.Add(paragraph2);

            // Create the text frame for the address.
            _addressFrame = section.AddTextFrame();
            _addressFrame.Height = "3.0cm";
            _addressFrame.Width = "7.0cm";
            _addressFrame.Left = ShapePosition.Left;
            _addressFrame.RelativeHorizontal = RelativeHorizontal.Margin;
            _addressFrame.Top = "5.0cm";
            _addressFrame.RelativeVertical = RelativeVertical.Page;

            // Show the sender in the address frame.
            new MezParagraph(_addressFrame.AddParagraph("PowerBooks Inc · Sample Street 42 · 56789 Cologne"))
                .Font(7).Bold(true).SpaceAfter(3);

            _document.AddParagraph().LineSpacingRule(LineSpacingRule.Exactly).LineSpacing("5.25cm");

            // Add the print date field.
            _document.AddParagraph()
            // We use SpaceBefore to move the first text line below the address field.
                .SpaceBefore(0)
                .Style("Reference")
                .AddFormattedText(new MezFormattedText("INVOICE").Bold(true))
                .AddTab()
                .AddText("Cologne, ")
                .AddDateField("dd.MM.yyyy");

            // Create the item table.
            _table = _document.AddTable("1cm;C|2.5cm;R|3cm;R|3.5cm;R|2cm;C|4cm;R")
                .Style("Table")
                .BorderColor(TableBorder)
                .BorderWidth(0.25)
                .BorderLeft(0.5)
                .BorderRight(0.5)
                .LeftIndent(0);

            //// Before you can add a row, you must define the columns.
            //var column = _table.AddColumn("1cm");
            //column.Format.Alignment = ParagraphAlignment.Center;

            //column = _table.AddColumn("2.5cm");
            //column.Format.Alignment = ParagraphAlignment.Right;

            //column = _table.AddColumn("3cm");
            //column.Format.Alignment = ParagraphAlignment.Right;

            //column = _table.AddColumn("3.5cm");
            //column.Format.Alignment = ParagraphAlignment.Right;

            //column = _table.AddColumn("2cm");
            //column.Format.Alignment = ParagraphAlignment.Center;

            //column = _table.AddColumn("4cm");
            //column.Format.Alignment = ParagraphAlignment.Right;

            // Create the header of the table.
            var row = _document.AddRow(new MezParagraph("Item")/*.Bold(true).Alignment(ParagraphAlignment.Left)*/,
                    new MezParagraph("Title and Author")/*.Alignment(ParagraphAlignment.Left)*/,
                    null, null, null,
                    new MezParagraph("Extended Price")/*.Alignment(ParagraphAlignment.Left)*/)
                .Heading(true)
                .Alignment(ParagraphAlignment.Center)
                .Bold(true)
                .ShadingColor(TableBlue);

            // Special settings not yet covered by MEZ.
            row.Row.Cells[0].Format.Font.Bold = false;
            row.Row.Cells[0].Format.Alignment = ParagraphAlignment.Left;
            row.Row.Cells[0].VerticalAlignment = VerticalAlignment.Bottom;
            row.Row.Cells[0].MergeDown = 1;
            row.Row.Cells[1].Format.Alignment = ParagraphAlignment.Left;
            row.Row.Cells[1].MergeRight = 3;
            row.Row.Cells[5].Format.Alignment = ParagraphAlignment.Left;
            row.Row.Cells[5].VerticalAlignment = VerticalAlignment.Bottom;
            row.Row.Cells[5].MergeDown = 1;

            row = _document.AddRow(null,
                    new MezParagraph("Quantity")/*.Alignment(ParagraphAlignment.Left)*/,
                    new MezParagraph("Unit Price")/*.Alignment(ParagraphAlignment.Left)*/,
                    new MezParagraph("Discount (%)")/*.Alignment(ParagraphAlignment.Left)*/,
                    new MezParagraph("Taxable")/*.Alignment(ParagraphAlignment.Left)*/,
                    null)
                .Heading(true)
                .Alignment(ParagraphAlignment.Center)
                .Bold(true)
                .ShadingColor(TableBlue);
            row.Row.Cells[1].Format.Alignment = ParagraphAlignment.Left;
            row.Row.Cells[2].Format.Alignment = ParagraphAlignment.Left;
            row.Row.Cells[3].Format.Alignment = ParagraphAlignment.Left;
            row.Row.Cells[4].Format.Alignment = ParagraphAlignment.Left;

            _table.Table.SetEdge(0, 0, 6, 2, Edge.Box, BorderStyle.Single, 0.75, Color.Empty);
        }

        /// <summary>
        /// Creates the dynamic parts of the invoice.
        /// </summary>
        void FillContent()
        {
            const double vat = 0.07;

            // Fill the address in the address text frame.
            var item = SelectItem("/invoice/to");
            var paragraph = _addressFrame.AddParagraph();
            paragraph.AddText(GetValue(item, "name/singleName"));
            paragraph.AddLineBreak();
            paragraph.AddText(GetValue(item, "address/line1"));
            paragraph.AddLineBreak();
            paragraph.AddText(GetValue(item, "address/postalCode") + " " + GetValue(item, "address/city"));

            // Iterate the invoice items.
            double totalExtendedPrice = 0;
            var iter = _navigator.Select("/invoice/items/*");
            while (iter.MoveNext())
            {
                item = iter.Current;
                var quantity = GetValueAsDouble(item, "quantity");
                var price = GetValueAsDouble(item, "price");
                var discount = GetValueAsDouble(item, "discount");

                var extendedPrice = quantity * price;
                extendedPrice = extendedPrice * (100 - discount) / 100;

                // Each item fills two rows.
                var row1 = _document.AddRow(new MezParagraph(GetValue(item, "itemNumber")),
                    // $THHO TODO Simplify this.
                    new MezParagraph()
                        .AddFormattedText(new MezFormattedText(GetValue(item, "title")).Style("Title"))
                        .AddFormattedText(new MezFormattedText(" by ").Italic(true))
                        .AddText(GetValue(item, "author")),
                    null, null, null,
                    new MezParagraph(extendedPrice.ToString("0.00") + " €"));
                var row2 = _document.AddRow(null,
                    new MezParagraph(GetValue(item, "quantity")),
                    new MezParagraph(price.ToString("0.00") + " €"),
                    new MezParagraph(discount > 0 ? (discount.ToString("0") + '%') : ""),
                    null,
                    new MezParagraph(price.ToString("0.00"))); // Hidden by MergeDown.
                // $THHO TODO Simplify this?
                row1.Row.TopPadding = 1.5;
                row1.Row.Cells[0].Shading.Color = TableGray;
                row1.Row.Cells[0].VerticalAlignment = VerticalAlignment.Center;
                row1.Row.Cells[0].MergeDown = 1;
                row1.Row.Cells[1].Format.Alignment = ParagraphAlignment.Left;
                row1.Row.Cells[1].MergeRight = 3;
                row1.Row.Cells[5].Shading.Color = TableGray;
                row1.Row.Cells[5].MergeDown = 1;

                //row1.Cells[0].AddParagraph(GetValue(item, "itemNumber"));
                //paragraph = row1.Cells[1].AddParagraph();
                //var formattedText = new FormattedText() { Style = "Title" };
                //formattedText.AddText(GetValue(item, "title"));
                //paragraph.Add(formattedText);
                //paragraph.AddFormattedText(" by ", TextFormat.Italic);
                //paragraph.AddText(GetValue(item, "author"));

                //row2.Cells[1].AddParagraph(GetValue(item, "quantity"));
                //row2.Cells[2].AddParagraph(price.ToString("0.00") + " €");
                //if (discount > 0)
                //    row2.Cells[3].AddParagraph(discount.ToString("0") + '%');
                //row2.Cells[4].AddParagraph();
                //row2.Cells[5].AddParagraph(price.ToString("0.00"));
                //row1.Cells[5].AddParagraph(extendedPrice.ToString("0.00") + " €");
                row1.Row.Cells[5].VerticalAlignment = VerticalAlignment.Bottom;
                totalExtendedPrice += extendedPrice;

                _table.Table.SetEdge(0, _table.Table.Rows.Count - 2, 6, 2, Edge.Box, BorderStyle.Single, 0.75);
            }

            // Add an invisible row as a space line to the table.
            var row = _document.AddRow();
            row.Row.Borders.Visible = false;

            // Add the total price row.
            row = _document.AddRow(new MezParagraph("Total Price")/*.Bold(true).Alignment(ParagraphAlignment.Right)*/,
                null, null, null, null,
                new MezParagraph(totalExtendedPrice.ToString("0.00") + " €")/*.Font("Segoe UI")*/);
            row.Row.Cells[0].Borders.Visible = false;
            //row.Cells[0].AddParagraph("Total Price");
            row.Row.Cells[0].Format.Font.Bold = true;
            row.Row.Cells[0].Format.Alignment = ParagraphAlignment.Right;
            row.Row.Cells[0].MergeRight = 4;
            //row.Cells[5].AddParagraph(totalExtendedPrice.ToString("0.00") + " €");
            row.Row.Cells[5].Format.Font.Name = "Segoe UI";

            // Add the VAT row.
            row = _document.AddRow(new MezParagraph("VAT (" + (vat * 100) + "%)")/*.Bold(true).Alignment(ParagraphAlignment.Right)*/,
                null, null, null, null,
                new MezParagraph((vat * totalExtendedPrice).ToString("0.00") + " €"));
            row.Row.Cells[0].Borders.Visible = false;
            //row.Cells[0].AddParagraph("VAT (7%)");
            row.Row.Cells[0].Format.Font.Bold = true;
            row.Row.Cells[0].Format.Alignment = ParagraphAlignment.Right;
            row.Row.Cells[0].MergeRight = 4;
            //row.Cells[5].AddParagraph((0.07 * totalExtendedPrice).ToString("0.00") + " €");

            // Add the additional fee row.
            row = _document.AddRow(new MezParagraph("Shipping and Handling")/*.Bold(true).Alignment(ParagraphAlignment.Right)*/,
                null, null, null, null,
                new MezParagraph(0.ToString("0.00") + " €"));
            row.Row.Cells[0].Borders.Visible = false;
            //row.Cells[0].AddParagraph("Shipping and Handling");
            //row.Cells[5].AddParagraph(0.ToString("0.00") + " €");
            row.Row.Cells[0].Format.Font.Bold = true;
            row.Row.Cells[0].Format.Alignment = ParagraphAlignment.Right;
            row.Row.Cells[0].MergeRight = 4;

            // Add the total due row.
            totalExtendedPrice += vat * totalExtendedPrice;
            row = _document.AddRow(new MezParagraph("Total Due")/*.Bold(true).Alignment(ParagraphAlignment.Right)*/,
                null, null, null, null,
                new MezParagraph(totalExtendedPrice.ToString("0.00") + " €")/*.Bold(true).Font("Segoe UI")*/);
            //row.Cells[0].AddParagraph("Total Due");
            row.Row.Cells[0].Borders.Visible = false;
            row.Row.Cells[0].Format.Font.Bold = true;
            row.Row.Cells[0].Format.Alignment = ParagraphAlignment.Right;
            row.Row.Cells[0].MergeRight = 4;
            //row.Cells[5].AddParagraph(totalExtendedPrice.ToString("0.00") + " €");
            row.Row.Cells[5].Format.Font.Name = "Segoe UI";
            row.Row.Cells[5].Format.Font.Bold = true;

            // Set the borders of the specified cell range.
            _table.Table.SetEdge(5, _table.Table.Rows.Count - 4, 1, 4, Edge.Box, BorderStyle.Single, 0.75);

            // Add the notes paragraph.
            item = SelectItem("/invoice");
            //paragraph.AddText(GetValue(item, "notes"));
            /*var paragraph2 =*/ _document.AddParagraph(GetValue(item, "notes"))
                .Alignment(ParagraphAlignment.Center)
                .SpaceBefore("1cm")
                .Borders(0.75)
                .BorderDistance(3)
                .Borders(TableBorder)
                .ShadingColor(TableGray);
            //paragraph2.Format.Alignment = ParagraphAlignment.Center;
            //paragraph2.Format.SpaceBefore = "1cm";
            //paragraph2.Format.Borders.Width = 0.75;
            //paragraph2.Format.Borders.Distance = 3;
            //paragraph2.Format.Borders.Color = TableBorder;
            //paragraph2.Format.Shading.Color = TableGray;
        }

        /// <summary>
        /// Selects a subtree in the XML data.
        /// </summary>
        XPathNavigator SelectItem(string path)
        {
            var iter = _navigator.Select(path);
            iter.MoveNext();
            return iter.Current;
        }

        /// <summary>
        /// Gets an element value from the XML data.
        /// </summary>
        static string GetValue(XPathNavigator nav, string name)
        {
            //nav = nav.Clone();
            var iter = nav.Select(name);
            iter.MoveNext();
            return iter.Current.Value;
        }

        /// <summary>
        /// Gets an element value as double from the XML data.
        /// </summary>
        static double GetValueAsDouble(XPathNavigator nav, string name)
        {
            try
            {
                var value = GetValue(nav, name);
                if (value.Length == 0)
                    return 0;
                return Double.Parse(value, CultureInfo.InvariantCulture);
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.Message);
            }
            return 0;
        }

        // Some pre-defined colors...
#if true
        // ... in RGB.
        readonly static Color TableBorder = new Color(81, 125, 192);
        readonly static Color TableBlue = new Color(235, 240, 249);
        readonly static Color TableGray = new Color(242, 242, 242);
#else
        // ... in CMYK.
        readonly static Color TableBorder = Color.FromCmyk(100, 50, 0, 30);
        readonly static Color TableBlue = Color.FromCmyk(0, 80, 50, 30);
        readonly static Color TableGray = Color.FromCmyk(30, 0, 0, 0, 100);
#endif
    }
}
