/*
Copyright © 2015-2016 Thomas Hövel Software, Troisdorf, Germany.

Permission is hereby granted, free of charge, to any person obtaining a
copy of this software and associated documentation files (the
"Software"), to deal in the Software without restriction, including
without limitation the rights to use, copy, modify, merge, publish,
distribute, sublicense, and/or sell copies of the Software, and to
permit persons to whom the Software is furnished to do so, subject to
the following conditions:

* The above copyright notice and this permission notice shall be
  included in all copies or substantial portions of the Software.

* THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND,
  EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF
  MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
  NONINFRINGEMENT. IN NO EVENT SHALL THE X CONSORTIUM BE LIABLE FOR ANY
  CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT,
  TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE
  SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

* Except as contained in this notice, the name of Thomas Hövel Software
  or any trademarks of Thomas Hövel Software shall not be used in
  advertising or otherwise to promote the sale, use or other dealings
  in this Software without prior written authorization
  from Thomas Hövel Software.

MigraDocMadeEZR and "MigraDoc Made EZR" are trademarks
of Thomas Hövel Software.
MigraDoc® is a registered trademark of empira Software GmbH and is
used here only to indicate the purpose of "MigraDoc Made EZR":
    To make using MigraDoc easier.
  
 */

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using MigraDoc.DocumentObjectModel;
using MigraDoc.DocumentObjectModel.Fields;
using MigraDoc.DocumentObjectModel.IO;
using MigraDoc.DocumentObjectModel.Shapes;
using MigraDoc.DocumentObjectModel.Shapes.Charts;
using MigraDoc.DocumentObjectModel.Tables;
using MigraDoc.Rendering;
using MigraDoc.RtfRendering;

namespace MigraDocMadeEZ
{
    /// <summary>
    /// MigraDocMadeEZR is the entry point to creating MigraDoc
    /// documents easier. Instantiating a new instance of
    /// MigraDocMadeEZR will automatically create a
    /// MigraDoc document with one section with initialized
    /// PageSetup, ready to receive contents using the methods
    /// of MigraDocMadeEZR or using MigraDoc methods directly.
    /// </summary>
    public class MigraDocMadeEZR
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="MigraDocMadeEZR"/> class.
        /// </summary>
        /// <param name="format">The page format.</param>
        public MigraDocMadeEZR(PageFormat format = PageFormat.A4)
        {
            Document = new Document();
            AddSection(format);
        }

        public Document Document { get; private set; }
        public Section Section { get { return Document.LastSection; } }

        public string InfoAuthor
        {
            get { return Document.Info.Author; }
            set { Document.Info.Author = value; }
        }

        public string InfoComment
        {
            get { return Document.Info.Comment; }
            set { Document.Info.Comment = value; }
        }

        public string InfoKeywords
        {
            get { return Document.Info.Keywords; }
            set { Document.Info.Keywords = value; }
        }

        public string InfoSubject
        {
            get { return Document.Info.Subject; }
            set { Document.Info.Subject = value; }
        }

        public string InfoTitle
        {
            get { return Document.Info.Title; }
            set { Document.Info.Title = value; }
        }

        #region Section Methods
        public Section AddSection(PageFormat format = PageFormat.A4)
        {
            Document.AddSection();
            // Magic: make sure we have a PageSetup without null values.
            Section.PageSetup = Document.DefaultPageSetup.Clone();
            Section.PageSetup.PageFormat = format;
            return Section;
        }

        #endregion

        #region Style Methods

        public MezStyle Style(string style)
        {
            var mdStyle = Document.Styles[style];
            return new MezStyle(mdStyle);
        }

        public MezStyle AddStyle(string style, string baseStyle = StyleNames.Normal)
        {
            var mdStyle = Document.Styles.AddStyle(style, baseStyle);
            return new MezStyle(mdStyle);
        }
        #endregion

        #region Paragraph Methods

        public void AddPageBreak()
        {
            Section.AddPageBreak();
        }

        public MezParagraph AddParagraph()
        {
            return new MezParagraph(Section.AddParagraph());
        }

        public MezParagraph AddParagraph(Paragraph para)
        {
            Section.Add(para);
            return new MezParagraph(para);
        }

        public MezParagraph AddParagraph(string text)
        {
            return new MezParagraph(text != null ?
                Section.AddParagraph(text) :
                Section.AddParagraph());
        }

        public MezParagraph AddParagraph(string style, string text)
        {
            return AddParagraph(text).Style(style);
        }

        public MezParagraph AddHeading1(string text)
        {
            return AddParagraph(StyleNames.Heading1, text);
        }

        public MezParagraph AddHeading2(string text)
        {
            return AddParagraph(StyleNames.Heading2, text);
        }

        public MezParagraph AddHeading3(string text)
        {
            return AddParagraph(StyleNames.Heading3, text);
        }

        public MezParagraph AddHeading4(string text)
        {
            return AddParagraph(StyleNames.Heading4, text);
        }

        public MezParagraph AddHeading5(string text)
        {
            return AddParagraph(StyleNames.Heading5, text);
        }

        public MezParagraph AddHeading6(string text)
        {
            return AddParagraph(StyleNames.Heading6, text);
        }

        public MezParagraph AddHeading7(string text)
        {
            return AddParagraph(StyleNames.Heading7, text);
        }

        public MezParagraph AddHeading8(string text)
        {
            return AddParagraph(StyleNames.Heading8, text);
        }

        public MezParagraph AddHeading9(string text)
        {
            return AddParagraph(StyleNames.Heading9, text);
        }

        #endregion

        #region Paragraph Methods II

        // TODO static or not?
        public /*static*/ MezParagraph NewParagraph()
        {
            return new MezParagraph(new Paragraph());
        }

        public /*static*/ MezFormattedText NewFormattedText()
        {
            return new MezFormattedText();
        }

        #endregion

        #region Hyperlink Methods
        /// <summary>
        /// Adds a new Hyperlink as a new paragraph. To insert a Hyperlink into a paragraph use the Hyperlink function of the paragraph.
        /// </summary>
        /// <param name="type">The type.</param>
        public MezHyperlink AddHyperLink(HyperlinkType type = HyperlinkType.Local)
        {
            var result = new MezHyperlink(new Hyperlink()).Type(type);
            Section.AddParagraph().Add(result);
            return result;
        }

        public MezHyperlink AddLocalLink(string bookmark, string text = null)
        {
            return AddHyperLink(HyperlinkType.Bookmark).Target(bookmark).AddText(text);
        }

        public MezHyperlink AddWebLink(string url, string text = null)
        {
            return AddHyperLink(HyperlinkType.Web).Target(url).AddText(text);
        }

        public MezHyperlink AddFileLink(string file, string text = null)
        {
            return AddHyperLink(HyperlinkType.File).Target(file).AddText(text);
        }
        #endregion

        #region Rendering Methods

        public void MakePdf(string filename, bool execute = false, bool unicode = true)
        {
            var doc = ClonedDocument();
            var renderer = new PdfDocumentRenderer(unicode) { Document = doc };
            renderer.RenderDocument();
            renderer.PdfDocument.Save(filename);
            if (execute)
                Process.Start(filename);
        }

        public void MakePdf(Stream stream, bool keepOpen = false)
        {
            var doc = ClonedDocument();
            var renderer = new PdfDocumentRenderer(true) { Document = doc };
            renderer.RenderDocument();
            renderer.PdfDocument.Save(stream, !keepOpen);
        }

        public void MakeRtf(string filename, bool execute = false, string workingDirectory = null)
        {
            var doc = ClonedDocument();
            var renderer = new RtfDocumentRenderer();
            renderer.Render(doc, filename, workingDirectory);
            if (execute)
                Process.Start(filename);
        }

        public void MakeRtf(Stream stream, bool keepOpen = false, string workingDirectory = null)
        {
            var doc = ClonedDocument();
            var renderer = new RtfDocumentRenderer();
#if MIGRADOC132
            if (!keepOpen)
            {
                renderer.Render(doc, stream, workingDirectory);
                return;
            }
            var str = renderer.RenderToString(doc, workingDirectory);

            StreamWriter strmWrtr = null;
            try
            {
                strmWrtr = new StreamWriter(stream, System.Text.Encoding.Default);
                strmWrtr.Write(str);
            }
            finally
            {
                if (strmWrtr != null)
                {
                    strmWrtr.Flush();
                    if (stream != null)
                    {
                        //if (!keepOpen)
                        //    strmWrtr.Close();
                        //else
                        stream.Position = 0; // Reset the stream position if the stream is kept open.
                    }
                }
            }
#else
            renderer.Render(doc, stream, !keepOpen, workingDirectory);
#endif
        }

        public void SaveMDDDL(string filename)
        {
            DdlWriter dw = new DdlWriter(filename);
            dw.WriteDocument(Document);
            dw.Close();
        }

        private Document ClonedDocument()
        {
#if MIGRADOC132
            // Workaround for bug in Clone() implementation (fixed with PDFsharp 1.50 beta 2).
            var str = DdlWriter.WriteToString(Document);
            var result = DdlReader.DocumentFromString(str);
            return result;
#else
            return Document.Clone();
#endif
        }
        #endregion

        #region Table Methods

        // TODO Column alignment.
        // TODO Helpers to format tables (colours, borders, ...).

        public MezTable AddTable(string columns)
        {
            var cols = columns.Split('|');
            return AddTable(cols);
        }

        public MezTable AddTable(params string[] columns)
        {
            var cols = new List<MezColumn>();
            foreach (var item in columns)
            {
                var elements = item.Split(';');
                // TODO Column alignment.
                if (elements.Length != 1)
                    throw new ArgumentException("AddTable: Invalid column specification \"" + item + "\": wrong item count.");
                var firstElement = elements[0];
                MezColumn col;
                if (firstElement.EndsWith("*"))
                {
                    // Star column.
                    uint stars;
                    if (uint.TryParse(firstElement.Substring(0, firstElement.Length - 1), out stars))
                    {
                        col = new MezStarColumn(stars);
                        cols.Add(col);
                    }
                    else
                        throw new ArgumentException("AddTable: Invalid column specification \"" + item + "\": invalid first element.");
                }
                else
                {
                    // Column with explicit width.
                    col = new MezUnitColumn(firstElement);
                    cols.Add(col);
                }

                if (elements.Length < 2)
                    continue;
                var secondElement = elements[1];
                switch (secondElement)
                {
                    case "L":
                        col.Alignment = ParagraphAlignment.Left;
                        break;
                    case "R":
                        col.Alignment = ParagraphAlignment.Right;
                        break;
                    case "C":
                        col.Alignment = ParagraphAlignment.Center;
                        break;
                    case "J":
                        col.Alignment = ParagraphAlignment.Justify;
                        break;
                }

                if (elements.Length < 3)
                    continue;
                var thirdElement = elements[2];
                col.Style = thirdElement;
            }
            return AddTable(cols.ToArray());
        }

        public MezTable AddTable(params MezColumn[] columns)
        {
            var result = Section.AddTable();
            var bodyWidth = BodyWidth();
            //int starColumns = 0;
            uint starFactors = 0;
            Unit fixedWidth = 0;

            // Evaluate the columns.
            foreach (var item in columns)
            {
                var star = item as MezStarColumn;
                if (star != null)
                {
                    //++starColumns;
                    starFactors += star.StarFactor;
                    continue;
                }
                var unit = item as MezUnitColumn;
                if (unit != null)
                {
                    fixedWidth += unit.Unit;
                    continue;
                }
                throw new NotImplementedException("Unsupported column type in \"AddTable(params MezColumn[] columns)\".");
            }

            // Create the columns.
            Unit starWidth = bodyWidth - fixedWidth;
            // Do not allow negative widths. Use 0 instead. Should we throw an exception?
            if (starWidth < 0)
                starWidth = 0;

            foreach (var item in columns)
            {
                var star = item as MezStarColumn;
                var unit = item as MezUnitColumn;
                Column col = null;
                if (star != null)
                {
                    col = result.AddColumn(starWidth * star.StarFactor / starFactors);
                }
                else if (unit != null)
                {
                    col = result.AddColumn(unit.Unit);
                }
                if (col != null)
                {
                    col.Style = item.Style;
                    if (item.Alignment.HasValue)
                        col.Format.Alignment = item.Alignment.Value;
                }
            }
            return new MezTable(result);
        }

        public MezRow AddRow(params string[] columns)
        {
            var table = Section.LastTable;
            var row = table.AddRow();
            // Should we throw an exception if columns.Length > table.Columns.Count?
            int end = Math.Min(columns.Length, table.Columns.Count);
            for (int i = 0; i < end; ++i)
            {
                row[i].AddParagraph(columns[i]);
            }
            return new MezRow(row);
        }

        public MezRow AddRow(params object[] columns)
        {
            var table = Section.LastTable;
            var row = table.AddRow();
            // Should we throw an exception if columns.Length > table.Columns.Count?
            int end = Math.Min(columns.Length, table.Columns.Count);
            for (int i = 0; i < end; ++i)
            {
                var str = columns[i] as string;
                if (str != null)
                {
                    row[i].AddParagraph(str);
                    continue;
                }
                var image = columns[i] as Image;
                if (image != null)
                {
                    row[i].Add(image);
                    continue;
                }
                var chart = columns[i] as Chart;
                if (chart != null)
                {
                    row[i].Add(chart);
                    continue;
                }
                var paragraph = columns[i] as Paragraph ?? columns[i] as MezParagraph;
                if (paragraph != null)
                {
                    row[i].Add(paragraph);
                    continue;
                }
                var textframe = columns[i] as TextFrame;
                if (textframe != null)
                {
                    row[i].Add(textframe);
                    continue;
                }
                throw new NotImplementedException("Unsupported column type in \"AddRow(params object[] columns)\".");
            }
            return new MezRow(row);
        }

        #endregion

        #region Misc Methods

        public Unit BodyWidth()
        {
            Unit left = Section.PageSetup.LeftMargin;
            Unit right = Section.PageSetup.RightMargin;
            return Section.PageSetup.PageWidth - left - right;
        }

        public void Borders(Unit all)
        {
            Borders(all, all);
        }

        public void Borders(Unit horizontal, Unit vertical)
        {
            Borders(horizontal, vertical, horizontal, vertical);
        }

        public void Borders(Unit top, Unit right, Unit bottom, Unit left)
        {
            Section.PageSetup.TopMargin = top;
            Section.PageSetup.RightMargin = right;
            Section.PageSetup.BottomMargin = bottom;
            Section.PageSetup.LeftMargin = left;
        }

        // TODO More PageSetup methods?


        #endregion

        #region Image Methods
        // Image file name from Stream etc.
        public MezImage AddImage(string filename)
        {
            var result = new MezImage(Section.AddImage(filename));
            return result;
        }
        #endregion

        #region Header/Footer Methods
        // TODO Simplify standard headers/footers with left/centered/right text.

        // TODO Even and FirstPage header/footer.
        public HeaderFooter AddPrimaryHeader(string left, string centered, string right, bool clear = true)
        {
            var hf = Section.Headers.Primary;
            return AddHeaderFooter(hf, left, centered, right, clear);
        }

        public HeaderFooter AddPrimaryHeader(object left, object centered, object right, bool clear = true)
        {
            var hf = Section.Headers.Primary;
            return AddHeaderFooter(hf, left, centered, right, clear);
        }

        public HeaderFooter AddPrimaryFooter(string left, string centered, string right, bool clear = true)
        {
            var hf = Section.Footers.Primary;
            return AddHeaderFooter(hf, left, centered, right, clear);
        }

        public HeaderFooter AddPrimaryFooter(object left, object centered, object right, bool clear = true)
        {
            var hf = Section.Footers.Primary;
            return AddHeaderFooter(hf, left, centered, right, clear);
        }

        private HeaderFooter AddHeaderFooter(HeaderFooter hf, object left, object centered, object right, bool clear)
        {
            if (clear)
            {
                hf.Elements.Clear();
                hf.Format.TabStops.ClearAll();
                var bodywidth = BodyWidth();
                hf.Format.TabStops.AddTabStop(bodywidth / 2, TabAlignment.Center);
                hf.Format.TabStops.AddTabStop(bodywidth, TabAlignment.Right);
            }

            Paragraph paragraph = null;
            if (left != null)
                AddHeaderFooterItem(ref paragraph, left);

            if (centered != null || right != null)
            {
                paragraph = paragraph ?? new Paragraph();
                paragraph.AddTab();
                AddHeaderFooterItem(ref paragraph, centered);
                if (right != null)
                {
                    paragraph.AddTab();
                    AddHeaderFooterItem(ref paragraph, right);
                }
            }
            if (paragraph != null)
                hf.Add(paragraph);

            return hf;
        }

        private void AddHeaderFooterItem(ref Paragraph paragraph, object item)
        {
            var str = item as string;
            if (!string.IsNullOrEmpty(str))
            {
                paragraph = paragraph ?? new Paragraph();
                paragraph.AddText(str);
                return;
            }
            // Magic MezParagraph has an implicit conversion to Paragraph.
            var para = item as Paragraph ?? item as MezParagraph;
            if (para != null)
            {
                foreach (DocumentObject el in para.Elements)
                {
                    var cl = (DocumentObject)el.Clone();
                    var bf = cl as BookmarkField;
                    if (bf != null) { paragraph.Add(bf); continue; }
                    var pf = cl as PageField;
                    if (pf != null) { paragraph.Add(pf); continue; }
                    var prf = cl as PageRefField;
                    if (prf != null) { paragraph.Add(prf); continue; }
                    var npf = cl as NumPagesField;
                    if (npf != null) { paragraph.Add(npf); continue; }
                    var sf = cl as SectionField;
                    if (sf != null) { paragraph.Add(sf); continue; }
                    var spf = cl as SectionPagesField;
                    if (spf != null) { paragraph.Add(spf); continue; }
                    var df = cl as DateField;
                    if (df != null) { paragraph.Add(df); continue; }
                    var @if = cl as InfoField;
                    if (@if != null) { paragraph.Add(@if); continue; }
                    var f = cl as Footnote;
                    if (f != null) { paragraph.Add(f); continue; }
                    var t = cl as Text;
                    if (t != null) { paragraph.Add(t); continue; }
                    var ft = cl as FormattedText;
                    if (ft != null) { paragraph.Add(ft); continue; }
                    var hl = cl as Hyperlink;
                    if (hl != null) { paragraph.Add(hl); continue; }
                    var i = cl as Image;
                    if (i != null) { paragraph.Add(i); continue; }
                    var ch = cl as Character;
                    if (ch != null) { paragraph.Add(ch); continue; }

                    throw new ArgumentException("Invalid item \"" + item.GetType().Name + "\" for Header/Footer.");
                }
                return;
            }
            throw new ArgumentException("Invalid item \"" + item.GetType().Name + "\" for Header/Footer.");
        }

        #endregion


        #region TODO Methods


        #endregion
    }

    public class MezStyle
    {
        public Style Style { get; private set; }

        public MezStyle(Style style)
        {
            Style = style;
        }

        public MezStyle Font(string font, Unit? fontSize = null)
        {
            Style.Font.Name = font;
            if (fontSize.HasValue)
                Style.Font.Size = fontSize.Value;
            return this;
        }

        public MezStyle Font(Unit fontSize)
        {
            Style.Font.Size = fontSize.Value;
            return this;
        }

        public MezStyle SpaceBefore(Unit space)
        {
            Style.ParagraphFormat.SpaceBefore = space;
            return this;
        }

        public MezStyle SpaceAfter(Unit space)
        {
            Style.ParagraphFormat.SpaceAfter = space;
            return this;
        }

        public MezStyle Bold(bool value)
        {
            //Style.Image.Bold = value;
            Style.ParagraphFormat.Font.Bold = value;
            return this;
        }

        public MezStyle Italic(bool value)
        {
            //Style.Image.Italic = value;
            Style.ParagraphFormat.Font.Italic = value;
            return this;
        }

        public MezStyle Underline(bool value)
        {
            Style.ParagraphFormat.Font.Underline = value ?
                MigraDoc.DocumentObjectModel.Underline.Single :
                MigraDoc.DocumentObjectModel.Underline.None;
            return this;
        }

        public MezStyle Color(Color value)
        {
            Style.ParagraphFormat.Font.Color = value;
            return this;
        }

        public MezStyle LeftIndent(Unit unit)
        {
            Style.ParagraphFormat.LeftIndent = unit;
            return this;
        }

        public MezStyle RightIndent(Unit unit)
        {
            Style.ParagraphFormat.RightIndent = unit;
            return this;
        }

        public MezStyle FirstLineIndent(Unit unit)
        {
            Style.ParagraphFormat.FirstLineIndent = unit;
            return this;
        }

        public MezStyle SetTabStop(Unit unit, TabAlignment alignment, TabLeader leader, bool clearAll = false)
        {
            if (clearAll)
                Style.ParagraphFormat.TabStops.ClearAll();
            Style.ParagraphFormat.TabStops.AddTabStop(unit, alignment, leader);
            return this;
        }

        public MezStyle SetTabStop(Unit unit, TabAlignment alignment, bool clearAll = false)
        {
            if (clearAll)
                Style.ParagraphFormat.TabStops.ClearAll();
            Style.ParagraphFormat.TabStops.AddTabStop(unit, alignment);
            return this;
        }

        public MezStyle SetTabStop(Unit unit, TabLeader leader, bool clearAll = false)
        {
            if (clearAll)
                Style.ParagraphFormat.TabStops.ClearAll();
            Style.ParagraphFormat.TabStops.AddTabStop(unit, leader);
            return this;
        }

        public MezStyle SetTabStop(Unit unit, bool clearAll = false)
        {
            if (clearAll)
                Style.ParagraphFormat.TabStops.ClearAll();
            Style.ParagraphFormat.TabStops.AddTabStop(unit);
            return this;
        }

        public MezStyle LineSpacing(Unit unit, LineSpacingRule? rule = null)
        {
            Style.ParagraphFormat.LineSpacing = unit;
            if (rule.HasValue)
                Style.ParagraphFormat.LineSpacingRule = rule.Value;
            return this;
        }

        public MezStyle PageBreakBefore(bool pageBreakBefore)
        {
            Style.ParagraphFormat.PageBreakBefore = pageBreakBefore;
            return this;
        }

        public MezStyle KeepWithNext(bool keepWithNext)
        {
            Style.ParagraphFormat.KeepWithNext = keepWithNext;
            return this;
        }

        public MezStyle Alignment(ParagraphAlignment alignment)
        {
            Style.ParagraphFormat.Alignment = alignment;
            return this;
        }

        public MezStyle Borders(Unit width)
        {
            Style.ParagraphFormat.Borders.Width = width;
            return this;
        }

        public MezStyle Borders(Color color)
        {
            Style.ParagraphFormat.Borders.Color = color;
            return this;
        }

        public MezStyle Borders(Unit width, Color color)
        {
            Style.ParagraphFormat.Borders.Width = width;
            Style.ParagraphFormat.Borders.Color = color;
            return this;
        }

        public MezStyle BorderDistance(Unit distance)
        {
            Style.ParagraphFormat.Borders.Distance = distance;
            return this;
        }

        public MezStyle ShadingColor(Color color)
        {
            Style.ParagraphFormat.Shading.Color = color;
            return this;
        }
    }

    public abstract class MezColumn
    {
        public ParagraphAlignment? Alignment;

        public string Style;
    }

    sealed public class MezStarColumn : MezColumn
    {
        public uint StarFactor { get; private set; }

        public MezStarColumn(uint starFactor)
        {
            StarFactor = starFactor;
            var d = new Column();
            var x = d.Format.Alignment;
            // TODO Alignment.
        }
    }

    sealed public class MezUnitColumn : MezColumn
    {
        public Unit Unit { get; private set; }

        public MezUnitColumn(Unit unit)
        {
            Unit = unit;
        }
    }

    public class MezRow
    {
        public Row Row { get; private set; }

        public MezRow(Row row)
        {
            Row = row;
        }

        public MezRow Heading(bool value)
        {
            Row.HeadingFormat = value;
            return this;
        }

        public MezRow Style(string style)
        {
            Row.Style = style;
            return this;
        }

        public MezRow Alignment(ParagraphAlignment alignment)
        {
            Row.Format.Alignment = alignment;
            return this;
        }

        public MezRow Bold(bool b)
        {
            Row.Format.Font.Bold = b;
            return this;
        }

        public MezRow ShadingColor(Color color)
        {
            Row.Format.Shading.Color = color;
            return this;
        }
    }

    public class MezTable
    {
        public Table Table { get; private set; }

        public MezTable(Table table)
        {
            Table = table;
        }

        public MezTable BorderWidth(Unit width)
        {
            Table.Borders.Width = width;
            return this;
        }

        public MezTable Padding(Unit width)
        {
            TopPadding(width);
            RightPadding(width);
            BottomPadding(width);
            LeftPadding(width);
            return this;
        }

        public MezTable TopPadding(Unit width)
        {
            Table.TopPadding = width;
            return this;
        }

        public MezTable RightPadding(Unit width)
        {
            Table.RightPadding = width;
            return this;
        }

        public MezTable BottomPadding(Unit width)
        {
            Table.BottomPadding = width;
            return this;
        }

        public MezTable LeftPadding(Unit width)
        {
            Table.LeftPadding = width;
            return this;
        }

        public MezTable LeftIndent(Unit leftIndent)
        {
            Table.Rows.LeftIndent = leftIndent;
            return this;
        }

        public MezTable Style(string style)
        {
            Table.Style = style;
            return this;
        }

        public MezTable BorderColor(Color color)
        {
            Table.Borders.Color = color;
            return this;
        }

        public MezTable BorderLeft(Unit width)
        {
            Table.Borders.Left.Width = width;
            return this;
        }

        public MezTable BorderRight(Unit width)
        {
            Table.Borders.Right.Width = width;
            return this;
        }

        public MezTable BorderTop(Unit width)
        {
            Table.Borders.Top.Width = width;
            return this;
        }

        public MezTable BorderBottom(Unit width)
        {
            Table.Borders.Bottom.Width = width;
            return this;
        }
    }

    public class MezParagraph
    {
        public Paragraph Paragraph { get; private set; }

        public MezParagraph(Paragraph paragraph)
        {
            Paragraph = paragraph;
        }

        public MezParagraph()
            : this(new Paragraph())
        {
        }

        public MezParagraph(string text)
            : this()
        {
            Paragraph.AddText(text);
        }

        public static implicit operator Paragraph(MezParagraph para)
        {
            return para.Paragraph;
        }

        public MezParagraph Style(string style)
        {
            Paragraph.Style = style;
            return this;
        }

        public MezParagraph Bold(bool bold)
        {
            Paragraph.Format.Font.Bold = bold;
            return this;
        }

        // TODO More attributes.

        public MezParagraph AddText(string text)
        {
            Paragraph.AddText(text);
            return this;
        }

        public MezParagraph AddBookmark(string name)
        {
            // AddBookmark returns a BookmarkField. Not much can be done with it.
            Paragraph.AddBookmark(name);
            return this;
        }

        public MezParagraph AddFormattedText(FormattedText text)
        {
            Paragraph.Add(text);
            return this;
        }

        /// <summary>
        /// Adds the contents of a paragraph to the paragraph.
        /// Paragraphs cannot be nested, but using this method a paragraph can be used
        /// as a container for formatted text.
        /// </summary>
        public MezParagraph AddParagraph(Paragraph text)
        {
            foreach (DocumentObject item in text.Elements)
            {
                var cl = (DocumentObject)item.Clone();
                var bf = cl as BookmarkField;
                if (bf != null) { Paragraph.Add(bf); continue; }
                var pf = cl as PageField;
                if (pf != null) { Paragraph.Add(pf); continue; }
                var prf = cl as PageRefField;
                if (prf != null) { Paragraph.Add(prf); continue; }
                var npf = cl as NumPagesField;
                if (npf != null) { Paragraph.Add(npf); continue; }
                var sf = cl as SectionField;
                if (sf != null) { Paragraph.Add(sf); continue; }
                var spf = cl as SectionPagesField;
                if (spf != null) { Paragraph.Add(spf); continue; }
                var df = cl as DateField;
                if (df != null) { Paragraph.Add(df); continue; }
                var @if = cl as InfoField;
                if (@if != null) { Paragraph.Add(@if); continue; }
                var f = cl as Footnote;
                if (f != null) { Paragraph.Add(f); continue; }
                var t = cl as Text;
                if (t != null) { Paragraph.Add(t); continue; }
                var ft = cl as FormattedText;
                if (ft != null) { Paragraph.Add(ft); continue; }
                var hl = cl as Hyperlink;
                if (hl != null) { Paragraph.Add(hl); continue; }
                var i = cl as Image;
                if (i != null) { Paragraph.Add(i); continue; }
                var ch = cl as Character;
                if (ch != null) { Paragraph.Add(ch); continue; }

                throw new ArgumentException("Invalid item \"" + item.GetType().Name + "\" for Paragraph.");
            }
            return this;
        }

        public MezParagraph AddPageField()
        {
            Paragraph.AddPageField();
            return this;
        }

        public MezParagraph AddNumPagesField()
        {
            Paragraph.AddNumPagesField();
            return this;
        }

        public MezParagraph AddPageRefField(string name)
        {
            Paragraph.AddPageRefField(name);
            return this;
        }

        // TODO MoreAddXxx
        public MezParagraph AddTab()
        {
            Paragraph.AddTab();
            return this;
        }

        /// <summary>
        /// Adds a hyper link. Note that it returns the MezHyperlink, not the initial MezParagraph - chaining of Paragraph methods ends at AddHyperlink.
        /// </summary>
        public MezHyperlink AddHyperLink(string target, HyperlinkType type = HyperlinkType.Local)
        {
            var result = new MezHyperlink(Paragraph.AddHyperlink(target, type));
            return result;
        }

        public MezParagraph AddLineBreak()
        {
            Paragraph.AddLineBreak();
            return this;
        }

        public MezImage AddImage(string filename)
        {
            var result = new MezImage(Paragraph.AddImage(filename));
            return result;
        }

        public MezParagraph ShadingColor(Color color)
        {
            Paragraph.Format.Shading.Color = color;
            return this;
        }

        public MezParagraph Borders(Unit width, Color color)
        {
            Paragraph.Format.Borders.Width = width;
            Paragraph.Format.Borders.Color = color;
            return this;
        }

        public MezParagraph Borders(Unit width)
        {
            Paragraph.Format.Borders.Width = width;
            return this;
        }

        public MezParagraph Borders(Color color)
        {
            Paragraph.Format.Borders.Color = color;
            return this;
        }

        public MezParagraph BorderDistance(Unit distance)
        {
            Paragraph.Format.Borders.Distance = distance;
            return this;
        }

        public MezParagraph Alignment(ParagraphAlignment alignment)
        {
            Paragraph.Format.Alignment = alignment;
            return this;
        }

        public MezParagraph LeftIndent(Unit indent)
        {
            Paragraph.Format.LeftIndent = indent;
            return this;
        }

        public MezParagraph RightIndent(Unit indent)
        {
            Paragraph.Format.RightIndent = indent;
            return this;
        }

        public MezParagraph FirstLineIndent(string indent)
        {
            Paragraph.Format.FirstLineIndent = indent;
            return this;
        }

        public MezParagraph Color(Color color)
        {
            Paragraph.Format.Font.Color = color;
            return this;
        }

        public MezParagraph Font(Unit fontSize)
        {
            Paragraph.Format.Font.Size = fontSize;
            return this;
        }

        public MezParagraph Font(string fontName, Unit? fontSize = null)
        {
            Paragraph.Format.Font.Name = fontName;
            if (fontSize != null)
                Paragraph.Format.Font.Size = fontSize.Value;
            return this;
        }

        public MezParagraph SpaceAfter(Unit space)
        {
            Paragraph.Format.SpaceAfter = space;
            return this;
        }

        public MezParagraph SpaceBefore(Unit space)
        {
            Paragraph.Format.SpaceBefore = space;
            return this;
        }

        public MezParagraph AddDateField(string date)
        {
            Paragraph.AddDateField(date);
            return this;
        }
    }

    public class MezFormattedText
    {
        public FormattedText Ft { get; private set; }

        public MezFormattedText(FormattedText ft)
        {
            Ft = ft;
        }

        public MezFormattedText()
            : this(new FormattedText())
        { }

        public MezFormattedText(string text)
            : this()
        {
            Ft.AddText(text);
        }

        public static implicit operator FormattedText(MezFormattedText ft)
        {
            return ft.Ft;
        }

        public MezFormattedText Bold(bool bold)
        {
            Ft.Bold = bold;
            return this;
        }

        public MezFormattedText Italic(bool italic)
        {
            Ft.Italic = italic;
            return this;
        }

        public MezFormattedText Subscript(bool subscript)
        {
            Ft.Subscript = subscript;
            return this;
        }

        public MezFormattedText Superscript(bool superscript)
        {
            Ft.Superscript = superscript;
            return this;
        }

        public MezFormattedText Underline(Underline underline)
        {
            Ft.Underline = underline;
            return this;
        }

        public MezFormattedText Color(Color color)
        {
            Ft.Color = color;
            return this;
        }

        public MezFormattedText Font(string fontname, Unit? fontsize = null)
        {
            Ft.Font.Name = fontname;
            if (fontsize.HasValue)
                Ft.Font.Size = fontsize.Value;
            return this;
        }

        public MezFormattedText AddText(string text)
        {
            Ft.AddText(text);
            return this;
        }

        public MezFormattedText AddFormattedText(FormattedText text)
        {
            Ft.Add(text);
            return this;
        }

        public MezFormattedText Font(Unit fontsize)
        {
            Ft.Font.Size = fontsize;
            return this;
        }

        // TODO MoreAddXxx
    }

    public abstract class MezShape
    {
        public Shape Shape { get; private set; }

        public MezShape(Shape shape)
        {
            Shape = shape;
        }

        public static implicit operator Shape(MezShape shape)
        {
            return shape.Shape;
        }
    }

    public class MezImage : MezShape
    {
        public Image Image
        {
            get { return (Image)Shape; }
        }

        public MezImage(Image image)
            : base(image)
        {
        }

        public MezImage(string filename)
            : this(new Image(filename))
        {
        }

        public static implicit operator Image(MezImage image)
        {
            return image.Image;
        }

        public MezImage FitToSize(Unit width, Unit height)
        {
            Unit neededHeight = Image.Height * width / Image.Width;
            Unit neededWidth = Image.Width * height / Image.Height;
            if (neededHeight < height)
            {
                Image.Width = width;
                Image.Height = neededHeight;
            }
            else if (neededWidth < width)
            {
                Image.Width = neededWidth;
                Image.Height = height;
            }
            return this;
        }

        public MezImage StretchToSize(Unit width, Unit height)
        {
            Image.Width = width;
            Image.Height = height;
            return this;
        }

        public MezImage Width(Unit width)
        {
            Image.Width = width;
            return this;
        }

        public MezImage Height(Unit height)
        {
            Image.Height = height;
            return this;
        }

        public MezImage LockAspectRatio(bool lockAspectRatio)
        {
            Image.LockAspectRatio = lockAspectRatio;
            return this;
        }

        public MezImage RelativeVertical(RelativeVertical relativeVertical)
        {
            Image.RelativeVertical = relativeVertical;
            return this;
        }

        public MezImage RelativeHorizontal(RelativeHorizontal relativeHorizontal)
        {
            Image.RelativeHorizontal = relativeHorizontal;
            return this;
        }

        public MezImage Top(ShapePosition top)
        {
            Image.Top = top;
            return this;
        }

        public MezImage Left(ShapePosition left)
        {
            Image.Left = left;
            return this;
        }

        public MezImage WrapFormatStyle(WrapStyle style)
        {
            Image.WrapFormat.Style = style;
            return this;
        }
    }

    public class MezTextFrame : MezShape
    {
        public TextFrame TextFrame
        {
            get { return (TextFrame)Shape; }
        }

        public MezTextFrame(TextFrame textFrame)
            : base(textFrame)
        {
        }

        public static implicit operator TextFrame(MezTextFrame textFrame)
        {
            return textFrame.TextFrame;
        }
    }

    public class MezHyperlink
    {
        public Hyperlink Hyperlink { get; private set; }

        public MezHyperlink(Hyperlink hyperlink)
        {
            Hyperlink = hyperlink;
        }

        public static implicit operator Hyperlink(MezHyperlink hyperlink)
        {
            return hyperlink.Hyperlink;
        }

        public MezHyperlink Target(string target)
        {
            Hyperlink.Name = target;
            return this;
        }

        public MezHyperlink Type(HyperlinkType type)
        {
            Hyperlink.Type = type;
            return this;
        }

        public MezHyperlink AddText(string text)
        {
            if (!string.IsNullOrEmpty(text))
                Hyperlink.AddText(text);
            return this;
        }

        public MezHyperlink AddParagraph(Paragraph paragraph)
        {
            // Magic: We don't add the paragraph, we clone the contents of the paragraph.
            foreach (DocumentObject item in paragraph.Elements)
            {
                Hyperlink.Elements.Add((DocumentObject)item.Clone());
            }
            return this;
        }

        public MezHyperlink AddFormattedText(FormattedText formattedText)
        {
            Hyperlink.Add(formattedText);
            return this;
        }
    }

    // Template
    //public class MezFoo
    //{
    //    public Image Image { get; private set; }

    //    public MezFoo(Image Image)
    //    {
    //        Image = Image;
    //    }

    //    public static implicit operator Image(MezFoo foo)
    //    {
    //        return foo.Image;
    //    }
    //}


}
