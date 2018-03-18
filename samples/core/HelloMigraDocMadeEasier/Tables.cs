using MigraDoc.DocumentObjectModel;
using MigraDoc.DocumentObjectModel.Tables;
using MigraDocMadeEZ;

namespace HelloMigraDoc
{
    /// <summary>
    /// This class demonstrates MigraDoc tables.
    /// </summary>
    public class Tables
    {
        public static void DefineTables(MigraDocMadeEZR document)
        {
            var paragraph = document.AddParagraph("Heading1", "Table Overview");
            paragraph.AddBookmark("Tables");

            DemonstrateSimpleTable(document);
            DemonstrateAlignment(document);
            DemonstrateCellMerge(document);
        }

        public static void DemonstrateSimpleTable(MigraDocMadeEZR document)
        {
            document.Section.AddParagraph("Simple Tables", "Heading2");

            var table = new Table();
            table.Borders.Width = 0.75;

            var column = table.AddColumn(Unit.FromCentimeter(2));
            column.Format.Alignment = ParagraphAlignment.Center;

            table.AddColumn(Unit.FromCentimeter(5));

            var row = table.AddRow();
            row.Shading.Color = Colors.PaleGoldenrod;
            var cell = row.Cells[0];
            cell.AddParagraph("Itemus");
            cell = row.Cells[1];
            cell.AddParagraph("Descriptum");

            row = table.AddRow();
            cell = row.Cells[0];
            cell.AddParagraph("1");
            cell = row.Cells[1];
            cell.AddParagraph(FillerText.ShortText);

            row = table.AddRow();
            cell = row.Cells[0];
            cell.AddParagraph("2");
            cell = row.Cells[1];
            cell.AddParagraph(FillerText.Text);

            table.SetEdge(0, 0, 2, 3, Edge.Box, BorderStyle.Single, 1.5, Colors.Black);

            document.Section.Add(table);
        }

        public static void DemonstrateAlignment(MigraDocMadeEZR document)
        {
            document.Section.AddParagraph("Cell Alignment", "Heading2");

            var table = document.Section.AddTable();
            table.Borders.Visible = true;
            table.Format.Shading.Color = Colors.LavenderBlush;
            table.Shading.Color = Colors.Salmon;
            table.TopPadding = 5;
            table.BottomPadding = 5;

            var column = table.AddColumn();
            column.Format.Alignment = ParagraphAlignment.Left;

            column = table.AddColumn();
            column.Format.Alignment = ParagraphAlignment.Center;

            column = table.AddColumn();
            column.Format.Alignment = ParagraphAlignment.Right;

            table.Rows.Height = 35;

            var row = table.AddRow();
            row.VerticalAlignment = VerticalAlignment.Top;
            row.Cells[0].AddParagraph("Text");
            row.Cells[1].AddParagraph("Text");
            row.Cells[2].AddParagraph("Text");

            row = table.AddRow();
            row.VerticalAlignment = VerticalAlignment.Center;
            row.Cells[0].AddParagraph("Text");
            row.Cells[1].AddParagraph("Text");
            row.Cells[2].AddParagraph("Text");

            row = table.AddRow();
            row.VerticalAlignment = VerticalAlignment.Bottom;
            row.Cells[0].AddParagraph("Text");
            row.Cells[1].AddParagraph("Text");
            row.Cells[2].AddParagraph("Text");
        }

        public static void DemonstrateCellMerge(MigraDocMadeEZR document)
        {
            document.Section.AddParagraph("Cell Merge", "Heading2");

            var table = document.Section.AddTable();
            table.Borders.Visible = true;
            table.TopPadding = 5;
            table.BottomPadding = 5;

            var column = table.AddColumn();
            column.Format.Alignment = ParagraphAlignment.Left;

            column = table.AddColumn();
            column.Format.Alignment = ParagraphAlignment.Center;

            column = table.AddColumn();
            column.Format.Alignment = ParagraphAlignment.Right;

            table.Rows.Height = 35;

            var row = table.AddRow();
            row.Cells[0].AddParagraph("Merge Right");
            row.Cells[0].MergeRight = 1;

            row = table.AddRow();
            row.VerticalAlignment = VerticalAlignment.Bottom;
            row.Cells[0].MergeDown = 1;
            row.Cells[0].VerticalAlignment = VerticalAlignment.Bottom;
            row.Cells[0].AddParagraph("Merge Down");

            table.AddRow();
        }
    }
}
