using System;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Xml;
using Xunit;

namespace Nedev.FileConverters.DocxToDoc.Tests.Format
{
    public class DocxReaderTests
    {
        private byte[] CreateDummyDocx()
        {
            using var ms = new MemoryStream();
            using (var archive = new ZipArchive(ms, ZipArchiveMode.Create, true))
            {
                var entry = archive.CreateEntry("word/document.xml");
                using var entryStream = entry.Open();
                using var writer = new StreamWriter(entryStream);
                // Valid minimal XML
                writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:body><w:p><w:r><w:t>Hello World</w:t></w:r></w:p></w:body></w:document>");
            }
            return ms.ToArray();
        }

        [Fact]
        public void ReadDocument_ValidStream_FindsText()
        {
            // Arrange
            byte[] dummyData = CreateDummyDocx();
            using var ms = new MemoryStream(dummyData);
            using var reader = new Nedev.FileConverters.DocxToDoc.Format.DocxReader(ms);

            // Act
            var model = reader.ReadDocument();

            // Assert
            Assert.Equal("Hello World\r", model.TextBuffer); // Includes the paragraph return
            Assert.Single(model.Paragraphs);
            Assert.Single(model.Paragraphs[0].Runs);
            Assert.Equal("Hello World", model.Paragraphs[0].Runs[0].Text);
        }

        [Fact]
        public void ReadDocument_MissingDocumentXml_ThrowsFileNotFound()
        {
            // Arrange
            using var ms = new MemoryStream();
            using (var archive = new ZipArchive(ms, ZipArchiveMode.Create, true))
            {
                // Create an empty zip
            }
            byte[] emptyZip = ms.ToArray();
            using var testStream = new MemoryStream(emptyZip);
            using var reader = new Nedev.FileConverters.DocxToDoc.Format.DocxReader(testStream);

            // Act & Assert
            Assert.Throws<FileNotFoundException>(() => 
            {
                reader.ReadDocument();
            });
        }

        [Fact]
        public void ReadDocument_WithFormatting_ParsesProperties()
        {
            // Arrange
            using var ms = new MemoryStream();
            using (var archive = new ZipArchive(ms, ZipArchiveMode.Create, true))
            {
                var entry = archive.CreateEntry("word/document.xml");
                using var entryStream = entry.Open();
                using var writer = new StreamWriter(entryStream);
                writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" +
                             "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:body><w:p>" +
                             "<w:r><w:rPr><w:b/><w:i w:val=\"1\"/><w:sz w:val=\"24\"/></w:rPr><w:t>BoldItalic12pt</w:t></w:r>" +
                             "</w:p></w:body></w:document>");
            }
            
            using var testStream = new MemoryStream(ms.ToArray());
            using var reader = new Nedev.FileConverters.DocxToDoc.Format.DocxReader(testStream);

            // Act
            var model = reader.ReadDocument();

            // Assert
            Assert.Single(model.Paragraphs);
            Assert.Single(model.Paragraphs[0].Runs);
            
            var run = model.Paragraphs[0].Runs[0];
            Assert.Equal("BoldItalic12pt", run.Text);
            Assert.True(run.Properties.IsBold);
            Assert.True(run.Properties.IsItalic);
            Assert.False(run.Properties.IsStrike);
            Assert.Equal(24, run.Properties.FontSize);
        }

        [Fact]
        public void ReadDocument_WithParagraphSpacing_ParsesSpacingProperties()
        {
            using var ms = new MemoryStream();
            using (var archive = new ZipArchive(ms, ZipArchiveMode.Create, true))
            {
                var entry = archive.CreateEntry("word/document.xml");
                using var entryStream = entry.Open();
                using var writer = new StreamWriter(entryStream);
                writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" +
                             "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:body><w:p><w:pPr>" +
                             "<w:spacing w:before=\"120\" w:after=\"240\" w:line=\"480\" w:lineRule=\"auto\"/>" +
                             "</w:pPr><w:r><w:t>Spaced paragraph</w:t></w:r></w:p></w:body></w:document>");
            }

            using var testStream = new MemoryStream(ms.ToArray());
            using var reader = new Nedev.FileConverters.DocxToDoc.Format.DocxReader(testStream);

            var model = reader.ReadDocument();

            var paragraph = Assert.Single(model.Paragraphs);
            Assert.Equal(120, paragraph.Properties.SpacingBeforeTwips);
            Assert.Equal(240, paragraph.Properties.SpacingAfterTwips);
            Assert.Equal(480, paragraph.Properties.LineSpacing);
            Assert.Equal("auto", paragraph.Properties.LineSpacingRule);
        }

        [Fact]
        public void ReadDocument_WithParagraphIndent_ParsesIndentProperties()
        {
            using var ms = new MemoryStream();
            using (var archive = new ZipArchive(ms, ZipArchiveMode.Create, true))
            {
                var entry = archive.CreateEntry("word/document.xml");
                using var entryStream = entry.Open();
                using var writer = new StreamWriter(entryStream);
                writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" +
                             "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:body><w:p><w:pPr>" +
                             "<w:ind w:left=\"720\" w:right=\"360\" w:firstLine=\"240\"/>" +
                             "</w:pPr><w:r><w:t>Indented paragraph</w:t></w:r></w:p></w:body></w:document>");
            }

            using var testStream = new MemoryStream(ms.ToArray());
            using var reader = new Nedev.FileConverters.DocxToDoc.Format.DocxReader(testStream);

            var model = reader.ReadDocument();

            var paragraph = Assert.Single(model.Paragraphs);
            Assert.Equal(720, paragraph.Properties.LeftIndentTwips);
            Assert.Equal(360, paragraph.Properties.RightIndentTwips);
            Assert.Equal(240, paragraph.Properties.FirstLineIndentTwips);
        }

        [Fact]
        public void ReadDocument_WithTableCellWidth_ParsesCellWidth()
        {
            using var ms = new MemoryStream();
            using (var archive = new ZipArchive(ms, ZipArchiveMode.Create, true))
            {
                var entry = archive.CreateEntry("word/document.xml");
                using var entryStream = entry.Open();
                using var writer = new StreamWriter(entryStream);
                writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" +
                             "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:body>" +
                             "<w:tbl><w:tr><w:tc><w:tcPr><w:tcW w:w=\"3200\" w:type=\"dxa\"/></w:tcPr><w:p><w:r><w:t>Cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl>" +
                             "</w:body></w:document>");
            }

            using var testStream = new MemoryStream(ms.ToArray());
            using var reader = new Nedev.FileConverters.DocxToDoc.Format.DocxReader(testStream);

            var model = reader.ReadDocument();

            var table = Assert.IsType<Nedev.FileConverters.DocxToDoc.Model.TableModel>(Assert.Single(model.Content));
            Assert.Equal(3200, table.Rows[0].Cells[0].Width);
        }

        [Fact]
        public void ReadDocument_WithTableGridAndGridSpan_UsesGridWidthsForCellsWithoutTcW()
        {
            using var ms = new MemoryStream();
            using (var archive = new ZipArchive(ms, ZipArchiveMode.Create, true))
            {
                var entry = archive.CreateEntry("word/document.xml");
                using var entryStream = entry.Open();
                using var writer = new StreamWriter(entryStream);
                writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" +
                             "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:body>" +
                             "<w:tbl><w:tblGrid><w:gridCol w:w=\"1200\"/><w:gridCol w:w=\"1800\"/><w:gridCol w:w=\"900\"/></w:tblGrid>" +
                             "<w:tr>" +
                             "<w:tc><w:tcPr><w:gridSpan w:val=\"2\"/></w:tcPr><w:p><w:r><w:t>A</w:t></w:r></w:p></w:tc>" +
                             "<w:tc><w:p><w:r><w:t>B</w:t></w:r></w:p></w:tc>" +
                             "</w:tr></w:tbl></w:body></w:document>");
            }

            using var testStream = new MemoryStream(ms.ToArray());
            using var reader = new Nedev.FileConverters.DocxToDoc.Format.DocxReader(testStream);

            var model = reader.ReadDocument();

            var table = Assert.IsType<Nedev.FileConverters.DocxToDoc.Model.TableModel>(Assert.Single(model.Content));
            Assert.Equal(new[] { 1200, 1800, 900 }, table.GridColumnWidths);
            Assert.Equal(3000, table.Rows[0].Cells[0].Width);
            Assert.Equal(2, table.Rows[0].Cells[0].GridSpan);
            Assert.Equal(900, table.Rows[0].Cells[1].Width);
        }

        [Fact]
        public void ReadDocument_WithMixedGridSpanAndCellMargins_ParsesDerivedWidthsAndPadding()
        {
            using var ms = new MemoryStream();
            using (var archive = new ZipArchive(ms, ZipArchiveMode.Create, true))
            {
                var entry = archive.CreateEntry("word/document.xml");
                using var entryStream = entry.Open();
                using var writer = new StreamWriter(entryStream);
                writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" +
                             "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:body>" +
                             "<w:tbl><w:tblPr><w:tblCellSpacing w:w=\"90\" w:type=\"dxa\"/><w:tblCellMar><w:left w:w=\"120\" w:type=\"dxa\"/><w:right w:w=\"180\" w:type=\"dxa\"/><w:top w:w=\"40\" w:type=\"dxa\"/><w:bottom w:w=\"70\" w:type=\"dxa\"/></w:tblCellMar></w:tblPr>" +
                             "<w:tblGrid><w:gridCol w:w=\"1000\"/><w:gridCol w:w=\"1500\"/><w:gridCol w:w=\"2000\"/><w:gridCol w:w=\"800\"/></w:tblGrid>" +
                             "<w:tr>" +
                             "<w:tc><w:tcPr><w:tcW w:w=\"2500\" w:type=\"dxa\"/><w:gridSpan w:val=\"2\"/><w:vAlign w:val=\"bottom\"/><w:tcMar><w:left w:w=\"60\" w:type=\"dxa\"/><w:right w:w=\"90\" w:type=\"dxa\"/><w:top w:w=\"20\" w:type=\"dxa\"/><w:bottom w:w=\"30\" w:type=\"dxa\"/></w:tcMar></w:tcPr><w:p><w:r><w:t>A</w:t></w:r></w:p></w:tc>" +
                             "<w:tc><w:p><w:r><w:t>B</w:t></w:r></w:p></w:tc>" +
                             "<w:tc><w:p><w:r><w:t>C</w:t></w:r></w:p></w:tc>" +
                             "</w:tr></w:tbl></w:body></w:document>");
            }

            using var testStream = new MemoryStream(ms.ToArray());
            using var reader = new Nedev.FileConverters.DocxToDoc.Format.DocxReader(testStream);

            var model = reader.ReadDocument();

            var table = Assert.IsType<Nedev.FileConverters.DocxToDoc.Model.TableModel>(Assert.Single(model.Content));
            Assert.Equal(90, table.CellSpacingTwips);
            Assert.Equal(120, table.DefaultCellPaddingLeftTwips);
            Assert.Equal(180, table.DefaultCellPaddingRightTwips);
            Assert.Equal(40, table.DefaultCellPaddingTopTwips);
            Assert.Equal(70, table.DefaultCellPaddingBottomTwips);
            Assert.Equal(2500, table.Rows[0].Cells[0].Width);
            Assert.Equal(2, table.Rows[0].Cells[0].GridSpan);
            Assert.Equal(Nedev.FileConverters.DocxToDoc.Model.TableCellVerticalAlignment.Bottom, table.Rows[0].Cells[0].VerticalAlignment);
            Assert.Equal(60, table.Rows[0].Cells[0].PaddingLeftTwips);
            Assert.Equal(90, table.Rows[0].Cells[0].PaddingRightTwips);
            Assert.Equal(20, table.Rows[0].Cells[0].PaddingTopTwips);
            Assert.Equal(30, table.Rows[0].Cells[0].PaddingBottomTwips);
            Assert.Equal(2000, table.Rows[0].Cells[1].Width);
            Assert.Equal(800, table.Rows[0].Cells[2].Width);
        }

        [Fact]
        public void ReadDocument_WithTableAndCellBorders_ParsesBorderThickness()
        {
            using var ms = new MemoryStream();
            using (var archive = new ZipArchive(ms, ZipArchiveMode.Create, true))
            {
                var entry = archive.CreateEntry("word/document.xml");
                using var entryStream = entry.Open();
                using var writer = new StreamWriter(entryStream);
                writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" +
                             "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:body>" +
                             "<w:tbl><w:tblPr><w:tblBorders><w:left w:val=\"single\" w:sz=\"8\"/><w:right w:val=\"single\" w:sz=\"12\"/><w:top w:val=\"single\" w:sz=\"16\"/><w:bottom w:val=\"single\" w:sz=\"20\"/></w:tblBorders></w:tblPr>" +
                             "<w:tr>" +
                             "<w:tc><w:tcPr><w:tcBorders><w:left w:val=\"single\" w:sz=\"24\"/><w:bottom w:val=\"single\" w:sz=\"28\"/></w:tcBorders></w:tcPr><w:p><w:r><w:t>A</w:t></w:r></w:p></w:tc>" +
                             "</w:tr></w:tbl></w:body></w:document>");
            }

            using var testStream = new MemoryStream(ms.ToArray());
            using var reader = new Nedev.FileConverters.DocxToDoc.Format.DocxReader(testStream);

            var model = reader.ReadDocument();

            var table = Assert.IsType<Nedev.FileConverters.DocxToDoc.Model.TableModel>(Assert.Single(model.Content));
            var cell = Assert.Single(table.Rows[0].Cells);

            Assert.Equal(20, table.DefaultBorderLeftTwips);
            Assert.Equal(30, table.DefaultBorderRightTwips);
            Assert.Equal(40, table.DefaultBorderTopTwips);
            Assert.Equal(50, table.DefaultBorderBottomTwips);
            Assert.Equal(60, cell.BorderLeftTwips);
            Assert.Equal(70, cell.BorderBottomTwips);
            Assert.Equal(0, cell.BorderRightTwips);
            Assert.Equal(0, cell.BorderTopTwips);
        }

        [Fact]
        public void ReadDocument_WithRowHeight_ParsesHeightAndRule()
        {
            using var ms = new MemoryStream();
            using (var archive = new ZipArchive(ms, ZipArchiveMode.Create, true))
            {
                var entry = archive.CreateEntry("word/document.xml");
                using var entryStream = entry.Open();
                using var writer = new StreamWriter(entryStream);
                writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" +
                             "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:body>" +
                             "<w:tbl><w:tr><w:trPr><w:trHeight w:val=\"1440\" w:hRule=\"exact\"/></w:trPr>" +
                             "<w:tc><w:p><w:r><w:t>A</w:t></w:r></w:p></w:tc></w:tr></w:tbl></w:body></w:document>");
            }

            using var testStream = new MemoryStream(ms.ToArray());
            using var reader = new Nedev.FileConverters.DocxToDoc.Format.DocxReader(testStream);

            var model = reader.ReadDocument();

            var table = Assert.IsType<Nedev.FileConverters.DocxToDoc.Model.TableModel>(Assert.Single(model.Content));
            var row = Assert.Single(table.Rows);

            Assert.Equal(1440, row.HeightTwips);
            Assert.Equal(Nedev.FileConverters.DocxToDoc.Model.TableRowHeightRule.Exact, row.HeightRule);
        }
    }
}
