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
        public void ReadDocument_WithTabbedAndBrokenRun_AppendsAllRunFragments()
        {
            using var ms = new MemoryStream();
            using (var archive = new ZipArchive(ms, ZipArchiveMode.Create, true))
            {
                var entry = archive.CreateEntry("word/document.xml");
                using var entryStream = entry.Open();
                using var writer = new StreamWriter(entryStream);
                writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" +
                             "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:body><w:p>" +
                             "<w:r><w:t>Hello</w:t><w:tab/><w:t>World</w:t><w:br/><w:t>Next</w:t><w:cr/><w:t>Line</w:t></w:r>" +
                             "</w:p></w:body></w:document>");
            }

            using var testStream = new MemoryStream(ms.ToArray());
            using var reader = new Nedev.FileConverters.DocxToDoc.Format.DocxReader(testStream);

            var model = reader.ReadDocument();

            var run = Assert.Single(Assert.Single(model.Paragraphs).Runs);
            Assert.Equal("Hello\tWorld\vNext\vLine", run.Text);
            Assert.Equal("Hello\tWorld\vNext\vLine\r", model.TextBuffer);
        }

        [Fact]
        public void ReadDocument_WithHyperlinkTabbedAndBrokenRun_AppendsDisplayText()
        {
            using var ms = new MemoryStream();
            using (var archive = new ZipArchive(ms, ZipArchiveMode.Create, true))
            {
                var entry = archive.CreateEntry("word/document.xml");
                using (var entryStream = entry.Open())
                using (var writer = new StreamWriter(entryStream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" +
                                 "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"><w:body><w:p>" +
                                 "<w:hyperlink r:id=\"rId1\"><w:r><w:t>Click</w:t><w:tab/><w:t>Here</w:t><w:br/><w:t>Now</w:t></w:r></w:hyperlink>" +
                                 "</w:p></w:body></w:document>");
                }

                var relsEntry = archive.CreateEntry("word/_rels/document.xml.rels");
                using var relsStream = relsEntry.Open();
                using var relsWriter = new StreamWriter(relsStream);
                relsWriter.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                                 "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
                                 "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink\" Target=\"https://example.com\" TargetMode=\"External\"/>" +
                                 "</Relationships>");
            }

            using var testStream = new MemoryStream(ms.ToArray());
            using var reader = new Nedev.FileConverters.DocxToDoc.Format.DocxReader(testStream);

            var model = reader.ReadDocument();

            var run = Assert.Single(Assert.Single(model.Paragraphs).Runs);
            Assert.NotNull(run.Hyperlink);
            Assert.Equal("Click\tHere\vNow", run.Text);
            Assert.Equal("Click\tHere\vNow", run.Hyperlink!.DisplayText);
            Assert.Equal("https://example.com", run.Hyperlink.TargetUrl);
            Assert.Equal("Click\tHere\vNow\r", model.TextBuffer);
        }

        [Fact]
        public void ReadDocument_WithSpecialHyphenCharacters_PreservesRunAndHyperlinkText()
        {
            using var ms = new MemoryStream();
            using (var archive = new ZipArchive(ms, ZipArchiveMode.Create, true))
            {
                var entry = archive.CreateEntry("word/document.xml");
                using (var entryStream = entry.Open())
                using (var writer = new StreamWriter(entryStream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" +
                                 "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"><w:body><w:p>" +
                                 "<w:r><w:t>co</w:t><w:noBreakHyphen/><w:t>op</w:t><w:softHyphen/><w:t>erate</w:t></w:r>" +
                                 "<w:r><w:t xml:space=\"preserve\"> </w:t></w:r>" +
                                 "<w:hyperlink r:id=\"rId1\"><w:r><w:t>re</w:t><w:noBreakHyphen/><w:t>enter</w:t><w:softHyphen/><w:t>ing</w:t></w:r></w:hyperlink>" +
                                 "</w:p></w:body></w:document>");
                }

                var relsEntry = archive.CreateEntry("word/_rels/document.xml.rels");
                using var relsStream = relsEntry.Open();
                using var relsWriter = new StreamWriter(relsStream);
                relsWriter.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                                 "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
                                 "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink\" Target=\"https://example.com/hyphen\" TargetMode=\"External\"/>" +
                                 "</Relationships>");
            }

            using var testStream = new MemoryStream(ms.ToArray());
            using var reader = new Nedev.FileConverters.DocxToDoc.Format.DocxReader(testStream);

            var model = reader.ReadDocument();
            var paragraph = Assert.Single(model.Paragraphs);
            var hyperlinkRun = Assert.Single(paragraph.Runs, r => r.Hyperlink != null);

            Assert.Equal("co\u2011op\u00ADerate", paragraph.Runs[0].Text);
            Assert.Equal("re\u2011enter\u00ADing", hyperlinkRun.Text);
            Assert.Equal("re\u2011enter\u00ADing", hyperlinkRun.Hyperlink!.DisplayText);
            Assert.Equal("https://example.com/hyphen", hyperlinkRun.Hyperlink.TargetUrl);
            Assert.Equal("co\u2011op\u00ADerate re\u2011enter\u00ADing\r", model.TextBuffer);
        }

        [Fact]
        public void ReadDocument_WithSymbolRuns_PreservesSymbolCharactersAndFonts()
        {
            using var ms = new MemoryStream();
            using (var archive = new ZipArchive(ms, ZipArchiveMode.Create, true))
            {
                var entry = archive.CreateEntry("word/document.xml");
                using (var entryStream = entry.Open())
                using (var writer = new StreamWriter(entryStream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" +
                                 "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"><w:body><w:p>" +
                                 "<w:r><w:sym w:font=\"Wingdings\" w:char=\"F0FC\"/></w:r>" +
                                 "<w:hyperlink r:id=\"rId1\"><w:r><w:sym w:font=\"Wingdings\" w:char=\"F0FC\"/></w:r></w:hyperlink>" +
                                 "</w:p></w:body></w:document>");
                }

                var relsEntry = archive.CreateEntry("word/_rels/document.xml.rels");
                using var relsStream = relsEntry.Open();
                using var relsWriter = new StreamWriter(relsStream);
                relsWriter.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                                 "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
                                 "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink\" Target=\"https://example.com/symbol\" TargetMode=\"External\"/>" +
                                 "</Relationships>");
            }

            using var testStream = new MemoryStream(ms.ToArray());
            using var reader = new Nedev.FileConverters.DocxToDoc.Format.DocxReader(testStream);

            var model = reader.ReadDocument();
            var paragraph = Assert.Single(model.Paragraphs);
            var symbolRun = paragraph.Runs[0];
            var hyperlinkRun = Assert.Single(paragraph.Runs, r => r.Hyperlink != null);

            Assert.Equal("\uF0FC", symbolRun.Text);
            Assert.Equal("Wingdings", symbolRun.Properties.FontName);
            Assert.Equal("\uF0FC", hyperlinkRun.Text);
            Assert.Equal("Wingdings", hyperlinkRun.Properties.FontName);
            Assert.Equal("\uF0FC", hyperlinkRun.Hyperlink!.DisplayText);
            Assert.Equal("https://example.com/symbol", hyperlinkRun.Hyperlink.TargetUrl);
            Assert.Equal("\uF0FC\uF0FC\r", model.TextBuffer);
        }

        [Fact]
        public void ReadDocument_WithMixedTextAndSymbolInSingleRun_SplitsRunsByEffectiveFont()
        {
            using var ms = new MemoryStream();
            using (var archive = new ZipArchive(ms, ZipArchiveMode.Create, true))
            {
                var entry = archive.CreateEntry("word/document.xml");
                using (var entryStream = entry.Open())
                using (var writer = new StreamWriter(entryStream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" +
                                 "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"><w:body><w:p>" +
                                 "<w:r><w:rPr><w:rFonts w:ascii=\"Calibri\"/></w:rPr><w:t>A</w:t><w:sym w:font=\"Wingdings\" w:char=\"F0FC\"/><w:t>B</w:t></w:r>" +
                                 "<w:hyperlink r:id=\"rId1\"><w:r><w:rPr><w:rFonts w:ascii=\"Calibri\"/></w:rPr><w:t>C</w:t><w:sym w:font=\"Wingdings\" w:char=\"F0FC\"/><w:t>D</w:t></w:r></w:hyperlink>" +
                                 "</w:p></w:body></w:document>");
                }

                var relsEntry = archive.CreateEntry("word/_rels/document.xml.rels");
                using var relsStream = relsEntry.Open();
                using var relsWriter = new StreamWriter(relsStream);
                relsWriter.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                                 "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
                                 "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink\" Target=\"https://example.com/mixed-symbol\" TargetMode=\"External\"/>" +
                                 "</Relationships>");
            }

            using var testStream = new MemoryStream(ms.ToArray());
            using var reader = new Nedev.FileConverters.DocxToDoc.Format.DocxReader(testStream);

            var model = reader.ReadDocument();
            var runs = Assert.Single(model.Paragraphs).Runs;

            Assert.Equal(6, runs.Count);
            Assert.Equal("A", runs[0].Text);
            Assert.Equal("Calibri", runs[0].Properties.FontName);
            Assert.Null(runs[0].Hyperlink);

            Assert.Equal("\uF0FC", runs[1].Text);
            Assert.Equal("Wingdings", runs[1].Properties.FontName);
            Assert.Null(runs[1].Hyperlink);

            Assert.Equal("B", runs[2].Text);
            Assert.Equal("Calibri", runs[2].Properties.FontName);
            Assert.Null(runs[2].Hyperlink);

            Assert.Equal("C", runs[3].Text);
            Assert.Equal("Calibri", runs[3].Properties.FontName);
            Assert.NotNull(runs[3].Hyperlink);

            Assert.Equal("\uF0FC", runs[4].Text);
            Assert.Equal("Wingdings", runs[4].Properties.FontName);
            Assert.NotNull(runs[4].Hyperlink);

            Assert.Equal("D", runs[5].Text);
            Assert.Equal("Calibri", runs[5].Properties.FontName);
            Assert.NotNull(runs[5].Hyperlink);
            Assert.Equal("C\uF0FCD", runs[5].Hyperlink!.DisplayText);
            Assert.Equal("https://example.com/mixed-symbol", runs[5].Hyperlink!.TargetUrl);
            Assert.Equal("A\uF0FCBC\uF0FCD\r", model.TextBuffer);
        }

        [Fact]
        public void ReadDocument_WithPositionedTabs_PreservesRunAndHyperlinkText()
        {
            using var ms = new MemoryStream();
            using (var archive = new ZipArchive(ms, ZipArchiveMode.Create, true))
            {
                var entry = archive.CreateEntry("word/document.xml");
                using (var entryStream = entry.Open())
                using (var writer = new StreamWriter(entryStream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" +
                                 "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"><w:body><w:p>" +
                                 "<w:r><w:t>Left</w:t><w:ptab w:alignment=\"center\" w:relativeTo=\"margin\"/><w:t>Right</w:t></w:r>" +
                                 "<w:r><w:t xml:space=\"preserve\"> </w:t></w:r>" +
                                 "<w:hyperlink r:id=\"rId1\"><w:r><w:t>Top</w:t><w:ptab w:alignment=\"right\" w:relativeTo=\"margin\"/><w:t>Bottom</w:t></w:r></w:hyperlink>" +
                                 "</w:p></w:body></w:document>");
                }

                var relsEntry = archive.CreateEntry("word/_rels/document.xml.rels");
                using var relsStream = relsEntry.Open();
                using var relsWriter = new StreamWriter(relsStream);
                relsWriter.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                                 "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
                                 "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink\" Target=\"https://example.com/ptab\" TargetMode=\"External\"/>" +
                                 "</Relationships>");
            }

            using var testStream = new MemoryStream(ms.ToArray());
            using var reader = new Nedev.FileConverters.DocxToDoc.Format.DocxReader(testStream);

            var model = reader.ReadDocument();
            var paragraph = Assert.Single(model.Paragraphs);
            var hyperlinkRun = Assert.Single(paragraph.Runs, r => r.Hyperlink != null);

            Assert.Equal("Left\tRight", paragraph.Runs[0].Text);
            Assert.Equal("Top\tBottom", hyperlinkRun.Text);
            Assert.Equal("Top\tBottom", hyperlinkRun.Hyperlink!.DisplayText);
            Assert.Equal("https://example.com/ptab", hyperlinkRun.Hyperlink.TargetUrl);
            Assert.Equal("Left\tRight Top\tBottom\r", model.TextBuffer);
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
            Assert.Equal(Nedev.FileConverters.DocxToDoc.Model.TableWidthUnit.Dxa, table.Rows[0].Cells[0].WidthUnit);
        }

        [Fact]
        public void ReadDocument_WithTableCellPctWidth_ParsesCellWidthAndUnit()
        {
            using var ms = new MemoryStream();
            using (var archive = new ZipArchive(ms, ZipArchiveMode.Create, true))
            {
                var entry = archive.CreateEntry("word/document.xml");
                using var entryStream = entry.Open();
                using var writer = new StreamWriter(entryStream);
                writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" +
                             "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:body>" +
                             "<w:tbl><w:tr><w:tc><w:tcPr><w:tcW w:w=\"2500\" w:type=\"pct\"/></w:tcPr><w:p><w:r><w:t>Cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl>" +
                             "</w:body></w:document>");
            }

            using var testStream = new MemoryStream(ms.ToArray());
            using var reader = new Nedev.FileConverters.DocxToDoc.Format.DocxReader(testStream);

            var model = reader.ReadDocument();

            var table = Assert.IsType<Nedev.FileConverters.DocxToDoc.Model.TableModel>(Assert.Single(model.Content));
            Assert.Equal(2500, table.Rows[0].Cells[0].Width);
            Assert.Equal(Nedev.FileConverters.DocxToDoc.Model.TableWidthUnit.Pct, table.Rows[0].Cells[0].WidthUnit);
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
        public void ReadDocument_WithInsideTableBorders_ParsesInsideBorderThickness()
        {
            using var ms = new MemoryStream();
            using (var archive = new ZipArchive(ms, ZipArchiveMode.Create, true))
            {
                var entry = archive.CreateEntry("word/document.xml");
                using var entryStream = entry.Open();
                using var writer = new StreamWriter(entryStream);
                writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" +
                             "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:body>" +
                             "<w:tbl><w:tblPr><w:tblBorders><w:insideH w:val=\"single\" w:sz=\"10\"/><w:insideV w:val=\"single\" w:sz=\"14\"/></w:tblBorders></w:tblPr>" +
                             "<w:tr><w:tc><w:p><w:r><w:t>A</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>B</w:t></w:r></w:p></w:tc></w:tr>" +
                             "</w:tbl></w:body></w:document>");
            }

            using var testStream = new MemoryStream(ms.ToArray());
            using var reader = new Nedev.FileConverters.DocxToDoc.Format.DocxReader(testStream);

            var model = reader.ReadDocument();

            var table = Assert.IsType<Nedev.FileConverters.DocxToDoc.Model.TableModel>(Assert.Single(model.Content));

            Assert.Equal(25, table.DefaultInsideHorizontalBorderTwips);
            Assert.Equal(35, table.DefaultInsideVerticalBorderTwips);
        }

        [Fact]
        public void ReadDocument_WithCellBorderNone_PreservesExplicitBorderOverride()
        {
            using var ms = new MemoryStream();
            using (var archive = new ZipArchive(ms, ZipArchiveMode.Create, true))
            {
                var entry = archive.CreateEntry("word/document.xml");
                using var entryStream = entry.Open();
                using var writer = new StreamWriter(entryStream);
                writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" +
                             "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:body>" +
                             "<w:tbl><w:tblPr><w:tblBorders><w:insideV w:val=\"single\" w:sz=\"14\"/></w:tblBorders></w:tblPr>" +
                             "<w:tr>" +
                             "<w:tc><w:tcPr><w:tcBorders><w:right w:val=\"none\"/></w:tcBorders></w:tcPr><w:p><w:r><w:t>A</w:t></w:r></w:p></w:tc>" +
                             "<w:tc><w:p><w:r><w:t>B</w:t></w:r></w:p></w:tc>" +
                             "</w:tr></w:tbl></w:body></w:document>");
            }

            using var testStream = new MemoryStream(ms.ToArray());
            using var reader = new Nedev.FileConverters.DocxToDoc.Format.DocxReader(testStream);

            var model = reader.ReadDocument();

            var table = Assert.IsType<Nedev.FileConverters.DocxToDoc.Model.TableModel>(Assert.Single(model.Content));
            var firstCell = table.Rows[0].Cells[0];

            Assert.True(firstCell.HasRightBorderOverride);
            Assert.Equal(0, firstCell.BorderRightTwips);
        }

        [Fact]
        public void ReadDocument_WithCellZeroMargins_PreservesExplicitPaddingOverrides()
        {
            using var ms = new MemoryStream();
            using (var archive = new ZipArchive(ms, ZipArchiveMode.Create, true))
            {
                var entry = archive.CreateEntry("word/document.xml");
                using var entryStream = entry.Open();
                using var writer = new StreamWriter(entryStream);
                writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" +
                             "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:body>" +
                             "<w:tbl><w:tblPr><w:tblCellMar><w:left w:w=\"120\" w:type=\"dxa\"/><w:right w:w=\"180\" w:type=\"dxa\"/><w:top w:w=\"40\" w:type=\"dxa\"/><w:bottom w:w=\"70\" w:type=\"dxa\"/></w:tblCellMar></w:tblPr>" +
                             "<w:tr>" +
                             "<w:tc><w:tcPr><w:tcMar><w:left w:w=\"0\" w:type=\"dxa\"/><w:right w:w=\"0\" w:type=\"dxa\"/><w:top w:w=\"0\" w:type=\"dxa\"/><w:bottom w:w=\"0\" w:type=\"dxa\"/></w:tcMar></w:tcPr><w:p><w:r><w:t>A</w:t></w:r></w:p></w:tc>" +
                             "</w:tr></w:tbl></w:body></w:document>");
            }

            using var testStream = new MemoryStream(ms.ToArray());
            using var reader = new Nedev.FileConverters.DocxToDoc.Format.DocxReader(testStream);

            var model = reader.ReadDocument();

            var table = Assert.IsType<Nedev.FileConverters.DocxToDoc.Model.TableModel>(Assert.Single(model.Content));
            var cell = Assert.Single(table.Rows[0].Cells);

            Assert.True(cell.HasLeftPaddingOverride);
            Assert.True(cell.HasRightPaddingOverride);
            Assert.True(cell.HasTopPaddingOverride);
            Assert.True(cell.HasBottomPaddingOverride);
            Assert.Equal(0, cell.PaddingLeftTwips);
            Assert.Equal(0, cell.PaddingRightTwips);
            Assert.Equal(0, cell.PaddingTopTwips);
            Assert.Equal(0, cell.PaddingBottomTwips);
        }

        [Fact]
        public void ReadDocument_WithPreferredTableWidth_ParsesWidthValueAndUnit()
        {
            using var ms = new MemoryStream();
            using (var archive = new ZipArchive(ms, ZipArchiveMode.Create, true))
            {
                var entry = archive.CreateEntry("word/document.xml");
                using var entryStream = entry.Open();
                using var writer = new StreamWriter(entryStream);
                writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" +
                             "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:body>" +
                             "<w:tbl><w:tblPr><w:tblW w:w=\"2500\" w:type=\"pct\"/></w:tblPr>" +
                             "<w:tr><w:tc><w:p><w:r><w:t>A</w:t></w:r></w:p></w:tc></w:tr></w:tbl></w:body></w:document>");
            }

            using var testStream = new MemoryStream(ms.ToArray());
            using var reader = new Nedev.FileConverters.DocxToDoc.Format.DocxReader(testStream);

            var model = reader.ReadDocument();

            var table = Assert.IsType<Nedev.FileConverters.DocxToDoc.Model.TableModel>(Assert.Single(model.Content));
            Assert.Equal(2500, table.PreferredWidthValue);
            Assert.Equal(Nedev.FileConverters.DocxToDoc.Model.TableWidthUnit.Pct, table.PreferredWidthUnit);
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
