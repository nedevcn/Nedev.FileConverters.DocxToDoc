using System;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
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
        public void ReadDocument_WithDoubleStrike_ParsesAsStrike()
        {
            using var ms = new MemoryStream();
            using (var archive = new ZipArchive(ms, ZipArchiveMode.Create, true))
            {
                var entry = archive.CreateEntry("word/document.xml");
                using var entryStream = entry.Open();
                using var writer = new StreamWriter(entryStream);
                writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" +
                             "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:body><w:p>" +
                             "<w:r><w:rPr><w:dstrike/></w:rPr><w:t>DoubleStrike</w:t></w:r>" +
                             "</w:p></w:body></w:document>");
            }

            using var testStream = new MemoryStream(ms.ToArray());
            using var reader = new Nedev.FileConverters.DocxToDoc.Format.DocxReader(testStream);

            var model = reader.ReadDocument();

            var run = Assert.Single(Assert.Single(model.Paragraphs).Runs);
            Assert.Equal("DoubleStrike", run.Text);
            Assert.True(run.Properties.IsStrike);
        }

        [Fact]
        public void ReadDocument_WithComplexScriptFormatting_ParsesProperties()
        {
            using var ms = new MemoryStream();
            using (var archive = new ZipArchive(ms, ZipArchiveMode.Create, true))
            {
                var entry = archive.CreateEntry("word/document.xml");
                using var entryStream = entry.Open();
                using var writer = new StreamWriter(entryStream);
                writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" +
                             "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:body><w:p>" +
                             "<w:r><w:rPr><w:bCs/><w:iCs/><w:szCs w:val=\"30\"/></w:rPr><w:t>ComplexScript</w:t></w:r>" +
                             "</w:p></w:body></w:document>");
            }

            using var testStream = new MemoryStream(ms.ToArray());
            using var reader = new Nedev.FileConverters.DocxToDoc.Format.DocxReader(testStream);

            var model = reader.ReadDocument();

            var run = Assert.Single(Assert.Single(model.Paragraphs).Runs);
            Assert.Equal("ComplexScript", run.Text);
            Assert.True(run.Properties.IsBold);
            Assert.True(run.Properties.IsItalic);
            Assert.Equal(30, run.Properties.FontSize);
        }

        [Fact]
        public void ReadDocument_WithExtendedUnderlineValues_ParsesClosestSupportedTypes()
        {
            using var ms = new MemoryStream();
            using (var archive = new ZipArchive(ms, ZipArchiveMode.Create, true))
            {
                var entry = archive.CreateEntry("word/document.xml");
                using var entryStream = entry.Open();
                using var writer = new StreamWriter(entryStream);
                writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" +
                             "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:body><w:p>" +
                             "<w:r><w:rPr><w:u w:val=\"words\"/></w:rPr><w:t>Words</w:t></w:r>" +
                             "<w:r><w:rPr><w:u w:val=\"dotDash\"/></w:rPr><w:t> Dash</w:t></w:r>" +
                             "<w:r><w:rPr><w:u w:val=\"wavyDouble\"/></w:rPr><w:t> Wave</w:t></w:r>" +
                             "</w:p></w:body></w:document>");
            }

            using var testStream = new MemoryStream(ms.ToArray());
            using var reader = new Nedev.FileConverters.DocxToDoc.Format.DocxReader(testStream);

            var model = reader.ReadDocument();
            var runs = Assert.Single(model.Paragraphs).Runs;

            Assert.Equal(Nedev.FileConverters.DocxToDoc.Model.UnderlineType.Single, runs[0].Properties.Underline);
            Assert.Equal(Nedev.FileConverters.DocxToDoc.Model.UnderlineType.Dashed, runs[1].Properties.Underline);
            Assert.Equal(Nedev.FileConverters.DocxToDoc.Model.UnderlineType.Wave, runs[2].Properties.Underline);
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
        public void ReadDocument_WithLastRenderedPageBreak_AppendsFormFeed()
        {
            using var ms = new MemoryStream();
            using (var archive = new ZipArchive(ms, ZipArchiveMode.Create, true))
            {
                var entry = archive.CreateEntry("word/document.xml");
                using var entryStream = entry.Open();
                using var writer = new StreamWriter(entryStream);
                writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" +
                             "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:body><w:p>" +
                             "<w:r><w:t>A</w:t><w:lastRenderedPageBreak/><w:t>B</w:t></w:r>" +
                             "</w:p></w:body></w:document>");
            }

            using var testStream = new MemoryStream(ms.ToArray());
            using var reader = new Nedev.FileConverters.DocxToDoc.Format.DocxReader(testStream);

            var model = reader.ReadDocument();

            var run = Assert.Single(Assert.Single(model.Paragraphs).Runs);
            Assert.Equal("A\fB", run.Text);
            Assert.Equal("A\fB\r", model.TextBuffer);
        }

        [Fact]
        public void ReadDocument_WithColumnBreak_AppendsColumnBreakCharacter()
        {
            using var ms = new MemoryStream();
            using (var archive = new ZipArchive(ms, ZipArchiveMode.Create, true))
            {
                var entry = archive.CreateEntry("word/document.xml");
                using var entryStream = entry.Open();
                using var writer = new StreamWriter(entryStream);
                writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" +
                             "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:body><w:p>" +
                             "<w:r><w:t>A</w:t><w:br w:type=\"column\"/><w:t>B</w:t></w:r>" +
                             "</w:p></w:body></w:document>");
            }

            using var testStream = new MemoryStream(ms.ToArray());
            using var reader = new Nedev.FileConverters.DocxToDoc.Format.DocxReader(testStream);

            var model = reader.ReadDocument();

            var run = Assert.Single(Assert.Single(model.Paragraphs).Runs);
            char columnBreak = '\x000E';
            Assert.Equal($"A{columnBreak}B", run.Text);
            Assert.Equal($"A{columnBreak}B\r", model.TextBuffer);
        }

        [Fact]
        public void ReadDocument_WithTextWrappingBreakClearAll_AppendsClearAllBreakCharacter()
        {
            using var ms = new MemoryStream();
            using (var archive = new ZipArchive(ms, ZipArchiveMode.Create, true))
            {
                var entry = archive.CreateEntry("word/document.xml");
                using var entryStream = entry.Open();
                using var writer = new StreamWriter(entryStream);
                writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" +
                             "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:body><w:p>" +
                             "<w:r><w:t>A</w:t><w:br w:type=\"textWrapping\" w:clear=\"all\"/><w:t>B</w:t></w:r>" +
                             "</w:p></w:body></w:document>");
            }

            using var testStream = new MemoryStream(ms.ToArray());
            using var reader = new Nedev.FileConverters.DocxToDoc.Format.DocxReader(testStream);

            var model = reader.ReadDocument();

            var run = Assert.Single(Assert.Single(model.Paragraphs).Runs);
            char clearAllBreak = '\x001E';
            Assert.Equal($"A{clearAllBreak}B", run.Text);
            Assert.Equal($"A{clearAllBreak}B\r", model.TextBuffer);
        }

        [Theory]
        [InlineData("left", '\x001C')]
        [InlineData("right", '\x001D')]
        public void ReadDocument_WithTextWrappingBreakClearSide_AppendsExpectedClearBreakCharacter(string clearValue, char expectedBreak)
        {
            using var ms = new MemoryStream();
            using (var archive = new ZipArchive(ms, ZipArchiveMode.Create, true))
            {
                var entry = archive.CreateEntry("word/document.xml");
                using var entryStream = entry.Open();
                using var writer = new StreamWriter(entryStream);
                writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" +
                             "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:body><w:p>" +
                             $"<w:r><w:t>A</w:t><w:br w:type=\"textWrapping\" w:clear=\"{clearValue}\"/><w:t>B</w:t></w:r>" +
                             "</w:p></w:body></w:document>");
            }

            using var testStream = new MemoryStream(ms.ToArray());
            using var reader = new Nedev.FileConverters.DocxToDoc.Format.DocxReader(testStream);

            var model = reader.ReadDocument();

            var run = Assert.Single(Assert.Single(model.Paragraphs).Runs);
            Assert.Equal($"A{expectedBreak}B", run.Text);
            Assert.Equal($"A{expectedBreak}B\r", model.TextBuffer);
        }

        [Fact]
        public void ReadDocument_WithPageBreakBefore_ParsesParagraphProperty()
        {
            using var ms = new MemoryStream();
            using (var archive = new ZipArchive(ms, ZipArchiveMode.Create, true))
            {
                var entry = archive.CreateEntry("word/document.xml");
                using var entryStream = entry.Open();
                using var writer = new StreamWriter(entryStream);
                writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" +
                             "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:body>" +
                             "<w:p><w:r><w:t>First</w:t></w:r></w:p>" +
                             "<w:p><w:pPr><w:pageBreakBefore/></w:pPr><w:r><w:t>Second</w:t></w:r></w:p>" +
                             "</w:body></w:document>");
            }

            using var testStream = new MemoryStream(ms.ToArray());
            using var reader = new Nedev.FileConverters.DocxToDoc.Format.DocxReader(testStream);

            var model = reader.ReadDocument();

            Assert.False(model.Paragraphs[0].Properties.PageBreakBefore);
            Assert.True(model.Paragraphs[1].Properties.PageBreakBefore);
        }

        [Fact]
        public void ReadDocument_WithPageBreakBeforeDisabled_ParsesFalseParagraphProperty()
        {
            using var ms = new MemoryStream();
            using (var archive = new ZipArchive(ms, ZipArchiveMode.Create, true))
            {
                var entry = archive.CreateEntry("word/document.xml");
                using var entryStream = entry.Open();
                using var writer = new StreamWriter(entryStream);
                writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" +
                             "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:body>" +
                             "<w:p><w:r><w:t>First</w:t></w:r></w:p>" +
                             "<w:p><w:pPr><w:pageBreakBefore w:val=\"false\"/></w:pPr><w:r><w:t>Second</w:t></w:r></w:p>" +
                             "</w:body></w:document>");
            }

            using var testStream = new MemoryStream(ms.ToArray());
            using var reader = new Nedev.FileConverters.DocxToDoc.Format.DocxReader(testStream);

            var model = reader.ReadDocument();

            Assert.False(model.Paragraphs[0].Properties.PageBreakBefore);
            Assert.False(model.Paragraphs[1].Properties.PageBreakBefore);
        }

        [Fact]
        public void ReadDocument_WithDrawingTextBoxContent_PreservesVisibleText()
        {
            using var ms = new MemoryStream();
            using (var archive = new ZipArchive(ms, ZipArchiveMode.Create, true))
            {
                var entry = archive.CreateEntry("word/document.xml");
                using var entryStream = entry.Open();
                using var writer = new StreamWriter(entryStream);
                writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" +
                             "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" " +
                             "xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\">" +
                             "<w:body><w:p>" +
                             "<w:r><w:t xml:space=\"preserve\">Before </w:t></w:r>" +
                             "<w:r><w:drawing><wp:anchor><w:txbxContent><w:p><w:r><w:t>Box</w:t></w:r></w:p></w:txbxContent></wp:anchor></w:drawing></w:r>" +
                             "<w:r><w:t> after</w:t></w:r>" +
                             "</w:p></w:body></w:document>");
            }

            using var testStream = new MemoryStream(ms.ToArray());
            using var reader = new Nedev.FileConverters.DocxToDoc.Format.DocxReader(testStream);

            var model = reader.ReadDocument();

            var paragraph = Assert.Single(model.Paragraphs);
            Assert.Collection(
                paragraph.Runs.FindAll(run => run.Text.Length > 0),
                run => Assert.Equal("Before ", run.Text),
                run => Assert.Equal("Box", run.Text),
                run => Assert.Equal(" after", run.Text));
            Assert.Equal("Before Box after\r", model.TextBuffer);
        }

        [Fact]
        public void ReadDocument_WithPlainTextAltChunk_PreservesVisibleParagraphsInOrder()
        {
            using var ms = new MemoryStream();
            using (var archive = new ZipArchive(ms, ZipArchiveMode.Create, true))
            {
                var entry = archive.CreateEntry("word/document.xml");
                using (var entryStream = entry.Open())
                using (var writer = new StreamWriter(entryStream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" +
                                 "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"><w:body>" +
                                 "<w:p><w:r><w:t>Before</w:t></w:r></w:p>" +
                                 "<w:altChunk r:id=\"rIdChunk\"/>" +
                                 "<w:p><w:r><w:t>After</w:t></w:r></w:p>" +
                                 "</w:body></w:document>");
                }

                var relsEntry = archive.CreateEntry("word/_rels/document.xml.rels");
                using (var relsStream = relsEntry.Open())
                using (var writer = new StreamWriter(relsStream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                                 "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
                                 "<Relationship Id=\"rIdChunk\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/aFChunk\" Target=\"afchunk1.txt\"/>" +
                                 "</Relationships>");
                }

                var chunkEntry = archive.CreateEntry("word/afchunk1.txt");
                using (var chunkStream = chunkEntry.Open())
                using (var writer = new StreamWriter(chunkStream))
                {
                    writer.Write("Chunk line 1\r\nChunk line 2");
                }
            }

            using var testStream = new MemoryStream(ms.ToArray());
            using var reader = new Nedev.FileConverters.DocxToDoc.Format.DocxReader(testStream);

            var model = reader.ReadDocument();

            Assert.Equal(
                new[] { "Before", "Chunk line 1", "Chunk line 2", "After" },
                model.Paragraphs.Select(paragraph => string.Concat(paragraph.Runs.Select(run => run.Text))).ToArray());
            Assert.Equal("Before\rChunk line 1\rChunk line 2\rAfter\r", model.TextBuffer);
        }

        [Fact]
        public void ReadDocument_WithHtmlAltChunk_PreservesVisibleParagraphsInOrder()
        {
            using var ms = new MemoryStream();
            using (var archive = new ZipArchive(ms, ZipArchiveMode.Create, true))
            {
                var entry = archive.CreateEntry("word/document.xml");
                using (var entryStream = entry.Open())
                using (var writer = new StreamWriter(entryStream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" +
                                 "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"><w:body>" +
                                 "<w:p><w:r><w:t>Lead</w:t></w:r></w:p>" +
                                 "<w:altChunk r:id=\"rIdChunk\"/>" +
                                 "<w:p><w:r><w:t>Tail</w:t></w:r></w:p>" +
                                 "</w:body></w:document>");
                }

                var relsEntry = archive.CreateEntry("word/_rels/document.xml.rels");
                using (var relsStream = relsEntry.Open())
                using (var writer = new StreamWriter(relsStream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                                 "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
                                 "<Relationship Id=\"rIdChunk\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/aFChunk\" Target=\"afchunk2.html\"/>" +
                                 "</Relationships>");
                }

                var chunkEntry = archive.CreateEntry("word/afchunk2.html");
                using (var chunkStream = chunkEntry.Open())
                using (var writer = new StreamWriter(chunkStream))
                {
                    writer.Write("<html><body><p>Alpha</p><div>Beta <b>Bold</b></div><p>Gamma<br/>Delta</p></body></html>");
                }
            }

            using var testStream = new MemoryStream(ms.ToArray());
            using var reader = new Nedev.FileConverters.DocxToDoc.Format.DocxReader(testStream);

            var model = reader.ReadDocument();

            Assert.Equal(
                new[] { "Lead", "Alpha", "Beta Bold", "Gamma", "Delta", "Tail" },
                model.Paragraphs.Select(paragraph => string.Concat(paragraph.Runs.Select(run => run.Text))).ToArray());
            Assert.Equal("Lead\rAlpha\rBeta Bold\rGamma\rDelta\rTail\r", model.TextBuffer);
        }

        [Fact]
        public void ReadDocument_WithAltChunkMissingRelationship_IgnoresChunkAndPreservesSurroundingParagraphs()
        {
            using var ms = new MemoryStream();
            using (var archive = new ZipArchive(ms, ZipArchiveMode.Create, true))
            {
                var entry = archive.CreateEntry("word/document.xml");
                using (var entryStream = entry.Open())
                using (var writer = new StreamWriter(entryStream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" +
                                 "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"><w:body>" +
                                 "<w:p><w:r><w:t>Before</w:t></w:r></w:p>" +
                                 "<w:altChunk r:id=\"rIdMissing\"/>" +
                                 "<w:p><w:r><w:t>After</w:t></w:r></w:p>" +
                                 "</w:body></w:document>");
                }
            }

            using var testStream = new MemoryStream(ms.ToArray());
            using var reader = new Nedev.FileConverters.DocxToDoc.Format.DocxReader(testStream);

            var model = reader.ReadDocument();

            Assert.Equal(
                new[] { "Before", "After" },
                model.Paragraphs.Select(paragraph => string.Concat(paragraph.Runs.Select(run => run.Text))).ToArray());
            Assert.Equal("Before\rAfter\r", model.TextBuffer);
        }

        [Fact]
        public void ReadDocument_WithAltChunkMissingTargetPart_IgnoresChunkAndPreservesSurroundingParagraphs()
        {
            using var ms = new MemoryStream();
            using (var archive = new ZipArchive(ms, ZipArchiveMode.Create, true))
            {
                var entry = archive.CreateEntry("word/document.xml");
                using (var entryStream = entry.Open())
                using (var writer = new StreamWriter(entryStream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" +
                                 "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"><w:body>" +
                                 "<w:p><w:r><w:t>Before</w:t></w:r></w:p>" +
                                 "<w:altChunk r:id=\"rIdChunk\"/>" +
                                 "<w:p><w:r><w:t>After</w:t></w:r></w:p>" +
                                 "</w:body></w:document>");
                }

                var relsEntry = archive.CreateEntry("word/_rels/document.xml.rels");
                using (var relsStream = relsEntry.Open())
                using (var writer = new StreamWriter(relsStream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                                 "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
                                 "<Relationship Id=\"rIdChunk\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/aFChunk\" Target=\"afchunk-missing.txt\"/>" +
                                 "</Relationships>");
                }
            }

            using var testStream = new MemoryStream(ms.ToArray());
            using var reader = new Nedev.FileConverters.DocxToDoc.Format.DocxReader(testStream);

            var model = reader.ReadDocument();

            Assert.Equal(
                new[] { "Before", "After" },
                model.Paragraphs.Select(paragraph => string.Concat(paragraph.Runs.Select(run => run.Text))).ToArray());
            Assert.Equal("Before\rAfter\r", model.TextBuffer);
        }

        [Fact]
        public void ReadDocument_WithRtfAltChunk_PreservesVisibleParagraphsInOrder()
        {
            using var ms = new MemoryStream();
            using (var archive = new ZipArchive(ms, ZipArchiveMode.Create, true))
            {
                var entry = archive.CreateEntry("word/document.xml");
                using (var entryStream = entry.Open())
                using (var writer = new StreamWriter(entryStream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" +
                                 "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"><w:body>" +
                                 "<w:p><w:r><w:t>Before</w:t></w:r></w:p>" +
                                 "<w:altChunk r:id=\"rIdChunk\"/>" +
                                 "<w:p><w:r><w:t>After</w:t></w:r></w:p>" +
                                 "</w:body></w:document>");
                }

                var relsEntry = archive.CreateEntry("word/_rels/document.xml.rels");
                using (var relsStream = relsEntry.Open())
                using (var writer = new StreamWriter(relsStream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                                 "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
                                 "<Relationship Id=\"rIdChunk\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/aFChunk\" Target=\"afchunk3.rtf\"/>" +
                                 "</Relationships>");
                }

                var chunkEntry = archive.CreateEntry("word/afchunk3.rtf");
                using (var chunkStream = chunkEntry.Open())
                using (var writer = new StreamWriter(chunkStream))
                {
                    writer.Write(@"{\rtf1\ansi Chunk line 1\par Chunk line 2}");
                }
            }

            using var testStream = new MemoryStream(ms.ToArray());
            using var reader = new Nedev.FileConverters.DocxToDoc.Format.DocxReader(testStream);

            var model = reader.ReadDocument();

            Assert.Equal(
                new[] { "Before", "Chunk line 1", "Chunk line 2", "After" },
                model.Paragraphs.Select(paragraph => string.Concat(paragraph.Runs.Select(run => run.Text))).ToArray());
            Assert.Equal("Before\rChunk line 1\rChunk line 2\rAfter\r", model.TextBuffer);
        }

        [Fact]
        public void ReadDocument_WithRtfAltChunkUnicodeEscapes_PreservesVisibleParagraphsInOrder()
        {
            using var ms = new MemoryStream();
            using (var archive = new ZipArchive(ms, ZipArchiveMode.Create, true))
            {
                var entry = archive.CreateEntry("word/document.xml");
                using (var entryStream = entry.Open())
                using (var writer = new StreamWriter(entryStream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" +
                                 "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"><w:body>" +
                                 "<w:p><w:r><w:t>Lead</w:t></w:r></w:p>" +
                                 "<w:altChunk r:id=\"rIdChunk\"/>" +
                                 "<w:p><w:r><w:t>Tail</w:t></w:r></w:p>" +
                                 "</w:body></w:document>");
                }

                var relsEntry = archive.CreateEntry("word/_rels/document.xml.rels");
                using (var relsStream = relsEntry.Open())
                using (var writer = new StreamWriter(relsStream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                                 "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
                                 "<Relationship Id=\"rIdChunk\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/aFChunk\" Target=\"afchunk5.rtf\"/>" +
                                 "</Relationships>");
                }

                var chunkEntry = archive.CreateEntry("word/afchunk5.rtf");
                using (var chunkStream = chunkEntry.Open())
                using (var writer = new StreamWriter(chunkStream))
                {
                    writer.Write(@"{\rtf1\ansi\uc1 Alpha \u945? beta\par Gamma}");
                }
            }

            using var testStream = new MemoryStream(ms.ToArray());
            using var reader = new Nedev.FileConverters.DocxToDoc.Format.DocxReader(testStream);

            var model = reader.ReadDocument();

            Assert.Equal(
                new[] { "Lead", "Alpha α beta", "Gamma", "Tail" },
                model.Paragraphs.Select(paragraph => string.Concat(paragraph.Runs.Select(run => run.Text))).ToArray());
            Assert.Equal("Lead\rAlpha α beta\rGamma\rTail\r", model.TextBuffer);
        }

        [Fact]
        public void ReadDocument_WithMalformedRtfAltChunk_DowngradesToPlainTextParagraphs()
        {
            using var ms = new MemoryStream();
            using (var archive = new ZipArchive(ms, ZipArchiveMode.Create, true))
            {
                var entry = archive.CreateEntry("word/document.xml");
                using (var entryStream = entry.Open())
                using (var writer = new StreamWriter(entryStream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" +
                                 "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"><w:body>" +
                                 "<w:p><w:r><w:t>Before</w:t></w:r></w:p>" +
                                 "<w:altChunk r:id=\"rIdChunk\"/>" +
                                 "<w:p><w:r><w:t>After</w:t></w:r></w:p>" +
                                 "</w:body></w:document>");
                }

                var relsEntry = archive.CreateEntry("word/_rels/document.xml.rels");
                using (var relsStream = relsEntry.Open())
                using (var writer = new StreamWriter(relsStream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                                 "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
                                 "<Relationship Id=\"rIdChunk\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/aFChunk\" Target=\"afchunk6.rtf\"/>" +
                                 "</Relationships>");
                }

                var chunkEntry = archive.CreateEntry("word/afchunk6.rtf");
                using (var chunkStream = chunkEntry.Open())
                using (var writer = new StreamWriter(chunkStream))
                {
                    writer.Write("Chunk line 1\r\nChunk line 2");
                }
            }

            using var testStream = new MemoryStream(ms.ToArray());
            using var reader = new Nedev.FileConverters.DocxToDoc.Format.DocxReader(testStream);

            var model = reader.ReadDocument();

            Assert.Equal(
                new[] { "Before", "Chunk line 1", "Chunk line 2", "After" },
                model.Paragraphs.Select(paragraph => string.Concat(paragraph.Runs.Select(run => run.Text))).ToArray());
            Assert.Equal("Before\rChunk line 1\rChunk line 2\rAfter\r", model.TextBuffer);
        }

        [Fact]
        public void ReadDocument_WithAltChunkBetweenParagraphAndTable_PreservesMixedContentOrder()
        {
            using var ms = new MemoryStream();
            using (var archive = new ZipArchive(ms, ZipArchiveMode.Create, true))
            {
                var entry = archive.CreateEntry("word/document.xml");
                using (var entryStream = entry.Open())
                using (var writer = new StreamWriter(entryStream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" +
                                 "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"><w:body>" +
                                 "<w:p><w:r><w:t>Before</w:t></w:r></w:p>" +
                                 "<w:altChunk r:id=\"rIdChunk\"/>" +
                                 "<w:tbl><w:tr><w:tc><w:p><w:r><w:t>Cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl>" +
                                 "<w:p><w:r><w:t>After</w:t></w:r></w:p>" +
                                 "</w:body></w:document>");
                }

                var relsEntry = archive.CreateEntry("word/_rels/document.xml.rels");
                using (var relsStream = relsEntry.Open())
                using (var writer = new StreamWriter(relsStream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                                 "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
                                 "<Relationship Id=\"rIdChunk\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/aFChunk\" Target=\"afchunk4.txt\"/>" +
                                 "</Relationships>");
                }

                var chunkEntry = archive.CreateEntry("word/afchunk4.txt");
                using (var chunkStream = chunkEntry.Open())
                using (var writer = new StreamWriter(chunkStream))
                {
                    writer.Write("Chunk line 1\r\nChunk line 2");
                }
            }

            using var testStream = new MemoryStream(ms.ToArray());
            using var reader = new Nedev.FileConverters.DocxToDoc.Format.DocxReader(testStream);

            var model = reader.ReadDocument();

            Assert.Collection(
                model.Content,
                block => Assert.Equal("Before", string.Concat(Assert.IsType<Nedev.FileConverters.DocxToDoc.Model.ParagraphModel>(block).Runs.Select(run => run.Text))),
                block => Assert.Equal("Chunk line 1", string.Concat(Assert.IsType<Nedev.FileConverters.DocxToDoc.Model.ParagraphModel>(block).Runs.Select(run => run.Text))),
                block => Assert.Equal("Chunk line 2", string.Concat(Assert.IsType<Nedev.FileConverters.DocxToDoc.Model.ParagraphModel>(block).Runs.Select(run => run.Text))),
                block =>
                {
                    var table = Assert.IsType<Nedev.FileConverters.DocxToDoc.Model.TableModel>(block);
                    Assert.Equal("Cell", table.Rows[0].Cells[0].Paragraphs[0].Runs[0].Text);
                },
                block => Assert.Equal("After", string.Concat(Assert.IsType<Nedev.FileConverters.DocxToDoc.Model.ParagraphModel>(block).Runs.Select(run => run.Text))));

            Assert.Equal(
                new[] { "Before", "Chunk line 1", "Chunk line 2", "After" },
                model.Paragraphs.Select(paragraph => string.Concat(paragraph.Runs.Select(run => run.Text))).ToArray());
            Assert.Equal("Before\rChunk line 1\rChunk line 2\rCell\rAfter\r", model.TextBuffer);
        }

        [Fact]
        public void ReadDocument_WithRtfAltChunkBetweenParagraphAndTable_PreservesMixedContentOrder()
        {
            using var ms = new MemoryStream();
            using (var archive = new ZipArchive(ms, ZipArchiveMode.Create, true))
            {
                var entry = archive.CreateEntry("word/document.xml");
                using (var entryStream = entry.Open())
                using (var writer = new StreamWriter(entryStream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" +
                                 "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"><w:body>" +
                                 "<w:p><w:r><w:t>Before</w:t></w:r></w:p>" +
                                 "<w:altChunk r:id=\"rIdChunk\"/>" +
                                 "<w:tbl><w:tr><w:tc><w:p><w:r><w:t>Cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl>" +
                                 "<w:p><w:r><w:t>After</w:t></w:r></w:p>" +
                                 "</w:body></w:document>");
                }

                var relsEntry = archive.CreateEntry("word/_rels/document.xml.rels");
                using (var relsStream = relsEntry.Open())
                using (var writer = new StreamWriter(relsStream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                                 "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
                                 "<Relationship Id=\"rIdChunk\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/aFChunk\" Target=\"afchunk7.rtf\"/>" +
                                 "</Relationships>");
                }

                var chunkEntry = archive.CreateEntry("word/afchunk7.rtf");
                using (var chunkStream = chunkEntry.Open())
                using (var writer = new StreamWriter(chunkStream))
                {
                    writer.Write(@"{\rtf1\ansi Chunk line 1\par Chunk line 2}");
                }
            }

            using var testStream = new MemoryStream(ms.ToArray());
            using var reader = new Nedev.FileConverters.DocxToDoc.Format.DocxReader(testStream);

            var model = reader.ReadDocument();

            Assert.Collection(
                model.Content,
                block => Assert.Equal("Before", string.Concat(Assert.IsType<Nedev.FileConverters.DocxToDoc.Model.ParagraphModel>(block).Runs.Select(run => run.Text))),
                block => Assert.Equal("Chunk line 1", string.Concat(Assert.IsType<Nedev.FileConverters.DocxToDoc.Model.ParagraphModel>(block).Runs.Select(run => run.Text))),
                block => Assert.Equal("Chunk line 2", string.Concat(Assert.IsType<Nedev.FileConverters.DocxToDoc.Model.ParagraphModel>(block).Runs.Select(run => run.Text))),
                block =>
                {
                    var table = Assert.IsType<Nedev.FileConverters.DocxToDoc.Model.TableModel>(block);
                    Assert.Equal("Cell", table.Rows[0].Cells[0].Paragraphs[0].Runs[0].Text);
                },
                block => Assert.Equal("After", string.Concat(Assert.IsType<Nedev.FileConverters.DocxToDoc.Model.ParagraphModel>(block).Runs.Select(run => run.Text))));

            Assert.Equal(
                new[] { "Before", "Chunk line 1", "Chunk line 2", "After" },
                model.Paragraphs.Select(paragraph => string.Concat(paragraph.Runs.Select(run => run.Text))).ToArray());
            Assert.Equal("Before\rChunk line 1\rChunk line 2\rCell\rAfter\r", model.TextBuffer);
        }

        [Fact]
        public void ReadDocument_WithEmbeddedDocxAltChunk_PreservesVisibleParagraphsInOrder()
        {
            using var ms = new MemoryStream();
            using (var archive = new ZipArchive(ms, ZipArchiveMode.Create, true))
            {
                var entry = archive.CreateEntry("word/document.xml");
                using (var entryStream = entry.Open())
                using (var writer = new StreamWriter(entryStream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" +
                                 "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"><w:body>" +
                                 "<w:p><w:r><w:t>Before</w:t></w:r></w:p>" +
                                 "<w:altChunk r:id=\"rIdChunk\"/>" +
                                 "<w:p><w:r><w:t>After</w:t></w:r></w:p>" +
                                 "</w:body></w:document>");
                }

                var relsEntry = archive.CreateEntry("word/_rels/document.xml.rels");
                using (var relsStream = relsEntry.Open())
                using (var writer = new StreamWriter(relsStream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                                 "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
                                 "<Relationship Id=\"rIdChunk\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/aFChunk\" Target=\"afchunk8.docx\"/>" +
                                 "</Relationships>");
                }

                var chunkEntry = archive.CreateEntry("word/afchunk8.docx");
                using (var chunkStream = chunkEntry.Open())
                {
                    byte[] chunkBytes = CreateEmbeddedDocxChunk(includeTable: false);
                    chunkStream.Write(chunkBytes, 0, chunkBytes.Length);
                }
            }

            using var testStream = new MemoryStream(ms.ToArray());
            using var reader = new Nedev.FileConverters.DocxToDoc.Format.DocxReader(testStream);

            var model = reader.ReadDocument();

            Assert.Equal(
                new[] { "Before", "Inner lead", "Inner tail", "After" },
                model.Paragraphs.Select(paragraph => string.Concat(paragraph.Runs.Select(run => run.Text))).ToArray());
            Assert.Equal("Before\rInner lead\rInner tail\rAfter\r", model.TextBuffer);
        }

        [Fact]
        public void ReadDocument_WithEmbeddedDocxAltChunkContainingTable_PreservesTableBlocksAndVisibleContentOrder()
        {
            using var ms = new MemoryStream();
            using (var archive = new ZipArchive(ms, ZipArchiveMode.Create, true))
            {
                var entry = archive.CreateEntry("word/document.xml");
                using (var entryStream = entry.Open())
                using (var writer = new StreamWriter(entryStream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" +
                                 "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"><w:body>" +
                                 "<w:p><w:r><w:t>Before</w:t></w:r></w:p>" +
                                 "<w:altChunk r:id=\"rIdChunk\"/>" +
                                 "<w:p><w:r><w:t>After</w:t></w:r></w:p>" +
                                 "</w:body></w:document>");
                }

                var relsEntry = archive.CreateEntry("word/_rels/document.xml.rels");
                using (var relsStream = relsEntry.Open())
                using (var writer = new StreamWriter(relsStream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                                 "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
                                 "<Relationship Id=\"rIdChunk\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/aFChunk\" Target=\"afchunk9.docx\"/>" +
                                 "</Relationships>");
                }

                var chunkEntry = archive.CreateEntry("word/afchunk9.docx");
                using (var chunkStream = chunkEntry.Open())
                {
                    byte[] chunkBytes = CreateEmbeddedDocxChunk(includeTable: true);
                    chunkStream.Write(chunkBytes, 0, chunkBytes.Length);
                }
            }

            using var testStream = new MemoryStream(ms.ToArray());
            using var reader = new Nedev.FileConverters.DocxToDoc.Format.DocxReader(testStream);

            var model = reader.ReadDocument();

            Assert.Collection(
                model.Content,
                block => Assert.Equal("Before", string.Concat(Assert.IsType<Nedev.FileConverters.DocxToDoc.Model.ParagraphModel>(block).Runs.Select(run => run.Text))),
                block => Assert.Equal("Inner lead", string.Concat(Assert.IsType<Nedev.FileConverters.DocxToDoc.Model.ParagraphModel>(block).Runs.Select(run => run.Text))),
                block =>
                {
                    var table = Assert.IsType<Nedev.FileConverters.DocxToDoc.Model.TableModel>(block);
                    Assert.Equal("Cell", table.Rows[0].Cells[0].Paragraphs[0].Runs[0].Text);
                },
                block => Assert.Equal("Inner tail", string.Concat(Assert.IsType<Nedev.FileConverters.DocxToDoc.Model.ParagraphModel>(block).Runs.Select(run => run.Text))),
                block => Assert.Equal("After", string.Concat(Assert.IsType<Nedev.FileConverters.DocxToDoc.Model.ParagraphModel>(block).Runs.Select(run => run.Text))));

            Assert.Equal(
                new[] { "Before", "Inner lead", "Inner tail", "After" },
                model.Paragraphs.Select(paragraph => string.Concat(paragraph.Runs.Select(run => run.Text))).ToArray());
            Assert.Equal("Before\rInner lead\rCell\rInner tail\rAfter\r", model.TextBuffer);
        }

        [Fact]
        public void ReadDocument_WithEmbeddedDocxAltChunkMissingTargetPart_IgnoresChunkAndPreservesSurroundingParagraphs()
        {
            using var ms = new MemoryStream();
            using (var archive = new ZipArchive(ms, ZipArchiveMode.Create, true))
            {
                var entry = archive.CreateEntry("word/document.xml");
                using (var entryStream = entry.Open())
                using (var writer = new StreamWriter(entryStream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" +
                                 "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"><w:body>" +
                                 "<w:p><w:r><w:t>Before</w:t></w:r></w:p>" +
                                 "<w:altChunk r:id=\"rIdChunk\"/>" +
                                 "<w:p><w:r><w:t>After</w:t></w:r></w:p>" +
                                 "</w:body></w:document>");
                }

                var relsEntry = archive.CreateEntry("word/_rels/document.xml.rels");
                using (var relsStream = relsEntry.Open())
                using (var writer = new StreamWriter(relsStream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                                 "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
                                 "<Relationship Id=\"rIdChunk\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/aFChunk\" Target=\"afchunk-missing.docx\"/>" +
                                 "</Relationships>");
                }
            }

            using var testStream = new MemoryStream(ms.ToArray());
            using var reader = new Nedev.FileConverters.DocxToDoc.Format.DocxReader(testStream);

            var model = reader.ReadDocument();

            Assert.Equal(
                new[] { "Before", "After" },
                model.Paragraphs.Select(paragraph => string.Concat(paragraph.Runs.Select(run => run.Text))).ToArray());
            Assert.Equal("Before\rAfter\r", model.TextBuffer);
        }

        [Fact]
        public void ReadDocument_WithMalformedEmbeddedDocxAltChunk_IgnoresChunkAndPreservesSurroundingParagraphs()
        {
            using var ms = new MemoryStream();
            using (var archive = new ZipArchive(ms, ZipArchiveMode.Create, true))
            {
                var entry = archive.CreateEntry("word/document.xml");
                using (var entryStream = entry.Open())
                using (var writer = new StreamWriter(entryStream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" +
                                 "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"><w:body>" +
                                 "<w:p><w:r><w:t>Before</w:t></w:r></w:p>" +
                                 "<w:altChunk r:id=\"rIdChunk\"/>" +
                                 "<w:p><w:r><w:t>After</w:t></w:r></w:p>" +
                                 "</w:body></w:document>");
                }

                var relsEntry = archive.CreateEntry("word/_rels/document.xml.rels");
                using (var relsStream = relsEntry.Open())
                using (var writer = new StreamWriter(relsStream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                                 "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
                                 "<Relationship Id=\"rIdChunk\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/aFChunk\" Target=\"afchunk10.docx\"/>" +
                                 "</Relationships>");
                }

                var chunkEntry = archive.CreateEntry("word/afchunk10.docx");
                using var chunkStream = chunkEntry.Open();
                byte[] chunkBytes = Encoding.UTF8.GetBytes("not-a-zip-package");
                chunkStream.Write(chunkBytes, 0, chunkBytes.Length);
            }

            using var testStream = new MemoryStream(ms.ToArray());
            using var reader = new Nedev.FileConverters.DocxToDoc.Format.DocxReader(testStream);

            var model = reader.ReadDocument();

            Assert.Equal(
                new[] { "Before", "After" },
                model.Paragraphs.Select(paragraph => string.Concat(paragraph.Runs.Select(run => run.Text))).ToArray());
            Assert.Equal("Before\rAfter\r", model.TextBuffer);
        }

        [Fact]
        public void ReadDocument_WithEmbeddedDocxAltChunkMissingDocumentXml_IgnoresChunkAndPreservesSurroundingParagraphs()
        {
            using var ms = new MemoryStream();
            using (var archive = new ZipArchive(ms, ZipArchiveMode.Create, true))
            {
                var entry = archive.CreateEntry("word/document.xml");
                using (var entryStream = entry.Open())
                using (var writer = new StreamWriter(entryStream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" +
                                 "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"><w:body>" +
                                 "<w:p><w:r><w:t>Before</w:t></w:r></w:p>" +
                                 "<w:altChunk r:id=\"rIdChunk\"/>" +
                                 "<w:p><w:r><w:t>After</w:t></w:r></w:p>" +
                                 "</w:body></w:document>");
                }

                var relsEntry = archive.CreateEntry("word/_rels/document.xml.rels");
                using (var relsStream = relsEntry.Open())
                using (var writer = new StreamWriter(relsStream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                                 "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
                                 "<Relationship Id=\"rIdChunk\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/aFChunk\" Target=\"afchunk11.docx\"/>" +
                                 "</Relationships>");
                }

                var chunkEntry = archive.CreateEntry("word/afchunk11.docx");
                using (var chunkStream = chunkEntry.Open())
                {
                    byte[] chunkBytes = CreateEmbeddedDocxChunkWithoutDocumentXml();
                    chunkStream.Write(chunkBytes, 0, chunkBytes.Length);
                }
            }

            using var testStream = new MemoryStream(ms.ToArray());
            using var reader = new Nedev.FileConverters.DocxToDoc.Format.DocxReader(testStream);

            var model = reader.ReadDocument();

            Assert.Equal(
                new[] { "Before", "After" },
                model.Paragraphs.Select(paragraph => string.Concat(paragraph.Runs.Select(run => run.Text))).ToArray());
            Assert.Equal("Before\rAfter\r", model.TextBuffer);
        }

        [Fact]
        public void ReadDocument_WithEmbeddedDocxAltChunkBetweenParagraphAndTable_PreservesMixedContentOrder()
        {
            using var ms = new MemoryStream();
            using (var archive = new ZipArchive(ms, ZipArchiveMode.Create, true))
            {
                var entry = archive.CreateEntry("word/document.xml");
                using (var entryStream = entry.Open())
                using (var writer = new StreamWriter(entryStream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" +
                                 "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"><w:body>" +
                                 "<w:p><w:r><w:t>Before</w:t></w:r></w:p>" +
                                 "<w:altChunk r:id=\"rIdChunk\"/>" +
                                 "<w:tbl><w:tr><w:tc><w:p><w:r><w:t>Outer cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl>" +
                                 "<w:p><w:r><w:t>After</w:t></w:r></w:p>" +
                                 "</w:body></w:document>");
                }

                var relsEntry = archive.CreateEntry("word/_rels/document.xml.rels");
                using (var relsStream = relsEntry.Open())
                using (var writer = new StreamWriter(relsStream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                                 "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
                                 "<Relationship Id=\"rIdChunk\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/aFChunk\" Target=\"afchunk12.docx\"/>" +
                                 "</Relationships>");
                }

                var chunkEntry = archive.CreateEntry("word/afchunk12.docx");
                using (var chunkStream = chunkEntry.Open())
                {
                    byte[] chunkBytes = CreateEmbeddedDocxChunk(includeTable: true);
                    chunkStream.Write(chunkBytes, 0, chunkBytes.Length);
                }
            }

            using var testStream = new MemoryStream(ms.ToArray());
            using var reader = new Nedev.FileConverters.DocxToDoc.Format.DocxReader(testStream);

            var model = reader.ReadDocument();

            Assert.Collection(
                model.Content,
                block => Assert.Equal("Before", string.Concat(Assert.IsType<Nedev.FileConverters.DocxToDoc.Model.ParagraphModel>(block).Runs.Select(run => run.Text))),
                block => Assert.Equal("Inner lead", string.Concat(Assert.IsType<Nedev.FileConverters.DocxToDoc.Model.ParagraphModel>(block).Runs.Select(run => run.Text))),
                block =>
                {
                    var table = Assert.IsType<Nedev.FileConverters.DocxToDoc.Model.TableModel>(block);
                    Assert.Equal("Cell", table.Rows[0].Cells[0].Paragraphs[0].Runs[0].Text);
                },
                block => Assert.Equal("Inner tail", string.Concat(Assert.IsType<Nedev.FileConverters.DocxToDoc.Model.ParagraphModel>(block).Runs.Select(run => run.Text))),
                block =>
                {
                    var table = Assert.IsType<Nedev.FileConverters.DocxToDoc.Model.TableModel>(block);
                    Assert.Equal("Outer cell", table.Rows[0].Cells[0].Paragraphs[0].Runs[0].Text);
                },
                block => Assert.Equal("After", string.Concat(Assert.IsType<Nedev.FileConverters.DocxToDoc.Model.ParagraphModel>(block).Runs.Select(run => run.Text))));

            Assert.Equal(
                new[] { "Before", "Inner lead", "Inner tail", "After" },
                model.Paragraphs.Select(paragraph => string.Concat(paragraph.Runs.Select(run => run.Text))).ToArray());
            Assert.Equal("Before\rInner lead\rCell\rInner tail\rOuter cell\rAfter\r", model.TextBuffer);
        }

        [Fact]
        public void ReadDocument_WithEmbeddedDocxAltChunkInsideTableCell_PreservesTableBlocksInsideCellContent()
        {
            using var ms = new MemoryStream();
            using (var archive = new ZipArchive(ms, ZipArchiveMode.Create, true))
            {
                var entry = archive.CreateEntry("word/document.xml");
                using (var entryStream = entry.Open())
                using (var writer = new StreamWriter(entryStream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" +
                                 "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"><w:body>" +
                                 "<w:p><w:r><w:t>Before</w:t></w:r></w:p>" +
                                 "<w:tbl><w:tr><w:tc><w:altChunk r:id=\"rIdChunk\"/></w:tc></w:tr></w:tbl>" +
                                 "<w:p><w:r><w:t>After</w:t></w:r></w:p>" +
                                 "</w:body></w:document>");
                }

                var relsEntry = archive.CreateEntry("word/_rels/document.xml.rels");
                using (var relsStream = relsEntry.Open())
                using (var writer = new StreamWriter(relsStream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                                 "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
                                 "<Relationship Id=\"rIdChunk\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/aFChunk\" Target=\"afchunk13.docx\"/>" +
                                 "</Relationships>");
                }

                var chunkEntry = archive.CreateEntry("word/afchunk13.docx");
                using (var chunkStream = chunkEntry.Open())
                {
                    byte[] chunkBytes = CreateEmbeddedDocxChunk(includeTable: true);
                    chunkStream.Write(chunkBytes, 0, chunkBytes.Length);
                }
            }

            using var testStream = new MemoryStream(ms.ToArray());
            using var reader = new Nedev.FileConverters.DocxToDoc.Format.DocxReader(testStream);

            var model = reader.ReadDocument();

            Assert.Collection(
                model.Content,
                block => Assert.Equal("Before", string.Concat(Assert.IsType<Nedev.FileConverters.DocxToDoc.Model.ParagraphModel>(block).Runs.Select(run => run.Text))),
                block =>
                {
                    var table = Assert.IsType<Nedev.FileConverters.DocxToDoc.Model.TableModel>(block);
                    var cell = table.Rows[0].Cells[0];
                    Assert.Collection(
                        cell.Content,
                        cellBlock => Assert.Equal("Inner lead", string.Concat(Assert.IsType<Nedev.FileConverters.DocxToDoc.Model.ParagraphModel>(cellBlock).Runs.Select(run => run.Text))),
                        cellBlock =>
                        {
                            var nestedTable = Assert.IsType<Nedev.FileConverters.DocxToDoc.Model.TableModel>(cellBlock);
                            Assert.Equal("Cell", nestedTable.Rows[0].Cells[0].Paragraphs[0].Runs[0].Text);
                        },
                        cellBlock => Assert.Equal("Inner tail", string.Concat(Assert.IsType<Nedev.FileConverters.DocxToDoc.Model.ParagraphModel>(cellBlock).Runs.Select(run => run.Text))));
                    Assert.Equal(
                        new[] { "Inner lead", "Cell", "Inner tail" },
                        cell.Paragraphs.Select(paragraph => string.Concat(paragraph.Runs.Select(run => run.Text))).ToArray());
                },
                block => Assert.Equal("After", string.Concat(Assert.IsType<Nedev.FileConverters.DocxToDoc.Model.ParagraphModel>(block).Runs.Select(run => run.Text))));

            Assert.Equal(
                new[] { "Before", "After" },
                model.Paragraphs.Select(paragraph => string.Concat(paragraph.Runs.Select(run => run.Text))).ToArray());
            Assert.Equal("Before\rInner lead\rCell\rInner tail\rAfter\r", model.TextBuffer);
        }

        [Fact]
        public void ReadDocument_WithEmbeddedDocxAltChunkBetweenCellParagraphs_PreservesMixedBlockOrderInsideCellContent()
        {
            using var ms = new MemoryStream();
            using (var archive = new ZipArchive(ms, ZipArchiveMode.Create, true))
            {
                var entry = archive.CreateEntry("word/document.xml");
                using (var entryStream = entry.Open())
                using (var writer = new StreamWriter(entryStream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" +
                                 "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"><w:body>" +
                                 "<w:p><w:r><w:t>Before</w:t></w:r></w:p>" +
                                 "<w:tbl><w:tr><w:tc>" +
                                 "<w:p><w:r><w:t>Cell lead</w:t></w:r></w:p>" +
                                 "<w:altChunk r:id=\"rIdChunk\"/>" +
                                 "<w:p><w:r><w:t>Cell tail</w:t></w:r></w:p>" +
                                 "</w:tc></w:tr></w:tbl>" +
                                 "<w:p><w:r><w:t>After</w:t></w:r></w:p>" +
                                 "</w:body></w:document>");
                }

                var relsEntry = archive.CreateEntry("word/_rels/document.xml.rels");
                using (var relsStream = relsEntry.Open())
                using (var writer = new StreamWriter(relsStream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                                 "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
                                 "<Relationship Id=\"rIdChunk\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/aFChunk\" Target=\"afchunk14.docx\"/>" +
                                 "</Relationships>");
                }

                var chunkEntry = archive.CreateEntry("word/afchunk14.docx");
                using (var chunkStream = chunkEntry.Open())
                {
                    byte[] chunkBytes = CreateEmbeddedDocxChunk(includeTable: true);
                    chunkStream.Write(chunkBytes, 0, chunkBytes.Length);
                }
            }

            using var testStream = new MemoryStream(ms.ToArray());
            using var reader = new Nedev.FileConverters.DocxToDoc.Format.DocxReader(testStream);

            var model = reader.ReadDocument();

            Assert.Collection(
                model.Content,
                block => Assert.Equal("Before", string.Concat(Assert.IsType<Nedev.FileConverters.DocxToDoc.Model.ParagraphModel>(block).Runs.Select(run => run.Text))),
                block =>
                {
                    var table = Assert.IsType<Nedev.FileConverters.DocxToDoc.Model.TableModel>(block);
                    var cell = table.Rows[0].Cells[0];
                    Assert.Collection(
                        cell.Content,
                        cellBlock => Assert.Equal("Cell lead", string.Concat(Assert.IsType<Nedev.FileConverters.DocxToDoc.Model.ParagraphModel>(cellBlock).Runs.Select(run => run.Text))),
                        cellBlock => Assert.Equal("Inner lead", string.Concat(Assert.IsType<Nedev.FileConverters.DocxToDoc.Model.ParagraphModel>(cellBlock).Runs.Select(run => run.Text))),
                        cellBlock =>
                        {
                            var nestedTable = Assert.IsType<Nedev.FileConverters.DocxToDoc.Model.TableModel>(cellBlock);
                            Assert.Equal("Cell", nestedTable.Rows[0].Cells[0].Paragraphs[0].Runs[0].Text);
                        },
                        cellBlock => Assert.Equal("Inner tail", string.Concat(Assert.IsType<Nedev.FileConverters.DocxToDoc.Model.ParagraphModel>(cellBlock).Runs.Select(run => run.Text))),
                        cellBlock => Assert.Equal("Cell tail", string.Concat(Assert.IsType<Nedev.FileConverters.DocxToDoc.Model.ParagraphModel>(cellBlock).Runs.Select(run => run.Text))));
                    Assert.Equal(
                        new[] { "Cell lead", "Inner lead", "Cell", "Inner tail", "Cell tail" },
                        cell.Paragraphs.Select(paragraph => string.Concat(paragraph.Runs.Select(run => run.Text))).ToArray());
                },
                block => Assert.Equal("After", string.Concat(Assert.IsType<Nedev.FileConverters.DocxToDoc.Model.ParagraphModel>(block).Runs.Select(run => run.Text))));

            Assert.Equal(
                new[] { "Before", "After" },
                model.Paragraphs.Select(paragraph => string.Concat(paragraph.Runs.Select(run => run.Text))).ToArray());
            Assert.Equal("Before\rCell lead\rInner lead\rCell\rInner tail\rCell tail\rAfter\r", model.TextBuffer);
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

        private static byte[] CreateEmbeddedDocxChunk(bool includeTable)
        {
            using var ms = new MemoryStream();
            using (var archive = new ZipArchive(ms, ZipArchiveMode.Create, true))
            {
                var entry = archive.CreateEntry("word/document.xml");
                using (var entryStream = entry.Open())
                using (var writer = new StreamWriter(entryStream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" +
                                 "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:body>" +
                                 "<w:p><w:r><w:t>Inner lead</w:t></w:r></w:p>");

                    if (includeTable)
                    {
                        writer.Write("<w:tbl><w:tr><w:tc><w:p><w:r><w:t>Cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl>");
                    }

                    writer.Write("<w:p><w:r><w:t>Inner tail</w:t></w:r></w:p>" +
                                 "</w:body></w:document>");
                }
            }

            return ms.ToArray();
        }

        private static byte[] CreateEmbeddedDocxChunkWithoutDocumentXml()
        {
            using var ms = new MemoryStream();
            using (var archive = new ZipArchive(ms, ZipArchiveMode.Create, true))
            {
                var entry = archive.CreateEntry("[Content_Types].xml");
                using var entryStream = entry.Open();
                using var writer = new StreamWriter(entryStream);
                writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\"/>");
            }

            return ms.ToArray();
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
        public void ReadDocument_WithHorizontalMerge_CombinesContinueCellsIntoGridSpanUsingGridWidths()
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
                             "<w:tc><w:tcPr><w:hMerge w:val=\"restart\"/></w:tcPr><w:p><w:r><w:t>A</w:t></w:r></w:p></w:tc>" +
                             "<w:tc><w:tcPr><w:tcW w:w=\"2500\" w:type=\"pct\"/><w:hMerge/></w:tcPr><w:p><w:r><w:t>B</w:t></w:r></w:p></w:tc>" +
                             "<w:tc><w:p><w:r><w:t>C</w:t></w:r></w:p></w:tc>" +
                             "</w:tr></w:tbl></w:body></w:document>");
            }

            using var testStream = new MemoryStream(ms.ToArray());
            using var reader = new Nedev.FileConverters.DocxToDoc.Format.DocxReader(testStream);

            var model = reader.ReadDocument();

            var table = Assert.IsType<Nedev.FileConverters.DocxToDoc.Model.TableModel>(Assert.Single(model.Content));
            var row = Assert.Single(table.Rows);
            Assert.Equal(2, row.Cells.Count);
            Assert.Equal(2, row.Cells[0].GridSpan);
            Assert.Equal(3000, row.Cells[0].Width);
            Assert.Equal(900, row.Cells[1].Width);
        }

        [Fact]
        public void ReadDocument_WithHorizontalMergeWithoutGrid_UsesDxaWidthFallbackForContinueCell()
        {
            using var ms = new MemoryStream();
            using (var archive = new ZipArchive(ms, ZipArchiveMode.Create, true))
            {
                var entry = archive.CreateEntry("word/document.xml");
                using var entryStream = entry.Open();
                using var writer = new StreamWriter(entryStream);
                writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" +
                             "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:body>" +
                             "<w:tbl><w:tr>" +
                             "<w:tc><w:tcPr><w:tcW w:w=\"1000\" w:type=\"dxa\"/><w:hMerge w:val=\"restart\"/></w:tcPr><w:p><w:r><w:t>A</w:t></w:r></w:p></w:tc>" +
                             "<w:tc><w:tcPr><w:tcW w:w=\"2000\" w:type=\"dxa\"/><w:hMerge/></w:tcPr><w:p><w:r><w:t>B</w:t></w:r></w:p></w:tc>" +
                             "<w:tc><w:tcPr><w:tcW w:w=\"900\" w:type=\"dxa\"/></w:tcPr><w:p><w:r><w:t>C</w:t></w:r></w:p></w:tc>" +
                             "</w:tr></w:tbl></w:body></w:document>");
            }

            using var testStream = new MemoryStream(ms.ToArray());
            using var reader = new Nedev.FileConverters.DocxToDoc.Format.DocxReader(testStream);

            var model = reader.ReadDocument();

            var table = Assert.IsType<Nedev.FileConverters.DocxToDoc.Model.TableModel>(Assert.Single(model.Content));
            var row = Assert.Single(table.Rows);
            Assert.Equal(2, row.Cells.Count);
            Assert.Equal(2, row.Cells[0].GridSpan);
            Assert.Equal(3000, row.Cells[0].Width);
            Assert.Equal(900, row.Cells[1].Width);
        }

        [Fact]
        public void ReadDocument_WithHorizontalMergeWithoutGrid_AccumulatesPctWidthWhenAnchorIsPct()
        {
            using var ms = new MemoryStream();
            using (var archive = new ZipArchive(ms, ZipArchiveMode.Create, true))
            {
                var entry = archive.CreateEntry("word/document.xml");
                using var entryStream = entry.Open();
                using var writer = new StreamWriter(entryStream);
                writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" +
                             "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:body>" +
                             "<w:tbl><w:tr>" +
                             "<w:tc><w:tcPr><w:tcW w:w=\"2000\" w:type=\"pct\"/><w:hMerge w:val=\"restart\"/></w:tcPr><w:p><w:r><w:t>A</w:t></w:r></w:p></w:tc>" +
                             "<w:tc><w:tcPr><w:tcW w:w=\"1500\" w:type=\"pct\"/><w:hMerge/></w:tcPr><w:p><w:r><w:t>B</w:t></w:r></w:p></w:tc>" +
                             "<w:tc><w:tcPr><w:tcW w:w=\"500\" w:type=\"pct\"/></w:tcPr><w:p><w:r><w:t>C</w:t></w:r></w:p></w:tc>" +
                             "</w:tr></w:tbl></w:body></w:document>");
            }

            using var testStream = new MemoryStream(ms.ToArray());
            using var reader = new Nedev.FileConverters.DocxToDoc.Format.DocxReader(testStream);

            var model = reader.ReadDocument();

            var table = Assert.IsType<Nedev.FileConverters.DocxToDoc.Model.TableModel>(Assert.Single(model.Content));
            var row = Assert.Single(table.Rows);
            Assert.Equal(2, row.Cells.Count);
            Assert.Equal(2, row.Cells[0].GridSpan);
            Assert.Equal(Nedev.FileConverters.DocxToDoc.Model.TableWidthUnit.Pct, row.Cells[0].WidthUnit);
            Assert.Equal(3500, row.Cells[0].Width);
            Assert.Equal(500, row.Cells[1].Width);
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
        public void ReadDocument_WithCellVerticalAlignBoth_MapsToCenter()
        {
            using var ms = new MemoryStream();
            using (var archive = new ZipArchive(ms, ZipArchiveMode.Create, true))
            {
                var entry = archive.CreateEntry("word/document.xml");
                using var entryStream = entry.Open();
                using var writer = new StreamWriter(entryStream);
                writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" +
                             "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:body>" +
                             "<w:tbl><w:tr>" +
                             "<w:tc><w:tcPr><w:vAlign w:val=\"both\"/></w:tcPr><w:p><w:r><w:t>A</w:t></w:r></w:p></w:tc>" +
                             "</w:tr></w:tbl></w:body></w:document>");
            }

            using var testStream = new MemoryStream(ms.ToArray());
            using var reader = new Nedev.FileConverters.DocxToDoc.Format.DocxReader(testStream);

            var model = reader.ReadDocument();

            var table = Assert.IsType<Nedev.FileConverters.DocxToDoc.Model.TableModel>(Assert.Single(model.Content));
            var cell = Assert.Single(Assert.Single(table.Rows).Cells);
            Assert.Equal(Nedev.FileConverters.DocxToDoc.Model.TableCellVerticalAlignment.Center, cell.VerticalAlignment);
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
        public void ReadDocument_WithCellBorderWithoutSize_UsesDefaultBorderWidth()
        {
            using var ms = new MemoryStream();
            using (var archive = new ZipArchive(ms, ZipArchiveMode.Create, true))
            {
                var entry = archive.CreateEntry("word/document.xml");
                using var entryStream = entry.Open();
                using var writer = new StreamWriter(entryStream);
                writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" +
                             "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:body>" +
                             "<w:tbl><w:tr>" +
                             "<w:tc><w:tcPr><w:tcBorders><w:left w:val=\"single\"/></w:tcBorders></w:tcPr><w:p><w:r><w:t>A</w:t></w:r></w:p></w:tc>" +
                             "</w:tr></w:tbl></w:body></w:document>");
            }

            using var testStream = new MemoryStream(ms.ToArray());
            using var reader = new Nedev.FileConverters.DocxToDoc.Format.DocxReader(testStream);

            var model = reader.ReadDocument();

            var table = Assert.IsType<Nedev.FileConverters.DocxToDoc.Model.TableModel>(Assert.Single(model.Content));
            var cell = Assert.Single(table.Rows[0].Cells);

            Assert.True(cell.HasLeftBorderOverride);
            Assert.Equal(Nedev.FileConverters.DocxToDoc.Model.BorderStyle.Single, cell.BorderLeftStyle);
            Assert.Equal(10, cell.BorderLeftTwips);
        }

        [Fact]
        public void ReadDocument_WithCellBorderThick_MapsToSingleBorderStyle()
        {
            using var ms = new MemoryStream();
            using (var archive = new ZipArchive(ms, ZipArchiveMode.Create, true))
            {
                var entry = archive.CreateEntry("word/document.xml");
                using var entryStream = entry.Open();
                using var writer = new StreamWriter(entryStream);
                writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" +
                             "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:body>" +
                             "<w:tbl><w:tr>" +
                             "<w:tc><w:tcPr><w:tcBorders><w:left w:val=\"thick\" w:sz=\"20\"/></w:tcBorders></w:tcPr><w:p><w:r><w:t>A</w:t></w:r></w:p></w:tc>" +
                             "</w:tr></w:tbl></w:body></w:document>");
            }

            using var testStream = new MemoryStream(ms.ToArray());
            using var reader = new Nedev.FileConverters.DocxToDoc.Format.DocxReader(testStream);

            var model = reader.ReadDocument();

            var table = Assert.IsType<Nedev.FileConverters.DocxToDoc.Model.TableModel>(Assert.Single(model.Content));
            var cell = Assert.Single(table.Rows[0].Cells);

            Assert.True(cell.HasLeftBorderOverride);
            Assert.Equal(Nedev.FileConverters.DocxToDoc.Model.BorderStyle.Single, cell.BorderLeftStyle);
            Assert.Equal(50, cell.BorderLeftTwips);
        }

        [Fact]
        public void ReadDocument_WithCellBorderDotDash_MapsToDashedBorderStyle()
        {
            using var ms = new MemoryStream();
            using (var archive = new ZipArchive(ms, ZipArchiveMode.Create, true))
            {
                var entry = archive.CreateEntry("word/document.xml");
                using var entryStream = entry.Open();
                using var writer = new StreamWriter(entryStream);
                writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" +
                             "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:body>" +
                             "<w:tbl><w:tr>" +
                             "<w:tc><w:tcPr><w:tcBorders><w:left w:val=\"dotDash\" w:sz=\"16\"/></w:tcBorders></w:tcPr><w:p><w:r><w:t>A</w:t></w:r></w:p></w:tc>" +
                             "</w:tr></w:tbl></w:body></w:document>");
            }

            using var testStream = new MemoryStream(ms.ToArray());
            using var reader = new Nedev.FileConverters.DocxToDoc.Format.DocxReader(testStream);

            var model = reader.ReadDocument();

            var table = Assert.IsType<Nedev.FileConverters.DocxToDoc.Model.TableModel>(Assert.Single(model.Content));
            var cell = Assert.Single(table.Rows[0].Cells);

            Assert.True(cell.HasLeftBorderOverride);
            Assert.Equal(Nedev.FileConverters.DocxToDoc.Model.BorderStyle.Dashed, cell.BorderLeftStyle);
            Assert.Equal(40, cell.BorderLeftTwips);
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
        public void ReadDocument_WithCellNilMargins_PreservesExplicitZeroPaddingOverrides()
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
                             "<w:tc><w:tcPr><w:tcMar><w:left w:type=\"nil\"/><w:right w:type=\"nil\"/><w:top w:type=\"nil\"/><w:bottom w:type=\"nil\"/></w:tcMar></w:tcPr><w:p><w:r><w:t>A</w:t></w:r></w:p></w:tc>" +
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
        public void ReadDocument_WithCellWidthNil_ParsesAsAutoWithoutNumericWidth()
        {
            using var ms = new MemoryStream();
            using (var archive = new ZipArchive(ms, ZipArchiveMode.Create, true))
            {
                var entry = archive.CreateEntry("word/document.xml");
                using var entryStream = entry.Open();
                using var writer = new StreamWriter(entryStream);
                writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" +
                             "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:body>" +
                             "<w:tbl><w:tr><w:tc><w:tcPr><w:tcW w:type=\"nil\" w:w=\"2500\"/></w:tcPr><w:p><w:r><w:t>A</w:t></w:r></w:p></w:tc></w:tr></w:tbl></w:body></w:document>");
            }

            using var testStream = new MemoryStream(ms.ToArray());
            using var reader = new Nedev.FileConverters.DocxToDoc.Format.DocxReader(testStream);

            var model = reader.ReadDocument();

            var table = Assert.IsType<Nedev.FileConverters.DocxToDoc.Model.TableModel>(Assert.Single(model.Content));
            var cell = Assert.Single(Assert.Single(table.Rows).Cells);
            Assert.Equal(Nedev.FileConverters.DocxToDoc.Model.TableWidthUnit.Auto, cell.WidthUnit);
            Assert.Equal(0, cell.Width);
        }

        [Fact]
        public void ReadDocument_WithPreferredTableWidthNil_ParsesAsAutoWithoutNumericWidth()
        {
            using var ms = new MemoryStream();
            using (var archive = new ZipArchive(ms, ZipArchiveMode.Create, true))
            {
                var entry = archive.CreateEntry("word/document.xml");
                using var entryStream = entry.Open();
                using var writer = new StreamWriter(entryStream);
                writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" +
                             "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:body>" +
                             "<w:tbl><w:tblPr><w:tblW w:w=\"5000\" w:type=\"nil\"/></w:tblPr>" +
                             "<w:tr><w:tc><w:p><w:r><w:t>A</w:t></w:r></w:p></w:tc></w:tr></w:tbl></w:body></w:document>");
            }

            using var testStream = new MemoryStream(ms.ToArray());
            using var reader = new Nedev.FileConverters.DocxToDoc.Format.DocxReader(testStream);

            var model = reader.ReadDocument();

            var table = Assert.IsType<Nedev.FileConverters.DocxToDoc.Model.TableModel>(Assert.Single(model.Content));
            Assert.Equal(Nedev.FileConverters.DocxToDoc.Model.TableWidthUnit.Auto, table.PreferredWidthUnit);
            Assert.Equal(0, table.PreferredWidthValue);
        }

        [Fact]
        public void ReadDocument_WithVerticalMerge_ParsesRestartContinueAndExplicitNone()
        {
            using var ms = new MemoryStream();
            using (var archive = new ZipArchive(ms, ZipArchiveMode.Create, true))
            {
                var entry = archive.CreateEntry("word/document.xml");
                using var entryStream = entry.Open();
                using var writer = new StreamWriter(entryStream);
                writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" +
                             "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:body>" +
                             "<w:tbl>" +
                             "<w:tr><w:tc><w:tcPr><w:vMerge w:val=\"restart\"/></w:tcPr><w:p><w:r><w:t>A1</w:t></w:r></w:p></w:tc></w:tr>" +
                             "<w:tr><w:tc><w:tcPr><w:vMerge/></w:tcPr><w:p><w:r><w:t>A2</w:t></w:r></w:p></w:tc></w:tr>" +
                             "<w:tr><w:tc><w:tcPr><w:vMerge w:val=\"false\"/></w:tcPr><w:p><w:r><w:t>A3</w:t></w:r></w:p></w:tc></w:tr>" +
                             "</w:tbl>" +
                             "</w:body></w:document>");
            }

            using var testStream = new MemoryStream(ms.ToArray());
            using var reader = new Nedev.FileConverters.DocxToDoc.Format.DocxReader(testStream);

            var model = reader.ReadDocument();

            var table = Assert.IsType<Nedev.FileConverters.DocxToDoc.Model.TableModel>(Assert.Single(model.Content));
            Assert.Equal(3, table.Rows.Count);
            Assert.Equal(Nedev.FileConverters.DocxToDoc.Model.TableCellVerticalMerge.Restart, table.Rows[0].Cells[0].VerticalMerge);
            Assert.Equal(Nedev.FileConverters.DocxToDoc.Model.TableCellVerticalMerge.Continue, table.Rows[1].Cells[0].VerticalMerge);
            Assert.Equal(Nedev.FileConverters.DocxToDoc.Model.TableCellVerticalMerge.None, table.Rows[2].Cells[0].VerticalMerge);
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

        [Fact]
        public void ReadDocument_WithRowHeaderAndCantSplit_ParsesRowFlags()
        {
            using var ms = new MemoryStream();
            using (var archive = new ZipArchive(ms, ZipArchiveMode.Create, true))
            {
                var entry = archive.CreateEntry("word/document.xml");
                using var entryStream = entry.Open();
                using var writer = new StreamWriter(entryStream);
                writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" +
                             "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:body>" +
                             "<w:tbl>" +
                             "<w:tr><w:trPr><w:tblHeader/><w:cantSplit/></w:trPr><w:tc><w:p><w:r><w:t>H</w:t></w:r></w:p></w:tc></w:tr>" +
                             "<w:tr><w:trPr><w:tblHeader w:val=\"0\"/><w:cantSplit w:val=\"false\"/></w:trPr><w:tc><w:p><w:r><w:t>N</w:t></w:r></w:p></w:tc></w:tr>" +
                             "</w:tbl></w:body></w:document>");
            }

            using var testStream = new MemoryStream(ms.ToArray());
            using var reader = new Nedev.FileConverters.DocxToDoc.Format.DocxReader(testStream);

            var model = reader.ReadDocument();

            var table = Assert.IsType<Nedev.FileConverters.DocxToDoc.Model.TableModel>(Assert.Single(model.Content));
            Assert.Equal(2, table.Rows.Count);
            Assert.True(table.Rows[0].IsHeader);
            Assert.True(table.Rows[0].CannotSplit);
            Assert.False(table.Rows[1].IsHeader);
            Assert.False(table.Rows[1].CannotSplit);
        }
    }
}
