using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Text;
using System.Threading.Tasks;
using Nedev.FileConverters.DocxToDoc.Format;
using OpenMcdf;
using Xunit;

namespace Nedev.FileConverters.DocxToDoc.Tests
{
    public class ConverterMetricsTests
    {
        [Fact]
        public void Convert_WithEmbeddedImage_LogsCorrectImageCount()
        {
            // Arrange
            var logger = new TestLogger();
            var converter = new DocxToDocConverter(logger);
            byte[] docxData = CreateDocxWithImage();
            using var inputStream = new MemoryStream(docxData);
            using var outputStream = new MemoryStream();

            // Act
            converter.Convert(inputStream, outputStream);

            // Assert
            Assert.True(outputStream.Length > 0);
            Assert.Contains(logger.InfoMessages, message => message == "Images: 1");
        }

        private static byte[] CreateDocxWithImage()
        {
            using var ms = new MemoryStream();
            using (var archive = new ZipArchive(ms, ZipArchiveMode.Create, true))
            {
                var contentTypesEntry = archive.CreateEntry("[Content_Types].xml");
                using (var stream = contentTypesEntry.Open())
                using (var writer = new StreamWriter(stream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                        "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">" +
                        "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>" +
                        "<Default Extension=\"xml\" ContentType=\"application/xml\"/>" +
                        "<Default Extension=\"png\" ContentType=\"image/png\"/>" +
                        "<Override PartName=\"/word/document.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml\"/>" +
                        "</Types>");
                }

                var relsEntry = archive.CreateEntry("_rels/.rels");
                using (var stream = relsEntry.Open())
                using (var writer = new StreamWriter(stream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                        "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
                        "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"word/document.xml\"/>" +
                        "</Relationships>");
                }

                var docRelsEntry = archive.CreateEntry("word/_rels/document.xml.rels");
                using (var stream = docRelsEntry.Open())
                using (var writer = new StreamWriter(stream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                        "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
                        "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/image\" Target=\"media/image1.png\"/>" +
                        "</Relationships>");
                }

                var imageEntry = archive.CreateEntry("word/media/image1.png");
                using (var stream = imageEntry.Open())
                {
                    byte[] pngData = new byte[]
                    {
                        0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A,
                        0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44, 0x52,
                        0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01,
                        0x08, 0x02, 0x00, 0x00, 0x00, 0x90, 0x77, 0x53,
                        0xDE, 0x00, 0x00, 0x00, 0x0C, 0x49, 0x44, 0x41,
                        0x54, 0x08, 0xD7, 0x63, 0xF8, 0x0F, 0x00, 0x00,
                        0x01, 0x01, 0x00, 0x05, 0x18, 0xD8, 0x4D, 0x00,
                        0x00, 0x00, 0x00, 0x49, 0x45, 0x4E, 0x44, 0xAE,
                        0x42, 0x60, 0x82
                    };
                    stream.Write(pngData, 0, pngData.Length);
                }

                var docEntry = archive.CreateEntry("word/document.xml");
                using (var stream = docEntry.Open())
                using (var writer = new StreamWriter(stream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                        "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" " +
                        "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" " +
                        "xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" " +
                        "xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\">" +
                        "<w:body><w:p>" +
                        "<w:r><w:t>Before image</w:t></w:r>" +
                        "<w:r><w:drawing><wp:inline><wp:extent cx=\"914400\" cy=\"914400\"/>" +
                        "<a:graphic><a:graphicData><pic:pic xmlns:pic=\"http://schemas.openxmlformats.org/drawingml/2006/picture\">" +
                        "<pic:blipFill><a:blip r:embed=\"rId1\"/></pic:blipFill>" +
                        "</pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>" +
                        "<w:r><w:t>After image</w:t></w:r>" +
                        "</w:p></w:body></w:document>");
                }
            }

            return ms.ToArray();
        }

        [Theory]
        [InlineData("afchunk-header.docx")]
        [InlineData("afchunk-footer.docx")]
        [InlineData("source.docx")]
        public void Convert_WithRepositorySampleDocx_ProducesValidCompoundFile(string fileName)
        {
            var converter = new DocxToDocConverter();
            string repositoryRoot = ResolveRepositoryRoot();
            string inputPath = Path.Combine(repositoryRoot, "samples", "generated-docx", fileName);
            Assert.True(File.Exists(inputPath));

            string tempDir = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(tempDir);
            string outputPath = Path.Combine(tempDir, Path.ChangeExtension(fileName, ".doc"));

            try
            {
                converter.Convert(inputPath, outputPath);

                Assert.True(File.Exists(outputPath));
                Assert.True(new FileInfo(outputPath).Length > 1024);

                using var fs = File.OpenRead(outputPath);
                using var compoundFile = new CompoundFile(fs);
                Assert.True(compoundFile.RootStorage.TryGetStream("WordDocument", out var wordDocumentStream));
                Assert.True(compoundFile.RootStorage.TryGetStream("1Table", out var tableStream));

                byte[] wordDocumentData = wordDocumentStream.GetData();
                byte[] tableData = tableStream.GetData();

                Assert.True(wordDocumentData.Length > 1536);
                Assert.True(tableData.Length > 0);
                Assert.True(BitConverter.ToInt32(wordDocumentData, 64) > 0);
            }
            finally
            {
                try
                {
                    Directory.Delete(tempDir, true);
                }
                catch
                {
                }
            }
        }

        [Fact]
        public void Convert_WithRepositorySampleContainingTable_WritesTapx()
        {
            var converter = new DocxToDocConverter();
            byte[] docx = CreateDocxWithTable();
            using var input = new MemoryStream(docx);
            string outputPath = CreateTempOutputPath("table");
            try
            {
                using (var output = File.Create(outputPath))
                {
                    converter.Convert(input, output);
                }
                using var docStream = File.OpenRead(outputPath);
                using var compoundFile = new CompoundFile(docStream);
                Assert.True(compoundFile.RootStorage.TryGetStream("WordDocument", out var wordDocumentStream));

                byte[] wordDocumentData = wordDocumentStream.GetData();
                int lcbPlcfbteTapx = BitConverter.ToInt32(wordDocumentData, 154 + (Fib.TapxPairIndex * 8) + 4);
                Assert.True(lcbPlcfbteTapx > 0);
            }
            finally
            {
                TryDeleteFile(outputPath);
            }
        }

        [Fact]
        public void Convert_WithRepositorySampleContainingHeaderFooter_WritesHeaderStory()
        {
            var converter = new DocxToDocConverter();
            byte[] docx = CreateDocxWithDefaultHeader();
            using var input = new MemoryStream(docx);
            string outputPath = CreateTempOutputPath("header-footer");
            try
            {
                using (var output = File.Create(outputPath))
                {
                    converter.Convert(input, output);
                }
                using var docStream = File.OpenRead(outputPath);
                using var compoundFile = new CompoundFile(docStream);
                Assert.True(compoundFile.RootStorage.TryGetStream("WordDocument", out var wordDocumentStream));

                byte[] wordDocumentData = wordDocumentStream.GetData();
                int ccpHdd = BitConverter.ToInt32(wordDocumentData, 72);
                int lcbPlcfHdd = BitConverter.ToInt32(wordDocumentData, 154 + (Fib.HeaderStoryPairIndex * 8) + 4);
                Assert.True(ccpHdd > 0);
                Assert.True(lcbPlcfHdd > 0);
            }
            finally
            {
                TryDeleteFile(outputPath);
            }
        }

        [Fact]
        public void Convert_WithCommentReplyDocx_WritesCommentStory()
        {
            var converter = new DocxToDocConverter();
            byte[] docx = CreateDocxWithCommentReplyRange();
            using var input = new MemoryStream(docx);
            using var output = new MemoryStream();

            converter.Convert(input, output);
            output.Position = 0;

            using var compoundFile = new CompoundFile(output);
            Assert.True(compoundFile.RootStorage.TryGetStream("WordDocument", out var wordDocumentStream));

            byte[] wordDocumentData = wordDocumentStream.GetData();
            int ccpAtn = BitConverter.ToInt32(wordDocumentData, 76);
            int lcbPlcfandRef = BitConverter.ToInt32(wordDocumentData, 154 + (Fib.CommentReferencePairIndex * 8) + 4);
            int lcbPlcfandTxt = BitConverter.ToInt32(wordDocumentData, 154 + (Fib.CommentTextPairIndex * 8) + 4);

            Assert.True(ccpAtn > 0);
            Assert.True(lcbPlcfandRef > 0);
            Assert.True(lcbPlcfandTxt > 0);
            string allStoryText = ExtractAllStoryText(wordDocumentData);
            Assert.Contains("Root comment", allStoryText);
            Assert.Contains("Reply comment", allStoryText);
        }

        [Fact]
        public void Convert_WithAltChunkAndTable_PreservesVisibleOrderInMainStory()
        {
            var converter = new DocxToDocConverter();
            byte[] docx = CreateDocxWithAltChunkAndTable();
            using var input = new MemoryStream(docx);
            using var output = new MemoryStream();

            converter.Convert(input, output);
            output.Position = 0;

            using var compoundFile = new CompoundFile(output);
            Assert.True(compoundFile.RootStorage.TryGetStream("WordDocument", out var wordDocumentStream));
            byte[] wordDocumentData = wordDocumentStream.GetData();
            string mainStoryText = ExtractMainStoryText(wordDocumentData);

            int beforeIndex = mainStoryText.IndexOf("Before", StringComparison.Ordinal);
            int chunkLine1Index = mainStoryText.IndexOf("Chunk line 1", StringComparison.Ordinal);
            int chunkLine2Index = mainStoryText.IndexOf("Chunk line 2", StringComparison.Ordinal);
            int cellIndex = mainStoryText.IndexOf("Cell", StringComparison.Ordinal);
            int afterIndex = mainStoryText.IndexOf("After", StringComparison.Ordinal);

            Assert.True(beforeIndex >= 0);
            Assert.True(chunkLine1Index > beforeIndex);
            Assert.True(chunkLine2Index > chunkLine1Index);
            Assert.True(cellIndex > chunkLine2Index);
            Assert.True(afterIndex > cellIndex);
        }

        [Fact]
        public void Convert_WithHeaderTableAndAltChunk_WritesHeaderStoryText()
        {
            var converter = new DocxToDocConverter();
            byte[] docx = CreateDocxWithHeaderTableAndAltChunk();
            using var input = new MemoryStream(docx);
            using var output = new MemoryStream();

            converter.Convert(input, output);
            output.Position = 0;

            using var compoundFile = new CompoundFile(output);
            Assert.True(compoundFile.RootStorage.TryGetStream("WordDocument", out var wordDocumentStream));
            byte[] wordDocumentData = wordDocumentStream.GetData();

            int ccpHdd = BitConverter.ToInt32(wordDocumentData, 72);
            Assert.True(ccpHdd > 0);

            string allStoryText = ExtractAllStoryText(wordDocumentData);
            Assert.Contains("HeaderLead", allStoryText);
            Assert.Contains("HeaderChunk", allStoryText);
            Assert.Contains("HeaderCell", allStoryText);
        }

        [Fact]
        public void Convert_WithMissingAltChunkRelationship_PreservesSurroundingText()
        {
            var converter = new DocxToDocConverter();
            byte[] docx = CreateDocxWithMissingAltChunkRelationship();
            using var input = new MemoryStream(docx);
            using var output = new MemoryStream();

            converter.Convert(input, output);
            output.Position = 0;

            using var compoundFile = new CompoundFile(output);
            Assert.True(compoundFile.RootStorage.TryGetStream("WordDocument", out var wordDocumentStream));
            string mainStoryText = ExtractMainStoryText(wordDocumentStream.GetData());

            int beforeIndex = mainStoryText.IndexOf("Before", StringComparison.Ordinal);
            int afterIndex = mainStoryText.IndexOf("After", StringComparison.Ordinal);

            Assert.True(beforeIndex >= 0);
            Assert.True(afterIndex > beforeIndex);
            Assert.DoesNotContain("Chunk", mainStoryText);
        }

        [Fact]
        public void Convert_WithMissingHeaderRelationship_DoesNotWriteHeaderStory()
        {
            var converter = new DocxToDocConverter();
            byte[] docx = CreateDocxWithMissingHeaderRelationship();
            using var input = new MemoryStream(docx);
            using var output = new MemoryStream();

            converter.Convert(input, output);
            output.Position = 0;

            using var compoundFile = new CompoundFile(output);
            Assert.True(compoundFile.RootStorage.TryGetStream("WordDocument", out var wordDocumentStream));
            byte[] wordDocumentData = wordDocumentStream.GetData();
            int ccpHdd = BitConverter.ToInt32(wordDocumentData, 72);
            int lcbPlcfHdd = BitConverter.ToInt32(wordDocumentData, 154 + (Fib.HeaderStoryPairIndex * 8) + 4);

            Assert.Equal(0, ccpHdd);
            Assert.Equal(0, lcbPlcfHdd);
        }

        [Fact]
        public void Convert_WithMalformedHeaderXml_ThrowsConversionExceptionWithParsingStage()
        {
            var converter = new DocxToDocConverter();
            byte[] docx = CreateDocxWithMalformedHeaderXml();
            using var input = new MemoryStream(docx);
            using var output = new MemoryStream();

            var ex = Assert.Throws<ConversionException>(() => converter.Convert(input, output));
            Assert.Equal(ConversionStage.Parsing, ex.Stage);
        }

        [Fact]
        public void Convert_WithFootnoteReferenceButMissingFootnotesPart_DoesNotWriteFootnoteStory()
        {
            var converter = new DocxToDocConverter();
            byte[] docx = CreateDocxWithFootnoteReferenceWithoutFootnotesPart();
            using var input = new MemoryStream(docx);
            using var output = new MemoryStream();

            converter.Convert(input, output);
            output.Position = 0;

            using var compoundFile = new CompoundFile(output);
            Assert.True(compoundFile.RootStorage.TryGetStream("WordDocument", out var wordDocumentStream));
            byte[] wordDocumentData = wordDocumentStream.GetData();

            int ccpFtn = BitConverter.ToInt32(wordDocumentData, 68);
            int lcbPlcffndRef = BitConverter.ToInt32(wordDocumentData, 154 + (Fib.FootnoteReferencePairIndex * 8) + 4);
            int lcbPlcffndTxt = BitConverter.ToInt32(wordDocumentData, 154 + (Fib.FootnoteTextPairIndex * 8) + 4);
            Assert.Equal(0, ccpFtn);
            Assert.Equal(0, lcbPlcffndRef);
            Assert.Equal(0, lcbPlcffndTxt);
        }

        [Fact]
        public void Convert_WithValidFootnotesPart_WritesFootnoteStory()
        {
            var converter = new DocxToDocConverter();
            byte[] docx = CreateDocxWithValidFootnotesPart();
            using var input = new MemoryStream(docx);
            using var output = new MemoryStream();

            converter.Convert(input, output);
            output.Position = 0;

            using var compoundFile = new CompoundFile(output);
            Assert.True(compoundFile.RootStorage.TryGetStream("WordDocument", out var wordDocumentStream));
            byte[] wordDocumentData = wordDocumentStream.GetData();

            int ccpFtn = BitConverter.ToInt32(wordDocumentData, 68);
            Assert.True(ccpFtn > 0);
            string allStoryText = ExtractAllStoryText(wordDocumentData);
            Assert.Contains("Footnote text", allStoryText);
        }

        [Fact]
        public void Convert_WithMalformedFootnotesXml_ThrowsConversionExceptionWithParsingStage()
        {
            var converter = new DocxToDocConverter();
            byte[] docx = CreateDocxWithMalformedFootnotesXml();
            using var input = new MemoryStream(docx);
            using var output = new MemoryStream();

            var ex = Assert.Throws<ConversionException>(() => converter.Convert(input, output));
            Assert.Equal(ConversionStage.Parsing, ex.Stage);
        }

        [Fact]
        public void Convert_WithValidEndnotesPart_WritesEndnoteStory()
        {
            var converter = new DocxToDocConverter();
            byte[] docx = CreateDocxWithValidEndnotesPart();
            using var input = new MemoryStream(docx);
            using var output = new MemoryStream();

            converter.Convert(input, output);
            output.Position = 0;

            using var compoundFile = new CompoundFile(output);
            Assert.True(compoundFile.RootStorage.TryGetStream("WordDocument", out var wordDocumentStream));
            byte[] wordDocumentData = wordDocumentStream.GetData();

            int ccpEdn = BitConverter.ToInt32(wordDocumentData, 80);
            Assert.True(ccpEdn > 0);
            string allStoryText = ExtractAllStoryText(wordDocumentData);
            Assert.Contains("Endnote text", allStoryText);
        }

        [Fact]
        public void Convert_WithMalformedEndnotesXml_ThrowsConversionExceptionWithParsingStage()
        {
            var converter = new DocxToDocConverter();
            byte[] docx = CreateDocxWithMalformedEndnotesXml();
            using var input = new MemoryStream(docx);
            using var output = new MemoryStream();

            var ex = Assert.Throws<ConversionException>(() => converter.Convert(input, output));
            Assert.Equal(ConversionStage.Parsing, ex.Stage);
        }

        [Fact]
        public void Convert_WithMalformedCommentsExtendedAndMalformedHeader_ThrowsConversionExceptionWithParsingStage()
        {
            var converter = new DocxToDocConverter();
            byte[] docx = CreateDocxWithMalformedCommentsExtendedAndMalformedHeader();
            using var input = new MemoryStream(docx);
            using var output = new MemoryStream();

            var ex = Assert.Throws<ConversionException>(() => converter.Convert(input, output));
            Assert.Equal(ConversionStage.Parsing, ex.Stage);
        }

        [Fact]
        public void Convert_WithValidCommentsButMalformedHeader_ThrowsConversionExceptionWithParsingStage()
        {
            var converter = new DocxToDocConverter();
            byte[] docx = CreateDocxWithValidCommentsAndMalformedHeader();
            using var input = new MemoryStream(docx);
            using var output = new MemoryStream();

            var ex = Assert.Throws<ConversionException>(() => converter.Convert(input, output));
            Assert.Equal(ConversionStage.Parsing, ex.Stage);
        }

        [Fact]
        public void Convert_WithComplexValidDocxAndWriteFailure_ThrowsWritingStage()
        {
            var converter = new DocxToDocConverter();
            byte[] docx = CreateDocxWithHeaderTableAndAltChunk();
            using var input = new MemoryStream(docx);
            using var output = new ThrowOnWriteStream();

            var ex = Assert.Throws<ConversionException>(() => converter.Convert(input, output));
            Assert.Equal(ConversionStage.Writing, ex.Stage);
        }

        [Fact]
        public void Convert_WithComplexValidDocxAndFlushFailure_ThrowsWritingStage()
        {
            var converter = new DocxToDocConverter();
            byte[] docx = CreateDocxWithHeaderTableAndAltChunk();
            using var input = new MemoryStream(docx);
            using var output = new ThrowOnFlushStream();

            var ex = Assert.Throws<ConversionException>(() => converter.Convert(input, output));
            Assert.Equal(ConversionStage.Writing, ex.Stage);
        }

        [Fact]
        public void Convert_WithMalformedDocxAndWriteFailure_StillThrowsParsingStage()
        {
            var converter = new DocxToDocConverter();
            byte[] docx = CreateDocxWithMalformedHeaderXml();
            using var input = new MemoryStream(docx);
            using var output = new ThrowOnWriteStream();

            var ex = Assert.Throws<ConversionException>(() => converter.Convert(input, output));
            Assert.Equal(ConversionStage.Parsing, ex.Stage);
        }

        [Fact]
        public async Task ConvertAsync_WithComplexValidDocxAndWriteFailure_ThrowsWritingStage()
        {
            var converter = new DocxToDocConverter();
            byte[] docx = CreateDocxWithHeaderTableAndAltChunk();
            using var input = new MemoryStream(docx);
            using var output = new ThrowOnWriteStream();

            var ex = await Assert.ThrowsAsync<ConversionException>(() => converter.ConvertAsync(input, output));
            Assert.Equal(ConversionStage.Writing, ex.Stage);
        }

        [Fact]
        public async Task ConvertAsync_WithComplexValidDocxAndFlushFailure_ThrowsWritingStage()
        {
            var converter = new DocxToDocConverter();
            byte[] docx = CreateDocxWithHeaderTableAndAltChunk();
            using var input = new MemoryStream(docx);
            using var output = new ThrowOnFlushStream();

            var ex = await Assert.ThrowsAsync<ConversionException>(() => converter.ConvertAsync(input, output));
            Assert.Equal(ConversionStage.Writing, ex.Stage);
        }

        [Fact]
        public async Task ConvertAsync_WithMalformedDocxAndWriteFailure_StillThrowsParsingStage()
        {
            var converter = new DocxToDocConverter();
            byte[] docx = CreateDocxWithMalformedHeaderXml();
            using var input = new MemoryStream(docx);
            using var output = new ThrowOnWriteStream();

            var ex = await Assert.ThrowsAsync<ConversionException>(() => converter.ConvertAsync(input, output));
            Assert.Equal(ConversionStage.Parsing, ex.Stage);
        }

        [Fact]
        public void Convert_WithMissingCommentsExtended_PreservesCommentEmission()
        {
            var converter = new DocxToDocConverter();
            byte[] docx = CreateDocxWithCommentsWithoutExtended();
            using var input = new MemoryStream(docx);
            using var output = new MemoryStream();

            converter.Convert(input, output);
            output.Position = 0;

            using var compoundFile = new CompoundFile(output);
            Assert.True(compoundFile.RootStorage.TryGetStream("WordDocument", out var wordDocumentStream));
            byte[] wordDocumentData = wordDocumentStream.GetData();

            int ccpAtn = BitConverter.ToInt32(wordDocumentData, 76);
            Assert.True(ccpAtn > 0);

            string allStoryText = ExtractAllStoryText(wordDocumentData);
            Assert.Contains("Comment without extended", allStoryText);
        }

        [Fact]
        public void Convert_WithMalformedCommentsExtended_ThrowsConversionExceptionWithParsingStage()
        {
            var converter = new DocxToDocConverter();
            byte[] docx = CreateDocxWithMalformedCommentsExtended();
            using var input = new MemoryStream(docx);
            using var output = new MemoryStream();

            var ex = Assert.Throws<ConversionException>(() => converter.Convert(input, output));
            Assert.Equal(ConversionStage.Parsing, ex.Stage);
        }

        [Fact]
        public void Convert_WithHeaderAltChunkMissingHeaderRels_PreservesSurroundingHeaderText()
        {
            var converter = new DocxToDocConverter();
            byte[] docx = CreateDocxWithHeaderAltChunkMissingHeaderRels();
            using var input = new MemoryStream(docx);
            using var output = new MemoryStream();

            converter.Convert(input, output);
            output.Position = 0;

            using var compoundFile = new CompoundFile(output);
            Assert.True(compoundFile.RootStorage.TryGetStream("WordDocument", out var wordDocumentStream));
            string allStoryText = ExtractAllStoryText(wordDocumentStream.GetData());

            Assert.Contains("HeaderBefore", allStoryText);
            Assert.Contains("HeaderAfter", allStoryText);
            Assert.DoesNotContain("HeaderChunk", allStoryText);
        }

        [Fact]
        public void Convert_WithMalformedEmbeddedDocxAltChunk_PreservesSurroundingMainStoryText()
        {
            var converter = new DocxToDocConverter();
            byte[] docx = CreateDocxWithMalformedEmbeddedDocxAltChunk();
            using var input = new MemoryStream(docx);
            using var output = new MemoryStream();

            converter.Convert(input, output);
            output.Position = 0;

            using var compoundFile = new CompoundFile(output);
            Assert.True(compoundFile.RootStorage.TryGetStream("WordDocument", out var wordDocumentStream));
            string mainStoryText = ExtractMainStoryText(wordDocumentStream.GetData());

            int beforeIndex = mainStoryText.IndexOf("Before", StringComparison.Ordinal);
            int afterIndex = mainStoryText.IndexOf("After", StringComparison.Ordinal);
            Assert.True(beforeIndex >= 0);
            Assert.True(afterIndex > beforeIndex);
            Assert.DoesNotContain("Inner", mainStoryText);
        }

        [Fact]
        public void Convert_WithPageBreakBefore_EmitsPageBreakBetweenParagraphs()
        {
            var converter = new DocxToDocConverter();
            byte[] docx = CreateDocxWithPageBreakBeforeParagraph();
            using var input = new MemoryStream(docx);
            using var output = new MemoryStream();

            converter.Convert(input, output);
            output.Position = 0;

            using var compoundFile = new CompoundFile(output);
            Assert.True(compoundFile.RootStorage.TryGetStream("WordDocument", out var wordDocumentStream));
            string mainStoryText = ExtractMainStoryText(wordDocumentStream.GetData());

            int firstIndex = mainStoryText.IndexOf("First", StringComparison.Ordinal);
            int secondIndex = mainStoryText.IndexOf("Second", StringComparison.Ordinal);
            int pageBreakIndex = mainStoryText.IndexOf('\f');

            Assert.True(firstIndex >= 0);
            Assert.True(pageBreakIndex > firstIndex);
            Assert.True(secondIndex > pageBreakIndex);
        }

        [Fact]
        public void Convert_WithPageBreakBeforeDisabled_DoesNotEmitPageBreakBetweenParagraphs()
        {
            var converter = new DocxToDocConverter();
            byte[] docx = CreateDocxWithPageBreakBeforeDisabledParagraph();
            using var input = new MemoryStream(docx);
            using var output = new MemoryStream();

            converter.Convert(input, output);
            output.Position = 0;

            using var compoundFile = new CompoundFile(output);
            Assert.True(compoundFile.RootStorage.TryGetStream("WordDocument", out var wordDocumentStream));
            string mainStoryText = ExtractMainStoryText(wordDocumentStream.GetData());

            int firstIndex = mainStoryText.IndexOf("First", StringComparison.Ordinal);
            int secondIndex = mainStoryText.IndexOf("Second", StringComparison.Ordinal);
            int pageBreakIndex = mainStoryText.IndexOf('\f');

            Assert.True(firstIndex >= 0);
            Assert.True(secondIndex > firstIndex);
            Assert.Equal(-1, pageBreakIndex);
        }

        private static byte[] CreateDocxWithCommentReplyRange()
        {
            using var ms = new MemoryStream();
            using (var archive = new ZipArchive(ms, ZipArchiveMode.Create, true))
            {
                var commentsEntry = archive.CreateEntry("word/comments.xml");
                using (var stream = commentsEntry.Open())
                using (var writer = new StreamWriter(stream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                        "<w:comments xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\">" +
                        "<w:comment w:id=\"0\" w:author=\"Root\" w:initials=\"RT\"><w:p w14:paraId=\"11111111\"><w:r><w:t>Root comment</w:t></w:r></w:p></w:comment>" +
                        "<w:comment w:id=\"1\" w:author=\"Reply\" w:initials=\"RP\"><w:p w14:paraId=\"22222222\"><w:r><w:t>Reply comment</w:t></w:r></w:p></w:comment>" +
                        "</w:comments>");
                }

                var commentsExtendedEntry = archive.CreateEntry("word/commentsExtended.xml");
                using (var stream = commentsExtendedEntry.Open())
                using (var writer = new StreamWriter(stream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                        "<w15:commentsEx xmlns:w15=\"http://schemas.microsoft.com/office/word/2012/wordml\">" +
                        "<w15:commentEx w15:paraId=\"11111111\"/>" +
                        "<w15:commentEx w15:paraId=\"22222222\" w15:paraIdParent=\"11111111\"/>" +
                        "</w15:commentsEx>");
                }

                var documentEntry = archive.CreateEntry("word/document.xml");
                using (var stream = documentEntry.Open())
                using (var writer = new StreamWriter(stream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                        "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">" +
                        "<w:body><w:p><w:r><w:t>A</w:t></w:r><w:commentRangeStart w:id=\"0\"/><w:r><w:t>B</w:t></w:r><w:commentRangeEnd w:id=\"0\"/><w:r><w:commentReference w:id=\"0\"/></w:r><w:r><w:t>C</w:t></w:r></w:p></w:body>" +
                        "</w:document>");
                }
            }

            return ms.ToArray();
        }

        private static byte[] CreateDocxWithTable()
        {
            using var ms = new MemoryStream();
            using (var archive = new ZipArchive(ms, ZipArchiveMode.Create, true))
            {
                var documentEntry = archive.CreateEntry("word/document.xml");
                using var stream = documentEntry.Open();
                using var writer = new StreamWriter(stream);
                writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                    "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">" +
                    "<w:body>" +
                    "<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w=\"2400\"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w=\"2400\" w:type=\"dxa\"/></w:tcPr><w:p><w:r><w:t>Cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl>" +
                    "<w:p><w:r><w:t>After</w:t></w:r></w:p>" +
                    "</w:body>" +
                    "</w:document>");
            }

            return ms.ToArray();
        }

        private static byte[] CreateDocxWithDefaultHeader()
        {
            using var ms = new MemoryStream();
            using (var archive = new ZipArchive(ms, ZipArchiveMode.Create, true))
            {
                var relsEntry = archive.CreateEntry("word/_rels/document.xml.rels");
                using (var stream = relsEntry.Open())
                using (var writer = new StreamWriter(stream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                        "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
                        "<Relationship Id=\"rIdHeader1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/header\" Target=\"header1.xml\"/>" +
                        "</Relationships>");
                }

                var headerEntry = archive.CreateEntry("word/header1.xml");
                using (var stream = headerEntry.Open())
                using (var writer = new StreamWriter(stream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                        "<w:hdr xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">" +
                        "<w:p><w:r><w:t>Header Text</w:t></w:r></w:p>" +
                        "</w:hdr>");
                }

                var documentEntry = archive.CreateEntry("word/document.xml");
                using (var stream = documentEntry.Open())
                using (var writer = new StreamWriter(stream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                        "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">" +
                        "<w:body>" +
                        "<w:p><w:r><w:t>Body</w:t></w:r></w:p>" +
                        "<w:sectPr><w:headerReference w:type=\"default\" r:id=\"rIdHeader1\"/></w:sectPr>" +
                        "</w:body>" +
                        "</w:document>");
                }
            }

            return ms.ToArray();
        }

        private static byte[] CreateDocxWithAltChunkAndTable()
        {
            using var ms = new MemoryStream();
            using (var archive = new ZipArchive(ms, ZipArchiveMode.Create, true))
            {
                var relsEntry = archive.CreateEntry("word/_rels/document.xml.rels");
                using (var relsStream = relsEntry.Open())
                using (var writer = new StreamWriter(relsStream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                                 "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
                                 "<Relationship Id=\"rIdChunk\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/aFChunk\" Target=\"afchunk-main.txt\"/>" +
                                 "</Relationships>");
                }

                var chunkEntry = archive.CreateEntry("word/afchunk-main.txt");
                using (var chunkStream = chunkEntry.Open())
                using (var writer = new StreamWriter(chunkStream))
                {
                    writer.Write("Chunk line 1\r\nChunk line 2");
                }

                var documentEntry = archive.CreateEntry("word/document.xml");
                using (var stream = documentEntry.Open())
                using (var writer = new StreamWriter(stream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                        "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">" +
                        "<w:body>" +
                        "<w:p><w:r><w:t>Before</w:t></w:r></w:p>" +
                        "<w:altChunk r:id=\"rIdChunk\"/>" +
                        "<w:tbl><w:tblGrid><w:gridCol w:w=\"2400\"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w=\"2400\" w:type=\"dxa\"/></w:tcPr><w:p><w:r><w:t>Cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl>" +
                        "<w:p><w:r><w:t>After</w:t></w:r></w:p>" +
                        "</w:body>" +
                        "</w:document>");
                }
            }

            return ms.ToArray();
        }

        private static byte[] CreateDocxWithHeaderTableAndAltChunk()
        {
            using var ms = new MemoryStream();
            using (var archive = new ZipArchive(ms, ZipArchiveMode.Create, true))
            {
                var relsEntry = archive.CreateEntry("word/_rels/document.xml.rels");
                using (var relsStream = relsEntry.Open())
                using (var writer = new StreamWriter(relsStream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                        "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
                        "<Relationship Id=\"rIdHeader1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/header\" Target=\"header1.xml\"/>" +
                        "</Relationships>");
                }

                var headerRelsEntry = archive.CreateEntry("word/_rels/header1.xml.rels");
                using (var relsStream = headerRelsEntry.Open())
                using (var writer = new StreamWriter(relsStream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                        "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
                        "<Relationship Id=\"rIdHeaderChunk\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/aFChunk\" Target=\"afchunk-header.txt\"/>" +
                        "</Relationships>");
                }

                var headerChunkEntry = archive.CreateEntry("word/afchunk-header.txt");
                using (var chunkStream = headerChunkEntry.Open())
                using (var writer = new StreamWriter(chunkStream))
                {
                    writer.Write("HeaderChunk");
                }

                var headerEntry = archive.CreateEntry("word/header1.xml");
                using (var stream = headerEntry.Open())
                using (var writer = new StreamWriter(stream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                        "<w:hdr xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">" +
                        "<w:p><w:r><w:t>HeaderLead</w:t></w:r></w:p>" +
                        "<w:altChunk r:id=\"rIdHeaderChunk\"/>" +
                        "<w:tbl><w:tblGrid><w:gridCol w:w=\"2400\"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w=\"2400\" w:type=\"dxa\"/></w:tcPr><w:p><w:r><w:t>HeaderCell</w:t></w:r></w:p></w:tc></w:tr></w:tbl>" +
                        "</w:hdr>");
                }

                var documentEntry = archive.CreateEntry("word/document.xml");
                using (var stream = documentEntry.Open())
                using (var writer = new StreamWriter(stream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                        "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">" +
                        "<w:body>" +
                        "<w:p><w:r><w:t>Body</w:t></w:r></w:p>" +
                        "<w:sectPr><w:headerReference w:type=\"default\" r:id=\"rIdHeader1\"/></w:sectPr>" +
                        "</w:body>" +
                        "</w:document>");
                }
            }

            return ms.ToArray();
        }

        private static byte[] CreateDocxWithMissingAltChunkRelationship()
        {
            using var ms = new MemoryStream();
            using (var archive = new ZipArchive(ms, ZipArchiveMode.Create, true))
            {
                var documentEntry = archive.CreateEntry("word/document.xml");
                using var stream = documentEntry.Open();
                using var writer = new StreamWriter(stream);
                writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                    "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">" +
                    "<w:body>" +
                    "<w:p><w:r><w:t>Before</w:t></w:r></w:p>" +
                    "<w:altChunk r:id=\"rIdMissing\"/>" +
                    "<w:p><w:r><w:t>After</w:t></w:r></w:p>" +
                    "</w:body>" +
                    "</w:document>");
            }

            return ms.ToArray();
        }

        private static byte[] CreateDocxWithMissingHeaderRelationship()
        {
            using var ms = new MemoryStream();
            using (var archive = new ZipArchive(ms, ZipArchiveMode.Create, true))
            {
                var documentEntry = archive.CreateEntry("word/document.xml");
                using var stream = documentEntry.Open();
                using var writer = new StreamWriter(stream);
                writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                    "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">" +
                    "<w:body>" +
                    "<w:p><w:r><w:t>Body</w:t></w:r></w:p>" +
                    "<w:sectPr><w:headerReference w:type=\"default\" r:id=\"rIdMissingHeader\"/></w:sectPr>" +
                    "</w:body>" +
                    "</w:document>");
            }

            return ms.ToArray();
        }

        private static byte[] CreateDocxWithMalformedHeaderXml()
        {
            using var ms = new MemoryStream();
            using (var archive = new ZipArchive(ms, ZipArchiveMode.Create, true))
            {
                var relsEntry = archive.CreateEntry("word/_rels/document.xml.rels");
                using (var relsStream = relsEntry.Open())
                using (var writer = new StreamWriter(relsStream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                        "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
                        "<Relationship Id=\"rIdHeader1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/header\" Target=\"header1.xml\"/>" +
                        "</Relationships>");
                }

                var headerEntry = archive.CreateEntry("word/header1.xml");
                using (var stream = headerEntry.Open())
                using (var writer = new StreamWriter(stream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><w:hdr xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:p>");
                }

                var documentEntry = archive.CreateEntry("word/document.xml");
                using (var stream = documentEntry.Open())
                using (var writer = new StreamWriter(stream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                        "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">" +
                        "<w:body>" +
                        "<w:p><w:r><w:t>Body</w:t></w:r></w:p>" +
                        "<w:sectPr><w:headerReference w:type=\"default\" r:id=\"rIdHeader1\"/></w:sectPr>" +
                        "</w:body>" +
                        "</w:document>");
                }
            }

            return ms.ToArray();
        }

        private static byte[] CreateDocxWithCommentsWithoutExtended()
        {
            using var ms = new MemoryStream();
            using (var archive = new ZipArchive(ms, ZipArchiveMode.Create, true))
            {
                var commentsEntry = archive.CreateEntry("word/comments.xml");
                using (var stream = commentsEntry.Open())
                using (var writer = new StreamWriter(stream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                        "<w:comments xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">" +
                        "<w:comment w:id=\"0\" w:author=\"A\"><w:p><w:r><w:t>Comment without extended</w:t></w:r></w:p></w:comment>" +
                        "</w:comments>");
                }

                var documentEntry = archive.CreateEntry("word/document.xml");
                using (var stream = documentEntry.Open())
                using (var writer = new StreamWriter(stream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                        "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">" +
                        "<w:body><w:p><w:r><w:t>A</w:t></w:r><w:commentRangeStart w:id=\"0\"/><w:r><w:t>B</w:t></w:r><w:commentRangeEnd w:id=\"0\"/><w:r><w:commentReference w:id=\"0\"/></w:r></w:p></w:body>" +
                        "</w:document>");
                }
            }

            return ms.ToArray();
        }

        private static byte[] CreateDocxWithMalformedCommentsExtended()
        {
            using var ms = new MemoryStream();
            using (var archive = new ZipArchive(ms, ZipArchiveMode.Create, true))
            {
                var commentsEntry = archive.CreateEntry("word/comments.xml");
                using (var stream = commentsEntry.Open())
                using (var writer = new StreamWriter(stream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                        "<w:comments xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\">" +
                        "<w:comment w:id=\"0\" w:author=\"A\"><w:p w14:paraId=\"11111111\"><w:r><w:t>Comment malformed ext</w:t></w:r></w:p></w:comment>" +
                        "</w:comments>");
                }

                var commentsExtendedEntry = archive.CreateEntry("word/commentsExtended.xml");
                using (var stream = commentsExtendedEntry.Open())
                using (var writer = new StreamWriter(stream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><w15:commentsEx xmlns:w15=\"http://schemas.microsoft.com/office/word/2012/wordml\"><w15:commentEx");
                }

                var documentEntry = archive.CreateEntry("word/document.xml");
                using (var stream = documentEntry.Open())
                using (var writer = new StreamWriter(stream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                        "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">" +
                        "<w:body><w:p><w:r><w:t>A</w:t></w:r><w:commentRangeStart w:id=\"0\"/><w:r><w:t>B</w:t></w:r><w:commentRangeEnd w:id=\"0\"/><w:r><w:commentReference w:id=\"0\"/></w:r></w:p></w:body>" +
                        "</w:document>");
                }
            }

            return ms.ToArray();
        }

        private static byte[] CreateDocxWithHeaderAltChunkMissingHeaderRels()
        {
            using var ms = new MemoryStream();
            using (var archive = new ZipArchive(ms, ZipArchiveMode.Create, true))
            {
                var relsEntry = archive.CreateEntry("word/_rels/document.xml.rels");
                using (var relsStream = relsEntry.Open())
                using (var writer = new StreamWriter(relsStream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                        "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
                        "<Relationship Id=\"rIdHeader1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/header\" Target=\"header1.xml\"/>" +
                        "</Relationships>");
                }

                var headerChunkEntry = archive.CreateEntry("word/afchunk-header-missingrels.txt");
                using (var chunkStream = headerChunkEntry.Open())
                using (var writer = new StreamWriter(chunkStream))
                {
                    writer.Write("HeaderChunk");
                }

                var headerEntry = archive.CreateEntry("word/header1.xml");
                using (var stream = headerEntry.Open())
                using (var writer = new StreamWriter(stream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                        "<w:hdr xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">" +
                        "<w:p><w:r><w:t>HeaderBefore</w:t></w:r></w:p>" +
                        "<w:altChunk r:id=\"rIdHeaderChunk\"/>" +
                        "<w:p><w:r><w:t>HeaderAfter</w:t></w:r></w:p>" +
                        "</w:hdr>");
                }

                var documentEntry = archive.CreateEntry("word/document.xml");
                using (var stream = documentEntry.Open())
                using (var writer = new StreamWriter(stream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                        "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">" +
                        "<w:body><w:p><w:r><w:t>Body</w:t></w:r></w:p><w:sectPr><w:headerReference w:type=\"default\" r:id=\"rIdHeader1\"/></w:sectPr></w:body>" +
                        "</w:document>");
                }
            }

            return ms.ToArray();
        }

        private static byte[] CreateDocxWithMalformedEmbeddedDocxAltChunk()
        {
            using var ms = new MemoryStream();
            using (var archive = new ZipArchive(ms, ZipArchiveMode.Create, true))
            {
                var relsEntry = archive.CreateEntry("word/_rels/document.xml.rels");
                using (var relsStream = relsEntry.Open())
                using (var writer = new StreamWriter(relsStream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                        "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
                        "<Relationship Id=\"rIdChunk\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/aFChunk\" Target=\"afchunk-bad.docx\"/>" +
                        "</Relationships>");
                }

                var badChunkEntry = archive.CreateEntry("word/afchunk-bad.docx");
                using (var stream = badChunkEntry.Open())
                {
                    byte[] badBytes = Encoding.UTF8.GetBytes("not-a-zip-package");
                    stream.Write(badBytes, 0, badBytes.Length);
                }

                var documentEntry = archive.CreateEntry("word/document.xml");
                using (var stream = documentEntry.Open())
                using (var writer = new StreamWriter(stream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                        "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">" +
                        "<w:body><w:p><w:r><w:t>Before</w:t></w:r></w:p><w:altChunk r:id=\"rIdChunk\"/><w:p><w:r><w:t>After</w:t></w:r></w:p></w:body>" +
                        "</w:document>");
                }
            }

            return ms.ToArray();
        }

        private static byte[] CreateDocxWithPageBreakBeforeParagraph()
        {
            using var ms = new MemoryStream();
            using (var archive = new ZipArchive(ms, ZipArchiveMode.Create, true))
            {
                var documentEntry = archive.CreateEntry("word/document.xml");
                using (var stream = documentEntry.Open())
                using (var writer = new StreamWriter(stream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                        "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">" +
                        "<w:body>" +
                        "<w:p><w:r><w:t>First</w:t></w:r></w:p>" +
                        "<w:p><w:pPr><w:pageBreakBefore/></w:pPr><w:r><w:t>Second</w:t></w:r></w:p>" +
                        "</w:body>" +
                        "</w:document>");
                }
            }

            return ms.ToArray();
        }

        private static byte[] CreateDocxWithPageBreakBeforeDisabledParagraph()
        {
            using var ms = new MemoryStream();
            using (var archive = new ZipArchive(ms, ZipArchiveMode.Create, true))
            {
                var documentEntry = archive.CreateEntry("word/document.xml");
                using (var stream = documentEntry.Open())
                using (var writer = new StreamWriter(stream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                        "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">" +
                        "<w:body>" +
                        "<w:p><w:r><w:t>First</w:t></w:r></w:p>" +
                        "<w:p><w:pPr><w:pageBreakBefore w:val=\"false\"/></w:pPr><w:r><w:t>Second</w:t></w:r></w:p>" +
                        "</w:body>" +
                        "</w:document>");
                }
            }

            return ms.ToArray();
        }

        private static byte[] CreateDocxWithFootnoteReferenceWithoutFootnotesPart()
        {
            using var ms = new MemoryStream();
            using (var archive = new ZipArchive(ms, ZipArchiveMode.Create, true))
            {
                var documentEntry = archive.CreateEntry("word/document.xml");
                using var stream = documentEntry.Open();
                using var writer = new StreamWriter(stream);
                writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                    "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">" +
                    "<w:body><w:p><w:r><w:t>A</w:t></w:r><w:r><w:footnoteReference w:id=\"2\"/></w:r></w:p></w:body>" +
                    "</w:document>");
            }

            return ms.ToArray();
        }

        private static byte[] CreateDocxWithValidFootnotesPart()
        {
            using var ms = new MemoryStream();
            using (var archive = new ZipArchive(ms, ZipArchiveMode.Create, true))
            {
                var footnotesEntry = archive.CreateEntry("word/footnotes.xml");
                using (var stream = footnotesEntry.Open())
                using (var writer = new StreamWriter(stream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                        "<w:footnotes xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">" +
                        "<w:footnote w:id=\"2\"><w:p><w:r><w:t>Footnote text</w:t></w:r></w:p></w:footnote>" +
                        "</w:footnotes>");
                }

                var documentEntry = archive.CreateEntry("word/document.xml");
                using (var stream = documentEntry.Open())
                using (var writer = new StreamWriter(stream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                        "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">" +
                        "<w:body><w:p><w:r><w:t>A</w:t></w:r><w:r><w:footnoteReference w:id=\"2\"/></w:r></w:p></w:body>" +
                        "</w:document>");
                }
            }

            return ms.ToArray();
        }

        private static byte[] CreateDocxWithMalformedFootnotesXml()
        {
            using var ms = new MemoryStream();
            using (var archive = new ZipArchive(ms, ZipArchiveMode.Create, true))
            {
                var footnotesEntry = archive.CreateEntry("word/footnotes.xml");
                using (var stream = footnotesEntry.Open())
                using (var writer = new StreamWriter(stream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><w:footnotes xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:footnote");
                }

                var documentEntry = archive.CreateEntry("word/document.xml");
                using (var stream = documentEntry.Open())
                using (var writer = new StreamWriter(stream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                        "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:body><w:p><w:r><w:t>A</w:t></w:r></w:p></w:body></w:document>");
                }
            }

            return ms.ToArray();
        }

        private static byte[] CreateDocxWithValidEndnotesPart()
        {
            using var ms = new MemoryStream();
            using (var archive = new ZipArchive(ms, ZipArchiveMode.Create, true))
            {
                var endnotesEntry = archive.CreateEntry("word/endnotes.xml");
                using (var stream = endnotesEntry.Open())
                using (var writer = new StreamWriter(stream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                        "<w:endnotes xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">" +
                        "<w:endnote w:id=\"2\"><w:p><w:r><w:t>Endnote text</w:t></w:r></w:p></w:endnote>" +
                        "</w:endnotes>");
                }

                var documentEntry = archive.CreateEntry("word/document.xml");
                using (var stream = documentEntry.Open())
                using (var writer = new StreamWriter(stream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                        "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">" +
                        "<w:body><w:p><w:r><w:t>A</w:t></w:r><w:r><w:endnoteReference w:id=\"2\"/></w:r></w:p></w:body>" +
                        "</w:document>");
                }
            }

            return ms.ToArray();
        }

        private static byte[] CreateDocxWithMalformedEndnotesXml()
        {
            using var ms = new MemoryStream();
            using (var archive = new ZipArchive(ms, ZipArchiveMode.Create, true))
            {
                var endnotesEntry = archive.CreateEntry("word/endnotes.xml");
                using (var stream = endnotesEntry.Open())
                using (var writer = new StreamWriter(stream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><w:endnotes xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:endnote");
                }

                var documentEntry = archive.CreateEntry("word/document.xml");
                using (var stream = documentEntry.Open())
                using (var writer = new StreamWriter(stream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                        "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:body><w:p><w:r><w:t>A</w:t></w:r></w:p></w:body></w:document>");
                }
            }

            return ms.ToArray();
        }

        private static byte[] CreateDocxWithMalformedCommentsExtendedAndMalformedHeader()
        {
            using var ms = new MemoryStream();
            using (var archive = new ZipArchive(ms, ZipArchiveMode.Create, true))
            {
                var commentsEntry = archive.CreateEntry("word/comments.xml");
                using (var stream = commentsEntry.Open())
                using (var writer = new StreamWriter(stream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                        "<w:comments xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\">" +
                        "<w:comment w:id=\"0\" w:author=\"A\"><w:p w14:paraId=\"11111111\"><w:r><w:t>Comment text</w:t></w:r></w:p></w:comment>" +
                        "</w:comments>");
                }

                var commentsExtendedEntry = archive.CreateEntry("word/commentsExtended.xml");
                using (var stream = commentsExtendedEntry.Open())
                using (var writer = new StreamWriter(stream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><w15:commentsEx xmlns:w15=\"http://schemas.microsoft.com/office/word/2012/wordml\"><w15:commentEx");
                }

                var relsEntry = archive.CreateEntry("word/_rels/document.xml.rels");
                using (var relsStream = relsEntry.Open())
                using (var writer = new StreamWriter(relsStream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                        "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
                        "<Relationship Id=\"rIdHeader1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/header\" Target=\"header1.xml\"/>" +
                        "</Relationships>");
                }

                var headerEntry = archive.CreateEntry("word/header1.xml");
                using (var stream = headerEntry.Open())
                using (var writer = new StreamWriter(stream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><w:hdr xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:p>");
                }

                var documentEntry = archive.CreateEntry("word/document.xml");
                using (var stream = documentEntry.Open())
                using (var writer = new StreamWriter(stream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                        "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">" +
                        "<w:body><w:p><w:r><w:t>Body</w:t></w:r></w:p><w:sectPr><w:headerReference w:type=\"default\" r:id=\"rIdHeader1\"/></w:sectPr></w:body>" +
                        "</w:document>");
                }
            }

            return ms.ToArray();
        }

        private static byte[] CreateDocxWithValidCommentsAndMalformedHeader()
        {
            using var ms = new MemoryStream();
            using (var archive = new ZipArchive(ms, ZipArchiveMode.Create, true))
            {
                var commentsEntry = archive.CreateEntry("word/comments.xml");
                using (var stream = commentsEntry.Open())
                using (var writer = new StreamWriter(stream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                        "<w:comments xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">" +
                        "<w:comment w:id=\"0\" w:author=\"A\"><w:p><w:r><w:t>Comment text</w:t></w:r></w:p></w:comment>" +
                        "</w:comments>");
                }

                var relsEntry = archive.CreateEntry("word/_rels/document.xml.rels");
                using (var relsStream = relsEntry.Open())
                using (var writer = new StreamWriter(relsStream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                        "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
                        "<Relationship Id=\"rIdHeader1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/header\" Target=\"header1.xml\"/>" +
                        "</Relationships>");
                }

                var headerEntry = archive.CreateEntry("word/header1.xml");
                using (var stream = headerEntry.Open())
                using (var writer = new StreamWriter(stream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><w:hdr xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:p>");
                }

                var documentEntry = archive.CreateEntry("word/document.xml");
                using (var stream = documentEntry.Open())
                using (var writer = new StreamWriter(stream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                        "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">" +
                        "<w:body><w:p><w:r><w:t>Body</w:t></w:r></w:p><w:sectPr><w:headerReference w:type=\"default\" r:id=\"rIdHeader1\"/></w:sectPr></w:body>" +
                        "</w:document>");
                }
            }

            return ms.ToArray();
        }

        private static string CreateTempOutputPath(string fileStem)
        {
            string tempDirectory = Path.Combine(Path.GetTempPath(), "Nedev.FileConverters.DocxToDoc.Tests", Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(tempDirectory);
            return Path.Combine(tempDirectory, fileStem + ".doc");
        }

        private static void TryDeleteFile(string path)
        {
            try
            {
                if (File.Exists(path))
                {
                    Directory.Delete(Path.GetDirectoryName(path)!, true);
                }
            }
            catch
            {
            }
        }

        private static string ResolveRepositoryRoot()
        {
            string? current = AppContext.BaseDirectory;
            while (!string.IsNullOrEmpty(current))
            {
                string solutionPath = Path.Combine(current, "Nedev.FileConverters.DocxToDoc.sln");
                if (File.Exists(solutionPath))
                {
                    return current;
                }

                current = Directory.GetParent(current)?.FullName;
            }

            throw new DirectoryNotFoundException("Repository root was not found.");
        }

        private static string ExtractMainStoryText(byte[] wordDocumentData)
        {
            int ccpText = BitConverter.ToInt32(wordDocumentData, 64);
            int textLength = Math.Max(0, Math.Min(ccpText, wordDocumentData.Length - 1536));
            if (textLength == 0)
            {
                return string.Empty;
            }

            byte[] textBytes = new byte[textLength];
            Array.Copy(wordDocumentData, 1536, textBytes, 0, textLength);
            return Encoding.GetEncoding(1252).GetString(textBytes);
        }

        private static string ExtractAllStoryText(byte[] wordDocumentData)
        {
            int ccpText = BitConverter.ToInt32(wordDocumentData, 64);
            int ccpFtn = BitConverter.ToInt32(wordDocumentData, 68);
            int ccpHdd = BitConverter.ToInt32(wordDocumentData, 72);
            int ccpAtn = BitConverter.ToInt32(wordDocumentData, 76);
            int ccpEdn = BitConverter.ToInt32(wordDocumentData, 80);
            int totalLength = ccpText + ccpFtn + ccpHdd + ccpAtn + ccpEdn;
            int safeLength = Math.Max(0, Math.Min(totalLength, wordDocumentData.Length - 1536));
            if (safeLength == 0)
            {
                return string.Empty;
            }

            byte[] textBytes = new byte[safeLength];
            Array.Copy(wordDocumentData, 1536, textBytes, 0, safeLength);
            return Encoding.GetEncoding(1252).GetString(textBytes);
        }

        private sealed class ThrowOnWriteStream : Stream
        {
            public override bool CanRead => false;
            public override bool CanSeek => false;
            public override bool CanWrite => true;
            public override long Length => 0;
            public override long Position { get; set; }
            public override void Flush()
            {
            }
            public override int Read(byte[] buffer, int offset, int count) => throw new NotSupportedException();
            public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
            public override void SetLength(long value) => throw new NotSupportedException();
            public override void Write(byte[] buffer, int offset, int count) => throw new IOException("write failed");
        }

        private sealed class ThrowOnFlushStream : Stream
        {
            public override bool CanRead => false;
            public override bool CanSeek => false;
            public override bool CanWrite => true;
            public override long Length => 0;
            public override long Position { get; set; }
            public override void Flush() => throw new IOException("flush failed");
            public override int Read(byte[] buffer, int offset, int count) => throw new NotSupportedException();
            public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
            public override void SetLength(long value) => throw new NotSupportedException();
            public override void Write(byte[] buffer, int offset, int count)
            {
            }
        }

        private sealed class TestLogger : ILogger
        {
            public List<string> InfoMessages { get; } = new List<string>();

            public void LogDebug(string message) { }

            public void LogInfo(string message)
            {
                InfoMessages.Add(message);
            }

            public void LogWarning(string message) { }

            public void LogError(string message) { }

            public void LogError(string message, Exception exception) { }
        }
    }
}
