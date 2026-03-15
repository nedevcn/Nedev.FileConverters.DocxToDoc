using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using Xunit;

namespace Nedev.FileConverters.DocxToDoc.Tests
{
    public class ErrorHandlingTests
    {
        [Fact]
        public void Convert_NullDocxPath_ThrowsArgumentNullException()
        {
            // Arrange
            var converter = new DocxToDocConverter();

            // Act & Assert
            Assert.Throws<ArgumentNullException>(() => converter.Convert(null!, "output.doc"));
        }

        [Fact]
        public void Convert_NullDocPath_ThrowsArgumentNullException()
        {
            // Arrange
            var converter = new DocxToDocConverter();

            // Act & Assert
            Assert.Throws<ArgumentNullException>(() => converter.Convert("input.docx", null!));
        }

        [Fact]
        public void Convert_EmptyDocxPath_ThrowsArgumentNullException()
        {
            // Arrange
            var converter = new DocxToDocConverter();

            // Act & Assert
            Assert.Throws<ArgumentNullException>(() => converter.Convert("", "output.doc"));
        }

        [Fact]
        public void Convert_NonExistentFile_ThrowsFileNotFoundException()
        {
            // Arrange
            var converter = new DocxToDocConverter();
            string nonExistentPath = Path.Combine(Path.GetTempPath(), $"nonexistent_{Guid.NewGuid()}.docx");

            // Act & Assert
            Assert.Throws<FileNotFoundException>(() => converter.Convert(nonExistentPath, "output.doc"));
        }

        [Fact]
        public void Convert_EmptyFile_ThrowsConversionException()
        {
            // Arrange
            var converter = new DocxToDocConverter();
            string emptyFilePath = Path.Combine(Path.GetTempPath(), $"empty_{Guid.NewGuid()}.docx");
            File.WriteAllText(emptyFilePath, "");

            try
            {
                // Act & Assert
                var ex = Assert.Throws<ConversionException>(() => converter.Convert(emptyFilePath, "output.doc"));
                Assert.Equal(ConversionStage.Validation, ex.Stage);
            }
            finally
            {
                File.Delete(emptyFilePath);
            }
        }

        [Fact]
        public void Convert_NullInputStream_ThrowsArgumentNullException()
        {
            // Arrange
            var converter = new DocxToDocConverter();

            // Act & Assert
            Assert.Throws<ArgumentNullException>(() => converter.Convert(null!, new MemoryStream()));
        }

        [Fact]
        public void Convert_NullOutputStream_ThrowsArgumentNullException()
        {
            // Arrange
            var converter = new DocxToDocConverter();

            // Act & Assert
            Assert.Throws<ArgumentNullException>(() => converter.Convert(new MemoryStream(), null!));
        }

        [Fact]
        public void Convert_NonReadableStream_ThrowsArgumentException()
        {
            // Arrange
            var converter = new DocxToDocConverter();
            var nonReadableStream = new TestStream(canRead: false, canWrite: true);

            // Act & Assert
            var ex = Assert.Throws<ArgumentException>(() => converter.Convert(nonReadableStream, new MemoryStream()));
            Assert.Contains("reading", ex.Message);
        }

        [Fact]
        public void Convert_NonWritableStream_ThrowsArgumentException()
        {
            // Arrange
            var converter = new DocxToDocConverter();
            var nonWritableStream = new TestStream(canRead: true, canWrite: false);

            // Act & Assert
            var ex = Assert.Throws<ArgumentException>(() => converter.Convert(new MemoryStream(), nonWritableStream));
            Assert.Contains("writing", ex.Message);
        }

        [Fact]
        public async Task ConvertAsync_NullDocxPath_ThrowsArgumentNullException()
        {
            // Arrange
            var converter = new DocxToDocConverter();

            // Act & Assert
            await Assert.ThrowsAsync<ArgumentNullException>(() => converter.ConvertAsync(null!, "output.doc"));
        }

        [Fact]
        public async Task ConvertAsync_NullDocPath_ThrowsArgumentNullException()
        {
            // Arrange
            var converter = new DocxToDocConverter();

            // Act & Assert
            await Assert.ThrowsAsync<ArgumentNullException>(() => converter.ConvertAsync("input.docx", null!));
        }

        [Fact]
        public async Task ConvertAsync_Cancellation_ThrowsOperationCanceledException()
        {
            // Arrange - create a valid DOCX file first
            var converter = new DocxToDocConverter();
            string tempDocx = Path.Combine(Path.GetTempPath(), $"test_{Guid.NewGuid()}.docx");
            CreateMinimalDocx(tempDocx);

            var cts = new CancellationTokenSource();
            cts.Cancel();

            try
            {
                // Act & Assert - should throw OperationCanceledException when cancelled
                await Assert.ThrowsAsync<OperationCanceledException>(() =>
                    converter.ConvertAsync(tempDocx, "output.doc", cts.Token));
            }
            finally
            {
                File.Delete(tempDocx);
            }
        }

        private void CreateMinimalDocx(string path)
        {
            using var fs = File.Create(path);
            using var archive = new System.IO.Compression.ZipArchive(fs, System.IO.Compression.ZipArchiveMode.Create);
            
            // Create [Content_Types].xml
            var contentTypesEntry = archive.CreateEntry("[Content_Types].xml");
            using (var stream = contentTypesEntry.Open())
            using (var writer = new StreamWriter(stream))
            {
                writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                    "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">" +
                    "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>" +
                    "<Default Extension=\"xml\" ContentType=\"application/xml\"/>" +
                    "<Override PartName=\"/word/document.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml\"/>" +
                    "</Types>");
            }

            // Create _rels/.rels
            var relsEntry = archive.CreateEntry("_rels/.rels");
            using (var stream = relsEntry.Open())
            using (var writer = new StreamWriter(stream))
            {
                writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                    "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
                    "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"word/document.xml\"/>" +
                    "</Relationships>");
            }

            // Create word/_rels/document.xml.rels
            var docRelsEntry = archive.CreateEntry("word/_rels/document.xml.rels");
            using (var stream = docRelsEntry.Open())
            using (var writer = new StreamWriter(stream))
            {
                writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                    "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
                    "</Relationships>");
            }

            // Create word/document.xml
            var docEntry = archive.CreateEntry("word/document.xml");
            using (var stream = docEntry.Open())
            using (var writer = new StreamWriter(stream))
            {
                writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                    "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">" +
                    "<w:body><w:p><w:r><w:t>Test</w:t></w:r></w:p></w:body>" +
                    "</w:document>");
            }
        }

        [Fact]
        public void ConversionException_ContainsCorrectStage()
        {
            // Arrange
            var innerException = new InvalidOperationException("Inner error");
            
            // Act
            var ex = new ConversionException(
                "Test message",
                "source.docx",
                "dest.doc",
                ConversionStage.Parsing,
                innerException);

            // Assert
            Assert.Equal("Test message", ex.Message);
            Assert.Equal("source.docx", ex.SourcePath);
            Assert.Equal("dest.doc", ex.DestinationPath);
            Assert.Equal(ConversionStage.Parsing, ex.Stage);
            Assert.Equal(innerException, ex.InnerException);
        }

        [Fact]
        public void NullLogger_DoesNotThrow()
        {
            // Arrange
            var logger = NullLogger.Instance;

            // Act & Assert - should not throw
            logger.LogDebug("debug");
            logger.LogInfo("info");
            logger.LogWarning("warning");
            logger.LogError("error");
            logger.LogError("error with exception", new Exception("test"));
        }

        [Fact]
        public void ConsoleLogger_DoesNotThrow()
        {
            // Arrange
            var logger = new ConsoleLogger();

            // Act & Assert - should not throw (output goes to console)
            logger.LogDebug("debug");
            logger.LogInfo("info");
            logger.LogWarning("warning");
            logger.LogError("error");
            logger.LogError("error with exception", new Exception("test"));
        }

        /// <summary>
        /// A test stream with configurable capabilities.
        /// </summary>
        private class TestStream : Stream
        {
            public TestStream(bool canRead, bool canWrite)
            {
                CanRead = canRead;
                CanWrite = canWrite;
            }

            public override bool CanRead { get; }
            public override bool CanSeek => false;
            public override bool CanWrite { get; }
            public override long Length => 0;
            public override long Position { get; set; }

            public override void Flush() { }
            public override int Read(byte[] buffer, int offset, int count) => 0;
            public override long Seek(long offset, SeekOrigin origin) => 0;
            public override void SetLength(long value) { }
            public override void Write(byte[] buffer, int offset, int count) { }
        }
    }
}
