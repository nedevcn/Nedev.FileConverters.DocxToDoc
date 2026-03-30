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
        public void Convert_InvalidDocxStream_ThrowsConversionExceptionWithParsingStage()
        {
            var converter = new DocxToDocConverter();
            using var invalidDocx = new MemoryStream(new byte[] { 1, 2, 3, 4, 5 });
            using var output = new MemoryStream();

            var ex = Assert.Throws<ConversionException>(() => converter.Convert(invalidDocx, output));
            Assert.Equal(ConversionStage.Parsing, ex.Stage);
        }

        [Fact]
        public void Convert_DocxMissingMainDocument_ThrowsConversionExceptionWithReadingStage()
        {
            var converter = new DocxToDocConverter();
            using var input = CreateDocxWithoutMainDocument();
            using var output = new MemoryStream();

            var ex = Assert.Throws<ConversionException>(() => converter.Convert(input, output));
            Assert.Equal(ConversionStage.Reading, ex.Stage);
        }

        [Fact]
        public void Convert_InvalidOutputPath_ThrowsConversionExceptionWithWritingStage()
        {
            var converter = new DocxToDocConverter();
            string tempDocx = Path.Combine(Path.GetTempPath(), $"valid_{Guid.NewGuid()}.docx");
            CreateMinimalDocx(tempDocx);

            string invalidOutputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString(), "out.doc");
            try
            {
                var ex = Assert.Throws<ConversionException>(() => converter.Convert(tempDocx, invalidOutputPath));
                Assert.Equal(ConversionStage.Writing, ex.Stage);
            }
            finally
            {
                File.Delete(tempDocx);
            }
        }

        [Fact]
        public async Task ConvertAsync_InvalidDocxStream_ThrowsConversionExceptionWithParsingStage()
        {
            var converter = new DocxToDocConverter();
            using var invalidDocx = new MemoryStream(new byte[] { 9, 8, 7, 6 });
            using var output = new MemoryStream();

            var ex = await Assert.ThrowsAsync<ConversionException>(() => converter.ConvertAsync(invalidDocx, output));
            Assert.Equal(ConversionStage.Parsing, ex.Stage);
        }

        [Fact]
        public async Task ConvertAsync_WriteFailure_ThrowsConversionExceptionWithWritingStage()
        {
            var converter = new DocxToDocConverter();
            string tempDocx = CreateTemporaryMinimalDocx();
            try
            {
                using var input = File.OpenRead(tempDocx);
                using var output = new ThrowOnWriteStream();

                var ex = await Assert.ThrowsAsync<ConversionException>(() => converter.ConvertAsync(input, output));
                Assert.Equal(ConversionStage.Writing, ex.Stage);
            }
            finally
            {
                File.Delete(tempDocx);
            }
        }

        [Fact]
        public void Convert_FinalizeFailure_ThrowsConversionExceptionWithFinalizingStage()
        {
            var converter = new DocxToDocConverter();
            using var input = CreateDocxAsMemoryStream();
            using var output = new ThrowOnFlushStream();

            var ex = Assert.Throws<ConversionException>(() => converter.Convert(input, output));
            Assert.Equal(ConversionStage.Finalizing, ex.Stage);
        }

        [Fact]
        public async Task ConvertAsync_FinalizeFailure_ThrowsConversionExceptionWithFinalizingStage()
        {
            var converter = new DocxToDocConverter();
            using var input = CreateDocxAsMemoryStream();
            using var output = new ThrowOnFlushStream();

            var ex = await Assert.ThrowsAsync<ConversionException>(() => converter.ConvertAsync(input, output));
            Assert.Equal(ConversionStage.Finalizing, ex.Stage);
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

        private string CreateTemporaryMinimalDocx()
        {
            string path = Path.Combine(Path.GetTempPath(), $"test_{Guid.NewGuid()}.docx");
            CreateMinimalDocx(path);
            return path;
        }

        private MemoryStream CreateDocxAsMemoryStream()
        {
            var stream = new MemoryStream();
            using (var archive = new System.IO.Compression.ZipArchive(stream, System.IO.Compression.ZipArchiveMode.Create, leaveOpen: true))
            {
                var contentTypesEntry = archive.CreateEntry("[Content_Types].xml");
                using (var contentTypesStream = contentTypesEntry.Open())
                using (var writer = new StreamWriter(contentTypesStream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                        "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">" +
                        "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>" +
                        "<Default Extension=\"xml\" ContentType=\"application/xml\"/>" +
                        "<Override PartName=\"/word/document.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml\"/>" +
                        "</Types>");
                }

                var relsEntry = archive.CreateEntry("_rels/.rels");
                using (var relsStream = relsEntry.Open())
                using (var writer = new StreamWriter(relsStream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                        "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
                        "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"word/document.xml\"/>" +
                        "</Relationships>");
                }

                var docRelsEntry = archive.CreateEntry("word/_rels/document.xml.rels");
                using (var docRelsStream = docRelsEntry.Open())
                using (var writer = new StreamWriter(docRelsStream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                        "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
                        "</Relationships>");
                }

                var docEntry = archive.CreateEntry("word/document.xml");
                using (var docStream = docEntry.Open())
                using (var writer = new StreamWriter(docStream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                        "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">" +
                        "<w:body><w:p><w:r><w:t>Finalize</w:t></w:r></w:p></w:body>" +
                        "</w:document>");
                }
            }

            stream.Position = 0;
            return stream;
        }

        private MemoryStream CreateDocxWithoutMainDocument()
        {
            var stream = new MemoryStream();
            using (var archive = new System.IO.Compression.ZipArchive(stream, System.IO.Compression.ZipArchiveMode.Create, leaveOpen: true))
            {
                var contentTypesEntry = archive.CreateEntry("[Content_Types].xml");
                using (var contentTypesStream = contentTypesEntry.Open())
                using (var writer = new StreamWriter(contentTypesStream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                        "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">" +
                        "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>" +
                        "<Default Extension=\"xml\" ContentType=\"application/xml\"/>" +
                        "<Override PartName=\"/word/document.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml\"/>" +
                        "</Types>");
                }

                var relsEntry = archive.CreateEntry("_rels/.rels");
                using (var relsStream = relsEntry.Open())
                using (var writer = new StreamWriter(relsStream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                        "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
                        "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"word/document.xml\"/>" +
                        "</Relationships>");
                }
            }

            stream.Position = 0;
            return stream;
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

        private class ThrowOnWriteStream : Stream
        {
            public override bool CanRead => false;
            public override bool CanSeek => false;
            public override bool CanWrite => true;
            public override long Length => 0;
            public override long Position { get; set; }

            public override void Flush() { }
            public override int Read(byte[] buffer, int offset, int count) => throw new NotSupportedException();
            public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
            public override void SetLength(long value) => throw new NotSupportedException();
            public override void Write(byte[] buffer, int offset, int count) => throw new IOException("write failed");
        }

        private class ThrowOnFlushStream : Stream
        {
            private readonly MemoryStream _inner = new MemoryStream();

            public override bool CanRead => _inner.CanRead;
            public override bool CanSeek => _inner.CanSeek;
            public override bool CanWrite => _inner.CanWrite;
            public override long Length => _inner.Length;
            public override long Position
            {
                get => _inner.Position;
                set => _inner.Position = value;
            }

            public override void Flush() => throw new IOException("flush failed");
            public override Task FlushAsync(CancellationToken cancellationToken) => Task.FromException(new IOException("flush failed"));
            public override int Read(byte[] buffer, int offset, int count) => _inner.Read(buffer, offset, count);
            public override long Seek(long offset, SeekOrigin origin) => _inner.Seek(offset, origin);
            public override void SetLength(long value) => _inner.SetLength(value);
            public override void Write(byte[] buffer, int offset, int count) => _inner.Write(buffer, offset, count);
            protected override void Dispose(bool disposing)
            {
                if (disposing)
                {
                    _inner.Dispose();
                }

                base.Dispose(disposing);
            }
        }
    }
}
