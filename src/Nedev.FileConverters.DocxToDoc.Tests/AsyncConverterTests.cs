using System;
using System.IO;
using System.IO.Compression;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Xunit;

namespace Nedev.FileConverters.DocxToDoc.Tests
{
    public class AsyncConverterTests
    {
        private byte[] CreateMinimalDocx()
        {
            using var ms = new MemoryStream();
            using (var archive = new ZipArchive(ms, ZipArchiveMode.Create, true))
            {
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

                // Create word/document.xml
                var docEntry = archive.CreateEntry("word/document.xml");
                using (var stream = docEntry.Open())
                using (var writer = new StreamWriter(stream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                        "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">" +
                        "<w:body>" +
                        "<w:p><w:r><w:t>Hello World</w:t></w:r></w:p>" +
                        "</w:body>" +
                        "</w:document>");
                }
            }
            return ms.ToArray();
        }

        [Fact]
        public async Task ConvertAsync_StreamToStream_Succeeds()
        {
            // Arrange
            var converter = new DocxToDocConverter();
            byte[] docxData = CreateMinimalDocx();
            using var inputStream = new MemoryStream(docxData);
            using var outputStream = new MemoryStream();

            // Act
            await converter.ConvertAsync(inputStream, outputStream);

            // Assert
            Assert.True(outputStream.Length > 0);
            Assert.True(outputStream.Length > 512); // Should be larger than FIB header
        }

        [Fact]
        public async Task ConvertAsync_FileToFile_Succeeds()
        {
            // Arrange
            var converter = new DocxToDocConverter();
            string tempDir = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString());
            Directory.CreateDirectory(tempDir);
            string inputFile = Path.Combine(tempDir, "test.docx");
            string outputFile = Path.Combine(tempDir, "test.doc");

            try
            {
                // Create test file
                byte[] docxData = CreateMinimalDocx();
                await File.WriteAllBytesAsync(inputFile, docxData);

                // Act
                await converter.ConvertAsync(inputFile, outputFile);

                // Assert
                Assert.True(File.Exists(outputFile));
                var fileInfo = new FileInfo(outputFile);
                Assert.True(fileInfo.Length > 512);
            }
            finally
            {
                // Cleanup
                try { Directory.Delete(tempDir, true); } catch { }
            }
        }

        [Fact]
        public async Task ConvertAsync_Cancellation_ThrowsOperationCanceledException()
        {
            // Arrange
            var converter = new DocxToDocConverter();
            byte[] docxData = CreateMinimalDocx();
            using var inputStream = new MemoryStream(docxData);
            using var outputStream = new MemoryStream();
            using var cts = new CancellationTokenSource();

            // Cancel immediately
            cts.Cancel();

            // Act & Assert
            await Assert.ThrowsAsync<OperationCanceledException>(async () =>
            {
                await converter.ConvertAsync(inputStream, outputStream, cts.Token);
            });
        }

        [Fact]
        public async Task ConvertAsync_NullInputStream_ThrowsArgumentNullException()
        {
            // Arrange
            var converter = new DocxToDocConverter();
            using var outputStream = new MemoryStream();

            // Act & Assert
            await Assert.ThrowsAsync<ArgumentNullException>(async () =>
            {
                await converter.ConvertAsync(null!, outputStream);
            });
        }

        [Fact]
        public async Task ConvertAsync_NullOutputStream_ThrowsArgumentNullException()
        {
            // Arrange
            var converter = new DocxToDocConverter();
            byte[] docxData = CreateMinimalDocx();
            using var inputStream = new MemoryStream(docxData);

            // Act & Assert
            await Assert.ThrowsAsync<ArgumentNullException>(async () =>
            {
                await converter.ConvertAsync(inputStream, null!);
            });
        }

        [Fact]
        public async Task ConvertAsync_MultipleConversions_SequentialExecution()
        {
            // Arrange
            var converter = new DocxToDocConverter();
            byte[] docxData = CreateMinimalDocx();
            int count = 3;
            var results = new List<long>();

            // Act - Run conversions sequentially (streams can't be shared)
            for (int i = 0; i < count; i++)
            {
                using var inputStream = new MemoryStream(docxData);
                using var outputStream = new MemoryStream();
                await converter.ConvertAsync(inputStream, outputStream);
                results.Add(outputStream.Length);
            }

            // Assert - all conversions should produce valid output
            Assert.Equal(count, results.Count);
            Assert.All(results, len => Assert.True(len > 512));
        }
    }
}
