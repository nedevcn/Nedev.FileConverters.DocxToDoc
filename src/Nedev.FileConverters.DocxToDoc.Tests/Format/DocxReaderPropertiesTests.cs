using System;
using System.IO;
using System.IO.Compression;
using Xunit;

namespace Nedev.FileConverters.DocxToDoc.Tests.Format
{
    public class DocxReaderPropertiesTests
    {
        private byte[] CreateDocxWithProperties()
        {
            using var ms = new MemoryStream();
            using (var archive = new ZipArchive(ms, ZipArchiveMode.Create, true))
            {
                var corePropsEntry = archive.CreateEntry("docProps/core.xml");
                using (var stream = corePropsEntry.Open())
                using (var writer = new StreamWriter(stream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                        "<cp:coreProperties xmlns:cp=\"http://schemas.openxmlformats.org/package/2006/metadata/core-properties\" " +
                        "xmlns:dc=\"http://purl.org/dc/elements/1.1/\" " +
                        "xmlns:dcterms=\"http://purl.org/dc/terms/\" " +
                        "xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\">" +
                        "<dc:title>Test Document Title</dc:title>" +
                        "<dc:subject>Test Subject</dc:subject>" +
                        "<dc:creator>John Doe</dc:creator>" +
                        "<cp:keywords>test, document, keywords</cp:keywords>" +
                        "<dc:description>This is a test document</dc:description>" +
                        "<cp:lastPrinted>2024-01-17T08:00:00Z</cp:lastPrinted>" +
                        "<cp:revision>5</cp:revision>" +
                        "<dcterms:created xsi:type=\"dcterms:W3CDTF\">2024-01-15T10:30:00Z</dcterms:created>" +
                        "<dcterms:modified xsi:type=\"dcterms:W3CDTF\">2024-01-16T14:20:00Z</dcterms:modified>" +
                        "</cp:coreProperties>");
                }

                var extendedPropsEntry = archive.CreateEntry("docProps/app.xml");
                using (var stream = extendedPropsEntry.Open())
                using (var writer = new StreamWriter(stream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                        "<Properties xmlns=\"http://schemas.openxmlformats.org/officeDocument/2006/extended-properties\" " +
                        "xmlns:vt=\"http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes\">" +
                        "<Application>Microsoft Word</Application>" +
                        "<Pages>7</Pages>" +
                        "<Words>123</Words>" +
                        "<Characters>456</Characters>" +
                        "<TotalTime>42</TotalTime>" +
                        "<Company>Test Company</Company>" +
                        "<Manager>Jane Smith</Manager>" +
                        "<Category>Test Category</Category>" +
                        "</Properties>");
                }

                var docEntry = archive.CreateEntry("word/document.xml");
                using (var stream = docEntry.Open())
                using (var writer = new StreamWriter(stream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                        "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">" +
                        "<w:body>" +
                        "<w:p><w:r><w:t>Test content</w:t></w:r></w:p>" +
                        "</w:body>" +
                        "</w:document>");
                }
            }

            return ms.ToArray();
        }

        [Fact]
        public void ReadDocument_WithProperties_ParsesCoreTextFields()
        {
            byte[] docxData = CreateDocxWithProperties();
            using var ms = new MemoryStream(docxData);
            using var reader = new Nedev.FileConverters.DocxToDoc.Format.DocxReader(ms);

            var model = reader.ReadDocument();

            Assert.Equal("Test Document Title", model.Properties.Title);
            Assert.Equal("Test Subject", model.Properties.Subject);
            Assert.Equal("John Doe", model.Properties.Author);
            Assert.Equal("test, document, keywords", model.Properties.Keywords);
            Assert.Equal("This is a test document", model.Properties.Comments);
            Assert.Equal(5, model.Properties.Revision);
        }

        [Fact]
        public void ReadDocument_WithProperties_ParsesSubjectFromCoreProperties()
        {
            byte[] docxData = CreateDocxWithProperties();
            using var ms = new MemoryStream(docxData);
            using var reader = new Nedev.FileConverters.DocxToDoc.Format.DocxReader(ms);

            var model = reader.ReadDocument();

            Assert.Equal("Test Subject", model.Properties.Subject);
        }

        [Fact]
        public void ReadDocument_WithProperties_ParsesCoreDateFields()
        {
            byte[] docxData = CreateDocxWithProperties();
            using var ms = new MemoryStream(docxData);
            using var reader = new Nedev.FileConverters.DocxToDoc.Format.DocxReader(ms);

            var model = reader.ReadDocument();

            Assert.Equal(new DateTime(2024, 1, 15, 10, 30, 0, DateTimeKind.Utc), model.Properties.Created!.Value.ToUniversalTime());
            Assert.Equal(new DateTime(2024, 1, 16, 14, 20, 0, DateTimeKind.Utc), model.Properties.Modified!.Value.ToUniversalTime());
            Assert.Equal(new DateTime(2024, 1, 17, 8, 0, 0, DateTimeKind.Utc), model.Properties.LastPrinted!.Value.ToUniversalTime());
        }

        [Fact]
        public void ReadDocument_WithProperties_ParsesLastPrintedFromCoreProperties()
        {
            byte[] docxData = CreateDocxWithProperties();
            using var ms = new MemoryStream(docxData);
            using var reader = new Nedev.FileConverters.DocxToDoc.Format.DocxReader(ms);

            var model = reader.ReadDocument();

            Assert.Equal(new DateTime(2024, 1, 17, 8, 0, 0, DateTimeKind.Utc), model.Properties.LastPrinted!.Value.ToUniversalTime());
        }

        [Fact]
        public void ReadDocument_WithProperties_ParsesExtendedIdentityFields()
        {
            byte[] docxData = CreateDocxWithProperties();
            using var ms = new MemoryStream(docxData);
            using var reader = new Nedev.FileConverters.DocxToDoc.Format.DocxReader(ms);

            var model = reader.ReadDocument();

            Assert.Equal("Test Company", model.Properties.Company);
            Assert.Equal("Jane Smith", model.Properties.Manager);
            Assert.Equal("Test Category", model.Properties.Category);
        }

        [Fact]
        public void ReadDocument_WithProperties_ParsesManagerFromExtendedProperties()
        {
            byte[] docxData = CreateDocxWithProperties();
            using var ms = new MemoryStream(docxData);
            using var reader = new Nedev.FileConverters.DocxToDoc.Format.DocxReader(ms);

            var model = reader.ReadDocument();

            Assert.Equal("Jane Smith", model.Properties.Manager);
        }

        [Fact]
        public void ReadDocument_WithProperties_ParsesExtendedStatisticFields()
        {
            byte[] docxData = CreateDocxWithProperties();
            using var ms = new MemoryStream(docxData);
            using var reader = new Nedev.FileConverters.DocxToDoc.Format.DocxReader(ms);

            var model = reader.ReadDocument();

            Assert.Equal(42, model.Properties.TotalEditingTime);
            Assert.Equal(7, model.Properties.Pages);
            Assert.Equal(123, model.Properties.Words);
            Assert.Equal(456, model.Properties.Characters);
        }

        [Fact]
        public void ReadDocument_WithProperties_ParsesTotalTimeFromExtendedProperties()
        {
            byte[] docxData = CreateDocxWithProperties();
            using var ms = new MemoryStream(docxData);
            using var reader = new Nedev.FileConverters.DocxToDoc.Format.DocxReader(ms);

            var model = reader.ReadDocument();

            Assert.Equal(42, model.Properties.TotalEditingTime);
        }
    }
}
