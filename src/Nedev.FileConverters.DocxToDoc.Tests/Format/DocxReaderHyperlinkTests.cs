using System;
using System.IO;
using System.IO.Compression;
using System.Text;
using Xunit;

namespace Nedev.FileConverters.DocxToDoc.Tests.Format
{
    public class DocxReaderHyperlinkTests
    {
        private byte[] CreateDocxWithHyperlink()
        {
            using var ms = new MemoryStream();
            using (var archive = new ZipArchive(ms, ZipArchiveMode.Create, true))
            {
                // Create word/_rels/document.xml.rels
                var docRelsEntry = archive.CreateEntry("word/_rels/document.xml.rels");
                using (var stream = docRelsEntry.Open())
                using (var writer = new StreamWriter(stream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                        "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
                        "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink\" Target=\"https://example.com\" TargetMode=\"External\"/>" +
                        "</Relationships>");
                }

                // Create word/document.xml with hyperlink
                var docEntry = archive.CreateEntry("word/document.xml");
                using (var stream = docEntry.Open())
                using (var writer = new StreamWriter(stream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                        "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" " +
                        "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">" +
                        "<w:body>" +
                        "<w:p>" +
                        "<w:r>" +
                        "<w:t>Visit </w:t>" +
                        "</w:r>" +
                        "<w:hyperlink r:id=\"rId1\" w:tooltip=\"Click to visit\">" +
                        "<w:r>" +
                        "<w:rPr>" +
                        "<w:rFonts w:hAnsi=\"Calibri\"/>" +
                        "<w:color w:val=\"0000FF\"/>" +
                        "<w:u w:val=\"single\"/>" +
                        "<w:sz w:val=\"28\"/>" +
                        "</w:rPr>" +
                        "<w:t>Example Website</w:t>" +
                        "</w:r>" +
                        "</w:hyperlink>" +
                        "<w:r>" +
                        "<w:t> for more info.</w:t>" +
                        "</w:r>" +
                        "</w:p>" +
                        "</w:body>" +
                        "</w:document>");
                }
            }
            return ms.ToArray();
        }

        [Fact]
        public void ReadDocument_WithHyperlink_ParsesHyperlinkData()
        {
            // Arrange
            byte[] docxData = CreateDocxWithHyperlink();
            using var ms = new MemoryStream(docxData);
            using var reader = new Nedev.FileConverters.DocxToDoc.Format.DocxReader(ms);

            // Act
            var model = reader.ReadDocument();

            // Assert
            Assert.Single(model.Paragraphs);
            // Paragraph should have runs from hyperlink
            Assert.True(model.Paragraphs[0].Runs.Count >= 1);

            // Find the hyperlink run
            var hyperlinkRun = model.Paragraphs[0].Runs.Find(r => r.Hyperlink != null);
            Assert.NotNull(hyperlinkRun);
            Assert.NotNull(hyperlinkRun.Hyperlink);
            Assert.Equal("rId1", hyperlinkRun.Hyperlink.RelationshipId);
            Assert.Equal("https://example.com", hyperlinkRun.Hyperlink.TargetUrl);
            Assert.Equal("Click to visit", hyperlinkRun.Hyperlink.Tooltip);
            // DisplayText may include more text depending on parsing
            Assert.Contains("Example Website", hyperlinkRun.Hyperlink.DisplayText);
        }

        [Fact]
        public void ReadDocument_WithHyperlink_ParsesFormatting()
        {
            // Arrange
            byte[] docxData = CreateDocxWithHyperlink();
            using var ms = new MemoryStream(docxData);
            using var reader = new Nedev.FileConverters.DocxToDoc.Format.DocxReader(ms);

            // Act
            var model = reader.ReadDocument();

            // Assert
            var hyperlinkRun = model.Paragraphs[0].Runs.Find(r => r.Hyperlink != null);
            Assert.NotNull(hyperlinkRun);
            Assert.NotNull(hyperlinkRun.Hyperlink);
            Assert.Equal("0000FF", hyperlinkRun.Properties.Color);
            Assert.Equal(Nedev.FileConverters.DocxToDoc.Model.UnderlineType.Single, hyperlinkRun.Properties.Underline);
            Assert.Equal(28, hyperlinkRun.Properties.FontSize);
            Assert.Equal("Calibri", hyperlinkRun.Properties.FontName);
        }
    }
}
