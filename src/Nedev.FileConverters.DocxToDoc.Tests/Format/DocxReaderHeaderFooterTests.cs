using System.IO;
using System.IO.Compression;
using System.Text;
using Xunit;
using Nedev.FileConverters.DocxToDoc.Model;

namespace Nedev.FileConverters.DocxToDoc.Tests.Format
{
    public class DocxReaderHeaderFooterTests
    {
        [Fact]
        public void ReadDocument_WithDefaultHeaderFooterReferences_ParsesPlainTextStories()
        {
            using var ms = new MemoryStream();
            using (var archive = new ZipArchive(ms, ZipArchiveMode.Create, true))
            {
                var documentEntry = archive.CreateEntry("word/document.xml");
                using (var documentStream = documentEntry.Open())
                using (var writer = new StreamWriter(documentStream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" +
                                 "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">" +
                                 "<w:body>" +
                                 "<w:p><w:r><w:t>AB</w:t></w:r></w:p>" +
                                 "<w:sectPr>" +
                                 "<w:headerReference w:type=\"default\" r:id=\"rIdHeader\"/>" +
                                 "<w:footerReference w:type=\"default\" r:id=\"rIdFooter\"/>" +
                                 "</w:sectPr>" +
                                 "</w:body>" +
                                 "</w:document>");
                }

                var relsEntry = archive.CreateEntry("word/_rels/document.xml.rels");
                using (var relsStream = relsEntry.Open())
                using (var writer = new StreamWriter(relsStream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" +
                                 "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
                                 "<Relationship Id=\"rIdHeader\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/header\" Target=\"header1.xml\" />" +
                                 "<Relationship Id=\"rIdFooter\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer\" Target=\"footer1.xml\" />" +
                                 "</Relationships>");
                }

                var headerEntry = archive.CreateEntry("word/header1.xml");
                using (var headerStream = headerEntry.Open())
                using (var writer = new StreamWriter(headerStream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" +
                                 "<w:hdr xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">" +
                                 "<w:p><w:r><w:t>Head</w:t></w:r></w:p>" +
                                 "<w:p><w:r><w:t>Line</w:t></w:r></w:p>" +
                                 "</w:hdr>");
                }

                var footerEntry = archive.CreateEntry("word/footer1.xml");
                using (var footerStream = footerEntry.Open())
                using (var writer = new StreamWriter(footerStream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" +
                                 "<w:ftr xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">" +
                                 "<w:p><w:r><w:t>Foot</w:t></w:r></w:p>" +
                                 "</w:ftr>");
                }
            }

            using var testStream = new MemoryStream(ms.ToArray());
            using var reader = new Nedev.FileConverters.DocxToDoc.Format.DocxReader(testStream);

            var model = reader.ReadDocument();

            var section = Assert.Single(model.Sections);
            Assert.Equal("rIdHeader", section.HeaderReference);
            Assert.Equal("rIdFooter", section.FooterReference);
            Assert.Equal("Head\rLine", section.DefaultHeaderText);
            Assert.Equal("Foot", section.DefaultFooterText);
            Assert.Equal("Head\rLine", section.HeaderText);
            Assert.Equal("Foot", section.FooterText);
            Assert.Equal("AB\r", model.TextBuffer);
        }

        [Fact]
        public void ReadDocument_WithoutDefaultHeaderFooterReferences_FallsBackToFirstAndEvenStories()
        {
            using var ms = new MemoryStream();
            using (var archive = new ZipArchive(ms, ZipArchiveMode.Create, true))
            {
                var documentEntry = archive.CreateEntry("word/document.xml");
                using (var documentStream = documentEntry.Open())
                using (var writer = new StreamWriter(documentStream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" +
                                 "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">" +
                                 "<w:body>" +
                                 "<w:p><w:r><w:t>AB</w:t></w:r></w:p>" +
                                 "<w:sectPr>" +
                                 "<w:headerReference w:type=\"first\" r:id=\"rIdHeaderFirst\"/>" +
                                 "<w:footerReference w:type=\"even\" r:id=\"rIdFooterEven\"/>" +
                                 "</w:sectPr>" +
                                 "</w:body>" +
                                 "</w:document>");
                }

                var relsEntry = archive.CreateEntry("word/_rels/document.xml.rels");
                using (var relsStream = relsEntry.Open())
                using (var writer = new StreamWriter(relsStream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" +
                                 "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
                                 "<Relationship Id=\"rIdHeaderFirst\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/header\" Target=\"header-first.xml\" />" +
                                 "<Relationship Id=\"rIdFooterEven\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer\" Target=\"footer-even.xml\" />" +
                                 "</Relationships>");
                }

                var settingsEntry = archive.CreateEntry("word/settings.xml");
                using (var settingsStream = settingsEntry.Open())
                using (var writer = new StreamWriter(settingsStream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" +
                                 "<w:settings xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">" +
                                 "<w:evenAndOddHeaders/>" +
                                 "</w:settings>");
                }

                var headerEntry = archive.CreateEntry("word/header-first.xml");
                using (var headerStream = headerEntry.Open())
                using (var writer = new StreamWriter(headerStream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" +
                                 "<w:hdr xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">" +
                                 "<w:p><w:r><w:t>First head</w:t></w:r></w:p>" +
                                 "</w:hdr>");
                }

                var footerEntry = archive.CreateEntry("word/footer-even.xml");
                using (var footerStream = footerEntry.Open())
                using (var writer = new StreamWriter(footerStream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" +
                                 "<w:ftr xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">" +
                                 "<w:p><w:r><w:t>Even foot</w:t></w:r></w:p>" +
                                 "</w:ftr>");
                }
            }

            using var testStream = new MemoryStream(ms.ToArray());
            using var reader = new Nedev.FileConverters.DocxToDoc.Format.DocxReader(testStream);

            var model = reader.ReadDocument();

            var section = Assert.Single(model.Sections);
            Assert.Null(section.HeaderReference);
            Assert.Null(section.FooterReference);
            Assert.Equal("First head", section.FirstPageHeaderText);
            Assert.Equal("Even foot", section.EvenPagesFooterText);
            Assert.Equal("First head", section.HeaderText);
            Assert.Equal("Even foot", section.FooterText);
        }

        [Fact]
        public void ReadDocument_WithMissingDefaultHeaderPart_FallsBackToAlternateStory()
        {
            using var ms = new MemoryStream();
            using (var archive = new ZipArchive(ms, ZipArchiveMode.Create, true))
            {
                var documentEntry = archive.CreateEntry("word/document.xml");
                using (var documentStream = documentEntry.Open())
                using (var writer = new StreamWriter(documentStream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" +
                                 "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">" +
                                 "<w:body>" +
                                 "<w:p><w:r><w:t>AB</w:t></w:r></w:p>" +
                                 "<w:sectPr>" +
                                 "<w:headerReference w:type=\"default\" r:id=\"rIdHeaderMissing\"/>" +
                                 "<w:headerReference w:type=\"first\" r:id=\"rIdHeaderFirst\"/>" +
                                 "<w:footerReference w:type=\"default\" r:id=\"rIdFooter\"/>" +
                                 "</w:sectPr>" +
                                 "</w:body>" +
                                 "</w:document>");
                }

                var relsEntry = archive.CreateEntry("word/_rels/document.xml.rels");
                using (var relsStream = relsEntry.Open())
                using (var writer = new StreamWriter(relsStream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" +
                                 "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
                                 "<Relationship Id=\"rIdHeaderMissing\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/header\" Target=\"header-missing.xml\" />" +
                                 "<Relationship Id=\"rIdHeaderFirst\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/header\" Target=\"header-first.xml\" />" +
                                 "<Relationship Id=\"rIdFooter\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer\" Target=\"footer1.xml\" />" +
                                 "</Relationships>");
                }

                var headerEntry = archive.CreateEntry("word/header-first.xml");
                using (var headerStream = headerEntry.Open())
                using (var writer = new StreamWriter(headerStream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" +
                                 "<w:hdr xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">" +
                                 "<w:p><w:r><w:t>Fallback head</w:t></w:r></w:p>" +
                                 "</w:hdr>");
                }

                var footerEntry = archive.CreateEntry("word/footer1.xml");
                using (var footerStream = footerEntry.Open())
                using (var writer = new StreamWriter(footerStream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" +
                                 "<w:ftr xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">" +
                                 "<w:p><w:r><w:t>Foot</w:t></w:r></w:p>" +
                                 "</w:ftr>");
                }
            }

            using var testStream = new MemoryStream(ms.ToArray());
            using var reader = new Nedev.FileConverters.DocxToDoc.Format.DocxReader(testStream);

            var model = reader.ReadDocument();

            var section = Assert.Single(model.Sections);
            Assert.Null(section.DefaultHeaderText);
            Assert.Equal("Fallback head", section.FirstPageHeaderText);
            Assert.Equal("Fallback head", section.HeaderText);
            Assert.Equal("Foot", section.FooterText);
        }

        [Fact]
        public void ReadDocument_WithTitlePageAndDedicatedFirstStories_PreservesDefaultAndFirstPageTexts()
        {
            using var ms = new MemoryStream();
            using (var archive = new ZipArchive(ms, ZipArchiveMode.Create, true))
            {
                var documentEntry = archive.CreateEntry("word/document.xml");
                using (var documentStream = documentEntry.Open())
                using (var writer = new StreamWriter(documentStream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" +
                                 "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">" +
                                 "<w:body>" +
                                 "<w:p><w:r><w:t>AB</w:t></w:r></w:p>" +
                                 "<w:sectPr>" +
                                 "<w:headerReference w:type=\"default\" r:id=\"rIdHeader\"/>" +
                                 "<w:footerReference w:type=\"default\" r:id=\"rIdFooter\"/>" +
                                 "<w:headerReference w:type=\"first\" r:id=\"rIdHeaderFirst\"/>" +
                                 "<w:footerReference w:type=\"first\" r:id=\"rIdFooterFirst\"/>" +
                                 "<w:titlePg/>" +
                                 "</w:sectPr>" +
                                 "</w:body>" +
                                 "</w:document>");
                }

                var relsEntry = archive.CreateEntry("word/_rels/document.xml.rels");
                using (var relsStream = relsEntry.Open())
                using (var writer = new StreamWriter(relsStream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" +
                                 "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
                                 "<Relationship Id=\"rIdHeader\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/header\" Target=\"header-default.xml\" />" +
                                 "<Relationship Id=\"rIdFooter\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer\" Target=\"footer-default.xml\" />" +
                                 "<Relationship Id=\"rIdHeaderFirst\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/header\" Target=\"header-first.xml\" />" +
                                 "<Relationship Id=\"rIdFooterFirst\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer\" Target=\"footer-first.xml\" />" +
                                 "</Relationships>");
                }

                var headerDefaultEntry = archive.CreateEntry("word/header-default.xml");
                using (var headerStream = headerDefaultEntry.Open())
                using (var writer = new StreamWriter(headerStream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" +
                                 "<w:hdr xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">" +
                                 "<w:p><w:r><w:t>Default head</w:t></w:r></w:p>" +
                                 "</w:hdr>");
                }

                var footerDefaultEntry = archive.CreateEntry("word/footer-default.xml");
                using (var footerStream = footerDefaultEntry.Open())
                using (var writer = new StreamWriter(footerStream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" +
                                 "<w:ftr xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">" +
                                 "<w:p><w:r><w:t>Default foot</w:t></w:r></w:p>" +
                                 "</w:ftr>");
                }

                var headerFirstEntry = archive.CreateEntry("word/header-first.xml");
                using (var headerStream = headerFirstEntry.Open())
                using (var writer = new StreamWriter(headerStream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" +
                                 "<w:hdr xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">" +
                                 "<w:p><w:r><w:t>First head</w:t></w:r></w:p>" +
                                 "</w:hdr>");
                }

                var footerFirstEntry = archive.CreateEntry("word/footer-first.xml");
                using (var footerStream = footerFirstEntry.Open())
                using (var writer = new StreamWriter(footerStream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" +
                                 "<w:ftr xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">" +
                                 "<w:p><w:r><w:t>First foot</w:t></w:r></w:p>" +
                                 "</w:ftr>");
                }
            }

            using var testStream = new MemoryStream(ms.ToArray());
            using var reader = new Nedev.FileConverters.DocxToDoc.Format.DocxReader(testStream);

            var model = reader.ReadDocument();

            var section = Assert.Single(model.Sections);
            Assert.True(section.DifferentFirstPage);
            Assert.Equal("Default head", section.DefaultHeaderText);
            Assert.Equal("Default foot", section.DefaultFooterText);
            Assert.Equal("First head", section.FirstPageHeaderText);
            Assert.Equal("First foot", section.FirstPageFooterText);
            Assert.Equal("Default head", section.HeaderText);
            Assert.Equal("Default foot", section.FooterText);
        }

        [Fact]
        public void ReadDocument_WithDefaultAndEvenStories_PreservesDedicatedEvenPageTexts()
        {
            using var ms = new MemoryStream();
            using (var archive = new ZipArchive(ms, ZipArchiveMode.Create, true))
            {
                var documentEntry = archive.CreateEntry("word/document.xml");
                using (var documentStream = documentEntry.Open())
                using (var writer = new StreamWriter(documentStream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" +
                                 "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">" +
                                 "<w:body>" +
                                 "<w:p><w:r><w:t>AB</w:t></w:r></w:p>" +
                                 "<w:sectPr>" +
                                 "<w:headerReference w:type=\"default\" r:id=\"rIdHeader\"/>" +
                                 "<w:footerReference w:type=\"default\" r:id=\"rIdFooter\"/>" +
                                 "<w:headerReference w:type=\"even\" r:id=\"rIdHeaderEven\"/>" +
                                 "<w:footerReference w:type=\"even\" r:id=\"rIdFooterEven\"/>" +
                                 "</w:sectPr>" +
                                 "</w:body>" +
                                 "</w:document>");
                }

                var relsEntry = archive.CreateEntry("word/_rels/document.xml.rels");
                using (var relsStream = relsEntry.Open())
                using (var writer = new StreamWriter(relsStream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" +
                                 "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
                                 "<Relationship Id=\"rIdHeader\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/header\" Target=\"header-default.xml\" />" +
                                 "<Relationship Id=\"rIdFooter\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer\" Target=\"footer-default.xml\" />" +
                                 "<Relationship Id=\"rIdHeaderEven\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/header\" Target=\"header-even.xml\" />" +
                                 "<Relationship Id=\"rIdFooterEven\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer\" Target=\"footer-even.xml\" />" +
                                 "</Relationships>");
                }

                var settingsEntry = archive.CreateEntry("word/settings.xml");
                using (var settingsStream = settingsEntry.Open())
                using (var writer = new StreamWriter(settingsStream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" +
                                 "<w:settings xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">" +
                                 "<w:evenAndOddHeaders/>" +
                                 "</w:settings>");
                }

                var headerDefaultEntry = archive.CreateEntry("word/header-default.xml");
                using (var headerStream = headerDefaultEntry.Open())
                using (var writer = new StreamWriter(headerStream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" +
                                 "<w:hdr xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">" +
                                 "<w:p><w:r><w:t>Default head</w:t></w:r></w:p>" +
                                 "</w:hdr>");
                }

                var footerDefaultEntry = archive.CreateEntry("word/footer-default.xml");
                using (var footerStream = footerDefaultEntry.Open())
                using (var writer = new StreamWriter(footerStream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" +
                                 "<w:ftr xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">" +
                                 "<w:p><w:r><w:t>Default foot</w:t></w:r></w:p>" +
                                 "</w:ftr>");
                }

                var headerEvenEntry = archive.CreateEntry("word/header-even.xml");
                using (var headerStream = headerEvenEntry.Open())
                using (var writer = new StreamWriter(headerStream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" +
                                 "<w:hdr xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">" +
                                 "<w:p><w:r><w:t>Even head</w:t></w:r></w:p>" +
                                 "</w:hdr>");
                }

                var footerEvenEntry = archive.CreateEntry("word/footer-even.xml");
                using (var footerStream = footerEvenEntry.Open())
                using (var writer = new StreamWriter(footerStream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" +
                                 "<w:ftr xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">" +
                                 "<w:p><w:r><w:t>Even foot</w:t></w:r></w:p>" +
                                 "</w:ftr>");
                }
            }

            using var testStream = new MemoryStream(ms.ToArray());
            using var reader = new Nedev.FileConverters.DocxToDoc.Format.DocxReader(testStream);

            var model = reader.ReadDocument();

            var section = Assert.Single(model.Sections);
            Assert.True(model.DifferentOddAndEvenPages);
            Assert.Equal("Default head", section.DefaultHeaderText);
            Assert.Equal("Default foot", section.DefaultFooterText);
            Assert.Equal("Even head", section.EvenPagesHeaderText);
            Assert.Equal("Even foot", section.EvenPagesFooterText);
            Assert.Equal("Default head", section.HeaderText);
            Assert.Equal("Default foot", section.FooterText);
        }

        [Fact]
        public void ReadDocument_WithoutEvenAndOddHeadersSetting_DoesNotPromoteEvenStoriesToSummaryText()
        {
            using var ms = new MemoryStream();
            using (var archive = new ZipArchive(ms, ZipArchiveMode.Create, true))
            {
                var documentEntry = archive.CreateEntry("word/document.xml");
                using (var documentStream = documentEntry.Open())
                using (var writer = new StreamWriter(documentStream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" +
                                 "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">" +
                                 "<w:body>" +
                                 "<w:p><w:r><w:t>AB</w:t></w:r></w:p>" +
                                 "<w:sectPr>" +
                                 "<w:headerReference w:type=\"even\" r:id=\"rIdHeaderEven\"/>" +
                                 "<w:footerReference w:type=\"even\" r:id=\"rIdFooterEven\"/>" +
                                 "</w:sectPr>" +
                                 "</w:body>" +
                                 "</w:document>");
                }

                var relsEntry = archive.CreateEntry("word/_rels/document.xml.rels");
                using (var relsStream = relsEntry.Open())
                using (var writer = new StreamWriter(relsStream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" +
                                 "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
                                 "<Relationship Id=\"rIdHeaderEven\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/header\" Target=\"header-even.xml\" />" +
                                 "<Relationship Id=\"rIdFooterEven\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer\" Target=\"footer-even.xml\" />" +
                                 "</Relationships>");
                }

                var headerEvenEntry = archive.CreateEntry("word/header-even.xml");
                using (var headerStream = headerEvenEntry.Open())
                using (var writer = new StreamWriter(headerStream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" +
                                 "<w:hdr xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">" +
                                 "<w:p><w:r><w:t>Even head</w:t></w:r></w:p>" +
                                 "</w:hdr>");
                }

                var footerEvenEntry = archive.CreateEntry("word/footer-even.xml");
                using (var footerStream = footerEvenEntry.Open())
                using (var writer = new StreamWriter(footerStream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" +
                                 "<w:ftr xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">" +
                                 "<w:p><w:r><w:t>Even foot</w:t></w:r></w:p>" +
                                 "</w:ftr>");
                }
            }

            using var testStream = new MemoryStream(ms.ToArray());
            using var reader = new Nedev.FileConverters.DocxToDoc.Format.DocxReader(testStream);

            var model = reader.ReadDocument();

            var section = Assert.Single(model.Sections);
            Assert.False(model.DifferentOddAndEvenPages);
            Assert.Equal("Even head", section.EvenPagesHeaderText);
            Assert.Equal("Even foot", section.EvenPagesFooterText);
            Assert.Null(section.HeaderText);
            Assert.Null(section.FooterText);
        }

        [Fact]
        public void ReadDocument_WithEvenAndOddHeadersSetting_PromotesEvenStoriesToSummaryText()
        {
            using var ms = new MemoryStream();
            using (var archive = new ZipArchive(ms, ZipArchiveMode.Create, true))
            {
                var documentEntry = archive.CreateEntry("word/document.xml");
                using (var documentStream = documentEntry.Open())
                using (var writer = new StreamWriter(documentStream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" +
                                 "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">" +
                                 "<w:body>" +
                                 "<w:p><w:r><w:t>AB</w:t></w:r></w:p>" +
                                 "<w:sectPr>" +
                                 "<w:headerReference w:type=\"even\" r:id=\"rIdHeaderEven\"/>" +
                                 "<w:footerReference w:type=\"even\" r:id=\"rIdFooterEven\"/>" +
                                 "</w:sectPr>" +
                                 "</w:body>" +
                                 "</w:document>");
                }

                var relsEntry = archive.CreateEntry("word/_rels/document.xml.rels");
                using (var relsStream = relsEntry.Open())
                using (var writer = new StreamWriter(relsStream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" +
                                 "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
                                 "<Relationship Id=\"rIdHeaderEven\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/header\" Target=\"header-even.xml\" />" +
                                 "<Relationship Id=\"rIdFooterEven\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer\" Target=\"footer-even.xml\" />" +
                                 "</Relationships>");
                }

                var settingsEntry = archive.CreateEntry("word/settings.xml");
                using (var settingsStream = settingsEntry.Open())
                using (var writer = new StreamWriter(settingsStream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" +
                                 "<w:settings xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">" +
                                 "<w:evenAndOddHeaders/>" +
                                 "</w:settings>");
                }

                var headerEvenEntry = archive.CreateEntry("word/header-even.xml");
                using (var headerStream = headerEvenEntry.Open())
                using (var writer = new StreamWriter(headerStream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" +
                                 "<w:hdr xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">" +
                                 "<w:p><w:r><w:t>Even head</w:t></w:r></w:p>" +
                                 "</w:hdr>");
                }

                var footerEvenEntry = archive.CreateEntry("word/footer-even.xml");
                using (var footerStream = footerEvenEntry.Open())
                using (var writer = new StreamWriter(footerStream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" +
                                 "<w:ftr xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">" +
                                 "<w:p><w:r><w:t>Even foot</w:t></w:r></w:p>" +
                                 "</w:ftr>");
                }
            }

            using var testStream = new MemoryStream(ms.ToArray());
            using var reader = new Nedev.FileConverters.DocxToDoc.Format.DocxReader(testStream);

            var model = reader.ReadDocument();

            var section = Assert.Single(model.Sections);
            Assert.True(model.DifferentOddAndEvenPages);
            Assert.Equal("Even head", section.EvenPagesHeaderText);
            Assert.Equal("Even foot", section.EvenPagesFooterText);
            Assert.Equal("Even head", section.HeaderText);
            Assert.Equal("Even foot", section.FooterText);
        }

        [Fact]
        public void ReadDocument_WithHeaderFieldStory_ParsesStructuredFieldAwareStory()
        {
            using var ms = new MemoryStream();
            using (var archive = new ZipArchive(ms, ZipArchiveMode.Create, true))
            {
                var documentEntry = archive.CreateEntry("word/document.xml");
                using (var documentStream = documentEntry.Open())
                using (var writer = new StreamWriter(documentStream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" +
                                 "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">" +
                                 "<w:body>" +
                                 "<w:p><w:r><w:t>AB</w:t></w:r></w:p>" +
                                 "<w:sectPr>" +
                                 "<w:headerReference w:type=\"default\" r:id=\"rIdHeader\"/>" +
                                 "</w:sectPr>" +
                                 "</w:body>" +
                                 "</w:document>");
                }

                var relsEntry = archive.CreateEntry("word/_rels/document.xml.rels");
                using (var relsStream = relsEntry.Open())
                using (var writer = new StreamWriter(relsStream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" +
                                 "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
                                 "<Relationship Id=\"rIdHeader\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/header\" Target=\"header1.xml\" />" +
                                 "</Relationships>");
                }

                var headerEntry = archive.CreateEntry("word/header1.xml");
                using (var headerStream = headerEntry.Open())
                using (var writer = new StreamWriter(headerStream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" +
                                 "<w:hdr xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">" +
                                 "<w:p>" +
                                 "<w:r><w:t>Page </w:t></w:r>" +
                                 "<w:fldSimple w:instr=\"PAGE\"><w:r><w:t>1</w:t></w:r></w:fldSimple>" +
                                 "<w:r><w:t> of </w:t></w:r>" +
                                 "<w:fldSimple w:instr=\"SECTIONPAGES\"><w:r><w:t>2</w:t></w:r></w:fldSimple>" +
                                 "</w:p>" +
                                 "</w:hdr>");
                }
            }

            using var testStream = new MemoryStream(ms.ToArray());
            using var reader = new Nedev.FileConverters.DocxToDoc.Format.DocxReader(testStream);

            var model = reader.ReadDocument();

            var section = Assert.Single(model.Sections);
            Assert.Equal("Page 1 of 2", section.DefaultHeaderText);

            Assert.NotNull(section.DefaultHeaderStory);
            var story = section.DefaultHeaderStory!;
            var paragraph = Assert.Single(story.Paragraphs);
            var significantRuns = paragraph.Runs.FindAll(run =>
                run.Text.Length > 0 ||
                run.Field != null ||
                run.IsFieldBegin ||
                run.IsFieldSeparate ||
                run.IsFieldEnd);

            Assert.Equal("Page ", significantRuns[0].Text);
            Assert.True(significantRuns[1].IsFieldBegin);
            Assert.Equal(FieldType.Page, significantRuns[1].Field?.Type);
            Assert.True(significantRuns[2].IsFieldSeparate);
            Assert.Equal(FieldType.Page, significantRuns[2].Field?.Type);
            Assert.Equal("1", significantRuns[3].Text);
            Assert.True(significantRuns[4].IsFieldEnd);
            Assert.Equal(FieldType.Page, significantRuns[4].Field?.Type);
            Assert.Equal(" of ", significantRuns[5].Text);
            Assert.True(significantRuns[6].IsFieldBegin);
            Assert.Equal(FieldType.SectionPages, significantRuns[6].Field?.Type);
            Assert.True(significantRuns[7].IsFieldSeparate);
            Assert.Equal(FieldType.SectionPages, significantRuns[7].Field?.Type);
            Assert.Equal("2", significantRuns[8].Text);
            Assert.True(significantRuns[9].IsFieldEnd);
            Assert.Equal(FieldType.SectionPages, significantRuns[9].Field?.Type);
        }

        [Fact]
        public void ReadDocument_WithHeaderHyperlinkStory_ParsesStructuredHyperlinkStory()
        {
            using var ms = new MemoryStream();
            using (var archive = new ZipArchive(ms, ZipArchiveMode.Create, true))
            {
                var documentEntry = archive.CreateEntry("word/document.xml");
                using (var documentStream = documentEntry.Open())
                using (var writer = new StreamWriter(documentStream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" +
                                 "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">" +
                                 "<w:body>" +
                                 "<w:p><w:r><w:t>AB</w:t></w:r></w:p>" +
                                 "<w:sectPr>" +
                                 "<w:headerReference w:type=\"default\" r:id=\"rIdHeader\"/>" +
                                 "</w:sectPr>" +
                                 "</w:body>" +
                                 "</w:document>");
                }

                var relsEntry = archive.CreateEntry("word/_rels/document.xml.rels");
                using (var relsStream = relsEntry.Open())
                using (var writer = new StreamWriter(relsStream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" +
                                 "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
                                 "<Relationship Id=\"rIdHeader\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/header\" Target=\"header1.xml\" />" +
                                 "</Relationships>");
                }

                var headerEntry = archive.CreateEntry("word/header1.xml");
                using (var headerStream = headerEntry.Open())
                using (var writer = new StreamWriter(headerStream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" +
                                 "<w:hdr xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">" +
                                 "<w:p>" +
                                 "<w:r><w:t>Go </w:t></w:r>" +
                                 "<w:hyperlink r:id=\"rIdLink\" w:tooltip=\"tip\">" +
                                 "<w:r><w:t>Ex</w:t></w:r>" +
                                 "<w:r><w:t>ample</w:t></w:r>" +
                                 "</w:hyperlink>" +
                                 "<w:r><w:t> now</w:t></w:r>" +
                                 "</w:p>" +
                                 "</w:hdr>");
                }

                var headerRelsEntry = archive.CreateEntry("word/_rels/header1.xml.rels");
                using (var headerRelsStream = headerRelsEntry.Open())
                using (var writer = new StreamWriter(headerRelsStream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" +
                                 "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
                                 "<Relationship Id=\"rIdLink\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink\" Target=\"https://example.com\" TargetMode=\"External\" />" +
                                 "</Relationships>");
                }
            }

            using var testStream = new MemoryStream(ms.ToArray());
            using var reader = new Nedev.FileConverters.DocxToDoc.Format.DocxReader(testStream);

            var model = reader.ReadDocument();

            var section = Assert.Single(model.Sections);
            Assert.Equal("Go Example now", section.DefaultHeaderText);

            Assert.NotNull(section.DefaultHeaderStory);
            var story = section.DefaultHeaderStory!;
            var paragraph = Assert.Single(story.Paragraphs);
            var significantRuns = paragraph.Runs.FindAll(run => run.Text.Length > 0);

            Assert.Equal("Go ", significantRuns[0].Text);
            Assert.Equal("Ex", significantRuns[1].Text);
            Assert.Equal("ample", significantRuns[2].Text);
            Assert.Equal(" now", significantRuns[3].Text);
            Assert.NotNull(significantRuns[1].Hyperlink);
            Assert.Same(significantRuns[1].Hyperlink, significantRuns[2].Hyperlink);
            Assert.Equal("Example", significantRuns[1].Hyperlink!.DisplayText);
            Assert.Equal("https://example.com", significantRuns[1].Hyperlink!.TargetUrl);
            Assert.Equal("tip", significantRuns[1].Hyperlink!.Tooltip);
        }

        [Fact]
        public void ReadDocument_WithMixedHeaderFieldAndHyperlinkStory_PreservesSummaryAndStructuredRuns()
        {
            using var ms = new MemoryStream();
            using (var archive = new ZipArchive(ms, ZipArchiveMode.Create, true))
            {
                var documentEntry = archive.CreateEntry("word/document.xml");
                using (var documentStream = documentEntry.Open())
                using (var writer = new StreamWriter(documentStream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" +
                                 "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">" +
                                 "<w:body>" +
                                 "<w:p><w:r><w:t>AB</w:t></w:r></w:p>" +
                                 "<w:sectPr>" +
                                 "<w:headerReference w:type=\"default\" r:id=\"rIdHeader\"/>" +
                                 "</w:sectPr>" +
                                 "</w:body>" +
                                 "</w:document>");
                }

                var relsEntry = archive.CreateEntry("word/_rels/document.xml.rels");
                using (var relsStream = relsEntry.Open())
                using (var writer = new StreamWriter(relsStream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" +
                                 "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
                                 "<Relationship Id=\"rIdHeader\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/header\" Target=\"header1.xml\" />" +
                                 "</Relationships>");
                }

                var headerEntry = archive.CreateEntry("word/header1.xml");
                using (var headerStream = headerEntry.Open())
                using (var writer = new StreamWriter(headerStream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" +
                                 "<w:hdr xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">" +
                                 "<w:p>" +
                                 "<w:r><w:t>Page </w:t></w:r>" +
                                 "<w:fldSimple w:instr=\"PAGE\"><w:r><w:t>1</w:t></w:r></w:fldSimple>" +
                                 "<w:r><w:t> | </w:t></w:r>" +
                                 "<w:hyperlink r:id=\"rIdLink\">" +
                                 "<w:r><w:t>Ex</w:t></w:r>" +
                                 "<w:r><w:t>ample</w:t></w:r>" +
                                 "</w:hyperlink>" +
                                 "</w:p>" +
                                 "</w:hdr>");
                }

                var headerRelsEntry = archive.CreateEntry("word/_rels/header1.xml.rels");
                using (var headerRelsStream = headerRelsEntry.Open())
                using (var writer = new StreamWriter(headerRelsStream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" +
                                 "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
                                 "<Relationship Id=\"rIdLink\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink\" Target=\"https://example.com\" TargetMode=\"External\" />" +
                                 "</Relationships>");
                }
            }

            using var testStream = new MemoryStream(ms.ToArray());
            using var reader = new Nedev.FileConverters.DocxToDoc.Format.DocxReader(testStream);

            var model = reader.ReadDocument();

            var section = Assert.Single(model.Sections);
            Assert.Equal("Page 1 | Example", section.DefaultHeaderText);
            Assert.Equal("Page 1 | Example", section.HeaderText);

            Assert.NotNull(section.DefaultHeaderStory);
            var story = section.DefaultHeaderStory!;
            var paragraph = Assert.Single(story.Paragraphs);
            var significantRuns = paragraph.Runs.FindAll(run =>
                run.Text.Length > 0 ||
                run.Field != null ||
                run.IsFieldBegin ||
                run.IsFieldSeparate ||
                run.IsFieldEnd);

            Assert.Equal("Page ", significantRuns[0].Text);
            Assert.True(significantRuns[1].IsFieldBegin);
            Assert.Equal(FieldType.Page, significantRuns[1].Field?.Type);
            Assert.True(significantRuns[2].IsFieldSeparate);
            Assert.Equal(FieldType.Page, significantRuns[2].Field?.Type);
            Assert.Equal("1", significantRuns[3].Text);
            Assert.True(significantRuns[4].IsFieldEnd);
            Assert.Equal(FieldType.Page, significantRuns[4].Field?.Type);
            Assert.Equal(" | ", significantRuns[5].Text);
            Assert.Equal("Ex", significantRuns[6].Text);
            Assert.Equal("ample", significantRuns[7].Text);
            Assert.NotNull(significantRuns[6].Hyperlink);
            Assert.Same(significantRuns[6].Hyperlink, significantRuns[7].Hyperlink);
            Assert.Equal("Example", significantRuns[6].Hyperlink!.DisplayText);
            Assert.Equal("https://example.com", significantRuns[6].Hyperlink!.TargetUrl);
        }

        [Fact]
        public void ReadDocument_WithHeaderInlineImageStory_ParsesStructuredImageRun()
        {
            using var ms = new MemoryStream();
            using (var archive = new ZipArchive(ms, ZipArchiveMode.Create, true))
            {
                var documentEntry = archive.CreateEntry("word/document.xml");
                using (var documentStream = documentEntry.Open())
                using (var writer = new StreamWriter(documentStream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" +
                                 "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">" +
                                 "<w:body>" +
                                 "<w:p><w:r><w:t>AB</w:t></w:r></w:p>" +
                                 "<w:sectPr>" +
                                 "<w:headerReference w:type=\"default\" r:id=\"rIdHeader\"/>" +
                                 "</w:sectPr>" +
                                 "</w:body>" +
                                 "</w:document>");
                }

                var relsEntry = archive.CreateEntry("word/_rels/document.xml.rels");
                using (var relsStream = relsEntry.Open())
                using (var writer = new StreamWriter(relsStream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" +
                                 "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
                                 "<Relationship Id=\"rIdHeader\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/header\" Target=\"header1.xml\" />" +
                                 "</Relationships>");
                }

                var headerEntry = archive.CreateEntry("word/header1.xml");
                using (var headerStream = headerEntry.Open())
                using (var writer = new StreamWriter(headerStream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" +
                                 "<w:hdr xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" " +
                                 "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" " +
                                 "xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" " +
                                 "xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\" " +
                                 "xmlns:pic=\"http://schemas.openxmlformats.org/drawingml/2006/picture\">" +
                                 "<w:p>" +
                                 "<w:r><w:t>Hi</w:t></w:r>" +
                                 "<w:r><w:drawing>" +
                                 "<wp:inline>" +
                                 "<wp:extent cx=\"914400\" cy=\"914400\"/>" +
                                 "<a:graphic><a:graphicData><pic:pic><pic:blipFill><a:blip r:embed=\"rIdImage\"/></pic:blipFill></pic:pic></a:graphicData></a:graphic>" +
                                 "</wp:inline>" +
                                 "</w:drawing></w:r>" +
                                 "<w:r><w:t>There</w:t></w:r>" +
                                 "</w:p>" +
                                 "</w:hdr>");
                }

                var headerRelsEntry = archive.CreateEntry("word/_rels/header1.xml.rels");
                using (var headerRelsStream = headerRelsEntry.Open())
                using (var writer = new StreamWriter(headerRelsStream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" +
                                 "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
                                 "<Relationship Id=\"rIdImage\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/image\" Target=\"media/image1.png\" />" +
                                 "</Relationships>");
                }

                var imageEntry = archive.CreateEntry("word/media/image1.png");
                using (var imageStream = imageEntry.Open())
                {
                    byte[] pngData = GetTestPngBytes();
                    imageStream.Write(pngData, 0, pngData.Length);
                }
            }

            using var testStream = new MemoryStream(ms.ToArray());
            using var reader = new Nedev.FileConverters.DocxToDoc.Format.DocxReader(testStream);

            var model = reader.ReadDocument();

            var section = Assert.Single(model.Sections);
            Assert.Equal("HiThere", section.DefaultHeaderText);
            Assert.Equal("HiThere", section.HeaderText);

            Assert.NotNull(section.DefaultHeaderStory);
            var story = section.DefaultHeaderStory!;
            var paragraph = Assert.Single(story.Paragraphs);
            var significantRuns = paragraph.Runs.FindAll(run => run.Text.Length > 0 || run.Image != null);
            Assert.Equal("Hi", significantRuns[0].Text);
            Assert.NotNull(significantRuns[1].Image);
            Assert.Equal(Nedev.FileConverters.DocxToDoc.Model.ImageLayoutType.Inline, significantRuns[1].Image!.LayoutType);
            Assert.NotNull(significantRuns[1].Image!.Data);
            Assert.True(significantRuns[1].Image!.Data!.Length > 0);
            Assert.Equal("image/png", significantRuns[1].Image!.ContentType);
            Assert.Equal(96, significantRuns[1].Image!.Width);
            Assert.Equal(96, significantRuns[1].Image!.Height);
            Assert.Equal("There", significantRuns[2].Text);
        }

        [Fact]
        public void ReadDocument_WithMixedHeaderFieldAndInlineImageStory_PreservesSummaryAndStructuredRuns()
        {
            using var ms = new MemoryStream();
            using (var archive = new ZipArchive(ms, ZipArchiveMode.Create, true))
            {
                var documentEntry = archive.CreateEntry("word/document.xml");
                using (var documentStream = documentEntry.Open())
                using (var writer = new StreamWriter(documentStream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" +
                                 "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">" +
                                 "<w:body>" +
                                 "<w:p><w:r><w:t>AB</w:t></w:r></w:p>" +
                                 "<w:sectPr>" +
                                 "<w:headerReference w:type=\"default\" r:id=\"rIdHeader\"/>" +
                                 "</w:sectPr>" +
                                 "</w:body>" +
                                 "</w:document>");
                }

                var relsEntry = archive.CreateEntry("word/_rels/document.xml.rels");
                using (var relsStream = relsEntry.Open())
                using (var writer = new StreamWriter(relsStream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" +
                                 "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
                                 "<Relationship Id=\"rIdHeader\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/header\" Target=\"header1.xml\" />" +
                                 "</Relationships>");
                }

                var headerEntry = archive.CreateEntry("word/header1.xml");
                using (var headerStream = headerEntry.Open())
                using (var writer = new StreamWriter(headerStream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" +
                                 "<w:hdr xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" " +
                                 "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" " +
                                 "xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" " +
                                 "xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\" " +
                                 "xmlns:pic=\"http://schemas.openxmlformats.org/drawingml/2006/picture\">" +
                                 "<w:p>" +
                                 "<w:r><w:t>Page </w:t></w:r>" +
                                 "<w:fldSimple w:instr=\"PAGE\"><w:r><w:t>1</w:t></w:r></w:fldSimple>" +
                                 "<w:r><w:t xml:space=\"preserve\"> </w:t></w:r>" +
                                 "<w:r><w:drawing>" +
                                 "<wp:inline>" +
                                 "<wp:extent cx=\"914400\" cy=\"914400\"/>" +
                                 "<a:graphic><a:graphicData><pic:pic><pic:blipFill><a:blip r:embed=\"rIdImage\"/></pic:blipFill></pic:pic></a:graphicData></a:graphic>" +
                                 "</wp:inline>" +
                                 "</w:drawing></w:r>" +
                                 "<w:r><w:t> tail</w:t></w:r>" +
                                 "</w:p>" +
                                 "</w:hdr>");
                }

                var headerRelsEntry = archive.CreateEntry("word/_rels/header1.xml.rels");
                using (var headerRelsStream = headerRelsEntry.Open())
                using (var writer = new StreamWriter(headerRelsStream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" +
                                 "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
                                 "<Relationship Id=\"rIdImage\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/image\" Target=\"media/image1.png\" />" +
                                 "</Relationships>");
                }

                var imageEntry = archive.CreateEntry("word/media/image1.png");
                using (var imageStream = imageEntry.Open())
                {
                    byte[] pngData = GetTestPngBytes();
                    imageStream.Write(pngData, 0, pngData.Length);
                }
            }

            using var testStream = new MemoryStream(ms.ToArray());
            using var reader = new Nedev.FileConverters.DocxToDoc.Format.DocxReader(testStream);

            var model = reader.ReadDocument();

            var section = Assert.Single(model.Sections);
            Assert.Equal("Page 1 tail", section.DefaultHeaderText);
            Assert.Equal("Page 1 tail", section.HeaderText);

            Assert.NotNull(section.DefaultHeaderStory);
            var story = section.DefaultHeaderStory!;
            var paragraph = Assert.Single(story.Paragraphs);
            var significantRuns = paragraph.Runs.FindAll(run =>
                run.Text.Length > 0 ||
                run.Field != null ||
                run.IsFieldBegin ||
                run.IsFieldSeparate ||
                run.IsFieldEnd ||
                run.Image != null);

            Assert.Equal("Page ", significantRuns[0].Text);
            Assert.True(significantRuns[1].IsFieldBegin);
            Assert.Equal(FieldType.Page, significantRuns[1].Field?.Type);
            Assert.True(significantRuns[2].IsFieldSeparate);
            Assert.Equal(FieldType.Page, significantRuns[2].Field?.Type);
            Assert.Equal("1", significantRuns[3].Text);
            Assert.True(significantRuns[4].IsFieldEnd);
            Assert.Equal(FieldType.Page, significantRuns[4].Field?.Type);
            Assert.Equal(" ", significantRuns[5].Text);
            Assert.NotNull(significantRuns[6].Image);
            Assert.Equal(Nedev.FileConverters.DocxToDoc.Model.ImageLayoutType.Inline, significantRuns[6].Image!.LayoutType);
            Assert.NotNull(significantRuns[6].Image!.Data);
            Assert.True(significantRuns[6].Image!.Data!.Length > 0);
            Assert.Equal("tail", significantRuns[7].Text);
        }

        [Fact]
        public void ReadDocument_WithMixedHeaderFieldAndFloatingImageStory_PreservesSummaryAndStructuredRuns()
        {
            using var ms = new MemoryStream();
            using (var archive = new ZipArchive(ms, ZipArchiveMode.Create, true))
            {
                var documentEntry = archive.CreateEntry("word/document.xml");
                using (var documentStream = documentEntry.Open())
                using (var writer = new StreamWriter(documentStream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" +
                                 "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">" +
                                 "<w:body>" +
                                 "<w:p><w:r><w:t>AB</w:t></w:r></w:p>" +
                                 "<w:sectPr>" +
                                 "<w:headerReference w:type=\"default\" r:id=\"rIdHeader\"/>" +
                                 "</w:sectPr>" +
                                 "</w:body>" +
                                 "</w:document>");
                }

                var relsEntry = archive.CreateEntry("word/_rels/document.xml.rels");
                using (var relsStream = relsEntry.Open())
                using (var writer = new StreamWriter(relsStream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" +
                                 "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
                                 "<Relationship Id=\"rIdHeader\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/header\" Target=\"header1.xml\" />" +
                                 "</Relationships>");
                }

                var headerEntry = archive.CreateEntry("word/header1.xml");
                using (var headerStream = headerEntry.Open())
                using (var writer = new StreamWriter(headerStream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" +
                                 "<w:hdr xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" " +
                                 "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" " +
                                 "xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" " +
                                 "xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\" " +
                                 "xmlns:pic=\"http://schemas.openxmlformats.org/drawingml/2006/picture\">" +
                                 "<w:p>" +
                                 "<w:r><w:t>Page </w:t></w:r>" +
                                 "<w:fldSimple w:instr=\"PAGE\"><w:r><w:t>1</w:t></w:r></w:fldSimple>" +
                                 "<w:r><w:t xml:space=\"preserve\"> </w:t></w:r>" +
                                 "<w:r><w:drawing>" +
                                 "<wp:anchor behindDoc=\"0\" allowOverlap=\"1\" distT=\"0\" distB=\"0\" distL=\"0\" distR=\"0\">" +
                                 "<wp:extent cx=\"914400\" cy=\"914400\"/>" +
                                 "<a:graphic><a:graphicData><pic:pic><pic:blipFill><a:blip r:embed=\"rIdImage\"/></pic:blipFill></pic:pic></a:graphicData></a:graphic>" +
                                 "</wp:anchor>" +
                                 "</w:drawing></w:r>" +
                                 "<w:r><w:t> tail</w:t></w:r>" +
                                 "</w:p>" +
                                 "</w:hdr>");
                }

                var headerRelsEntry = archive.CreateEntry("word/_rels/header1.xml.rels");
                using (var headerRelsStream = headerRelsEntry.Open())
                using (var writer = new StreamWriter(headerRelsStream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" +
                                 "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
                                 "<Relationship Id=\"rIdImage\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/image\" Target=\"media/image1.png\" />" +
                                 "</Relationships>");
                }

                var imageEntry = archive.CreateEntry("word/media/image1.png");
                using (var imageStream = imageEntry.Open())
                {
                    byte[] pngData = GetTestPngBytes();
                    imageStream.Write(pngData, 0, pngData.Length);
                }
            }

            using var testStream = new MemoryStream(ms.ToArray());
            using var reader = new Nedev.FileConverters.DocxToDoc.Format.DocxReader(testStream);

            var model = reader.ReadDocument();

            var section = Assert.Single(model.Sections);
            Assert.Equal("Page 1 tail", section.DefaultHeaderText);
            Assert.Equal("Page 1 tail", section.HeaderText);

            Assert.NotNull(section.DefaultHeaderStory);
            var story = section.DefaultHeaderStory!;
            var paragraph = Assert.Single(story.Paragraphs);
            var significantRuns = paragraph.Runs.FindAll(run =>
                run.Text.Length > 0 ||
                run.Field != null ||
                run.IsFieldBegin ||
                run.IsFieldSeparate ||
                run.IsFieldEnd ||
                run.Image != null);

            Assert.Equal("Page ", significantRuns[0].Text);
            Assert.True(significantRuns[1].IsFieldBegin);
            Assert.Equal(FieldType.Page, significantRuns[1].Field?.Type);
            Assert.True(significantRuns[2].IsFieldSeparate);
            Assert.Equal(FieldType.Page, significantRuns[2].Field?.Type);
            Assert.Equal("1", significantRuns[3].Text);
            Assert.True(significantRuns[4].IsFieldEnd);
            Assert.Equal(FieldType.Page, significantRuns[4].Field?.Type);
            Assert.Equal(" ", significantRuns[5].Text);
            Assert.NotNull(significantRuns[6].Image);
            var image = significantRuns[6].Image!;
            Assert.Equal(Nedev.FileConverters.DocxToDoc.Model.ImageLayoutType.Floating, image.LayoutType);
            Assert.Equal(Nedev.FileConverters.DocxToDoc.Model.ImageWrapType.Inline, image.WrapType);
            Assert.NotNull(image.Data);
            Assert.True(image.Data!.Length > 0);
            Assert.Equal("tail", significantRuns[7].Text);
        }

        [Fact]
        public void ReadDocument_WithMixedHeaderHyperlinkAndInlineImageStory_PreservesSummaryAndStructuredRuns()
        {
            using var ms = new MemoryStream();
            using (var archive = new ZipArchive(ms, ZipArchiveMode.Create, true))
            {
                var documentEntry = archive.CreateEntry("word/document.xml");
                using (var documentStream = documentEntry.Open())
                using (var writer = new StreamWriter(documentStream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" +
                                 "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">" +
                                 "<w:body>" +
                                 "<w:p><w:r><w:t>AB</w:t></w:r></w:p>" +
                                 "<w:sectPr>" +
                                 "<w:headerReference w:type=\"default\" r:id=\"rIdHeader\"/>" +
                                 "</w:sectPr>" +
                                 "</w:body>" +
                                 "</w:document>");
                }

                var relsEntry = archive.CreateEntry("word/_rels/document.xml.rels");
                using (var relsStream = relsEntry.Open())
                using (var writer = new StreamWriter(relsStream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" +
                                 "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
                                 "<Relationship Id=\"rIdHeader\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/header\" Target=\"header1.xml\" />" +
                                 "</Relationships>");
                }

                var headerEntry = archive.CreateEntry("word/header1.xml");
                using (var headerStream = headerEntry.Open())
                using (var writer = new StreamWriter(headerStream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" +
                                 "<w:hdr xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" " +
                                 "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" " +
                                 "xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" " +
                                 "xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\" " +
                                 "xmlns:pic=\"http://schemas.openxmlformats.org/drawingml/2006/picture\">" +
                                 "<w:p>" +
                                 "<w:r><w:t>Go </w:t></w:r>" +
                                 "<w:hyperlink r:id=\"rIdLink\">" +
                                 "<w:r><w:t>Example</w:t></w:r>" +
                                 "</w:hyperlink>" +
                                 "<w:r><w:t xml:space=\"preserve\"> </w:t></w:r>" +
                                 "<w:r><w:drawing>" +
                                 "<wp:inline>" +
                                 "<wp:extent cx=\"914400\" cy=\"914400\"/>" +
                                 "<a:graphic><a:graphicData><pic:pic><pic:blipFill><a:blip r:embed=\"rIdImage\"/></pic:blipFill></pic:pic></a:graphicData></a:graphic>" +
                                 "</wp:inline>" +
                                 "</w:drawing></w:r>" +
                                 "<w:r><w:t> now</w:t></w:r>" +
                                 "</w:p>" +
                                 "</w:hdr>");
                }

                var headerRelsEntry = archive.CreateEntry("word/_rels/header1.xml.rels");
                using (var headerRelsStream = headerRelsEntry.Open())
                using (var writer = new StreamWriter(headerRelsStream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" +
                                 "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
                                 "<Relationship Id=\"rIdLink\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink\" Target=\"https://example.com\" TargetMode=\"External\" />" +
                                 "<Relationship Id=\"rIdImage\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/image\" Target=\"media/image1.png\" />" +
                                 "</Relationships>");
                }

                var imageEntry = archive.CreateEntry("word/media/image1.png");
                using (var imageStream = imageEntry.Open())
                {
                    byte[] pngData = GetTestPngBytes();
                    imageStream.Write(pngData, 0, pngData.Length);
                }
            }

            using var testStream = new MemoryStream(ms.ToArray());
            using var reader = new Nedev.FileConverters.DocxToDoc.Format.DocxReader(testStream);

            var model = reader.ReadDocument();

            var section = Assert.Single(model.Sections);
            Assert.Equal("Go Example now", section.DefaultHeaderText);
            Assert.Equal("Go Example now", section.HeaderText);

            Assert.NotNull(section.DefaultHeaderStory);
            var story = section.DefaultHeaderStory!;
            var paragraph = Assert.Single(story.Paragraphs);
            var significantRuns = paragraph.Runs.FindAll(run => run.Text.Length > 0 || run.Image != null);

            Assert.Equal("Go ", significantRuns[0].Text);
            Assert.Equal("Example", significantRuns[1].Text);
            Assert.NotNull(significantRuns[1].Hyperlink);
            Assert.Equal("https://example.com", significantRuns[1].Hyperlink!.TargetUrl);
            Assert.Equal(" ", significantRuns[2].Text);
            Assert.NotNull(significantRuns[3].Image);
            Assert.Equal(Nedev.FileConverters.DocxToDoc.Model.ImageLayoutType.Inline, significantRuns[3].Image!.LayoutType);
            Assert.NotNull(significantRuns[3].Image!.Data);
            Assert.True(significantRuns[3].Image!.Data!.Length > 0);
            Assert.Equal("now", significantRuns[4].Text);
        }

        [Fact]
        public void ReadDocument_WithHeaderFloatingImageStory_PreservesFloatingImageRun()
        {
            using var ms = new MemoryStream();
            using (var archive = new ZipArchive(ms, ZipArchiveMode.Create, true))
            {
                var documentEntry = archive.CreateEntry("word/document.xml");
                using (var documentStream = documentEntry.Open())
                using (var writer = new StreamWriter(documentStream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" +
                                 "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">" +
                                 "<w:body>" +
                                 "<w:p><w:r><w:t>AB</w:t></w:r></w:p>" +
                                 "<w:sectPr>" +
                                 "<w:headerReference w:type=\"default\" r:id=\"rIdHeader\"/>" +
                                 "</w:sectPr>" +
                                 "</w:body>" +
                                 "</w:document>");
                }

                var relsEntry = archive.CreateEntry("word/_rels/document.xml.rels");
                using (var relsStream = relsEntry.Open())
                using (var writer = new StreamWriter(relsStream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" +
                                 "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
                                 "<Relationship Id=\"rIdHeader\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/header\" Target=\"header1.xml\" />" +
                                 "</Relationships>");
                }

                var headerEntry = archive.CreateEntry("word/header1.xml");
                using (var headerStream = headerEntry.Open())
                using (var writer = new StreamWriter(headerStream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" +
                                 "<w:hdr xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" " +
                                 "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" " +
                                 "xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" " +
                                 "xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\" " +
                                 "xmlns:pic=\"http://schemas.openxmlformats.org/drawingml/2006/picture\">" +
                                 "<w:p>" +
                                 "<w:r><w:t>Hi</w:t></w:r>" +
                                 "<w:r><w:drawing>" +
                                 "<wp:anchor behindDoc=\"0\" allowOverlap=\"1\" distT=\"0\" distB=\"0\" distL=\"0\" distR=\"0\">" +
                                 "<wp:extent cx=\"914400\" cy=\"914400\"/>" +
                                 "<a:graphic><a:graphicData><pic:pic><pic:blipFill><a:blip r:embed=\"rIdImage\"/></pic:blipFill></pic:pic></a:graphicData></a:graphic>" +
                                 "</wp:anchor>" +
                                 "</w:drawing></w:r>" +
                                 "<w:r><w:t>There</w:t></w:r>" +
                                 "</w:p>" +
                                 "</w:hdr>");
                }

                var headerRelsEntry = archive.CreateEntry("word/_rels/header1.xml.rels");
                using (var headerRelsStream = headerRelsEntry.Open())
                using (var writer = new StreamWriter(headerRelsStream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" +
                                 "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
                                 "<Relationship Id=\"rIdImage\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/image\" Target=\"media/image1.png\" />" +
                                 "</Relationships>");
                }

                var imageEntry = archive.CreateEntry("word/media/image1.png");
                using (var imageStream = imageEntry.Open())
                {
                    byte[] pngData = GetTestPngBytes();
                    imageStream.Write(pngData, 0, pngData.Length);
                }
            }

            using var testStream = new MemoryStream(ms.ToArray());
            using var reader = new Nedev.FileConverters.DocxToDoc.Format.DocxReader(testStream);

            var model = reader.ReadDocument();

            var section = Assert.Single(model.Sections);
            Assert.Equal("HiThere", section.DefaultHeaderText);
            Assert.Equal("HiThere", section.HeaderText);

            Assert.NotNull(section.DefaultHeaderStory);
            var story = section.DefaultHeaderStory!;
            var paragraph = Assert.Single(story.Paragraphs);
            var significantRuns = paragraph.Runs.FindAll(run => run.Text.Length > 0 || run.Image != null);

            Assert.Collection(
                significantRuns,
                run => Assert.Equal("Hi", run.Text),
                run =>
                {
                    Assert.NotNull(run.Image);
                    Assert.Equal(Nedev.FileConverters.DocxToDoc.Model.ImageLayoutType.Floating, run.Image!.LayoutType);
                    Assert.Equal(Nedev.FileConverters.DocxToDoc.Model.ImageWrapType.Inline, run.Image.WrapType);
                    Assert.NotNull(run.Image.Data);
                    Assert.True(run.Image.Data!.Length > 0);
                },
                run => Assert.Equal("There", run.Text));
        }

        [Fact]
        public void ReadDocument_WithHeaderTextBoxContent_PreservesSummaryAndStructuredRuns()
        {
            using var ms = new MemoryStream();
            using (var archive = new ZipArchive(ms, ZipArchiveMode.Create, true))
            {
                var documentEntry = archive.CreateEntry("word/document.xml");
                using (var documentStream = documentEntry.Open())
                using (var writer = new StreamWriter(documentStream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" +
                                 "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">" +
                                 "<w:body>" +
                                 "<w:p><w:r><w:t>AB</w:t></w:r></w:p>" +
                                 "<w:sectPr>" +
                                 "<w:headerReference w:type=\"default\" r:id=\"rIdHeader\"/>" +
                                 "</w:sectPr>" +
                                 "</w:body>" +
                                 "</w:document>");
                }

                var relsEntry = archive.CreateEntry("word/_rels/document.xml.rels");
                using (var relsStream = relsEntry.Open())
                using (var writer = new StreamWriter(relsStream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" +
                                 "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
                                 "<Relationship Id=\"rIdHeader\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/header\" Target=\"header1.xml\" />" +
                                 "</Relationships>");
                }

                var headerEntry = archive.CreateEntry("word/header1.xml");
                using (var headerStream = headerEntry.Open())
                using (var writer = new StreamWriter(headerStream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" +
                                 "<w:hdr xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" " +
                                 "xmlns:v=\"urn:schemas-microsoft-com:vml\">" +
                                 "<w:p>" +
                                 "<w:r><w:t xml:space=\"preserve\">Head </w:t></w:r>" +
                                 "<w:r><w:pict><v:shape><v:textbox><w:txbxContent><w:p><w:r><w:t>Box</w:t></w:r></w:p></w:txbxContent></v:textbox></v:shape></w:pict></w:r>" +
                                 "<w:r><w:t> tail</w:t></w:r>" +
                                 "</w:p>" +
                                 "</w:hdr>");
                }
            }

            using var testStream = new MemoryStream(ms.ToArray());
            using var reader = new Nedev.FileConverters.DocxToDoc.Format.DocxReader(testStream);

            var model = reader.ReadDocument();

            var section = Assert.Single(model.Sections);
            Assert.Equal("Head Box tail", section.DefaultHeaderText);
            Assert.Equal("Head Box tail", section.HeaderText);

            Assert.NotNull(section.DefaultHeaderStory);
            var story = section.DefaultHeaderStory!;
            var paragraph = Assert.Single(story.Paragraphs);
            Assert.Collection(
                paragraph.Runs.FindAll(run => run.Text.Length > 0),
                run => Assert.Equal("Head ", run.Text),
                run => Assert.Equal("Box", run.Text),
                run => Assert.Equal(" tail", run.Text));
        }

        [Fact]
        public void ReadDocument_WithMixedHeaderHyperlinkAndFloatingImageStory_PreservesSummaryAndStructuredRuns()
        {
            using var ms = new MemoryStream();
            using (var archive = new ZipArchive(ms, ZipArchiveMode.Create, true))
            {
                var documentEntry = archive.CreateEntry("word/document.xml");
                using (var documentStream = documentEntry.Open())
                using (var writer = new StreamWriter(documentStream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" +
                                 "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">" +
                                 "<w:body>" +
                                 "<w:p><w:r><w:t>AB</w:t></w:r></w:p>" +
                                 "<w:sectPr>" +
                                 "<w:headerReference w:type=\"default\" r:id=\"rIdHeader\"/>" +
                                 "</w:sectPr>" +
                                 "</w:body>" +
                                 "</w:document>");
                }

                var relsEntry = archive.CreateEntry("word/_rels/document.xml.rels");
                using (var relsStream = relsEntry.Open())
                using (var writer = new StreamWriter(relsStream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" +
                                 "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
                                 "<Relationship Id=\"rIdHeader\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/header\" Target=\"header1.xml\" />" +
                                 "</Relationships>");
                }

                var headerEntry = archive.CreateEntry("word/header1.xml");
                using (var headerStream = headerEntry.Open())
                using (var writer = new StreamWriter(headerStream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" +
                                 "<w:hdr xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" " +
                                 "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" " +
                                 "xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" " +
                                 "xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\" " +
                                 "xmlns:pic=\"http://schemas.openxmlformats.org/drawingml/2006/picture\">" +
                                 "<w:p>" +
                                 "<w:r><w:t>Go </w:t></w:r>" +
                                 "<w:hyperlink r:id=\"rIdLink\">" +
                                 "<w:r><w:t>Example</w:t></w:r>" +
                                 "</w:hyperlink>" +
                                 "<w:r><w:t xml:space=\"preserve\"> </w:t></w:r>" +
                                 "<w:r><w:drawing>" +
                                 "<wp:anchor behindDoc=\"0\" allowOverlap=\"1\" distT=\"0\" distB=\"0\" distL=\"0\" distR=\"0\">" +
                                 "<wp:extent cx=\"914400\" cy=\"914400\"/>" +
                                 "<a:graphic><a:graphicData><pic:pic><pic:blipFill><a:blip r:embed=\"rIdImage\"/></pic:blipFill></pic:pic></a:graphicData></a:graphic>" +
                                 "</wp:anchor>" +
                                 "</w:drawing></w:r>" +
                                 "<w:r><w:t> now</w:t></w:r>" +
                                 "</w:p>" +
                                 "</w:hdr>");
                }

                var headerRelsEntry = archive.CreateEntry("word/_rels/header1.xml.rels");
                using (var headerRelsStream = headerRelsEntry.Open())
                using (var writer = new StreamWriter(headerRelsStream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" +
                                 "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
                                 "<Relationship Id=\"rIdLink\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink\" Target=\"https://example.com\" TargetMode=\"External\" />" +
                                 "<Relationship Id=\"rIdImage\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/image\" Target=\"media/image1.png\" />" +
                                 "</Relationships>");
                }

                var imageEntry = archive.CreateEntry("word/media/image1.png");
                using (var imageStream = imageEntry.Open())
                {
                    byte[] pngData = GetTestPngBytes();
                    imageStream.Write(pngData, 0, pngData.Length);
                }
            }

            using var testStream = new MemoryStream(ms.ToArray());
            using var reader = new Nedev.FileConverters.DocxToDoc.Format.DocxReader(testStream);

            var model = reader.ReadDocument();

            var section = Assert.Single(model.Sections);
            Assert.Equal("Go Example now", section.DefaultHeaderText);
            Assert.Equal("Go Example now", section.HeaderText);

            Assert.NotNull(section.DefaultHeaderStory);
            var story = section.DefaultHeaderStory!;
            var paragraph = Assert.Single(story.Paragraphs);
            var significantRuns = paragraph.Runs.FindAll(run => run.Text.Length > 0 || run.Image != null);

            Assert.Equal("Go ", significantRuns[0].Text);
            Assert.Equal("Example", significantRuns[1].Text);
            Assert.NotNull(significantRuns[1].Hyperlink);
            Assert.Equal("https://example.com", significantRuns[1].Hyperlink!.TargetUrl);
            Assert.Equal(" ", significantRuns[2].Text);
            Assert.NotNull(significantRuns[3].Image);
            Assert.Equal(Nedev.FileConverters.DocxToDoc.Model.ImageLayoutType.Floating, significantRuns[3].Image!.LayoutType);
            Assert.NotNull(significantRuns[3].Image!.Data);
            Assert.True(significantRuns[3].Image!.Data!.Length > 0);
            Assert.Equal("now", significantRuns[4].Text);
        }

        [Fact]
        public void ReadDocument_WithHeaderParagraphAndTableStory_PreservesMixedBlocks()
        {
            using var ms = new MemoryStream();
            using (var archive = new ZipArchive(ms, ZipArchiveMode.Create, true))
            {
                var documentEntry = archive.CreateEntry("word/document.xml");
                using (var documentStream = documentEntry.Open())
                using (var writer = new StreamWriter(documentStream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" +
                                 "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">" +
                                 "<w:body>" +
                                 "<w:p><w:r><w:t>AB</w:t></w:r></w:p>" +
                                 "<w:sectPr>" +
                                 "<w:headerReference w:type=\"default\" r:id=\"rIdHeader\"/>" +
                                 "</w:sectPr>" +
                                 "</w:body>" +
                                 "</w:document>");
                }

                var relsEntry = archive.CreateEntry("word/_rels/document.xml.rels");
                using (var relsStream = relsEntry.Open())
                using (var writer = new StreamWriter(relsStream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" +
                                 "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
                                 "<Relationship Id=\"rIdHeader\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/header\" Target=\"header1.xml\" />" +
                                 "</Relationships>");
                }

                var headerEntry = archive.CreateEntry("word/header1.xml");
                using (var headerStream = headerEntry.Open())
                using (var writer = new StreamWriter(headerStream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" +
                                 "<w:hdr xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">" +
                                 "<w:p><w:r><w:t>Head</w:t></w:r></w:p>" +
                                 "<w:tbl>" +
                                 "<w:tblGrid><w:gridCol w:w=\"2400\"/></w:tblGrid>" +
                                 "<w:tr>" +
                                 "<w:tc>" +
                                 "<w:tcPr><w:tcW w:type=\"dxa\" w:w=\"2400\"/></w:tcPr>" +
                                 "<w:p><w:r><w:t>Cell</w:t></w:r></w:p>" +
                                 "</w:tc>" +
                                 "</w:tr>" +
                                 "</w:tbl>" +
                                 "</w:hdr>");
                }
            }

            using var testStream = new MemoryStream(ms.ToArray());
            using var reader = new Nedev.FileConverters.DocxToDoc.Format.DocxReader(testStream);

            var model = reader.ReadDocument();

            var section = Assert.Single(model.Sections);
            Assert.Equal("Head\rCell", section.DefaultHeaderText);
            Assert.Equal("Head\rCell", section.HeaderText);

            Assert.NotNull(section.DefaultHeaderStory);
            var story = section.DefaultHeaderStory!;
            Assert.Single(story.Paragraphs);
            Assert.Equal(2, story.Content.Count);

            var headingParagraph = Assert.IsType<ParagraphModel>(story.Content[0]);
            var headingRun = Assert.Single(headingParagraph.Runs.FindAll(run => run.Text.Length > 0));
            Assert.Equal("Head", headingRun.Text);

            var table = Assert.IsType<TableModel>(story.Content[1]);
            var row = Assert.Single(table.Rows);
            var cell = Assert.Single(row.Cells);
            var cellParagraph = Assert.Single(cell.Paragraphs);
            var cellRun = Assert.Single(cellParagraph.Runs.FindAll(run => run.Text.Length > 0));
            Assert.Equal("Cell", cellRun.Text);
        }

        [Fact]
        public void ReadDocument_WithHeaderTableAutoCellWidth_FallsBackToGridWidth()
        {
            using var ms = new MemoryStream();
            using (var archive = new ZipArchive(ms, ZipArchiveMode.Create, true))
            {
                var documentEntry = archive.CreateEntry("word/document.xml");
                using (var documentStream = documentEntry.Open())
                using (var writer = new StreamWriter(documentStream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" +
                                 "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">" +
                                 "<w:body><w:p><w:r><w:t>A</w:t></w:r></w:p><w:sectPr><w:headerReference w:type=\"default\" r:id=\"rIdHeader\"/></w:sectPr></w:body></w:document>");
                }

                var relsEntry = archive.CreateEntry("word/_rels/document.xml.rels");
                using (var relsStream = relsEntry.Open())
                using (var writer = new StreamWriter(relsStream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" +
                                 "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
                                 "<Relationship Id=\"rIdHeader\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/header\" Target=\"header1.xml\" />" +
                                 "</Relationships>");
                }

                var headerEntry = archive.CreateEntry("word/header1.xml");
                using (var headerStream = headerEntry.Open())
                using (var writer = new StreamWriter(headerStream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" +
                                 "<w:hdr xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">" +
                                 "<w:tbl><w:tblGrid><w:gridCol w:w=\"3600\"/></w:tblGrid><w:tr><w:tc>" +
                                 "<w:tcPr><w:tcW w:type=\"auto\" w:w=\"0\"/></w:tcPr><w:p><w:r><w:t>Cell</w:t></w:r></w:p>" +
                                 "</w:tc></w:tr></w:tbl>" +
                                 "</w:hdr>");
                }
            }

            using var testStream = new MemoryStream(ms.ToArray());
            using var reader = new Nedev.FileConverters.DocxToDoc.Format.DocxReader(testStream);

            var model = reader.ReadDocument();

            var section = Assert.Single(model.Sections);
            var story = section.DefaultHeaderStory;
            Assert.NotNull(story);
            var table = Assert.IsType<TableModel>(Assert.Single(story!.Content));
            var cell = Assert.Single(Assert.Single(table.Rows).Cells);
            Assert.Equal(3600, cell.Width);
        }

        [Fact]
        public void ReadDocument_WithHeaderAltChunk_PreservesVisibleParagraphsInOrder()
        {
            byte[] packageBytes = CreateDefaultHeaderAltChunkPackage("afchunk-header.txt", "Chunk line 1\r\nChunk line 2");

            using var testStream = new MemoryStream(packageBytes);
            using var reader = new Nedev.FileConverters.DocxToDoc.Format.DocxReader(testStream);

            var model = reader.ReadDocument();

            var section = Assert.Single(model.Sections);
            Assert.Equal("Lead\rChunk line 1\rChunk line 2\rTail", section.DefaultHeaderText);
            Assert.Equal("Lead\rChunk line 1\rChunk line 2\rTail", section.HeaderText);

            Assert.NotNull(section.DefaultHeaderStory);
            var story = section.DefaultHeaderStory!;
            Assert.Equal(
                new[] { "Lead", "Chunk line 1", "Chunk line 2", "Tail" },
                story.Content
                    .Select(block => string.Concat(Assert.IsType<ParagraphModel>(block).Runs.Select(run => run.Text)))
                    .ToArray());
        }

        [Fact]
        public void ReadDocument_WithHeaderHtmlAltChunk_PreservesVisibleParagraphsInOrder()
        {
            byte[] packageBytes = CreateDefaultHeaderAltChunkPackage(
                "afchunk-header.html",
                "<html><body><p>Chunk line 1</p><div>Chunk line 2</div></body></html>");

            using var testStream = new MemoryStream(packageBytes);
            using var reader = new Nedev.FileConverters.DocxToDoc.Format.DocxReader(testStream);

            var model = reader.ReadDocument();

            var section = Assert.Single(model.Sections);
            Assert.Equal("Lead\rChunk line 1\rChunk line 2\rTail", section.DefaultHeaderText);
            Assert.Equal("Lead\rChunk line 1\rChunk line 2\rTail", section.HeaderText);

            Assert.NotNull(section.DefaultHeaderStory);
            var story = section.DefaultHeaderStory!;
            Assert.Equal(
                new[] { "Lead", "Chunk line 1", "Chunk line 2", "Tail" },
                story.Content
                    .Select(block => string.Concat(Assert.IsType<ParagraphModel>(block).Runs.Select(run => run.Text)))
                    .ToArray());
        }

        [Fact]
        public void ReadDocument_WithHeaderRtfAltChunk_PreservesVisibleParagraphsInOrder()
        {
            byte[] packageBytes = CreateDefaultHeaderAltChunkPackage("afchunk-header.rtf", @"{\rtf1\ansi Chunk line 1\par Chunk line 2}");

            using var testStream = new MemoryStream(packageBytes);
            using var reader = new Nedev.FileConverters.DocxToDoc.Format.DocxReader(testStream);

            var model = reader.ReadDocument();

            var section = Assert.Single(model.Sections);
            Assert.Equal("Lead\rChunk line 1\rChunk line 2\rTail", section.DefaultHeaderText);
            Assert.Equal("Lead\rChunk line 1\rChunk line 2\rTail", section.HeaderText);

            Assert.NotNull(section.DefaultHeaderStory);
            var story = section.DefaultHeaderStory!;
            Assert.Equal(
                new[] { "Lead", "Chunk line 1", "Chunk line 2", "Tail" },
                story.Content
                    .Select(block => string.Concat(Assert.IsType<ParagraphModel>(block).Runs.Select(run => run.Text)))
                    .ToArray());
        }

        [Fact]
        public void ReadDocument_WithHeaderEmbeddedDocxAltChunk_PreservesVisibleParagraphsInOrder()
        {
            byte[] packageBytes = CreateDefaultHeaderAltChunkPackage("afchunk-header.docx", CreateEmbeddedDocxAltChunk(includeTable: false));

            using var testStream = new MemoryStream(packageBytes);
            using var reader = new Nedev.FileConverters.DocxToDoc.Format.DocxReader(testStream);

            var model = reader.ReadDocument();

            var section = Assert.Single(model.Sections);
            Assert.Equal("Lead\rInner lead\rInner tail\rTail", section.DefaultHeaderText);
            Assert.Equal("Lead\rInner lead\rInner tail\rTail", section.HeaderText);

            Assert.NotNull(section.DefaultHeaderStory);
            var story = section.DefaultHeaderStory!;
            Assert.Equal(
                new[] { "Lead", "Inner lead", "Inner tail", "Tail" },
                story.Content
                    .Select(block => string.Concat(Assert.IsType<ParagraphModel>(block).Runs.Select(run => run.Text)))
                    .ToArray());
        }

        [Fact]
        public void ReadDocument_WithHeaderEmbeddedDocxAltChunkContainingTable_PreservesTableBlocksAndVisibleContentOrder()
        {
            byte[] packageBytes = CreateHeaderAltChunkAndTablePackage("afchunk-header.docx", CreateEmbeddedDocxAltChunk(includeTable: true));

            using var testStream = new MemoryStream(packageBytes);
            using var reader = new Nedev.FileConverters.DocxToDoc.Format.DocxReader(testStream);

            var model = reader.ReadDocument();

            var section = Assert.Single(model.Sections);
            Assert.Equal("Lead\rInner lead\rCell\rInner tail\rCell\rTail", section.DefaultHeaderText);
            Assert.Equal("Lead\rInner lead\rCell\rInner tail\rCell\rTail", section.HeaderText);

            Assert.NotNull(section.DefaultHeaderStory);
            var story = section.DefaultHeaderStory!;
            Assert.Equal(6, story.Content.Count);

            Assert.Collection(
                story.Content,
                block => Assert.Equal("Lead", string.Concat(Assert.IsType<ParagraphModel>(block).Runs.Select(run => run.Text))),
                block => Assert.Equal("Inner lead", string.Concat(Assert.IsType<ParagraphModel>(block).Runs.Select(run => run.Text))),
                block =>
                {
                    var table = Assert.IsType<TableModel>(block);
                    var cellRun = Assert.Single(table.Rows[0].Cells[0].Paragraphs[0].Runs.FindAll(run => run.Text.Length > 0));
                    Assert.Equal("Cell", cellRun.Text);
                },
                block => Assert.Equal("Inner tail", string.Concat(Assert.IsType<ParagraphModel>(block).Runs.Select(run => run.Text))),
                block =>
                {
                    var table = Assert.IsType<TableModel>(block);
                    var cellRun = Assert.Single(table.Rows[0].Cells[0].Paragraphs[0].Runs.FindAll(run => run.Text.Length > 0));
                    Assert.Equal("Cell", cellRun.Text);
                },
                block => Assert.Equal("Tail", string.Concat(Assert.IsType<ParagraphModel>(block).Runs.Select(run => run.Text))));

            Assert.Equal(
                new[] { "Lead", "Inner lead", "Inner tail", "Tail" },
                story.Paragraphs.Select(paragraph => string.Concat(paragraph.Runs.Select(run => run.Text))).ToArray());
        }

        [Fact]
        public void ReadDocument_WithFooterEmbeddedDocxAltChunk_PreservesVisibleParagraphsInOrder()
        {
            byte[] packageBytes = CreateDefaultFooterAltChunkPackage("afchunk-footer.docx", CreateEmbeddedDocxAltChunk(includeTable: false));

            using var testStream = new MemoryStream(packageBytes);
            using var reader = new Nedev.FileConverters.DocxToDoc.Format.DocxReader(testStream);

            var model = reader.ReadDocument();

            var section = Assert.Single(model.Sections);
            Assert.Equal("Lead\rInner lead\rInner tail\rTail", section.DefaultFooterText);
            Assert.Equal("Lead\rInner lead\rInner tail\rTail", section.FooterText);

            Assert.NotNull(section.DefaultFooterStory);
            var story = section.DefaultFooterStory!;
            Assert.Equal(
                new[] { "Lead", "Inner lead", "Inner tail", "Tail" },
                story.Content
                    .Select(block => string.Concat(Assert.IsType<ParagraphModel>(block).Runs.Select(run => run.Text)))
                    .ToArray());
        }

        [Fact]
        public void ReadDocument_WithFooterEmbeddedDocxAltChunkContainingTable_PreservesTableBlocksAndVisibleContentOrder()
        {
            byte[] packageBytes = CreateFooterAltChunkAndTablePackage("afchunk-footer.docx", CreateEmbeddedDocxAltChunk(includeTable: true));

            using var testStream = new MemoryStream(packageBytes);
            using var reader = new Nedev.FileConverters.DocxToDoc.Format.DocxReader(testStream);

            var model = reader.ReadDocument();

            var section = Assert.Single(model.Sections);
            Assert.Equal("Lead\rInner lead\rCell\rInner tail\rCell\rTail", section.DefaultFooterText);
            Assert.Equal("Lead\rInner lead\rCell\rInner tail\rCell\rTail", section.FooterText);

            Assert.NotNull(section.DefaultFooterStory);
            var story = section.DefaultFooterStory!;
            Assert.Equal(6, story.Content.Count);

            Assert.Collection(
                story.Content,
                block => Assert.Equal("Lead", string.Concat(Assert.IsType<ParagraphModel>(block).Runs.Select(run => run.Text))),
                block => Assert.Equal("Inner lead", string.Concat(Assert.IsType<ParagraphModel>(block).Runs.Select(run => run.Text))),
                block =>
                {
                    var table = Assert.IsType<TableModel>(block);
                    var cellRun = Assert.Single(table.Rows[0].Cells[0].Paragraphs[0].Runs.FindAll(run => run.Text.Length > 0));
                    Assert.Equal("Cell", cellRun.Text);
                },
                block => Assert.Equal("Inner tail", string.Concat(Assert.IsType<ParagraphModel>(block).Runs.Select(run => run.Text))),
                block =>
                {
                    var table = Assert.IsType<TableModel>(block);
                    var cellRun = Assert.Single(table.Rows[0].Cells[0].Paragraphs[0].Runs.FindAll(run => run.Text.Length > 0));
                    Assert.Equal("Cell", cellRun.Text);
                },
                block => Assert.Equal("Tail", string.Concat(Assert.IsType<ParagraphModel>(block).Runs.Select(run => run.Text))));

            Assert.Equal(
                new[] { "Lead", "Inner lead", "Inner tail", "Tail" },
                story.Paragraphs.Select(paragraph => string.Concat(paragraph.Runs.Select(run => run.Text))).ToArray());
        }

        [Fact]
        public void ReadDocument_WithHeaderEmbeddedDocxAltChunkInsideTableCell_PreservesTableBlocksInsideCellContent()
        {
            byte[] packageBytes = CreateHeaderTableCellAltChunkPackage("afchunk-header-cell.docx", CreateEmbeddedDocxAltChunk(includeTable: true));

            using var testStream = new MemoryStream(packageBytes);
            using var reader = new Nedev.FileConverters.DocxToDoc.Format.DocxReader(testStream);

            var model = reader.ReadDocument();

            var section = Assert.Single(model.Sections);
            Assert.Equal("Lead\rInner lead\rCell\rInner tail\rTail", section.DefaultHeaderText);
            Assert.Equal("Lead\rInner lead\rCell\rInner tail\rTail", section.HeaderText);

            Assert.NotNull(section.DefaultHeaderStory);
            var story = section.DefaultHeaderStory!;
            Assert.Collection(
                story.Content,
                block => Assert.Equal("Lead", string.Concat(Assert.IsType<ParagraphModel>(block).Runs.Select(run => run.Text))),
                block =>
                {
                    var table = Assert.IsType<TableModel>(block);
                    var cell = table.Rows[0].Cells[0];
                    Assert.Collection(
                        cell.Content,
                        cellBlock => Assert.Equal("Inner lead", string.Concat(Assert.IsType<ParagraphModel>(cellBlock).Runs.Select(run => run.Text))),
                        cellBlock =>
                        {
                            var nestedTable = Assert.IsType<TableModel>(cellBlock);
                            Assert.Equal("Cell", nestedTable.Rows[0].Cells[0].Paragraphs[0].Runs[0].Text);
                        },
                        cellBlock => Assert.Equal("Inner tail", string.Concat(Assert.IsType<ParagraphModel>(cellBlock).Runs.Select(run => run.Text))));
                    Assert.Equal(
                        new[] { "Inner lead", "Cell", "Inner tail" },
                        cell.Paragraphs.Select(paragraph => string.Concat(paragraph.Runs.Select(run => run.Text))).ToArray());
                },
                block => Assert.Equal("Tail", string.Concat(Assert.IsType<ParagraphModel>(block).Runs.Select(run => run.Text))));

            Assert.Equal(
                new[] { "Lead", "Tail" },
                story.Paragraphs.Select(paragraph => string.Concat(paragraph.Runs.Select(run => run.Text))).ToArray());
        }

        [Fact]
        public void ReadDocument_WithFooterEmbeddedDocxAltChunkBetweenCellParagraphs_PreservesMixedBlockOrderInsideCellContentAndSummaryText()
        {
            byte[] packageBytes = CreateFooterTableCellAltChunkPackage(
                "afchunk-footer-cell.docx",
                CreateEmbeddedDocxAltChunk(includeTable: true),
                includeSurroundingCellParagraphs: true);

            using var testStream = new MemoryStream(packageBytes);
            using var reader = new Nedev.FileConverters.DocxToDoc.Format.DocxReader(testStream);

            var model = reader.ReadDocument();

            var section = Assert.Single(model.Sections);
            Assert.Equal("Lead\rCell lead\rInner lead\rCell\rInner tail\rCell tail\rTail", section.DefaultFooterText);
            Assert.Equal("Lead\rCell lead\rInner lead\rCell\rInner tail\rCell tail\rTail", section.FooterText);

            Assert.NotNull(section.DefaultFooterStory);
            var story = section.DefaultFooterStory!;
            Assert.Collection(
                story.Content,
                block => Assert.Equal("Lead", string.Concat(Assert.IsType<ParagraphModel>(block).Runs.Select(run => run.Text))),
                block =>
                {
                    var table = Assert.IsType<TableModel>(block);
                    var cell = table.Rows[0].Cells[0];
                    Assert.Collection(
                        cell.Content,
                        cellBlock => Assert.Equal("Cell lead", string.Concat(Assert.IsType<ParagraphModel>(cellBlock).Runs.Select(run => run.Text))),
                        cellBlock => Assert.Equal("Inner lead", string.Concat(Assert.IsType<ParagraphModel>(cellBlock).Runs.Select(run => run.Text))),
                        cellBlock =>
                        {
                            var nestedTable = Assert.IsType<TableModel>(cellBlock);
                            Assert.Equal("Cell", nestedTable.Rows[0].Cells[0].Paragraphs[0].Runs[0].Text);
                        },
                        cellBlock => Assert.Equal("Inner tail", string.Concat(Assert.IsType<ParagraphModel>(cellBlock).Runs.Select(run => run.Text))),
                        cellBlock => Assert.Equal("Cell tail", string.Concat(Assert.IsType<ParagraphModel>(cellBlock).Runs.Select(run => run.Text))));
                    Assert.Equal(
                        new[] { "Cell lead", "Inner lead", "Cell", "Inner tail", "Cell tail" },
                        cell.Paragraphs.Select(paragraph => string.Concat(paragraph.Runs.Select(run => run.Text))).ToArray());
                },
                block => Assert.Equal("Tail", string.Concat(Assert.IsType<ParagraphModel>(block).Runs.Select(run => run.Text))));

            Assert.Equal(
                new[] { "Lead", "Tail" },
                story.Paragraphs.Select(paragraph => string.Concat(paragraph.Runs.Select(run => run.Text))).ToArray());
        }

        [Fact]
        public void ReadDocument_WithFooterEmbeddedDocxAltChunkInsideTableCellMissingRelationship_IgnoresChunkAndPreservesCellParagraphs()
        {
            byte[] packageBytes = CreateFooterTableCellAltChunkPackage(
                "afchunk-footer-cell.docx",
                CreateEmbeddedDocxAltChunk(includeTable: true),
                includeChunkRelationship: false,
                includeSurroundingCellParagraphs: true);

            using var testStream = new MemoryStream(packageBytes);
            using var reader = new Nedev.FileConverters.DocxToDoc.Format.DocxReader(testStream);

            var model = reader.ReadDocument();

            var section = Assert.Single(model.Sections);
            Assert.Equal("Lead\rCell lead\rCell tail\rTail", section.DefaultFooterText);
            Assert.Equal("Lead\rCell lead\rCell tail\rTail", section.FooterText);

            Assert.NotNull(section.DefaultFooterStory);
            var story = section.DefaultFooterStory!;
            var outerTable = Assert.IsType<TableModel>(story.Content[1]);
            Assert.Equal(
                new[] { "Cell lead", "Cell tail" },
                outerTable.Rows[0].Cells[0].Paragraphs.Select(paragraph => string.Concat(paragraph.Runs.Select(run => run.Text))).ToArray());
        }

        [Fact]
        public void ReadDocument_WithFooterEmbeddedDocxAltChunkMissingRelationship_IgnoresChunkAndPreservesSurroundingParagraphs()
        {
            byte[] packageBytes = CreateDefaultFooterAltChunkPackage(
                "afchunk-footer.docx",
                CreateEmbeddedDocxAltChunk(includeTable: false),
                includeChunkRelationship: false);

            using var testStream = new MemoryStream(packageBytes);
            using var reader = new Nedev.FileConverters.DocxToDoc.Format.DocxReader(testStream);

            var model = reader.ReadDocument();

            var section = Assert.Single(model.Sections);
            Assert.Equal("Lead\rTail", section.DefaultFooterText);
            Assert.Equal("Lead\rTail", section.FooterText);

            Assert.NotNull(section.DefaultFooterStory);
            var story = section.DefaultFooterStory!;
            Assert.Equal(
                new[] { "Lead", "Tail" },
                story.Content
                    .Select(block => string.Concat(Assert.IsType<ParagraphModel>(block).Runs.Select(run => run.Text)))
                    .ToArray());
        }

        [Fact]
        public void ReadDocument_WithFooterEmbeddedDocxAltChunkMissingTargetPart_IgnoresChunkAndPreservesSurroundingParagraphs()
        {
            byte[] packageBytes = CreateDefaultFooterAltChunkPackage(
                "afchunk-footer.docx",
                CreateEmbeddedDocxAltChunk(includeTable: false),
                includeChunkPart: false);

            using var testStream = new MemoryStream(packageBytes);
            using var reader = new Nedev.FileConverters.DocxToDoc.Format.DocxReader(testStream);

            var model = reader.ReadDocument();

            var section = Assert.Single(model.Sections);
            Assert.Equal("Lead\rTail", section.DefaultFooterText);
            Assert.Equal("Lead\rTail", section.FooterText);

            Assert.NotNull(section.DefaultFooterStory);
            var story = section.DefaultFooterStory!;
            Assert.Equal(
                new[] { "Lead", "Tail" },
                story.Content
                    .Select(block => string.Concat(Assert.IsType<ParagraphModel>(block).Runs.Select(run => run.Text)))
                    .ToArray());
        }

        [Fact]
        public void ReadDocument_WithMalformedFooterEmbeddedDocxAltChunk_IgnoresChunkAndPreservesSurroundingParagraphs()
        {
            byte[] packageBytes = CreateDefaultFooterAltChunkPackage(
                "afchunk-footer.docx",
                Encoding.UTF8.GetBytes("not-a-zip-package"));

            using var testStream = new MemoryStream(packageBytes);
            using var reader = new Nedev.FileConverters.DocxToDoc.Format.DocxReader(testStream);

            var model = reader.ReadDocument();

            var section = Assert.Single(model.Sections);
            Assert.Equal("Lead\rTail", section.DefaultFooterText);
            Assert.Equal("Lead\rTail", section.FooterText);

            Assert.NotNull(section.DefaultFooterStory);
            var story = section.DefaultFooterStory!;
            Assert.Equal(
                new[] { "Lead", "Tail" },
                story.Content
                    .Select(block => string.Concat(Assert.IsType<ParagraphModel>(block).Runs.Select(run => run.Text)))
                    .ToArray());
        }

        [Fact]
        public void ReadDocument_WithFooterEmbeddedDocxAltChunkMissingDocumentXml_IgnoresChunkAndPreservesSurroundingParagraphs()
        {
            byte[] packageBytes = CreateDefaultFooterAltChunkPackage(
                "afchunk-footer.docx",
                CreateEmbeddedDocxAltChunkWithoutDocumentXml());

            using var testStream = new MemoryStream(packageBytes);
            using var reader = new Nedev.FileConverters.DocxToDoc.Format.DocxReader(testStream);

            var model = reader.ReadDocument();

            var section = Assert.Single(model.Sections);
            Assert.Equal("Lead\rTail", section.DefaultFooterText);
            Assert.Equal("Lead\rTail", section.FooterText);

            Assert.NotNull(section.DefaultFooterStory);
            var story = section.DefaultFooterStory!;
            Assert.Equal(
                new[] { "Lead", "Tail" },
                story.Content
                    .Select(block => string.Concat(Assert.IsType<ParagraphModel>(block).Runs.Select(run => run.Text)))
                    .ToArray());
        }

        [Fact]
        public void ReadDocument_WithoutEvenAndOddHeadersSetting_DoesNotPromoteEvenFooterEmbeddedDocxAltChunkToSummaryText()
        {
            byte[] packageBytes = CreateEvenFooterAltChunkPackage(
                "afchunk-footer-even.docx",
                CreateEmbeddedDocxAltChunk(includeTable: false),
                includeEvenAndOddHeaders: false);

            using var testStream = new MemoryStream(packageBytes);
            using var reader = new Nedev.FileConverters.DocxToDoc.Format.DocxReader(testStream);

            var model = reader.ReadDocument();

            var section = Assert.Single(model.Sections);
            Assert.False(model.DifferentOddAndEvenPages);
            Assert.Equal("Lead\rInner lead\rInner tail\rTail", section.EvenPagesFooterText);
            Assert.Null(section.FooterText);
        }

        [Fact]
        public void ReadDocument_WithEvenAndOddHeadersSetting_PromotesEvenFooterEmbeddedDocxAltChunkToSummaryText()
        {
            byte[] packageBytes = CreateEvenFooterAltChunkPackage(
                "afchunk-footer-even.docx",
                CreateEmbeddedDocxAltChunk(includeTable: false),
                includeEvenAndOddHeaders: true);

            using var testStream = new MemoryStream(packageBytes);
            using var reader = new Nedev.FileConverters.DocxToDoc.Format.DocxReader(testStream);

            var model = reader.ReadDocument();

            var section = Assert.Single(model.Sections);
            Assert.True(model.DifferentOddAndEvenPages);
            Assert.Equal("Lead\rInner lead\rInner tail\rTail", section.EvenPagesFooterText);
            Assert.Equal("Lead\rInner lead\rInner tail\rTail", section.FooterText);
        }

        [Fact]
        public void ReadDocument_WithoutEvenAndOddHeadersSetting_DoesNotPromoteEvenFooterEmbeddedDocxAltChunkInsideTableCellToSummaryText()
        {
            byte[] packageBytes = CreateEvenFooterTableCellAltChunkPackage(
                "afchunk-footer-even-cell.docx",
                CreateEmbeddedDocxAltChunk(includeTable: true),
                includeEvenAndOddHeaders: false);

            using var testStream = new MemoryStream(packageBytes);
            using var reader = new Nedev.FileConverters.DocxToDoc.Format.DocxReader(testStream);

            var model = reader.ReadDocument();

            var section = Assert.Single(model.Sections);
            Assert.False(model.DifferentOddAndEvenPages);
            Assert.Equal("Lead\rInner lead\rCell\rInner tail\rTail", section.EvenPagesFooterText);
            Assert.Null(section.FooterText);
        }

        [Fact]
        public void ReadDocument_WithEvenAndOddHeadersSetting_PromotesEvenFooterEmbeddedDocxAltChunkInsideTableCellToSummaryText()
        {
            byte[] packageBytes = CreateEvenFooterTableCellAltChunkPackage(
                "afchunk-footer-even-cell.docx",
                CreateEmbeddedDocxAltChunk(includeTable: true),
                includeEvenAndOddHeaders: true);

            using var testStream = new MemoryStream(packageBytes);
            using var reader = new Nedev.FileConverters.DocxToDoc.Format.DocxReader(testStream);

            var model = reader.ReadDocument();

            var section = Assert.Single(model.Sections);
            Assert.True(model.DifferentOddAndEvenPages);
            Assert.Equal("Lead\rInner lead\rCell\rInner tail\rTail", section.EvenPagesFooterText);
            Assert.Equal("Lead\rInner lead\rCell\rInner tail\rTail", section.FooterText);
        }

        [Fact]
        public void ReadDocument_WithHeaderEmbeddedDocxAltChunkMissingRelationship_IgnoresChunkAndPreservesSurroundingParagraphs()
        {
            byte[] packageBytes = CreateDefaultHeaderAltChunkPackage(
                "afchunk-header.docx",
                CreateEmbeddedDocxAltChunk(includeTable: false),
                includeChunkRelationship: false);

            using var testStream = new MemoryStream(packageBytes);
            using var reader = new Nedev.FileConverters.DocxToDoc.Format.DocxReader(testStream);

            var model = reader.ReadDocument();

            var section = Assert.Single(model.Sections);
            Assert.Equal("Lead\rTail", section.DefaultHeaderText);
            Assert.Equal("Lead\rTail", section.HeaderText);

            Assert.NotNull(section.DefaultHeaderStory);
            var story = section.DefaultHeaderStory!;
            Assert.Equal(
                new[] { "Lead", "Tail" },
                story.Content
                    .Select(block => string.Concat(Assert.IsType<ParagraphModel>(block).Runs.Select(run => run.Text)))
                    .ToArray());
        }

        [Fact]
        public void ReadDocument_WithHeaderEmbeddedDocxAltChunkMissingTargetPart_IgnoresChunkAndPreservesSurroundingParagraphs()
        {
            byte[] packageBytes = CreateDefaultHeaderAltChunkPackage(
                "afchunk-header.docx",
                CreateEmbeddedDocxAltChunk(includeTable: false),
                includeChunkPart: false);

            using var testStream = new MemoryStream(packageBytes);
            using var reader = new Nedev.FileConverters.DocxToDoc.Format.DocxReader(testStream);

            var model = reader.ReadDocument();

            var section = Assert.Single(model.Sections);
            Assert.Equal("Lead\rTail", section.DefaultHeaderText);
            Assert.Equal("Lead\rTail", section.HeaderText);

            Assert.NotNull(section.DefaultHeaderStory);
            var story = section.DefaultHeaderStory!;
            Assert.Equal(
                new[] { "Lead", "Tail" },
                story.Content
                    .Select(block => string.Concat(Assert.IsType<ParagraphModel>(block).Runs.Select(run => run.Text)))
                    .ToArray());
        }

        [Fact]
        public void ReadDocument_WithMalformedHeaderEmbeddedDocxAltChunk_IgnoresChunkAndPreservesSurroundingParagraphs()
        {
            byte[] packageBytes = CreateDefaultHeaderAltChunkPackage(
                "afchunk-header.docx",
                Encoding.UTF8.GetBytes("not-a-zip-package"));

            using var testStream = new MemoryStream(packageBytes);
            using var reader = new Nedev.FileConverters.DocxToDoc.Format.DocxReader(testStream);

            var model = reader.ReadDocument();

            var section = Assert.Single(model.Sections);
            Assert.Equal("Lead\rTail", section.DefaultHeaderText);
            Assert.Equal("Lead\rTail", section.HeaderText);

            Assert.NotNull(section.DefaultHeaderStory);
            var story = section.DefaultHeaderStory!;
            Assert.Equal(
                new[] { "Lead", "Tail" },
                story.Content
                    .Select(block => string.Concat(Assert.IsType<ParagraphModel>(block).Runs.Select(run => run.Text)))
                    .ToArray());
        }

        [Fact]
        public void ReadDocument_WithHeaderEmbeddedDocxAltChunkMissingDocumentXml_IgnoresChunkAndPreservesSurroundingParagraphs()
        {
            byte[] packageBytes = CreateDefaultHeaderAltChunkPackage(
                "afchunk-header.docx",
                CreateEmbeddedDocxAltChunkWithoutDocumentXml());

            using var testStream = new MemoryStream(packageBytes);
            using var reader = new Nedev.FileConverters.DocxToDoc.Format.DocxReader(testStream);

            var model = reader.ReadDocument();

            var section = Assert.Single(model.Sections);
            Assert.Equal("Lead\rTail", section.DefaultHeaderText);
            Assert.Equal("Lead\rTail", section.HeaderText);

            Assert.NotNull(section.DefaultHeaderStory);
            var story = section.DefaultHeaderStory!;
            Assert.Equal(
                new[] { "Lead", "Tail" },
                story.Content
                    .Select(block => string.Concat(Assert.IsType<ParagraphModel>(block).Runs.Select(run => run.Text)))
                    .ToArray());
        }

        [Fact]
        public void ReadDocument_WithHeaderAltChunkMissingRelationship_IgnoresChunkAndPreservesSurroundingParagraphs()
        {
            byte[] packageBytes = CreateDefaultHeaderAltChunkPackage(
                "afchunk-header.txt",
                "Chunk line 1\r\nChunk line 2",
                includeChunkRelationship: false);

            using var testStream = new MemoryStream(packageBytes);
            using var reader = new Nedev.FileConverters.DocxToDoc.Format.DocxReader(testStream);

            var model = reader.ReadDocument();

            var section = Assert.Single(model.Sections);
            Assert.Equal("Lead\rTail", section.DefaultHeaderText);
            Assert.Equal("Lead\rTail", section.HeaderText);

            Assert.NotNull(section.DefaultHeaderStory);
            var story = section.DefaultHeaderStory!;
            Assert.Equal(
                new[] { "Lead", "Tail" },
                story.Content
                    .Select(block => string.Concat(Assert.IsType<ParagraphModel>(block).Runs.Select(run => run.Text)))
                    .ToArray());
        }

        [Fact]
        public void ReadDocument_WithHeaderAltChunkMissingTargetPart_IgnoresChunkAndPreservesSurroundingParagraphs()
        {
            byte[] packageBytes = CreateDefaultHeaderAltChunkPackage(
                "afchunk-header.txt",
                "Chunk line 1\r\nChunk line 2",
                includeChunkPart: false);

            using var testStream = new MemoryStream(packageBytes);
            using var reader = new Nedev.FileConverters.DocxToDoc.Format.DocxReader(testStream);

            var model = reader.ReadDocument();

            var section = Assert.Single(model.Sections);
            Assert.Equal("Lead\rTail", section.DefaultHeaderText);
            Assert.Equal("Lead\rTail", section.HeaderText);

            Assert.NotNull(section.DefaultHeaderStory);
            var story = section.DefaultHeaderStory!;
            Assert.Equal(
                new[] { "Lead", "Tail" },
                story.Content
                    .Select(block => string.Concat(Assert.IsType<ParagraphModel>(block).Runs.Select(run => run.Text)))
                    .ToArray());
        }

        [Fact]
        public void ReadDocument_WithUnsupportedHeaderAltChunkType_IgnoresChunkAndPreservesSurroundingParagraphs()
        {
            byte[] packageBytes = CreateDefaultHeaderAltChunkPackage("afchunk-header.bin", "Chunk line 1\r\nChunk line 2");

            using var testStream = new MemoryStream(packageBytes);
            using var reader = new Nedev.FileConverters.DocxToDoc.Format.DocxReader(testStream);

            var model = reader.ReadDocument();

            var section = Assert.Single(model.Sections);
            Assert.Equal("Lead\rTail", section.DefaultHeaderText);
            Assert.Equal("Lead\rTail", section.HeaderText);

            Assert.NotNull(section.DefaultHeaderStory);
            var story = section.DefaultHeaderStory!;
            Assert.Equal(
                new[] { "Lead", "Tail" },
                story.Content
                    .Select(block => string.Concat(Assert.IsType<ParagraphModel>(block).Runs.Select(run => run.Text)))
                    .ToArray());
        }

        [Fact]
        public void ReadDocument_WithHeaderAltChunkBetweenParagraphAndTable_PreservesMixedContentOrder()
        {
            byte[] packageBytes = CreateHeaderAltChunkAndTablePackage("afchunk-header.txt", "Chunk line 1\r\nChunk line 2");

            using var testStream = new MemoryStream(packageBytes);
            using var reader = new Nedev.FileConverters.DocxToDoc.Format.DocxReader(testStream);

            var model = reader.ReadDocument();

            var section = Assert.Single(model.Sections);
            Assert.Equal("Lead\rChunk line 1\rChunk line 2\rCell\rTail", section.DefaultHeaderText);
            Assert.Equal("Lead\rChunk line 1\rChunk line 2\rCell\rTail", section.HeaderText);

            Assert.NotNull(section.DefaultHeaderStory);
            var story = section.DefaultHeaderStory!;
            Assert.Equal(5, story.Content.Count);

            Assert.Collection(
                story.Content,
                block => Assert.Equal("Lead", string.Concat(Assert.IsType<ParagraphModel>(block).Runs.Select(run => run.Text))),
                block => Assert.Equal("Chunk line 1", string.Concat(Assert.IsType<ParagraphModel>(block).Runs.Select(run => run.Text))),
                block => Assert.Equal("Chunk line 2", string.Concat(Assert.IsType<ParagraphModel>(block).Runs.Select(run => run.Text))),
                block =>
                {
                    var table = Assert.IsType<TableModel>(block);
                    var cellRun = Assert.Single(table.Rows[0].Cells[0].Paragraphs[0].Runs.FindAll(run => run.Text.Length > 0));
                    Assert.Equal("Cell", cellRun.Text);
                },
                block => Assert.Equal("Tail", string.Concat(Assert.IsType<ParagraphModel>(block).Runs.Select(run => run.Text))));
        }

        private static byte[] CreateDefaultHeaderAltChunkPackage(
            string chunkFileName,
            string chunkContent,
            bool includeChunkRelationship = true,
            bool includeChunkPart = true)
        {
            return CreateDefaultHeaderAltChunkPackage(
                chunkFileName,
                Encoding.UTF8.GetBytes(chunkContent),
                includeChunkRelationship,
                includeChunkPart);
        }

        private static byte[] CreateDefaultHeaderAltChunkPackage(
            string chunkFileName,
            byte[] chunkBytes,
            bool includeChunkRelationship = true,
            bool includeChunkPart = true)
        {
            using var ms = new MemoryStream();
            using (var archive = new ZipArchive(ms, ZipArchiveMode.Create, true))
            {
                var documentEntry = archive.CreateEntry("word/document.xml");
                using (var documentStream = documentEntry.Open())
                using (var writer = new StreamWriter(documentStream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" +
                                 "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">" +
                                 "<w:body>" +
                                 "<w:p><w:r><w:t>AB</w:t></w:r></w:p>" +
                                 "<w:sectPr>" +
                                 "<w:headerReference w:type=\"default\" r:id=\"rIdHeader\"/>" +
                                 "</w:sectPr>" +
                                 "</w:body>" +
                                 "</w:document>");
                }

                var relsEntry = archive.CreateEntry("word/_rels/document.xml.rels");
                using (var relsStream = relsEntry.Open())
                using (var writer = new StreamWriter(relsStream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" +
                                 "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
                                 "<Relationship Id=\"rIdHeader\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/header\" Target=\"header1.xml\" />" +
                                 "</Relationships>");
                }

                var headerEntry = archive.CreateEntry("word/header1.xml");
                using (var headerStream = headerEntry.Open())
                using (var writer = new StreamWriter(headerStream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" +
                                 "<w:hdr xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">" +
                                 "<w:p><w:r><w:t>Lead</w:t></w:r></w:p>" +
                                 "<w:altChunk r:id=\"rIdChunk\"/>" +
                                 "<w:p><w:r><w:t>Tail</w:t></w:r></w:p>" +
                                 "</w:hdr>");
                }

                var headerRelsEntry = archive.CreateEntry("word/_rels/header1.xml.rels");
                using (var relsStream = headerRelsEntry.Open())
                using (var writer = new StreamWriter(relsStream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                                 "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">");
                    if (includeChunkRelationship)
                    {
                        writer.Write($"<Relationship Id=\"rIdChunk\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/aFChunk\" Target=\"{chunkFileName}\"/>");
                    }

                    writer.Write("</Relationships>");
                }

                if (includeChunkPart)
                {
                    var chunkEntry = archive.CreateEntry($"word/{chunkFileName}");
                    using (var chunkStream = chunkEntry.Open())
                    {
                        chunkStream.Write(chunkBytes, 0, chunkBytes.Length);
                    }
                }
            }

            return ms.ToArray();
        }

        private static byte[] CreateHeaderAltChunkAndTablePackage(string chunkFileName, string chunkContent)
        {
            return CreateHeaderAltChunkAndTablePackage(chunkFileName, Encoding.UTF8.GetBytes(chunkContent));
        }

        private static byte[] CreateHeaderAltChunkAndTablePackage(string chunkFileName, byte[] chunkBytes)
        {
            using var ms = new MemoryStream();
            using (var archive = new ZipArchive(ms, ZipArchiveMode.Create, true))
            {
                var documentEntry = archive.CreateEntry("word/document.xml");
                using (var documentStream = documentEntry.Open())
                using (var writer = new StreamWriter(documentStream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" +
                                 "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">" +
                                 "<w:body>" +
                                 "<w:p><w:r><w:t>AB</w:t></w:r></w:p>" +
                                 "<w:sectPr>" +
                                 "<w:headerReference w:type=\"default\" r:id=\"rIdHeader\"/>" +
                                 "</w:sectPr>" +
                                 "</w:body>" +
                                 "</w:document>");
                }

                var relsEntry = archive.CreateEntry("word/_rels/document.xml.rels");
                using (var relsStream = relsEntry.Open())
                using (var writer = new StreamWriter(relsStream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" +
                                 "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
                                 "<Relationship Id=\"rIdHeader\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/header\" Target=\"header1.xml\" />" +
                                 "</Relationships>");
                }

                var headerEntry = archive.CreateEntry("word/header1.xml");
                using (var headerStream = headerEntry.Open())
                using (var writer = new StreamWriter(headerStream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" +
                                 "<w:hdr xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">" +
                                 "<w:p><w:r><w:t>Lead</w:t></w:r></w:p>" +
                                 "<w:altChunk r:id=\"rIdChunk\"/>" +
                                 "<w:tbl>" +
                                 "<w:tblGrid><w:gridCol w:w=\"2400\"/></w:tblGrid>" +
                                 "<w:tr>" +
                                 "<w:tc>" +
                                 "<w:tcPr><w:tcW w:type=\"dxa\" w:w=\"2400\"/></w:tcPr>" +
                                 "<w:p><w:r><w:t>Cell</w:t></w:r></w:p>" +
                                 "</w:tc>" +
                                 "</w:tr>" +
                                 "</w:tbl>" +
                                 "<w:p><w:r><w:t>Tail</w:t></w:r></w:p>" +
                                 "</w:hdr>");
                }

                var headerRelsEntry = archive.CreateEntry("word/_rels/header1.xml.rels");
                using (var relsStream = headerRelsEntry.Open())
                using (var writer = new StreamWriter(relsStream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                                 "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
                                 $"<Relationship Id=\"rIdChunk\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/aFChunk\" Target=\"{chunkFileName}\"/>" +
                                 "</Relationships>");
                }

                var chunkEntry = archive.CreateEntry($"word/{chunkFileName}");
                using (var chunkStream = chunkEntry.Open())
                {
                    chunkStream.Write(chunkBytes, 0, chunkBytes.Length);
                }
            }

            return ms.ToArray();
        }

        private static byte[] CreateDefaultFooterAltChunkPackage(
            string chunkFileName,
            byte[] chunkBytes,
            bool includeChunkRelationship = true,
            bool includeChunkPart = true)
        {
            using var ms = new MemoryStream();
            using (var archive = new ZipArchive(ms, ZipArchiveMode.Create, true))
            {
                var documentEntry = archive.CreateEntry("word/document.xml");
                using (var documentStream = documentEntry.Open())
                using (var writer = new StreamWriter(documentStream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" +
                                 "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">" +
                                 "<w:body>" +
                                 "<w:p><w:r><w:t>AB</w:t></w:r></w:p>" +
                                 "<w:sectPr>" +
                                 "<w:footerReference w:type=\"default\" r:id=\"rIdFooter\"/>" +
                                 "</w:sectPr>" +
                                 "</w:body>" +
                                 "</w:document>");
                }

                var relsEntry = archive.CreateEntry("word/_rels/document.xml.rels");
                using (var relsStream = relsEntry.Open())
                using (var writer = new StreamWriter(relsStream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" +
                                 "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
                                 "<Relationship Id=\"rIdFooter\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer\" Target=\"footer1.xml\" />" +
                                 "</Relationships>");
                }

                var footerEntry = archive.CreateEntry("word/footer1.xml");
                using (var footerStream = footerEntry.Open())
                using (var writer = new StreamWriter(footerStream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" +
                                 "<w:ftr xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">" +
                                 "<w:p><w:r><w:t>Lead</w:t></w:r></w:p>" +
                                 "<w:altChunk r:id=\"rIdChunk\"/>" +
                                 "<w:p><w:r><w:t>Tail</w:t></w:r></w:p>" +
                                 "</w:ftr>");
                }

                var footerRelsEntry = archive.CreateEntry("word/_rels/footer1.xml.rels");
                using (var relsStream = footerRelsEntry.Open())
                using (var writer = new StreamWriter(relsStream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                                 "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">");
                    if (includeChunkRelationship)
                    {
                        writer.Write($"<Relationship Id=\"rIdChunk\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/aFChunk\" Target=\"{chunkFileName}\"/>");
                    }

                    writer.Write("</Relationships>");
                }

                if (includeChunkPart)
                {
                    var chunkEntry = archive.CreateEntry($"word/{chunkFileName}");
                    using (var chunkStream = chunkEntry.Open())
                    {
                        chunkStream.Write(chunkBytes, 0, chunkBytes.Length);
                    }
                }
            }

            return ms.ToArray();
        }

        private static byte[] CreateFooterAltChunkAndTablePackage(string chunkFileName, byte[] chunkBytes)
        {
            using var ms = new MemoryStream();
            using (var archive = new ZipArchive(ms, ZipArchiveMode.Create, true))
            {
                var documentEntry = archive.CreateEntry("word/document.xml");
                using (var documentStream = documentEntry.Open())
                using (var writer = new StreamWriter(documentStream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" +
                                 "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">" +
                                 "<w:body>" +
                                 "<w:p><w:r><w:t>AB</w:t></w:r></w:p>" +
                                 "<w:sectPr>" +
                                 "<w:footerReference w:type=\"default\" r:id=\"rIdFooter\"/>" +
                                 "</w:sectPr>" +
                                 "</w:body>" +
                                 "</w:document>");
                }

                var relsEntry = archive.CreateEntry("word/_rels/document.xml.rels");
                using (var relsStream = relsEntry.Open())
                using (var writer = new StreamWriter(relsStream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" +
                                 "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
                                 "<Relationship Id=\"rIdFooter\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer\" Target=\"footer1.xml\" />" +
                                 "</Relationships>");
                }

                var footerEntry = archive.CreateEntry("word/footer1.xml");
                using (var footerStream = footerEntry.Open())
                using (var writer = new StreamWriter(footerStream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" +
                                 "<w:ftr xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">" +
                                 "<w:p><w:r><w:t>Lead</w:t></w:r></w:p>" +
                                 "<w:altChunk r:id=\"rIdChunk\"/>" +
                                 "<w:tbl>" +
                                 "<w:tblGrid><w:gridCol w:w=\"2400\"/></w:tblGrid>" +
                                 "<w:tr>" +
                                 "<w:tc>" +
                                 "<w:tcPr><w:tcW w:type=\"dxa\" w:w=\"2400\"/></w:tcPr>" +
                                 "<w:p><w:r><w:t>Cell</w:t></w:r></w:p>" +
                                 "</w:tc>" +
                                 "</w:tr>" +
                                 "</w:tbl>" +
                                 "<w:p><w:r><w:t>Tail</w:t></w:r></w:p>" +
                                 "</w:ftr>");
                }

                var footerRelsEntry = archive.CreateEntry("word/_rels/footer1.xml.rels");
                using (var relsStream = footerRelsEntry.Open())
                using (var writer = new StreamWriter(relsStream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                                 "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
                                 $"<Relationship Id=\"rIdChunk\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/aFChunk\" Target=\"{chunkFileName}\"/>" +
                                 "</Relationships>");
                }

                var chunkEntry = archive.CreateEntry($"word/{chunkFileName}");
                using (var chunkStream = chunkEntry.Open())
                {
                    chunkStream.Write(chunkBytes, 0, chunkBytes.Length);
                }
            }

            return ms.ToArray();
        }

        private static byte[] CreateHeaderTableCellAltChunkPackage(string chunkFileName, byte[] chunkBytes)
        {
            using var ms = new MemoryStream();
            using (var archive = new ZipArchive(ms, ZipArchiveMode.Create, true))
            {
                var documentEntry = archive.CreateEntry("word/document.xml");
                using (var documentStream = documentEntry.Open())
                using (var writer = new StreamWriter(documentStream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" +
                                 "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">" +
                                 "<w:body>" +
                                 "<w:p><w:r><w:t>AB</w:t></w:r></w:p>" +
                                 "<w:sectPr>" +
                                 "<w:headerReference w:type=\"default\" r:id=\"rIdHeader\"/>" +
                                 "</w:sectPr>" +
                                 "</w:body>" +
                                 "</w:document>");
                }

                var relsEntry = archive.CreateEntry("word/_rels/document.xml.rels");
                using (var relsStream = relsEntry.Open())
                using (var writer = new StreamWriter(relsStream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" +
                                 "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
                                 "<Relationship Id=\"rIdHeader\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/header\" Target=\"header1.xml\" />" +
                                 "</Relationships>");
                }

                var headerEntry = archive.CreateEntry("word/header1.xml");
                using (var headerStream = headerEntry.Open())
                using (var writer = new StreamWriter(headerStream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" +
                                 "<w:hdr xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">" +
                                 "<w:p><w:r><w:t>Lead</w:t></w:r></w:p>" +
                                 "<w:tbl>" +
                                 "<w:tr>" +
                                 "<w:tc>" +
                                 "<w:altChunk r:id=\"rIdChunk\"/>" +
                                 "</w:tc>" +
                                 "</w:tr>" +
                                 "</w:tbl>" +
                                 "<w:p><w:r><w:t>Tail</w:t></w:r></w:p>" +
                                 "</w:hdr>");
                }

                var headerRelsEntry = archive.CreateEntry("word/_rels/header1.xml.rels");
                using (var relsStream = headerRelsEntry.Open())
                using (var writer = new StreamWriter(relsStream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                                 "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
                                 $"<Relationship Id=\"rIdChunk\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/aFChunk\" Target=\"{chunkFileName}\"/>" +
                                 "</Relationships>");
                }

                var chunkEntry = archive.CreateEntry($"word/{chunkFileName}");
                using (var chunkStream = chunkEntry.Open())
                {
                    chunkStream.Write(chunkBytes, 0, chunkBytes.Length);
                }
            }

            return ms.ToArray();
        }

        private static byte[] CreateFooterTableCellAltChunkPackage(
            string chunkFileName,
            byte[] chunkBytes,
            bool includeChunkRelationship = true,
            bool includeChunkPart = true,
            bool includeSurroundingCellParagraphs = false)
        {
            using var ms = new MemoryStream();
            using (var archive = new ZipArchive(ms, ZipArchiveMode.Create, true))
            {
                var documentEntry = archive.CreateEntry("word/document.xml");
                using (var documentStream = documentEntry.Open())
                using (var writer = new StreamWriter(documentStream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" +
                                 "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">" +
                                 "<w:body>" +
                                 "<w:p><w:r><w:t>AB</w:t></w:r></w:p>" +
                                 "<w:sectPr>" +
                                 "<w:footerReference w:type=\"default\" r:id=\"rIdFooter\"/>" +
                                 "</w:sectPr>" +
                                 "</w:body>" +
                                 "</w:document>");
                }

                var relsEntry = archive.CreateEntry("word/_rels/document.xml.rels");
                using (var relsStream = relsEntry.Open())
                using (var writer = new StreamWriter(relsStream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" +
                                 "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
                                 "<Relationship Id=\"rIdFooter\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer\" Target=\"footer1.xml\" />" +
                                 "</Relationships>");
                }

                var footerEntry = archive.CreateEntry("word/footer1.xml");
                using (var footerStream = footerEntry.Open())
                using (var writer = new StreamWriter(footerStream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" +
                                 "<w:ftr xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">" +
                                 "<w:p><w:r><w:t>Lead</w:t></w:r></w:p>" +
                                 "<w:tbl><w:tr><w:tc>");

                    if (includeSurroundingCellParagraphs)
                    {
                        writer.Write("<w:p><w:r><w:t>Cell lead</w:t></w:r></w:p>");
                    }

                    writer.Write("<w:altChunk r:id=\"rIdChunk\"/>");

                    if (includeSurroundingCellParagraphs)
                    {
                        writer.Write("<w:p><w:r><w:t>Cell tail</w:t></w:r></w:p>");
                    }

                    writer.Write("</w:tc></w:tr></w:tbl>" +
                                 "<w:p><w:r><w:t>Tail</w:t></w:r></w:p>" +
                                 "</w:ftr>");
                }

                var footerRelsEntry = archive.CreateEntry("word/_rels/footer1.xml.rels");
                using (var relsStream = footerRelsEntry.Open())
                using (var writer = new StreamWriter(relsStream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                                 "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">");
                    if (includeChunkRelationship)
                    {
                        writer.Write($"<Relationship Id=\"rIdChunk\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/aFChunk\" Target=\"{chunkFileName}\"/>");
                    }

                    writer.Write("</Relationships>");
                }

                if (includeChunkPart)
                {
                    var chunkEntry = archive.CreateEntry($"word/{chunkFileName}");
                    using (var chunkStream = chunkEntry.Open())
                    {
                        chunkStream.Write(chunkBytes, 0, chunkBytes.Length);
                    }
                }
            }

            return ms.ToArray();
        }

        private static byte[] CreateEvenFooterAltChunkPackage(string chunkFileName, byte[] chunkBytes, bool includeEvenAndOddHeaders)
        {
            using var ms = new MemoryStream();
            using (var archive = new ZipArchive(ms, ZipArchiveMode.Create, true))
            {
                var documentEntry = archive.CreateEntry("word/document.xml");
                using (var documentStream = documentEntry.Open())
                using (var writer = new StreamWriter(documentStream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" +
                                 "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">" +
                                 "<w:body>" +
                                 "<w:p><w:r><w:t>AB</w:t></w:r></w:p>" +
                                 "<w:sectPr>" +
                                 "<w:footerReference w:type=\"even\" r:id=\"rIdFooterEven\"/>" +
                                 "</w:sectPr>" +
                                 "</w:body>" +
                                 "</w:document>");
                }

                var relsEntry = archive.CreateEntry("word/_rels/document.xml.rels");
                using (var relsStream = relsEntry.Open())
                using (var writer = new StreamWriter(relsStream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" +
                                 "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
                                 "<Relationship Id=\"rIdFooterEven\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer\" Target=\"footer-even.xml\" />" +
                                 "</Relationships>");
                }

                if (includeEvenAndOddHeaders)
                {
                    var settingsEntry = archive.CreateEntry("word/settings.xml");
                    using (var settingsStream = settingsEntry.Open())
                    using (var writer = new StreamWriter(settingsStream))
                    {
                        writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" +
                                     "<w:settings xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">" +
                                     "<w:evenAndOddHeaders/>" +
                                     "</w:settings>");
                    }
                }

                var footerEntry = archive.CreateEntry("word/footer-even.xml");
                using (var footerStream = footerEntry.Open())
                using (var writer = new StreamWriter(footerStream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" +
                                 "<w:ftr xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">" +
                                 "<w:p><w:r><w:t>Lead</w:t></w:r></w:p>" +
                                 "<w:altChunk r:id=\"rIdChunk\"/>" +
                                 "<w:p><w:r><w:t>Tail</w:t></w:r></w:p>" +
                                 "</w:ftr>");
                }

                var footerRelsEntry = archive.CreateEntry("word/_rels/footer-even.xml.rels");
                using (var relsStream = footerRelsEntry.Open())
                using (var writer = new StreamWriter(relsStream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                                 "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
                                 $"<Relationship Id=\"rIdChunk\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/aFChunk\" Target=\"{chunkFileName}\"/>" +
                                 "</Relationships>");
                }

                var chunkEntry = archive.CreateEntry($"word/{chunkFileName}");
                using (var chunkStream = chunkEntry.Open())
                {
                    chunkStream.Write(chunkBytes, 0, chunkBytes.Length);
                }
            }

            return ms.ToArray();
        }

        private static byte[] CreateEvenFooterTableCellAltChunkPackage(string chunkFileName, byte[] chunkBytes, bool includeEvenAndOddHeaders)
        {
            using var ms = new MemoryStream();
            using (var archive = new ZipArchive(ms, ZipArchiveMode.Create, true))
            {
                var documentEntry = archive.CreateEntry("word/document.xml");
                using (var documentStream = documentEntry.Open())
                using (var writer = new StreamWriter(documentStream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" +
                                 "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">" +
                                 "<w:body>" +
                                 "<w:p><w:r><w:t>AB</w:t></w:r></w:p>" +
                                 "<w:sectPr>" +
                                 "<w:footerReference w:type=\"even\" r:id=\"rIdFooterEven\"/>" +
                                 "</w:sectPr>" +
                                 "</w:body>" +
                                 "</w:document>");
                }

                var relsEntry = archive.CreateEntry("word/_rels/document.xml.rels");
                using (var relsStream = relsEntry.Open())
                using (var writer = new StreamWriter(relsStream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" +
                                 "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
                                 "<Relationship Id=\"rIdFooterEven\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer\" Target=\"footer-even.xml\" />" +
                                 "</Relationships>");
                }

                if (includeEvenAndOddHeaders)
                {
                    var settingsEntry = archive.CreateEntry("word/settings.xml");
                    using (var settingsStream = settingsEntry.Open())
                    using (var writer = new StreamWriter(settingsStream))
                    {
                        writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" +
                                     "<w:settings xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">" +
                                     "<w:evenAndOddHeaders/>" +
                                     "</w:settings>");
                    }
                }

                var footerEntry = archive.CreateEntry("word/footer-even.xml");
                using (var footerStream = footerEntry.Open())
                using (var writer = new StreamWriter(footerStream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" +
                                 "<w:ftr xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">" +
                                 "<w:p><w:r><w:t>Lead</w:t></w:r></w:p>" +
                                 "<w:tbl><w:tr><w:tc><w:altChunk r:id=\"rIdChunk\"/></w:tc></w:tr></w:tbl>" +
                                 "<w:p><w:r><w:t>Tail</w:t></w:r></w:p>" +
                                 "</w:ftr>");
                }

                var footerRelsEntry = archive.CreateEntry("word/_rels/footer-even.xml.rels");
                using (var relsStream = footerRelsEntry.Open())
                using (var writer = new StreamWriter(relsStream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                                 "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
                                 $"<Relationship Id=\"rIdChunk\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/aFChunk\" Target=\"{chunkFileName}\"/>" +
                                 "</Relationships>");
                }

                var chunkEntry = archive.CreateEntry($"word/{chunkFileName}");
                using (var chunkStream = chunkEntry.Open())
                {
                    chunkStream.Write(chunkBytes, 0, chunkBytes.Length);
                }
            }

            return ms.ToArray();
        }

        private static byte[] CreateEmbeddedDocxAltChunk(bool includeTable)
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

        private static byte[] CreateEmbeddedDocxAltChunkWithoutDocumentXml()
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

        private static byte[] GetTestPngBytes()
        {
            return new byte[]
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
        }
    }
}
