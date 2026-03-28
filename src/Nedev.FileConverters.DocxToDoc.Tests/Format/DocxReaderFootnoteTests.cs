using System.IO;
using System.IO.Compression;
using Xunit;

namespace Nedev.FileConverters.DocxToDoc.Tests.Format
{
    public class DocxReaderFootnoteTests
    {
        [Fact]
        public void ReadDocument_WithFootnoteReference_ParsesFootnoteTextAndReferenceCp()
        {
            using var ms = new MemoryStream();
            using (var archive = new ZipArchive(ms, ZipArchiveMode.Create, true))
            {
                var documentEntry = archive.CreateEntry("word/document.xml");
                using (var documentStream = documentEntry.Open())
                using (var writer = new StreamWriter(documentStream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" +
                                 "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:body><w:p>" +
                                 "<w:r><w:t>AB</w:t></w:r>" +
                                 "<w:r><w:footnoteReference w:id=\"2\"/></w:r>" +
                                 "<w:r><w:t>CD</w:t></w:r>" +
                                 "</w:p></w:body></w:document>");
                }

                var footnotesEntry = archive.CreateEntry("word/footnotes.xml");
                using (var footnotesStream = footnotesEntry.Open())
                using (var writer = new StreamWriter(footnotesStream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" +
                                 "<w:footnotes xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">" +
                                 "<w:footnote w:id=\"0\" w:type=\"separator\"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>" +
                                 "<w:footnote w:id=\"2\"><w:p><w:r><w:footnoteRef/></w:r><w:r><w:t>Note</w:t></w:r><w:r><w:tab/><w:t>More</w:t></w:r></w:p></w:footnote>" +
                                 "</w:footnotes>");
                }
            }

            using var testStream = new MemoryStream(ms.ToArray());
            using var reader = new Nedev.FileConverters.DocxToDoc.Format.DocxReader(testStream);

            var model = reader.ReadDocument();

            var footnote = Assert.Single(model.Footnotes);
            Assert.Equal("2", footnote.Id);
            Assert.Equal("Note\tMore", footnote.Text);
            Assert.Equal(2, footnote.ReferenceCp);
            Assert.Equal("ABCD\r", model.TextBuffer);
        }

        [Fact]
        public void ReadDocument_WithReservedAndMissingFootnoteDefinitions_IgnoresUnsupportedReferences()
        {
            using var ms = new MemoryStream();
            using (var archive = new ZipArchive(ms, ZipArchiveMode.Create, true))
            {
                var documentEntry = archive.CreateEntry("word/document.xml");
                using (var documentStream = documentEntry.Open())
                using (var writer = new StreamWriter(documentStream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" +
                                 "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:body><w:p>" +
                                 "<w:r><w:footnoteReference w:id=\"0\"/></w:r>" +
                                 "<w:r><w:t>AB</w:t></w:r>" +
                                 "<w:r><w:footnoteReference w:id=\"9\"/></w:r>" +
                                 "</w:p></w:body></w:document>");
                }

                var footnotesEntry = archive.CreateEntry("word/footnotes.xml");
                using (var footnotesStream = footnotesEntry.Open())
                using (var writer = new StreamWriter(footnotesStream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" +
                                 "<w:footnotes xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">" +
                                 "<w:footnote w:id=\"0\"><w:p><w:r><w:t>Reserved separator fallback</w:t></w:r></w:p></w:footnote>" +
                                 "<w:footnote w:id=\"2\"><w:p><w:r><w:t>Live footnote</w:t></w:r></w:p></w:footnote>" +
                                 "</w:footnotes>");
                }
            }

            using var testStream = new MemoryStream(ms.ToArray());
            using var reader = new Nedev.FileConverters.DocxToDoc.Format.DocxReader(testStream);

            var model = reader.ReadDocument();

            var footnote = Assert.Single(model.Footnotes);
            Assert.Equal("2", footnote.Id);
            Assert.Equal("Live footnote", footnote.Text);
            Assert.Equal(-1, footnote.ReferenceCp);
            Assert.Equal("AB\r", model.TextBuffer);
        }

        [Fact]
        public void ReadDocument_WithMultiParagraphFootnote_PreservesParagraphBreaksInNoteText()
        {
            using var ms = new MemoryStream();
            using (var archive = new ZipArchive(ms, ZipArchiveMode.Create, true))
            {
                var documentEntry = archive.CreateEntry("word/document.xml");
                using (var documentStream = documentEntry.Open())
                using (var writer = new StreamWriter(documentStream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" +
                                 "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:body><w:p>" +
                                 "<w:r><w:t>AB</w:t></w:r>" +
                                 "<w:r><w:footnoteReference w:id=\"2\"/></w:r>" +
                                 "<w:r><w:t>CD</w:t></w:r>" +
                                 "</w:p></w:body></w:document>");
                }

                var footnotesEntry = archive.CreateEntry("word/footnotes.xml");
                using (var footnotesStream = footnotesEntry.Open())
                using (var writer = new StreamWriter(footnotesStream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" +
                                 "<w:footnotes xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">" +
                                 "<w:footnote w:id=\"2\">" +
                                 "<w:p><w:r><w:footnoteRef/></w:r><w:r><w:t>First</w:t></w:r></w:p>" +
                                 "<w:p><w:r><w:t>Second</w:t></w:r></w:p>" +
                                 "</w:footnote>" +
                                 "</w:footnotes>");
                }
            }

            using var testStream = new MemoryStream(ms.ToArray());
            using var reader = new Nedev.FileConverters.DocxToDoc.Format.DocxReader(testStream);

            var model = reader.ReadDocument();

            var footnote = Assert.Single(model.Footnotes);
            Assert.Equal("2", footnote.Id);
            Assert.Equal("First\rSecond", footnote.Text);
            Assert.Equal(2, footnote.ReferenceCp);
            Assert.Equal("ABCD\r", model.TextBuffer);
        }

        [Fact]
        public void ReadDocument_WithTrailingEmptyFootnoteParagraph_PreservesTrailingParagraphBreak()
        {
            using var ms = new MemoryStream();
            using (var archive = new ZipArchive(ms, ZipArchiveMode.Create, true))
            {
                var documentEntry = archive.CreateEntry("word/document.xml");
                using (var documentStream = documentEntry.Open())
                using (var writer = new StreamWriter(documentStream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" +
                                 "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:body><w:p>" +
                                 "<w:r><w:t>AB</w:t></w:r>" +
                                 "<w:r><w:footnoteReference w:id=\"2\"/></w:r>" +
                                 "<w:r><w:t>CD</w:t></w:r>" +
                                 "</w:p></w:body></w:document>");
                }

                var footnotesEntry = archive.CreateEntry("word/footnotes.xml");
                using (var footnotesStream = footnotesEntry.Open())
                using (var writer = new StreamWriter(footnotesStream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" +
                                 "<w:footnotes xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">" +
                                 "<w:footnote w:id=\"2\">" +
                                 "<w:p><w:r><w:footnoteRef/></w:r><w:r><w:t>First</w:t></w:r></w:p>" +
                                 "<w:p/>" +
                                 "</w:footnote>" +
                                 "</w:footnotes>");
                }
            }

            using var testStream = new MemoryStream(ms.ToArray());
            using var reader = new Nedev.FileConverters.DocxToDoc.Format.DocxReader(testStream);

            var model = reader.ReadDocument();

            var footnote = Assert.Single(model.Footnotes);
            Assert.Equal("2", footnote.Id);
            Assert.Equal("First\r", footnote.Text);
            Assert.Equal(2, footnote.ReferenceCp);
            Assert.Equal("ABCD\r", model.TextBuffer);
        }

        [Fact]
        public void ReadDocument_WithFootnoteSeparatorStories_ParsesSpecialStoryText()
        {
            using var ms = new MemoryStream();
            using (var archive = new ZipArchive(ms, ZipArchiveMode.Create, true))
            {
                var documentEntry = archive.CreateEntry("word/document.xml");
                using (var documentStream = documentEntry.Open())
                using (var writer = new StreamWriter(documentStream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" +
                                 "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:body><w:p>" +
                                 "<w:r><w:t>AB</w:t></w:r>" +
                                 "</w:p></w:body></w:document>");
                }

                var footnotesEntry = archive.CreateEntry("word/footnotes.xml");
                using (var footnotesStream = footnotesEntry.Open())
                using (var writer = new StreamWriter(footnotesStream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" +
                                 "<w:footnotes xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">" +
                                 "<w:footnote w:id=\"0\" w:type=\"separator\">" +
                                 "<w:p><w:r><w:t>Sep</w:t></w:r></w:p>" +
                                 "<w:p><w:r><w:t>Line</w:t></w:r></w:p>" +
                                 "</w:footnote>" +
                                 "<w:footnote w:id=\"1\" w:type=\"continuationSeparator\">" +
                                 "<w:p><w:r><w:t>Continue</w:t></w:r></w:p>" +
                                 "</w:footnote>" +
                                 "<w:footnote w:id=\"2\"><w:p><w:r><w:t>Live footnote</w:t></w:r></w:p></w:footnote>" +
                                 "</w:footnotes>");
                }
            }

            using var testStream = new MemoryStream(ms.ToArray());
            using var reader = new Nedev.FileConverters.DocxToDoc.Format.DocxReader(testStream);

            var model = reader.ReadDocument();

            Assert.Equal("Sep\rLine", model.FootnoteSeparatorText);
            Assert.Equal("Continue", model.FootnoteContinuationSeparatorText);
            var footnote = Assert.Single(model.Footnotes);
            Assert.Equal("2", footnote.Id);
            Assert.Equal("Live footnote", footnote.Text);
        }

        [Fact]
        public void ReadDocument_WithEmptyAndUnknownSpecialFootnotes_PreservesEmptySeparatorAndIgnoresUnknownTypes()
        {
            using var ms = new MemoryStream();
            using (var archive = new ZipArchive(ms, ZipArchiveMode.Create, true))
            {
                var documentEntry = archive.CreateEntry("word/document.xml");
                using (var documentStream = documentEntry.Open())
                using (var writer = new StreamWriter(documentStream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" +
                                 "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:body><w:p>" +
                                 "<w:r><w:t>AB</w:t></w:r>" +
                                 "</w:p></w:body></w:document>");
                }

                var footnotesEntry = archive.CreateEntry("word/footnotes.xml");
                using (var footnotesStream = footnotesEntry.Open())
                using (var writer = new StreamWriter(footnotesStream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" +
                                 "<w:footnotes xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">" +
                                 "<w:footnote w:id=\"0\" w:type=\"separator\"><w:p/></w:footnote>" +
                                 "<w:footnote w:id=\"1\" w:type=\"customSeparator\"><w:p><w:r><w:t>Ignored</w:t></w:r></w:p></w:footnote>" +
                                 "<w:footnote w:id=\"2\"><w:p><w:r><w:t>Live footnote</w:t></w:r></w:p></w:footnote>" +
                                 "</w:footnotes>");
                }
            }

            using var testStream = new MemoryStream(ms.ToArray());
            using var reader = new Nedev.FileConverters.DocxToDoc.Format.DocxReader(testStream);

            var model = reader.ReadDocument();

            Assert.Equal(string.Empty, model.FootnoteSeparatorText);
            Assert.Null(model.FootnoteContinuationSeparatorText);
            var footnote = Assert.Single(model.Footnotes);
            Assert.Equal("2", footnote.Id);
            Assert.Equal("Live footnote", footnote.Text);
        }

        [Fact]
        public void ReadDocument_WithFootnoteContinuationNotice_ParsesSpecialStoryText()
        {
            using var ms = new MemoryStream();
            using (var archive = new ZipArchive(ms, ZipArchiveMode.Create, true))
            {
                var documentEntry = archive.CreateEntry("word/document.xml");
                using (var documentStream = documentEntry.Open())
                using (var writer = new StreamWriter(documentStream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" +
                                 "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:body><w:p>" +
                                 "<w:r><w:t>AB</w:t></w:r>" +
                                 "</w:p></w:body></w:document>");
                }

                var footnotesEntry = archive.CreateEntry("word/footnotes.xml");
                using (var footnotesStream = footnotesEntry.Open())
                using (var writer = new StreamWriter(footnotesStream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" +
                                 "<w:footnotes xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">" +
                                 "<w:footnote w:id=\"0\" w:type=\"continuationNotice\">" +
                                 "<w:p><w:r><w:t>Notice</w:t></w:r></w:p>" +
                                 "<w:p><w:r><w:t>More</w:t></w:r></w:p>" +
                                 "</w:footnote>" +
                                 "<w:footnote w:id=\"2\"><w:p><w:r><w:t>Live footnote</w:t></w:r></w:p></w:footnote>" +
                                 "</w:footnotes>");
                }
            }

            using var testStream = new MemoryStream(ms.ToArray());
            using var reader = new Nedev.FileConverters.DocxToDoc.Format.DocxReader(testStream);

            var model = reader.ReadDocument();

            Assert.Equal("Notice\rMore", model.FootnoteContinuationNoticeText);
            var footnote = Assert.Single(model.Footnotes);
            Assert.Equal("2", footnote.Id);
            Assert.Equal("Live footnote", footnote.Text);
        }

        [Fact]
        public void ReadDocument_WithFootnoteMultiFragmentCustomMark_ParsesReferenceCpAndVisibleMark()
        {
            using var ms = new MemoryStream();
            using (var archive = new ZipArchive(ms, ZipArchiveMode.Create, true))
            {
                var documentEntry = archive.CreateEntry("word/document.xml");
                using (var documentStream = documentEntry.Open())
                using (var writer = new StreamWriter(documentStream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" +
                                 "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:body><w:p>" +
                                 "<w:r><w:t>AB</w:t></w:r>" +
                                 "<w:r><w:footnoteReference w:id=\"2\" w:customMarkFollows=\"1\"/></w:r>" +
                                 "<w:r><w:rPr><w:vertAlign w:val=\"superscript\"/></w:rPr><w:t>[</w:t></w:r>" +
                                 "<w:r><w:rPr><w:rStyle w:val=\"FootnoteReference\"/></w:rPr><w:t>12</w:t></w:r>" +
                                 "<w:r><w:rPr><w:vertAlign w:val=\"superscript\"/></w:rPr><w:t>]</w:t></w:r>" +
                                 "<w:r><w:t>CD</w:t></w:r>" +
                                 "</w:p></w:body></w:document>");
                }

                var footnotesEntry = archive.CreateEntry("word/footnotes.xml");
                using (var footnotesStream = footnotesEntry.Open())
                using (var writer = new StreamWriter(footnotesStream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" +
                                 "<w:footnotes xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">" +
                                 "<w:footnote w:id=\"2\"><w:p><w:r><w:t>Note</w:t></w:r></w:p></w:footnote>" +
                                 "</w:footnotes>");
                }
            }

            using var testStream = new MemoryStream(ms.ToArray());
            using var reader = new Nedev.FileConverters.DocxToDoc.Format.DocxReader(testStream);

            var model = reader.ReadDocument();

            var footnote = Assert.Single(model.Footnotes);
            Assert.Equal(2, footnote.ReferenceCp);
            Assert.Equal("[12]", footnote.CustomMarkText);
            Assert.Equal("AB[12]CD\r", model.TextBuffer);
        }

        [Fact]
        public void ReadDocument_WithFootnoteCustomMarkButNoVisibleFollower_LeavesCustomMarkUnset()
        {
            using var ms = new MemoryStream();
            using (var archive = new ZipArchive(ms, ZipArchiveMode.Create, true))
            {
                var documentEntry = archive.CreateEntry("word/document.xml");
                using (var documentStream = documentEntry.Open())
                using (var writer = new StreamWriter(documentStream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" +
                                 "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:body><w:p>" +
                                 "<w:r><w:t>AB</w:t></w:r>" +
                                 "<w:r><w:footnoteReference w:id=\"2\" w:customMarkFollows=\"1\"/></w:r>" +
                                 "<w:r><w:t>CD</w:t></w:r>" +
                                 "</w:p></w:body></w:document>");
                }

                var footnotesEntry = archive.CreateEntry("word/footnotes.xml");
                using (var footnotesStream = footnotesEntry.Open())
                using (var writer = new StreamWriter(footnotesStream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" +
                                 "<w:footnotes xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">" +
                                 "<w:footnote w:id=\"2\"><w:p><w:r><w:t>Note</w:t></w:r></w:p></w:footnote>" +
                                 "</w:footnotes>");
                }
            }

            using var testStream = new MemoryStream(ms.ToArray());
            using var reader = new Nedev.FileConverters.DocxToDoc.Format.DocxReader(testStream);

            var model = reader.ReadDocument();

            var footnote = Assert.Single(model.Footnotes);
            Assert.Equal(2, footnote.ReferenceCp);
            Assert.Null(footnote.CustomMarkText);
            Assert.Equal("ABCD\r", model.TextBuffer);
        }
    }
}
