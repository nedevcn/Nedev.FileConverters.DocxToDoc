using System.IO;
using System.IO.Compression;
using Xunit;

namespace Nedev.FileConverters.DocxToDoc.Tests.Format
{
    public class DocxReaderEndnoteTests
    {
        [Fact]
        public void ReadDocument_WithEndnoteReference_ParsesEndnoteTextAndReferenceCp()
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
                                 "<w:r><w:endnoteReference w:id=\"2\"/></w:r>" +
                                 "<w:r><w:t>CD</w:t></w:r>" +
                                 "</w:p></w:body></w:document>");
                }

                var endnotesEntry = archive.CreateEntry("word/endnotes.xml");
                using (var endnotesStream = endnotesEntry.Open())
                using (var writer = new StreamWriter(endnotesStream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" +
                                 "<w:endnotes xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">" +
                                 "<w:endnote w:id=\"0\" w:type=\"separator\"><w:p><w:r><w:separator/></w:r></w:p></w:endnote>" +
                                 "<w:endnote w:id=\"2\"><w:p><w:r><w:endnoteRef/></w:r><w:r><w:t>Note</w:t></w:r><w:r><w:tab/><w:t>More</w:t></w:r></w:p></w:endnote>" +
                                 "</w:endnotes>");
                }
            }

            using var testStream = new MemoryStream(ms.ToArray());
            using var reader = new Nedev.FileConverters.DocxToDoc.Format.DocxReader(testStream);

            var model = reader.ReadDocument();

            var endnote = Assert.Single(model.Endnotes);
            Assert.Equal("2", endnote.Id);
            Assert.Equal("Note\tMore", endnote.Text);
            Assert.Equal(2, endnote.ReferenceCp);
            Assert.Equal("ABCD\r", model.TextBuffer);
        }

        [Fact]
        public void ReadDocument_WithReservedAndMissingEndnoteDefinitions_IgnoresUnsupportedReferences()
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
                                 "<w:r><w:endnoteReference w:id=\"0\"/></w:r>" +
                                 "<w:r><w:t>AB</w:t></w:r>" +
                                 "<w:r><w:endnoteReference w:id=\"9\"/></w:r>" +
                                 "</w:p></w:body></w:document>");
                }

                var endnotesEntry = archive.CreateEntry("word/endnotes.xml");
                using (var endnotesStream = endnotesEntry.Open())
                using (var writer = new StreamWriter(endnotesStream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" +
                                 "<w:endnotes xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">" +
                                 "<w:endnote w:id=\"0\"><w:p><w:r><w:t>Reserved separator fallback</w:t></w:r></w:p></w:endnote>" +
                                 "<w:endnote w:id=\"2\"><w:p><w:r><w:t>Live endnote</w:t></w:r></w:p></w:endnote>" +
                                 "</w:endnotes>");
                }
            }

            using var testStream = new MemoryStream(ms.ToArray());
            using var reader = new Nedev.FileConverters.DocxToDoc.Format.DocxReader(testStream);

            var model = reader.ReadDocument();

            var endnote = Assert.Single(model.Endnotes);
            Assert.Equal("2", endnote.Id);
            Assert.Equal("Live endnote", endnote.Text);
            Assert.Equal(-1, endnote.ReferenceCp);
            Assert.Equal("AB\r", model.TextBuffer);
        }

        [Fact]
        public void ReadDocument_WithMultiParagraphEndnote_PreservesParagraphBreaksInNoteText()
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
                                 "<w:r><w:endnoteReference w:id=\"2\"/></w:r>" +
                                 "<w:r><w:t>CD</w:t></w:r>" +
                                 "</w:p></w:body></w:document>");
                }

                var endnotesEntry = archive.CreateEntry("word/endnotes.xml");
                using (var endnotesStream = endnotesEntry.Open())
                using (var writer = new StreamWriter(endnotesStream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" +
                                 "<w:endnotes xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">" +
                                 "<w:endnote w:id=\"2\">" +
                                 "<w:p><w:r><w:endnoteRef/></w:r><w:r><w:t>First</w:t></w:r></w:p>" +
                                 "<w:p><w:r><w:t>Second</w:t></w:r></w:p>" +
                                 "</w:endnote>" +
                                 "</w:endnotes>");
                }
            }

            using var testStream = new MemoryStream(ms.ToArray());
            using var reader = new Nedev.FileConverters.DocxToDoc.Format.DocxReader(testStream);

            var model = reader.ReadDocument();

            var endnote = Assert.Single(model.Endnotes);
            Assert.Equal("2", endnote.Id);
            Assert.Equal("First\rSecond", endnote.Text);
            Assert.Equal(2, endnote.ReferenceCp);
            Assert.Equal("ABCD\r", model.TextBuffer);
        }

        [Fact]
        public void ReadDocument_WithEndnoteSeparatorStories_ParsesSpecialStoryText()
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

                var endnotesEntry = archive.CreateEntry("word/endnotes.xml");
                using (var endnotesStream = endnotesEntry.Open())
                using (var writer = new StreamWriter(endnotesStream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" +
                                 "<w:endnotes xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">" +
                                 "<w:endnote w:id=\"0\" w:type=\"separator\">" +
                                 "<w:p><w:r><w:t>End Sep</w:t></w:r></w:p>" +
                                 "</w:endnote>" +
                                 "<w:endnote w:id=\"1\" w:type=\"continuationSeparator\">" +
                                 "<w:p><w:r><w:t>End Continue</w:t></w:r></w:p>" +
                                 "</w:endnote>" +
                                 "<w:endnote w:id=\"2\"><w:p><w:r><w:t>Live endnote</w:t></w:r></w:p></w:endnote>" +
                                 "</w:endnotes>");
                }
            }

            using var testStream = new MemoryStream(ms.ToArray());
            using var reader = new Nedev.FileConverters.DocxToDoc.Format.DocxReader(testStream);

            var model = reader.ReadDocument();

            Assert.Equal("End Sep", model.EndnoteSeparatorText);
            Assert.Equal("End Continue", model.EndnoteContinuationSeparatorText);
            var endnote = Assert.Single(model.Endnotes);
            Assert.Equal("2", endnote.Id);
            Assert.Equal("Live endnote", endnote.Text);
        }

        [Fact]
        public void ReadDocument_WithEndnoteContinuationNotice_ParsesSpecialStoryText()
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

                var endnotesEntry = archive.CreateEntry("word/endnotes.xml");
                using (var endnotesStream = endnotesEntry.Open())
                using (var writer = new StreamWriter(endnotesStream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" +
                                 "<w:endnotes xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">" +
                                 "<w:endnote w:id=\"0\" w:type=\"continuationNotice\">" +
                                 "<w:p><w:r><w:t>End Notice</w:t></w:r></w:p>" +
                                 "</w:endnote>" +
                                 "<w:endnote w:id=\"2\"><w:p><w:r><w:t>Live endnote</w:t></w:r></w:p></w:endnote>" +
                                 "</w:endnotes>");
                }
            }

            using var testStream = new MemoryStream(ms.ToArray());
            using var reader = new Nedev.FileConverters.DocxToDoc.Format.DocxReader(testStream);

            var model = reader.ReadDocument();

            Assert.Equal("End Notice", model.EndnoteContinuationNoticeText);
            var endnote = Assert.Single(model.Endnotes);
            Assert.Equal("2", endnote.Id);
            Assert.Equal("Live endnote", endnote.Text);
        }

        [Fact]
        public void ReadDocument_WithEmptyAndUnknownSpecialEndnotes_PreservesEmptyContinuationNoticeAndIgnoresUnknownTypes()
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

                var endnotesEntry = archive.CreateEntry("word/endnotes.xml");
                using (var endnotesStream = endnotesEntry.Open())
                using (var writer = new StreamWriter(endnotesStream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" +
                                 "<w:endnotes xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">" +
                                 "<w:endnote w:id=\"0\" w:type=\"continuationNotice\"><w:p/></w:endnote>" +
                                 "<w:endnote w:id=\"1\" w:type=\"customNotice\"><w:p><w:r><w:t>Ignored</w:t></w:r></w:p></w:endnote>" +
                                 "<w:endnote w:id=\"2\"><w:p><w:r><w:t>Live endnote</w:t></w:r></w:p></w:endnote>" +
                                 "</w:endnotes>");
                }
            }

            using var testStream = new MemoryStream(ms.ToArray());
            using var reader = new Nedev.FileConverters.DocxToDoc.Format.DocxReader(testStream);

            var model = reader.ReadDocument();

            Assert.Equal(string.Empty, model.EndnoteContinuationNoticeText);
            Assert.Null(model.EndnoteSeparatorText);
            Assert.Null(model.EndnoteContinuationSeparatorText);
            var endnote = Assert.Single(model.Endnotes);
            Assert.Equal("2", endnote.Id);
            Assert.Equal("Live endnote", endnote.Text);
        }

        [Fact]
        public void ReadDocument_WithEndnoteMultiFragmentCustomMark_ParsesReferenceCpAndVisibleMark()
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
                                 "<w:r><w:endnoteReference w:id=\"2\" w:customMarkFollows=\"1\"/></w:r>" +
                                 "<w:r><w:rPr><w:vertAlign w:val=\"superscript\"/></w:rPr><w:t>(</w:t></w:r>" +
                                 "<w:r><w:rPr><w:rStyle w:val=\"EndnoteReference\"/></w:rPr><w:t>a</w:t></w:r>" +
                                 "<w:r><w:rPr><w:vertAlign w:val=\"superscript\"/></w:rPr><w:t>)</w:t></w:r>" +
                                 "<w:r><w:t>CD</w:t></w:r>" +
                                 "</w:p></w:body></w:document>");
                }

                var endnotesEntry = archive.CreateEntry("word/endnotes.xml");
                using (var endnotesStream = endnotesEntry.Open())
                using (var writer = new StreamWriter(endnotesStream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" +
                                 "<w:endnotes xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">" +
                                 "<w:endnote w:id=\"2\"><w:p><w:r><w:t>Note</w:t></w:r></w:p></w:endnote>" +
                                 "</w:endnotes>");
                }
            }

            using var testStream = new MemoryStream(ms.ToArray());
            using var reader = new Nedev.FileConverters.DocxToDoc.Format.DocxReader(testStream);

            var model = reader.ReadDocument();

            var endnote = Assert.Single(model.Endnotes);
            Assert.Equal(2, endnote.ReferenceCp);
            Assert.Equal("(a)", endnote.CustomMarkText);
            Assert.Equal("AB(a)CD\r", model.TextBuffer);
        }

        [Fact]
        public void ReadDocument_WithEndnoteCustomMarkButNoVisibleFollower_LeavesCustomMarkUnset()
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
                                 "<w:r><w:endnoteReference w:id=\"2\" w:customMarkFollows=\"1\"/></w:r>" +
                                 "<w:r><w:t>CD</w:t></w:r>" +
                                 "</w:p></w:body></w:document>");
                }

                var endnotesEntry = archive.CreateEntry("word/endnotes.xml");
                using (var endnotesStream = endnotesEntry.Open())
                using (var writer = new StreamWriter(endnotesStream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" +
                                 "<w:endnotes xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">" +
                                 "<w:endnote w:id=\"2\"><w:p><w:r><w:t>Note</w:t></w:r></w:p></w:endnote>" +
                                 "</w:endnotes>");
                }
            }

            using var testStream = new MemoryStream(ms.ToArray());
            using var reader = new Nedev.FileConverters.DocxToDoc.Format.DocxReader(testStream);

            var model = reader.ReadDocument();

            var endnote = Assert.Single(model.Endnotes);
            Assert.Equal(2, endnote.ReferenceCp);
            Assert.Null(endnote.CustomMarkText);
            Assert.Equal("ABCD\r", model.TextBuffer);
        }

        [Fact]
        public void ReadDocument_WithEndnoteWhitespaceOnlyCustomMarkFollower_LeavesCustomMarkUnset()
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
                                 "<w:r><w:endnoteReference w:id=\"2\" w:customMarkFollows=\"1\"/></w:r>" +
                                 "<w:r><w:rPr><w:vertAlign w:val=\"superscript\"/></w:rPr><w:t xml:space=\"preserve\"> </w:t></w:r>" +
                                 "<w:r><w:t>CD</w:t></w:r>" +
                                 "</w:p></w:body></w:document>");
                }

                var endnotesEntry = archive.CreateEntry("word/endnotes.xml");
                using (var endnotesStream = endnotesEntry.Open())
                using (var writer = new StreamWriter(endnotesStream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" +
                                 "<w:endnotes xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">" +
                                 "<w:endnote w:id=\"2\"><w:p><w:r><w:t>Note</w:t></w:r></w:p></w:endnote>" +
                                 "</w:endnotes>");
                }
            }

            using var testStream = new MemoryStream(ms.ToArray());
            using var reader = new Nedev.FileConverters.DocxToDoc.Format.DocxReader(testStream);

            var model = reader.ReadDocument();

            var endnote = Assert.Single(model.Endnotes);
            Assert.Null(endnote.CustomMarkText);
            Assert.Equal("AB CD\r", model.TextBuffer);
        }
    }
}
