using System;
using System.IO;
using System.IO.Compression;
using System.Text;
using Xunit;

namespace Nedev.FileConverters.DocxToDoc.Tests.Format
{
    public class DocxReaderBookmarkTests
    {
        private byte[] CreateDocxWithBookmark()
        {
            using var ms = new MemoryStream();
            using (var archive = new ZipArchive(ms, ZipArchiveMode.Create, true))
            {
                // Create word/document.xml with bookmark
                var docEntry = archive.CreateEntry("word/document.xml");
                using (var stream = docEntry.Open())
                using (var writer = new StreamWriter(stream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                        "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">" +
                        "<w:body>" +
                        "<w:p>" +
                        "<w:r>" +
                        "<w:t>Before bookmark </w:t>" +
                        "</w:r>" +
                        "<w:bookmarkStart w:id=\"0\" w:name=\"MyBookmark\"/>" +
                        "<w:r>" +
                        "<w:t>Bookmarked text</w:t>" +
                        "</w:r>" +
                        "<w:bookmarkEnd w:id=\"0\"/>" +
                        "<w:r>" +
                        "<w:t> After bookmark</w:t>" +
                        "</w:r>" +
                        "</w:p>" +
                        "</w:body>" +
                        "</w:document>");
                }
            }
            return ms.ToArray();
        }

        private byte[] CreateDocxWithMultipleBookmarks()
        {
            using var ms = new MemoryStream();
            using (var archive = new ZipArchive(ms, ZipArchiveMode.Create, true))
            {
                // Create word/document.xml with multiple bookmarks
                var docEntry = archive.CreateEntry("word/document.xml");
                using (var stream = docEntry.Open())
                using (var writer = new StreamWriter(stream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                        "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">" +
                        "<w:body>" +
                        "<w:p>" +
                        "<w:bookmarkStart w:id=\"0\" w:name=\"StartBookmark\"/>" +
                        "<w:r><w:t>First bookmark</w:t></w:r>" +
                        "<w:bookmarkEnd w:id=\"0\"/>" +
                        "</w:p>" +
                        "<w:p>" +
                        "<w:bookmarkStart w:id=\"1\" w:name=\"SecondBookmark\"/>" +
                        "<w:r><w:t>Second bookmark</w:t></w:r>" +
                        "<w:bookmarkEnd w:id=\"1\"/>" +
                        "</w:p>" +
                        "</w:body>" +
                        "</w:document>");
                }
            }
            return ms.ToArray();
        }

        private byte[] CreateDocxWithBookmarkContainingTabAndBreak()
        {
            using var ms = new MemoryStream();
            using (var archive = new ZipArchive(ms, ZipArchiveMode.Create, true))
            {
                var docEntry = archive.CreateEntry("word/document.xml");
                using (var stream = docEntry.Open())
                using (var writer = new StreamWriter(stream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                        "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">" +
                        "<w:body>" +
                        "<w:p>" +
                        "<w:bookmarkStart w:id=\"0\" w:name=\"MixedBookmark\"/>" +
                        "<w:r><w:t>A</w:t><w:tab/><w:t>B</w:t><w:br/><w:t>C</w:t></w:r>" +
                        "<w:bookmarkEnd w:id=\"0\"/>" +
                        "</w:p>" +
                        "</w:body>" +
                        "</w:document>");
                }
            }
            return ms.ToArray();
        }

        private byte[] CreateDocxWithBookmarkContainingSpecialHyphens()
        {
            using var ms = new MemoryStream();
            using (var archive = new ZipArchive(ms, ZipArchiveMode.Create, true))
            {
                var docEntry = archive.CreateEntry("word/document.xml");
                using (var stream = docEntry.Open())
                using (var writer = new StreamWriter(stream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                        "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">" +
                        "<w:body>" +
                        "<w:p>" +
                        "<w:bookmarkStart w:id=\"0\" w:name=\"HyphenBookmark\"/>" +
                        "<w:r><w:t>co</w:t><w:noBreakHyphen/><w:t>op</w:t><w:softHyphen/><w:t>erate</w:t></w:r>" +
                        "<w:bookmarkEnd w:id=\"0\"/>" +
                        "</w:p>" +
                        "</w:body>" +
                        "</w:document>");
                }
            }

            return ms.ToArray();
        }

        private byte[] CreateDocxWithBookmarkContainingSymbolRun()
        {
            using var ms = new MemoryStream();
            using (var archive = new ZipArchive(ms, ZipArchiveMode.Create, true))
            {
                var docEntry = archive.CreateEntry("word/document.xml");
                using (var stream = docEntry.Open())
                using (var writer = new StreamWriter(stream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                        "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">" +
                        "<w:body>" +
                        "<w:p>" +
                        "<w:bookmarkStart w:id=\"0\" w:name=\"SymbolBookmark\"/>" +
                        "<w:r><w:sym w:font=\"Wingdings\" w:char=\"F0FC\"/></w:r>" +
                        "<w:bookmarkEnd w:id=\"0\"/>" +
                        "</w:p>" +
                        "</w:body>" +
                        "</w:document>");
                }
            }

            return ms.ToArray();
        }

        private byte[] CreateDocxWithBookmarkContainingPositionedTab()
        {
            using var ms = new MemoryStream();
            using (var archive = new ZipArchive(ms, ZipArchiveMode.Create, true))
            {
                var docEntry = archive.CreateEntry("word/document.xml");
                using (var stream = docEntry.Open())
                using (var writer = new StreamWriter(stream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                        "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">" +
                        "<w:body>" +
                        "<w:p>" +
                        "<w:bookmarkStart w:id=\"0\" w:name=\"PtabBookmark\"/>" +
                        "<w:r><w:t>A</w:t><w:ptab w:alignment=\"center\" w:relativeTo=\"margin\"/><w:t>B</w:t></w:r>" +
                        "<w:bookmarkEnd w:id=\"0\"/>" +
                        "</w:p>" +
                        "</w:body>" +
                        "</w:document>");
                }
            }

            return ms.ToArray();
        }

        [Fact]
        public void ReadDocument_WithBookmark_ParsesBookmarkData()
        {
            // Arrange
            byte[] docxData = CreateDocxWithBookmark();
            using var ms = new MemoryStream(docxData);
            using var reader = new Nedev.FileConverters.DocxToDoc.Format.DocxReader(ms);

            // Act
            var model = reader.ReadDocument();

            // Assert
            Assert.Single(model.Bookmarks);
            Assert.Equal("MyBookmark", model.Bookmarks[0].Name);
            Assert.Equal("0", model.Bookmarks[0].Id);
        }

        [Fact]
        public void ReadDocument_WithMultipleBookmarks_ParsesAllBookmarks()
        {
            // Arrange
            byte[] docxData = CreateDocxWithMultipleBookmarks();
            using var ms = new MemoryStream(docxData);
            using var reader = new Nedev.FileConverters.DocxToDoc.Format.DocxReader(ms);

            // Act
            var model = reader.ReadDocument();

            // Assert
            Assert.Equal(2, model.Bookmarks.Count);
            Assert.Contains(model.Bookmarks, b => b.Name == "StartBookmark");
            Assert.Contains(model.Bookmarks, b => b.Name == "SecondBookmark");
        }

        [Fact]
        public void ReadDocument_WithBookmark_HasValidCpRange()
        {
            // Arrange
            byte[] docxData = CreateDocxWithBookmark();
            using var ms = new MemoryStream(docxData);
            using var reader = new Nedev.FileConverters.DocxToDoc.Format.DocxReader(ms);

            // Act
            var model = reader.ReadDocument();

            // Assert
            Assert.Single(model.Bookmarks);
            var bookmark = model.Bookmarks[0];
            Assert.True(bookmark.StartCp >= 0);
            Assert.True(bookmark.EndCp >= bookmark.StartCp);
        }

        [Fact]
        public void ReadDocument_WithBookmarkTabAndBreak_CountsControlCharactersInCpRange()
        {
            byte[] docxData = CreateDocxWithBookmarkContainingTabAndBreak();
            using var ms = new MemoryStream(docxData);
            using var reader = new Nedev.FileConverters.DocxToDoc.Format.DocxReader(ms);

            var model = reader.ReadDocument();

            var bookmark = Assert.Single(model.Bookmarks);
            Assert.Equal(5, bookmark.EndCp - bookmark.StartCp);
        }

        [Fact]
        public void ReadDocument_WithBookmarkSpecialHyphens_CountsInlineHyphenCharactersInCpRange()
        {
            byte[] docxData = CreateDocxWithBookmarkContainingSpecialHyphens();
            using var ms = new MemoryStream(docxData);
            using var reader = new Nedev.FileConverters.DocxToDoc.Format.DocxReader(ms);

            var model = reader.ReadDocument();

            var bookmark = Assert.Single(model.Bookmarks);
            Assert.Equal(11, bookmark.EndCp - bookmark.StartCp);
        }

        [Fact]
        public void ReadDocument_WithBookmarkSymbolRun_CountsSymbolCharacterInCpRange()
        {
            byte[] docxData = CreateDocxWithBookmarkContainingSymbolRun();
            using var ms = new MemoryStream(docxData);
            using var reader = new Nedev.FileConverters.DocxToDoc.Format.DocxReader(ms);

            var model = reader.ReadDocument();

            var bookmark = Assert.Single(model.Bookmarks);
            Assert.Equal(1, bookmark.EndCp - bookmark.StartCp);
        }

        [Fact]
        public void ReadDocument_WithBookmarkPositionedTab_CountsTabCharacterInCpRange()
        {
            byte[] docxData = CreateDocxWithBookmarkContainingPositionedTab();
            using var ms = new MemoryStream(docxData);
            using var reader = new Nedev.FileConverters.DocxToDoc.Format.DocxReader(ms);

            var model = reader.ReadDocument();

            var bookmark = Assert.Single(model.Bookmarks);
            Assert.Equal(3, bookmark.EndCp - bookmark.StartCp);
        }
    }
}
