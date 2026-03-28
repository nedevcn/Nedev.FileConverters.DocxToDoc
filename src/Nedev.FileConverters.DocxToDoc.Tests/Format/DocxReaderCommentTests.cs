using System;
using System.IO;
using System.IO.Compression;
using System.Text;
using Xunit;

namespace Nedev.FileConverters.DocxToDoc.Tests.Format
{
    public class DocxReaderCommentTests
    {
        private byte[] CreateDocxWithComments()
        {
            using var ms = new MemoryStream();
            using (var archive = new ZipArchive(ms, ZipArchiveMode.Create, true))
            {
                // Create word/comments.xml
                var commentsEntry = archive.CreateEntry("word/comments.xml");
                using (var stream = commentsEntry.Open())
                using (var writer = new StreamWriter(stream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                        "<w:comments xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" " +
                        "xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\">" +
                        "<w:comment w:id=\"0\" w:author=\"John Doe\" w:initials=\"JD\" w:date=\"2024-01-15T10:30:00Z\">" +
                        "<w:p><w:r><w:t>This is a test comment</w:t></w:r></w:p>" +
                        "</w:comment>" +
                        "<w:comment w:id=\"1\" w:author=\"Jane Smith\" w:initials=\"JS\" w:date=\"2024-01-16T14:20:00Z\" w:done=\"1\">" +
                        "<w:p><w:r><w:t>Another comment that is resolved</w:t></w:r></w:p>" +
                        "</w:comment>" +
                        "</w:comments>");
                }

                // Create word/document.xml
                var docEntry = archive.CreateEntry("word/document.xml");
                using (var stream = docEntry.Open())
                using (var writer = new StreamWriter(stream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                        "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">" +
                        "<w:body>" +
                        "<w:p><w:r><w:t>Document with comments</w:t></w:r></w:p>" +
                        "</w:body>" +
                        "</w:document>");
                }
            }
            return ms.ToArray();
        }

        private byte[] CreateDocxWithoutComments()
        {
            using var ms = new MemoryStream();
            using (var archive = new ZipArchive(ms, ZipArchiveMode.Create, true))
            {
                // Create word/document.xml without comments
                var docEntry = archive.CreateEntry("word/document.xml");
                using (var stream = docEntry.Open())
                using (var writer = new StreamWriter(stream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                        "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">" +
                        "<w:body>" +
                        "<w:p><w:r><w:t>Document without comments</w:t></w:r></w:p>" +
                        "</w:body>" +
                        "</w:document>");
                }
            }
            return ms.ToArray();
        }

        private byte[] CreateDocxWithFragmentedComment()
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
                        "<w:comment w:id=\"0\" w:author=\"John Doe\">" +
                        "<w:p><w:r><w:t>Left</w:t><w:tab/><w:t>Right</w:t><w:br/><w:t>Next</w:t></w:r></w:p>" +
                        "</w:comment>" +
                        "</w:comments>");
                }

                var docEntry = archive.CreateEntry("word/document.xml");
                using (var stream = docEntry.Open())
                using (var writer = new StreamWriter(stream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                        "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">" +
                        "<w:body><w:p><w:r><w:t>Document</w:t></w:r></w:p></w:body>" +
                        "</w:document>");
                }
            }
            return ms.ToArray();
        }

        private byte[] CreateDocxWithCommentReplies()
        {
            using var ms = new MemoryStream();
            using (var archive = new ZipArchive(ms, ZipArchiveMode.Create, true))
            {
                var commentsEntry = archive.CreateEntry("word/comments.xml");
                using (var stream = commentsEntry.Open())
                using (var writer = new StreamWriter(stream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                        "<w:comments xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" " +
                        "xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\">" +
                        "<w:comment w:id=\"0\" w:author=\"John Doe\" w:initials=\"JD\">" +
                        "<w:p w14:paraId=\"11111111\"><w:r><w:t>Parent comment</w:t></w:r></w:p>" +
                        "</w:comment>" +
                        "<w:comment w:id=\"1\" w:author=\"Jane Smith\" w:initials=\"JS\">" +
                        "<w:p w14:paraId=\"22222222\"><w:r><w:t>Reply comment</w:t></w:r></w:p>" +
                        "</w:comment>" +
                        "</w:comments>");
                }

                var commentsExtendedEntry = archive.CreateEntry("word/commentsExtended.xml");
                using (var stream = commentsExtendedEntry.Open())
                using (var writer = new StreamWriter(stream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                        "<w15:commentsEx xmlns:w15=\"http://schemas.microsoft.com/office/word/2012/wordml\">" +
                        "<w15:commentEx w15:paraId=\"11111111\" w15:done=\"1\"/>" +
                        "<w15:commentEx w15:paraId=\"22222222\" w15:paraIdParent=\"11111111\"/>" +
                        "</w15:commentsEx>");
                }

                var docEntry = archive.CreateEntry("word/document.xml");
                using (var stream = docEntry.Open())
                using (var writer = new StreamWriter(stream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                        "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">" +
                        "<w:body><w:p><w:r><w:t>Document</w:t></w:r></w:p></w:body>" +
                        "</w:document>");
                }
            }

            return ms.ToArray();
        }

        private byte[] CreateDocxWithPositionedTabComment()
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
                        "<w:comment w:id=\"0\" w:author=\"John Doe\">" +
                        "<w:p><w:r><w:t>Left</w:t><w:ptab w:alignment=\"center\" w:relativeTo=\"margin\"/><w:t>Right</w:t></w:r></w:p>" +
                        "</w:comment>" +
                        "</w:comments>");
                }

                var docEntry = archive.CreateEntry("word/document.xml");
                using (var stream = docEntry.Open())
                using (var writer = new StreamWriter(stream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                        "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">" +
                        "<w:body><w:p><w:r><w:t>Document</w:t></w:r></w:p></w:body>" +
                        "</w:document>");
                }
            }
            return ms.ToArray();
        }

        private byte[] CreateDocxWithMultiParagraphComment()
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
                        "<w:comment w:id=\"0\" w:author=\"John Doe\">" +
                        "<w:p><w:r><w:t>First</w:t></w:r></w:p>" +
                        "<w:p><w:r><w:t>Second</w:t></w:r></w:p>" +
                        "</w:comment>" +
                        "</w:comments>");
                }

                var docEntry = archive.CreateEntry("word/document.xml");
                using (var stream = docEntry.Open())
                using (var writer = new StreamWriter(stream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                        "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">" +
                        "<w:body><w:p><w:r><w:t>Document</w:t></w:r></w:p></w:body>" +
                        "</w:document>");
                }
            }

            return ms.ToArray();
        }

        private byte[] CreateDocxWithTrailingEmptyParagraphComment()
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
                        "<w:comment w:id=\"0\" w:author=\"John Doe\">" +
                        "<w:p><w:r><w:t>First</w:t></w:r></w:p>" +
                        "<w:p/>" +
                        "</w:comment>" +
                        "</w:comments>");
                }

                var docEntry = archive.CreateEntry("word/document.xml");
                using (var stream = docEntry.Open())
                using (var writer = new StreamWriter(stream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                        "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">" +
                        "<w:body><w:p><w:r><w:t>Document</w:t></w:r></w:p></w:body>" +
                        "</w:document>");
                }
            }

            return ms.ToArray();
        }

        private byte[] CreateDocxWithExplicitCommentRange()
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
                        "<w:comment w:id=\"0\" w:author=\"John Doe\">" +
                        "<w:p><w:r><w:t>Range comment</w:t></w:r></w:p>" +
                        "</w:comment>" +
                        "</w:comments>");
                }

                var docEntry = archive.CreateEntry("word/document.xml");
                using (var stream = docEntry.Open())
                using (var writer = new StreamWriter(stream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                        "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">" +
                        "<w:body>" +
                        "<w:p>" +
                        "<w:r><w:t>AA</w:t></w:r>" +
                        "<w:commentRangeStart w:id=\"0\"/>" +
                        "<w:r><w:t>BB</w:t></w:r>" +
                        "<w:commentRangeEnd w:id=\"0\"/>" +
                        "<w:r><w:commentReference w:id=\"0\"/></w:r>" +
                        "<w:r><w:t>CC</w:t></w:r>" +
                        "</w:p>" +
                        "</w:body>" +
                        "</w:document>");
                }
            }

            return ms.ToArray();
        }

        private byte[] CreateDocxWithCommentRangeAcrossHyperlinkAndBreak()
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
                        "<w:comment w:id=\"0\" w:author=\"John Doe\">" +
                        "<w:p><w:r><w:t>Complex comment</w:t></w:r></w:p>" +
                        "</w:comment>" +
                        "</w:comments>");
                }

                var docEntry = archive.CreateEntry("word/document.xml");
                using (var stream = docEntry.Open())
                using (var writer = new StreamWriter(stream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                        "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" " +
                        "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">" +
                        "<w:body>" +
                        "<w:p>" +
                        "<w:r><w:t>X</w:t></w:r>" +
                        "<w:commentRangeStart w:id=\"0\"/>" +
                        "<w:hyperlink r:id=\"rId1\"><w:r><w:t>Go</w:t></w:r></w:hyperlink>" +
                        "<w:r><w:br/><w:t>N</w:t></w:r>" +
                        "<w:commentRangeEnd w:id=\"0\"/>" +
                        "<w:r><w:t>Z</w:t></w:r>" +
                        "</w:p>" +
                        "</w:body>" +
                        "</w:document>");
                }

                var relsEntry = archive.CreateEntry("word/_rels/document.xml.rels");
                using (var stream = relsEntry.Open())
                using (var writer = new StreamWriter(stream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                        "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
                        "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink\" Target=\"https://example.com/comment-range\" TargetMode=\"External\"/>" +
                        "</Relationships>");
                }
            }

            return ms.ToArray();
        }

        private byte[] CreateDocxWithCollapsedCommentReference()
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
                        "<w:comment w:id=\"0\" w:author=\"John Doe\">" +
                        "<w:p><w:r><w:t>Collapsed comment</w:t></w:r></w:p>" +
                        "</w:comment>" +
                        "</w:comments>");
                }

                var docEntry = archive.CreateEntry("word/document.xml");
                using (var stream = docEntry.Open())
                using (var writer = new StreamWriter(stream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                        "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">" +
                        "<w:body>" +
                        "<w:p>" +
                        "<w:r><w:t>AB</w:t></w:r>" +
                        "<w:r><w:commentReference w:id=\"0\"/></w:r>" +
                        "<w:r><w:t>CD</w:t></w:r>" +
                        "</w:p>" +
                        "</w:body>" +
                        "</w:document>");
                }
            }

            return ms.ToArray();
        }

        private byte[] CreateDocxWithCommentRangeClosedByReference()
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
                        "<w:comment w:id=\"0\" w:author=\"John Doe\">" +
                        "<w:p><w:r><w:t>Reference-closed comment</w:t></w:r></w:p>" +
                        "</w:comment>" +
                        "</w:comments>");
                }

                var docEntry = archive.CreateEntry("word/document.xml");
                using (var stream = docEntry.Open())
                using (var writer = new StreamWriter(stream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                        "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">" +
                        "<w:body>" +
                        "<w:p>" +
                        "<w:r><w:t>A</w:t></w:r>" +
                        "<w:commentRangeStart w:id=\"0\"/>" +
                        "<w:r><w:t>B</w:t><w:tab/><w:t>C</w:t></w:r>" +
                        "<w:r><w:commentReference w:id=\"0\"/></w:r>" +
                        "<w:r><w:t>D</w:t></w:r>" +
                        "</w:p>" +
                        "</w:body>" +
                        "</w:document>");
                }
            }

            return ms.ToArray();
        }

        [Fact]
        public void ReadDocument_WithComments_ParsesAllComments()
        {
            // Arrange
            byte[] docxData = CreateDocxWithComments();
            using var ms = new MemoryStream(docxData);
            using var reader = new Nedev.FileConverters.DocxToDoc.Format.DocxReader(ms);

            // Act
            var model = reader.ReadDocument();

            // Assert
            Assert.Equal(2, model.Comments.Count);
        }

        [Fact]
        public void ReadDocument_WithComments_ParsesCommentId()
        {
            // Arrange
            byte[] docxData = CreateDocxWithComments();
            using var ms = new MemoryStream(docxData);
            using var reader = new Nedev.FileConverters.DocxToDoc.Format.DocxReader(ms);

            // Act
            var model = reader.ReadDocument();

            // Assert
            Assert.Contains(model.Comments, c => c.Id == "0");
            Assert.Contains(model.Comments, c => c.Id == "1");
        }

        [Fact]
        public void ReadDocument_WithComments_ParsesCommentAuthor()
        {
            // Arrange
            byte[] docxData = CreateDocxWithComments();
            using var ms = new MemoryStream(docxData);
            using var reader = new Nedev.FileConverters.DocxToDoc.Format.DocxReader(ms);

            // Act
            var model = reader.ReadDocument();

            // Assert
            Assert.Contains(model.Comments, c => c.Author == "John Doe");
            Assert.Contains(model.Comments, c => c.Author == "Jane Smith");
        }

        [Fact]
        public void ReadDocument_WithComments_ParsesCommentInitials()
        {
            // Arrange
            byte[] docxData = CreateDocxWithComments();
            using var ms = new MemoryStream(docxData);
            using var reader = new Nedev.FileConverters.DocxToDoc.Format.DocxReader(ms);

            // Act
            var model = reader.ReadDocument();

            // Assert
            Assert.Contains(model.Comments, c => c.Initials == "JD");
            Assert.Contains(model.Comments, c => c.Initials == "JS");
        }

        [Fact]
        public void ReadDocument_WithComments_ParsesCommentText()
        {
            // Arrange
            byte[] docxData = CreateDocxWithComments();
            using var ms = new MemoryStream(docxData);
            using var reader = new Nedev.FileConverters.DocxToDoc.Format.DocxReader(ms);

            // Act
            var model = reader.ReadDocument();

            // Assert
            Assert.Contains(model.Comments, c => c.Text == "This is a test comment");
            Assert.Contains(model.Comments, c => c.Text == "Another comment that is resolved");
        }

        [Fact]
        public void ReadDocument_WithComments_ParsesCommentDate()
        {
            // Arrange
            byte[] docxData = CreateDocxWithComments();
            using var ms = new MemoryStream(docxData);
            using var reader = new Nedev.FileConverters.DocxToDoc.Format.DocxReader(ms);

            // Act
            var model = reader.ReadDocument();

            // Assert
            Assert.All(model.Comments, c => Assert.NotNull(c.Date));
        }

        [Fact]
        public void ReadDocument_WithComments_ParsesDoneStatus()
        {
            // Arrange
            byte[] docxData = CreateDocxWithComments();
            using var ms = new MemoryStream(docxData);
            using var reader = new Nedev.FileConverters.DocxToDoc.Format.DocxReader(ms);

            // Act
            var model = reader.ReadDocument();

            // Assert
            Assert.Contains(model.Comments, c => c.IsDone == true);
            Assert.Contains(model.Comments, c => c.IsDone == false);
        }

        [Fact]
        public void ReadDocument_WithCommentsExtended_ParsesReplyParentAndDoneMetadata()
        {
            byte[] docxData = CreateDocxWithCommentReplies();
            using var ms = new MemoryStream(docxData);
            using var reader = new Nedev.FileConverters.DocxToDoc.Format.DocxReader(ms);

            var model = reader.ReadDocument();

            var parentComment = Assert.Single(model.Comments, comment => comment.Id == "0");
            var replyComment = Assert.Single(model.Comments, comment => comment.Id == "1");

            Assert.True(parentComment.IsDone);
            Assert.False(parentComment.IsReply);
            Assert.True(replyComment.IsReply);
            Assert.Equal("0", replyComment.ParentId);
            Assert.False(replyComment.IsDone);
        }

        [Fact]
        public void ReadDocument_WithoutComments_HasEmptyCommentsList()
        {
            // Arrange
            byte[] docxData = CreateDocxWithoutComments();
            using var ms = new MemoryStream(docxData);
            using var reader = new Nedev.FileConverters.DocxToDoc.Format.DocxReader(ms);

            // Act
            var model = reader.ReadDocument();

            // Assert
            Assert.Empty(model.Comments);
        }

        [Fact]
        public void ReadDocument_WithFragmentedComment_PreservesTabsAndBreaks()
        {
            byte[] docxData = CreateDocxWithFragmentedComment();
            using var ms = new MemoryStream(docxData);
            using var reader = new Nedev.FileConverters.DocxToDoc.Format.DocxReader(ms);

            var model = reader.ReadDocument();

            var comment = Assert.Single(model.Comments);
            Assert.Equal("Left\tRight\vNext", comment.Text);
        }

        [Fact]
        public void ReadDocument_WithPositionedTabComment_PreservesTabCharacter()
        {
            byte[] docxData = CreateDocxWithPositionedTabComment();
            using var ms = new MemoryStream(docxData);
            using var reader = new Nedev.FileConverters.DocxToDoc.Format.DocxReader(ms);

            var model = reader.ReadDocument();

            var comment = Assert.Single(model.Comments);
            Assert.Equal("Left\tRight", comment.Text);
        }

        [Fact]
        public void ReadDocument_WithMultiParagraphComment_PreservesParagraphBreaks()
        {
            byte[] docxData = CreateDocxWithMultiParagraphComment();
            using var ms = new MemoryStream(docxData);
            using var reader = new Nedev.FileConverters.DocxToDoc.Format.DocxReader(ms);

            var model = reader.ReadDocument();

            var comment = Assert.Single(model.Comments);
            Assert.Equal("First\rSecond", comment.Text);
        }

        [Fact]
        public void ReadDocument_WithTrailingEmptyCommentParagraph_PreservesTrailingParagraphBreak()
        {
            byte[] docxData = CreateDocxWithTrailingEmptyParagraphComment();
            using var ms = new MemoryStream(docxData);
            using var reader = new Nedev.FileConverters.DocxToDoc.Format.DocxReader(ms);

            var model = reader.ReadDocument();

            var comment = Assert.Single(model.Comments);
            Assert.Equal("First\r", comment.Text);
        }

        [Fact]
        public void ReadDocument_WithExplicitCommentRange_ParsesAnchorCpRange()
        {
            byte[] docxData = CreateDocxWithExplicitCommentRange();
            using var ms = new MemoryStream(docxData);
            using var reader = new Nedev.FileConverters.DocxToDoc.Format.DocxReader(ms);

            var model = reader.ReadDocument();

            var comment = Assert.Single(model.Comments);
            Assert.Equal("Range comment", comment.Text);
            Assert.Equal(2, comment.StartCp);
            Assert.Equal(4, comment.EndCp);
            Assert.Equal("AABBCC\r", model.TextBuffer);
        }

        [Fact]
        public void ReadDocument_WithCommentRangeAcrossHyperlinkAndBreak_CountsVisibleCharactersInAnchorRange()
        {
            byte[] docxData = CreateDocxWithCommentRangeAcrossHyperlinkAndBreak();
            using var ms = new MemoryStream(docxData);
            using var reader = new Nedev.FileConverters.DocxToDoc.Format.DocxReader(ms);

            var model = reader.ReadDocument();

            var comment = Assert.Single(model.Comments);
            Assert.Equal("Complex comment", comment.Text);
            Assert.Equal(1, comment.StartCp);
            Assert.Equal(5, comment.EndCp);
            Assert.Equal("XGo\vNZ\r", model.TextBuffer);
        }

        [Fact]
        public void ReadDocument_WithCollapsedCommentReference_UsesZeroWidthAnchorAtReferenceCp()
        {
            byte[] docxData = CreateDocxWithCollapsedCommentReference();
            using var ms = new MemoryStream(docxData);
            using var reader = new Nedev.FileConverters.DocxToDoc.Format.DocxReader(ms);

            var model = reader.ReadDocument();

            var comment = Assert.Single(model.Comments);
            Assert.Equal("Collapsed comment", comment.Text);
            Assert.Equal(2, comment.StartCp);
            Assert.Equal(2, comment.EndCp);
            Assert.Equal("ABCD\r", model.TextBuffer);
        }

        [Fact]
        public void ReadDocument_WithCommentRangeClosedByReference_ClosesAnchorAtReferenceCp()
        {
            byte[] docxData = CreateDocxWithCommentRangeClosedByReference();
            using var ms = new MemoryStream(docxData);
            using var reader = new Nedev.FileConverters.DocxToDoc.Format.DocxReader(ms);

            var model = reader.ReadDocument();

            var comment = Assert.Single(model.Comments);
            Assert.Equal("Reference-closed comment", comment.Text);
            Assert.Equal(1, comment.StartCp);
            Assert.Equal(4, comment.EndCp);
            Assert.Equal("AB\tCD\r", model.TextBuffer);
        }
    }
}
