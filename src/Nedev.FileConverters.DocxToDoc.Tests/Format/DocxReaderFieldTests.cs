using System;
using System.IO;
using System.IO.Compression;
using System.Text;
using Xunit;

namespace Nedev.FileConverters.DocxToDoc.Tests.Format
{
    public class DocxReaderFieldTests
    {
        private byte[] CreateDocxWithPageField()
        {
            using var ms = new MemoryStream();
            using (var archive = new ZipArchive(ms, ZipArchiveMode.Create, true))
            {
                // Create word/document.xml with page number field
                var docEntry = archive.CreateEntry("word/document.xml");
                using (var stream = docEntry.Open())
                using (var writer = new StreamWriter(stream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                        "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">" +
                        "<w:body>" +
                        "<w:p>" +
                        "<w:r>" +
                        "<w:t>Page </w:t>" +
                        "</w:r>" +
                        "<w:r>" +
                        "<w:fldChar w:fldCharType=\"begin\"/>" +
                        "</w:r>" +
                        "<w:r>" +
                        "<w:instrText>PAGE</w:instrText>" +
                        "</w:r>" +
                        "<w:r>" +
                        "<w:fldChar w:fldCharType=\"separate\"/>" +
                        "</w:r>" +
                        "<w:r>" +
                        "<w:t>1</w:t>" +
                        "</w:r>" +
                        "<w:r>" +
                        "<w:fldChar w:fldCharType=\"end\"/>" +
                        "</w:r>" +
                        "</w:p>" +
                        "</w:body>" +
                        "</w:document>");
                }
            }
            return ms.ToArray();
        }

        private byte[] CreateDocxWithDateField()
        {
            using var ms = new MemoryStream();
            using (var archive = new ZipArchive(ms, ZipArchiveMode.Create, true))
            {
                // Create word/document.xml with date field
                var docEntry = archive.CreateEntry("word/document.xml");
                using (var stream = docEntry.Open())
                using (var writer = new StreamWriter(stream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                        "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">" +
                        "<w:body>" +
                        "<w:p>" +
                        "<w:r>" +
                        "<w:t>Date: </w:t>" +
                        "</w:r>" +
                        "<w:r>" +
                        "<w:fldChar w:fldCharType=\"begin\"/>" +
                        "</w:r>" +
                        "<w:r>" +
                        "<w:instrText>DATE \\@ \"yyyy-MM-dd\"</w:instrText>" +
                        "</w:r>" +
                        "<w:r>" +
                        "<w:fldChar w:fldCharType=\"separate\"/>" +
                        "</w:r>" +
                        "<w:r>" +
                        "<w:t>2024-01-15</w:t>" +
                        "</w:r>" +
                        "<w:r>" +
                        "<w:fldChar w:fldCharType=\"end\"/>" +
                        "</w:r>" +
                        "</w:p>" +
                        "</w:body>" +
                        "</w:document>");
                }
            }
            return ms.ToArray();
        }

        private byte[] CreateDocxWithSplitDateFieldInstruction()
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
                        "<w:r><w:t>Date: </w:t></w:r>" +
                        "<w:r><w:fldChar w:fldCharType=\"begin\"/></w:r>" +
                        "<w:r><w:instrText>DATE </w:instrText><w:instrText>\\@ \"yyyy-MM-dd\"</w:instrText></w:r>" +
                        "<w:r><w:fldChar w:fldCharType=\"separate\"/></w:r>" +
                        "<w:r><w:t>2024-01-15</w:t></w:r>" +
                        "<w:r><w:fldChar w:fldCharType=\"end\"/></w:r>" +
                        "</w:p>" +
                        "</w:body>" +
                        "</w:document>");
                }
            }
            return ms.ToArray();
        }

        private byte[] CreateDocxWithSimplePageField()
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
                        "<w:r><w:t>Page </w:t></w:r>" +
                        "<w:fldSimple w:instr=\"PAGE\" w:fldLock=\"true\">" +
                        "<w:r><w:t>1</w:t></w:r>" +
                        "</w:fldSimple>" +
                        "</w:p>" +
                        "</w:body>" +
                        "</w:document>");
                }
            }

            return ms.ToArray();
        }

        private byte[] CreateDocxWithSimpleDateField()
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
                        "<w:r><w:t>Date: </w:t></w:r>" +
                        "<w:fldSimple w:instr=\"DATE \\@ &quot;yyyy-MM-dd&quot;\" w:dirty=\"true\">" +
                        "<w:r><w:t>2024-</w:t></w:r>" +
                        "<w:r><w:t>01-15</w:t></w:r>" +
                        "</w:fldSimple>" +
                        "</w:p>" +
                        "</w:body>" +
                        "</w:document>");
                }
            }

            return ms.ToArray();
        }

        [Fact]
        public void ReadDocument_WithPageField_ParsesFieldData()
        {
            // Arrange
            byte[] docxData = CreateDocxWithPageField();
            using var ms = new MemoryStream(docxData);
            using var reader = new Nedev.FileConverters.DocxToDoc.Format.DocxReader(ms);

            // Act
            var model = reader.ReadDocument();

            // Assert
            Assert.Single(model.Paragraphs);
            Assert.True(model.Paragraphs[0].Runs.Count >= 3);

            // Find the field begin run
            var fieldBeginRun = model.Paragraphs[0].Runs.Find(r => r.IsFieldBegin);
            Assert.NotNull(fieldBeginRun);
            Assert.NotNull(fieldBeginRun.Field);
                Assert.Equal("PAGE", fieldBeginRun.Field.Instruction);
                Assert.Equal(Nedev.FileConverters.DocxToDoc.Model.FieldType.Page, fieldBeginRun.Field.Type);
        }

        [Fact]
        public void ReadDocument_WithDateField_ParsesFieldData()
        {
            // Arrange
            byte[] docxData = CreateDocxWithDateField();
            using var ms = new MemoryStream(docxData);
            using var reader = new Nedev.FileConverters.DocxToDoc.Format.DocxReader(ms);

            // Act
            var model = reader.ReadDocument();

            // Assert
            Assert.Single(model.Paragraphs);

            // Find the field begin run
            var fieldBeginRun = model.Paragraphs[0].Runs.Find(r => r.IsFieldBegin);
            Assert.NotNull(fieldBeginRun);
            Assert.NotNull(fieldBeginRun.Field);
                Assert.Equal("DATE \\@ \"yyyy-MM-dd\"", fieldBeginRun.Field.Instruction);
                Assert.Equal(Nedev.FileConverters.DocxToDoc.Model.FieldType.Date, fieldBeginRun.Field.Type);
        }

        [Fact]
        public void ReadDocument_WithField_HasFieldMarkers()
        {
            // Arrange
            byte[] docxData = CreateDocxWithPageField();
            using var ms = new MemoryStream(docxData);
            using var reader = new Nedev.FileConverters.DocxToDoc.Format.DocxReader(ms);

            // Act
            var model = reader.ReadDocument();

            // Assert
            Assert.Single(model.Paragraphs);

            // Verify field markers exist
            Assert.Contains(model.Paragraphs[0].Runs, r => r.IsFieldBegin);
            Assert.Contains(model.Paragraphs[0].Runs, r => r.IsFieldSeparate);
            Assert.Contains(model.Paragraphs[0].Runs, r => r.IsFieldEnd);
        }

        [Fact]
        public void ReadDocument_WithSplitInstrText_ParsesFullInstruction()
        {
            byte[] docxData = CreateDocxWithSplitDateFieldInstruction();
            using var ms = new MemoryStream(docxData);
            using var reader = new Nedev.FileConverters.DocxToDoc.Format.DocxReader(ms);

            var model = reader.ReadDocument();

            var fieldBeginRun = model.Paragraphs[0].Runs.Find(r => r.IsFieldBegin);
            Assert.NotNull(fieldBeginRun);
            Assert.NotNull(fieldBeginRun.Field);
            Assert.Equal("DATE \\@ \"yyyy-MM-dd\"", fieldBeginRun.Field.Instruction);
            Assert.Equal(Nedev.FileConverters.DocxToDoc.Model.FieldType.Date, fieldBeginRun.Field.Type);
            Assert.Contains(model.Paragraphs[0].Runs, r => r.IsFieldSeparate);
        }

        [Fact]
        public void ReadDocument_WithSimplePageField_ExpandsSyntheticFieldMarkers()
        {
            byte[] docxData = CreateDocxWithSimplePageField();
            using var ms = new MemoryStream(docxData);
            using var reader = new Nedev.FileConverters.DocxToDoc.Format.DocxReader(ms);

            var model = reader.ReadDocument();
            var runs = model.Paragraphs[0].Runs;

            var fieldBeginRun = runs.Find(r => r.IsFieldBegin);
            var fieldSeparateRun = runs.Find(r => r.IsFieldSeparate);
            var fieldEndRun = runs.Find(r => r.IsFieldEnd);

            Assert.NotNull(fieldBeginRun);
            Assert.NotNull(fieldSeparateRun);
            Assert.NotNull(fieldEndRun);
            Assert.NotNull(fieldBeginRun.Field);
            Assert.Same(fieldBeginRun.Field, fieldSeparateRun.Field);
            Assert.Same(fieldBeginRun.Field, fieldEndRun.Field);
            Assert.Equal("PAGE", fieldBeginRun.Field.Instruction);
            Assert.Equal(Nedev.FileConverters.DocxToDoc.Model.FieldType.Page, fieldBeginRun.Field.Type);
            Assert.True(fieldBeginRun.Field.IsLocked);

            int beginIndex = runs.FindIndex(r => r.IsFieldBegin);
            int separateIndex = runs.FindIndex(r => r.IsFieldSeparate);
            int resultIndex = runs.FindIndex(r => r.Text == "1");
            int endIndex = runs.FindIndex(r => r.IsFieldEnd);

            Assert.True(beginIndex >= 0 && beginIndex < separateIndex);
            Assert.True(separateIndex < resultIndex);
            Assert.True(resultIndex < endIndex);
            Assert.Equal("Page 1\r", model.TextBuffer);
        }

        [Fact]
        public void ReadDocument_WithSimpleDateField_PreservesResultRunsAndTextBuffer()
        {
            byte[] docxData = CreateDocxWithSimpleDateField();
            using var ms = new MemoryStream(docxData);
            using var reader = new Nedev.FileConverters.DocxToDoc.Format.DocxReader(ms);

            var model = reader.ReadDocument();
            var runs = model.Paragraphs[0].Runs;
            var fieldBeginRun = runs.Find(r => r.IsFieldBegin);

            Assert.NotNull(fieldBeginRun);
            Assert.NotNull(fieldBeginRun.Field);
            Assert.Equal("DATE \\@ \"yyyy-MM-dd\"", fieldBeginRun.Field.Instruction);
            Assert.Equal(Nedev.FileConverters.DocxToDoc.Model.FieldType.Date, fieldBeginRun.Field.Type);
            Assert.True(fieldBeginRun.Field.IsDirty);

            var textRuns = runs.FindAll(r => !string.IsNullOrEmpty(r.Text));
            Assert.Collection(
                textRuns,
                run => Assert.Equal("Date: ", run.Text),
                run => Assert.Equal("2024-", run.Text),
                run => Assert.Equal("01-15", run.Text));
            Assert.Contains(runs, r => r.IsFieldSeparate);
            Assert.Contains(runs, r => r.IsFieldEnd);
            Assert.Equal("Date: 2024-01-15\r", model.TextBuffer);
        }
    }
}
