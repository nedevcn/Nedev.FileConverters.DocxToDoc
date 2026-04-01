using System;
using System.IO;
using System.Text;
using System.IO.Compression;
using Nedev.FileConverters.DocxToDoc.Format;
using Nedev.FileConverters.DocxToDoc.Model;
using Xunit;

namespace Nedev.FileConverters.DocxToDoc.Tests.Format
{
    public class DocWriterFormattingTests
    {
        [Fact]
        public void WriteDocBlocks_FormattingIncluded_CreatesChpx()
        {
            // Register encoding provider for Windows-1252 
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Arrange
            var model = new DocumentModel();
            var para = new ParagraphModel();
            var run = new RunModel { Text = "FormatTest" };
            run.Properties.IsBold = true;
            run.Properties.IsItalic = true;
            run.Properties.FontSize = 24; // 12pt
            
            para.Runs.Add(run);
            model.Content.Add(para);

            var writer = new DocWriter();
            using var ms = new MemoryStream();

            // Act
            writer.WriteDocBlocks(model, ms);
            ms.Position = 0;

            // Assert
            using var compoundFile = new OpenMcdf.CompoundFile(ms);
            Assert.True(compoundFile.RootStorage.TryGetStream("WordDocument", out var wordDocStream));
            Assert.True(compoundFile.RootStorage.TryGetStream("1Table", out var tableStream));
            
            byte[] tableData = tableStream.GetData();
            byte[] wordDocData = wordDocStream.GetData();

            // Just rough assertions that the streams are non-empty and larger than minimums 
            Assert.True(tableData.Length > 20); // Should contain Clx and PlcfbteChpx
            
            // Expected WordDoc min size: 1536 (Text Start) + 11 (Text length) 
            // Plus at least one 512 byte FKP since we have formatting
            Assert.True(wordDocData.Length >= 1536 + 11 + 512); 
        }

        [Fact]
        public void WriteDocBlocks_ParagraphFormatting_CreatesPapx()
        {
            // Register encoding provider for Windows-1252 
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Arrange
            var model = new DocumentModel();
            var para = new ParagraphModel();
            para.Runs.Add(new RunModel { Text = "ParaTest" });
            para.Properties.Alignment = ParagraphModel.Justification.Center;
            model.Content.Add(para);

            var writer = new DocWriter();
            using var ms = new MemoryStream();

            // Act
            writer.WriteDocBlocks(model, ms);
            ms.Position = 0;

            // Assert
            using var compoundFile = new OpenMcdf.CompoundFile(ms);
            Assert.True(compoundFile.RootStorage.TryGetStream("WordDocument", out var wordDocStream));
            Assert.True(compoundFile.RootStorage.TryGetStream("1Table", out var tableStream));
            
            byte[] tableData = tableStream.GetData();
            byte[] wordDocData = wordDocStream.GetData();

            // Check that the FIB was updated with non-zero Papx PLCF offsets
            // The PLCF for Papx uses Fib.PapxPairIndex in the simplified Fc/Lcb map.
            // Fib size is >= 154 + 744 = 898
            Assert.True(wordDocData.Length >= 898);
            int papxOffset = 154 + (Fib.PapxPairIndex * 8);
            int fcPlcfbtePapx = BitConverter.ToInt32(wordDocData, papxOffset);
            int lcbPlcfbtePapx = BitConverter.ToInt32(wordDocData, papxOffset + 4);
            
            Assert.NotEqual(0, fcPlcfbtePapx);
            Assert.True(lcbPlcfbtePapx > 0);
            
            // Verify that a PAPX FKP was written to WordDocument
            Assert.True(wordDocData.Length >= 1536 + 9 + 512); 
        }

        [Fact]
        public void WriteDocBlocks_ParagraphSpacing_WritesSpacingSprmsIntoPapx()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            var model = new DocumentModel();
            var para = new ParagraphModel();
            para.Runs.Add(new RunModel { Text = "SpacingTest" });
            para.Properties.SpacingBeforeTwips = 120;
            para.Properties.SpacingAfterTwips = 240;
            model.Content.Add(para);

            var writer = new DocWriter();
            using var ms = new MemoryStream();

            writer.WriteDocBlocks(model, ms);
            ms.Position = 0;

            using var compoundFile = new OpenMcdf.CompoundFile(ms);
            Assert.True(compoundFile.RootStorage.TryGetStream("WordDocument", out var wordDocStream));

            byte[] wordDocData = wordDocStream.GetData();

            Assert.True(ContainsSubsequence(wordDocData, new byte[] { 0x22, 0x26, 0x78, 0x00 }));
            Assert.True(ContainsSubsequence(wordDocData, new byte[] { 0x23, 0x26, 0xF0, 0x00 }));
        }

        [Fact]
        public void WriteDocBlocks_LineSpacing_WritesLineSpacingSprmIntoPapx()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            var model = new DocumentModel();
            var para = new ParagraphModel();
            para.Runs.Add(new RunModel { Text = "LineSpacingTest" });
            para.Properties.LineSpacing = 360;
            para.Properties.LineSpacingRule = "auto";
            model.Content.Add(para);

            var writer = new DocWriter();
            using var ms = new MemoryStream();

            writer.WriteDocBlocks(model, ms);
            ms.Position = 0;

            using var compoundFile = new OpenMcdf.CompoundFile(ms);
            Assert.True(compoundFile.RootStorage.TryGetStream("WordDocument", out var wordDocStream));

            byte[] wordDocData = wordDocStream.GetData();

            Assert.True(ContainsSubsequence(wordDocData, new byte[] { 0x24, 0x26, 0x68, 0x01 }));
        }

        [Fact]
        public void WriteDocBlocks_ParagraphIndent_WritesIndentSprmsIntoPapx()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            var model = new DocumentModel();
            var para = new ParagraphModel();
            para.Runs.Add(new RunModel { Text = "IndentTest" });
            para.Properties.LeftIndentTwips = 720;
            para.Properties.RightIndentTwips = 360;
            para.Properties.FirstLineIndentTwips = -240;
            model.Content.Add(para);

            var writer = new DocWriter();
            using var ms = new MemoryStream();

            writer.WriteDocBlocks(model, ms);
            ms.Position = 0;

            using var compoundFile = new OpenMcdf.CompoundFile(ms);
            Assert.True(compoundFile.RootStorage.TryGetStream("WordDocument", out var wordDocStream));

            byte[] wordDocData = wordDocStream.GetData();

            Assert.True(ContainsSubsequence(wordDocData, new byte[] { 0x0E, 0x84, 0x68, 0x01 }));
            Assert.True(ContainsSubsequence(wordDocData, new byte[] { 0x0F, 0x84, 0xD0, 0x02 }));
            Assert.True(ContainsSubsequence(wordDocData, new byte[] { 0x11, 0x84, 0x10, 0xFF }));
        }

        [Fact]
        public void WriteDocBlocks_ParagraphKeepFlags_WritesKeepSprmsIntoPapx()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            var model = new DocumentModel();
            var para = new ParagraphModel();
            para.Runs.Add(new RunModel { Text = "KeepFlagsTest" });
            para.Properties.KeepNext = true;
            para.Properties.KeepLines = true;
            model.Content.Add(para);

            var writer = new DocWriter();
            using var ms = new MemoryStream();

            writer.WriteDocBlocks(model, ms);
            ms.Position = 0;

            using var compoundFile = new OpenMcdf.CompoundFile(ms);
            Assert.True(compoundFile.RootStorage.TryGetStream("WordDocument", out var wordDocStream));

            byte[] wordDocData = wordDocStream.GetData();

            Assert.True(ContainsSubsequence(wordDocData, new byte[] { 0x05, 0x24, 0x01 }));
            Assert.True(ContainsSubsequence(wordDocData, new byte[] { 0x06, 0x24, 0x01 }));
        }

        [Fact]
        public void WriteDocBlocks_ParagraphWidowControl_WritesWidowControlSprmIntoPapx()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            var model = new DocumentModel();
            var para = new ParagraphModel();
            para.Runs.Add(new RunModel { Text = "WidowControlTest" });
            para.Properties.WidowControl = true;
            model.Content.Add(para);

            var writer = new DocWriter();
            using var ms = new MemoryStream();

            writer.WriteDocBlocks(model, ms);
            ms.Position = 0;

            using var compoundFile = new OpenMcdf.CompoundFile(ms);
            Assert.True(compoundFile.RootStorage.TryGetStream("WordDocument", out var wordDocStream));

            byte[] wordDocData = wordDocStream.GetData();

            Assert.True(ContainsSubsequence(wordDocData, new byte[] { 0x07, 0x24, 0x01 }));
        }

        [Fact]
        public void WriteDocBlocks_ParagraphContextualSpacing_WritesContextualSpacingSprmIntoPapx()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            var model = new DocumentModel();
            var para = new ParagraphModel();
            para.Runs.Add(new RunModel { Text = "ContextSpacingTest" });
            para.Properties.ContextualSpacing = true;
            para.Properties.ParagraphStyleId = "Heading1";
            model.Content.Add(para);

            var writer = new DocWriter();
            using var ms = new MemoryStream();

            writer.WriteDocBlocks(model, ms);
            ms.Position = 0;

            using var compoundFile = new OpenMcdf.CompoundFile(ms);
            Assert.True(compoundFile.RootStorage.TryGetStream("WordDocument", out var wordDocStream));

            byte[] wordDocData = wordDocStream.GetData();

            Assert.True(ContainsSubsequence(wordDocData, new byte[] { 0x44, 0x24, 0x01 }));
        }

        [Fact]
        public void WriteDocBlocks_ParagraphPageBreakBefore_WritesPageBreakBeforeSprmIntoPapx()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            var model = new DocumentModel();
            var para = new ParagraphModel();
            para.Runs.Add(new RunModel { Text = "PageBreakBeforeTest" });
            para.Properties.PageBreakBefore = true;
            model.Content.Add(para);

            var writer = new DocWriter();
            using var ms = new MemoryStream();

            writer.WriteDocBlocks(model, ms);
            ms.Position = 0;

            using var compoundFile = new OpenMcdf.CompoundFile(ms);
            Assert.True(compoundFile.RootStorage.TryGetStream("WordDocument", out var wordDocStream));

            byte[] wordDocData = wordDocStream.GetData();

            Assert.True(ContainsSubsequence(wordDocData, new byte[] { 0x08, 0x24, 0x01 }));
        }

        [Fact]
        public void WriteDocBlocks_RunColor_WritesColorSprmIntoChpx()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            var model = new DocumentModel();
            var para = new ParagraphModel();
            var run = new RunModel { Text = "ColorTest" };
            run.Properties.Color = "FF0000";
            run.Properties.ColorSpecified = true;
            para.Runs.Add(run);
            model.Content.Add(para);

            var writer = new DocWriter();
            using var ms = new MemoryStream();

            writer.WriteDocBlocks(model, ms);
            ms.Position = 0;

            using var compoundFile = new OpenMcdf.CompoundFile(ms);
            Assert.True(compoundFile.RootStorage.TryGetStream("WordDocument", out var wordDocStream));

            byte[] wordDocData = wordDocStream.GetData();

            Assert.True(ContainsSubsequence(wordDocData, new byte[] { 0x42, 0x2A, 0x06 }));
        }

        private static bool ContainsSubsequence(byte[] buffer, byte[] subsequence)
        {
            if (subsequence.Length == 0 || buffer.Length < subsequence.Length)
            {
                return false;
            }

            for (int index = 0; index <= buffer.Length - subsequence.Length; index++)
            {
                bool match = true;
                for (int innerIndex = 0; innerIndex < subsequence.Length; innerIndex++)
                {
                    if (buffer[index + innerIndex] != subsequence[innerIndex])
                    {
                        match = false;
                        break;
                    }
                }

                if (match)
                {
                    return true;
                }
            }

            return false;
        }
    }
}
