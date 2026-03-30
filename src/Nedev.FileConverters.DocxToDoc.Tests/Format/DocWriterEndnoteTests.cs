using System.IO;
using System.Text;
using Nedev.FileConverters.DocxToDoc.Format;
using Nedev.FileConverters.DocxToDoc.Model;
using Xunit;

namespace Nedev.FileConverters.DocxToDoc.Tests.Format
{
    public class DocWriterEndnoteTests
    {
        [Fact]
        public void WriteDocBlocks_WithEndnote_WritesReferenceAndEndnoteStory()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            var model = new DocumentModel();
            model.Content.Add(new ParagraphModel
            {
                Runs =
                {
                    new RunModel { Text = "ABCD" }
                }
            });
            model.Endnotes.Add(new EndnoteModel
            {
                Id = "2",
                Text = "Note",
                ReferenceCp = 2
            });

            var writer = new DocWriter();
            using var ms = new MemoryStream();

            writer.WriteDocBlocks(model, ms);
            ms.Position = 0;

            using var compoundFile = new OpenMcdf.CompoundFile(ms);
            Assert.True(compoundFile.RootStorage.TryGetStream("WordDocument", out var wordDocStream));
            Assert.True(compoundFile.RootStorage.TryGetStream("1Table", out var tableStream));

            var wordDocData = wordDocStream.GetData();
            var tableData = tableStream.GetData();
            string expectedText = "AB\x0002CD\r\x0002Note\r";
            var textBytes = new byte[expectedText.Length];
            System.Array.Copy(wordDocData, 1536, textBytes, 0, expectedText.Length);
            var extractedText = Encoding.GetEncoding(1252).GetString(textBytes);

            Assert.Equal(expectedText, extractedText);
            Assert.Equal(6, BitConverter.ToInt32(wordDocData, 64));
            Assert.Equal(6, BitConverter.ToInt32(wordDocData, 80));

            int fcPlcfendRef = BitConverter.ToInt32(wordDocData, 154 + (Fib.EndnoteReferencePairIndex * 8));
            int lcbPlcfendRef = BitConverter.ToInt32(wordDocData, 154 + (Fib.EndnoteReferencePairIndex * 8) + 4);
            int fcPlcfendTxt = BitConverter.ToInt32(wordDocData, 154 + (Fib.EndnoteTextPairIndex * 8));
            int lcbPlcfendTxt = BitConverter.ToInt32(wordDocData, 154 + (Fib.EndnoteTextPairIndex * 8) + 4);

            Assert.NotEqual(0, fcPlcfendRef);
            Assert.True(lcbPlcfendRef > 0);
            Assert.NotEqual(0, fcPlcfendTxt);
            Assert.True(lcbPlcfendTxt > 0);
            Assert.Equal(2, BitConverter.ToInt32(tableData, fcPlcfendRef));
            Assert.Equal(6, BitConverter.ToInt32(tableData, fcPlcfendRef + 4));
            Assert.Equal((ushort)1, BitConverter.ToUInt16(tableData, fcPlcfendRef + 8));
            Assert.Equal(0, BitConverter.ToInt32(tableData, fcPlcfendTxt));
            Assert.Equal(5, BitConverter.ToInt32(tableData, fcPlcfendTxt + 4));
            Assert.Equal(6, BitConverter.ToInt32(tableData, fcPlcfendTxt + 8));
        }

        [Fact]
        public void WriteDocBlocks_WithMultiParagraphEndnote_PreservesInternalParagraphMarksInStory()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            var model = new DocumentModel();
            model.Content.Add(new ParagraphModel
            {
                Runs =
                {
                    new RunModel { Text = "ABCD" }
                }
            });
            model.Endnotes.Add(new EndnoteModel
            {
                Id = "2",
                Text = "First\rSecond",
                ReferenceCp = 2
            });

            var writer = new DocWriter();
            using var ms = new MemoryStream();

            writer.WriteDocBlocks(model, ms);
            ms.Position = 0;

            using var compoundFile = new OpenMcdf.CompoundFile(ms);
            Assert.True(compoundFile.RootStorage.TryGetStream("WordDocument", out var wordDocStream));
            Assert.True(compoundFile.RootStorage.TryGetStream("1Table", out var tableStream));

            var wordDocData = wordDocStream.GetData();
            var tableData = tableStream.GetData();
            string expectedText = "AB\x0002CD\r\x0002First\rSecond\r";
            var textBytes = new byte[expectedText.Length];
            System.Array.Copy(wordDocData, 1536, textBytes, 0, expectedText.Length);
            var extractedText = Encoding.GetEncoding(1252).GetString(textBytes);

            Assert.Equal(expectedText, extractedText);
            Assert.Equal(6, BitConverter.ToInt32(wordDocData, 64));
            Assert.Equal(14, BitConverter.ToInt32(wordDocData, 80));

            int fcPlcfendTxt = BitConverter.ToInt32(wordDocData, 154 + (Fib.EndnoteTextPairIndex * 8));
            Assert.Equal(0, BitConverter.ToInt32(tableData, fcPlcfendTxt));
            Assert.Equal(13, BitConverter.ToInt32(tableData, fcPlcfendTxt + 4));
            Assert.Equal(14, BitConverter.ToInt32(tableData, fcPlcfendTxt + 8));
        }

        [Fact]
        public void WriteDocBlocks_WithTrailingEmptyEndnoteParagraph_PreservesFinalEmptyParagraph()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            var model = new DocumentModel();
            model.Content.Add(new ParagraphModel
            {
                Runs =
                {
                    new RunModel { Text = "ABCD" }
                }
            });
            model.Endnotes.Add(new EndnoteModel
            {
                Id = "2",
                Text = "First\r",
                ReferenceCp = 2
            });

            var writer = new DocWriter();
            using var ms = new MemoryStream();

            writer.WriteDocBlocks(model, ms);
            ms.Position = 0;

            using var compoundFile = new OpenMcdf.CompoundFile(ms);
            Assert.True(compoundFile.RootStorage.TryGetStream("WordDocument", out var wordDocStream));
            Assert.True(compoundFile.RootStorage.TryGetStream("1Table", out var tableStream));

            var wordDocData = wordDocStream.GetData();
            var tableData = tableStream.GetData();
            string expectedText = "AB\x0002CD\r\x0002First\r\r";
            var textBytes = new byte[expectedText.Length];
            System.Array.Copy(wordDocData, 1536, textBytes, 0, expectedText.Length);
            var extractedText = Encoding.GetEncoding(1252).GetString(textBytes);

            Assert.Equal(expectedText, extractedText);
            Assert.Equal(6, BitConverter.ToInt32(wordDocData, 64));
            Assert.Equal(8, BitConverter.ToInt32(wordDocData, 80));

            int fcPlcfendTxt = BitConverter.ToInt32(wordDocData, 154 + (Fib.EndnoteTextPairIndex * 8));
            Assert.Equal(0, BitConverter.ToInt32(tableData, fcPlcfendTxt));
            Assert.Equal(7, BitConverter.ToInt32(tableData, fcPlcfendTxt + 4));
            Assert.Equal(8, BitConverter.ToInt32(tableData, fcPlcfendTxt + 8));
        }

        [Fact]
        public void WriteDocBlocks_WithFootnoteCommentAndEndnote_PreservesIndependentStoryCpAccounting()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            var model = new DocumentModel();
            model.Content.Add(new ParagraphModel
            {
                Runs =
                {
                    new RunModel { Text = "ABCD" }
                }
            });
            model.Footnotes.Add(new FootnoteModel
            {
                Id = "2",
                Text = "Fn",
                ReferenceCp = 1
            });
            model.Comments.Add(new CommentModel
            {
                Id = "0",
                Text = "Cm",
                StartCp = 3,
                EndCp = 3
            });
            model.Endnotes.Add(new EndnoteModel
            {
                Id = "4",
                Text = "En",
                ReferenceCp = 4
            });

            var writer = new DocWriter();
            using var ms = new MemoryStream();

            writer.WriteDocBlocks(model, ms);
            ms.Position = 0;

            using var compoundFile = new OpenMcdf.CompoundFile(ms);
            Assert.True(compoundFile.RootStorage.TryGetStream("WordDocument", out var wordDocStream));
            Assert.True(compoundFile.RootStorage.TryGetStream("1Table", out var tableStream));

            var wordDocData = wordDocStream.GetData();
            var tableData = tableStream.GetData();
            string expectedText = "A\x0002BC\x0005D\x0002\r\x0002Fn\r\x0005Cm\r\x0002En\r";
            var textBytes = new byte[expectedText.Length];
            System.Array.Copy(wordDocData, 1536, textBytes, 0, expectedText.Length);
            var extractedText = Encoding.GetEncoding(1252).GetString(textBytes);

            Assert.Equal(expectedText, extractedText);
            Assert.Equal(8, BitConverter.ToInt32(wordDocData, 64));
            Assert.Equal(4, BitConverter.ToInt32(wordDocData, 68));
            Assert.Equal(4, BitConverter.ToInt32(wordDocData, 76));
            Assert.Equal(4, BitConverter.ToInt32(wordDocData, 80));

            int fcPlcffndRef = BitConverter.ToInt32(wordDocData, 154 + (Fib.FootnoteReferencePairIndex * 8));
            int fcPlcfandRef = BitConverter.ToInt32(wordDocData, 154 + (Fib.CommentReferencePairIndex * 8));
            int fcPlcfendRef = BitConverter.ToInt32(wordDocData, 154 + (Fib.EndnoteReferencePairIndex * 8));

            Assert.Equal(1, BitConverter.ToInt32(tableData, fcPlcffndRef));
            Assert.Equal(4, BitConverter.ToInt32(tableData, fcPlcfandRef));
            Assert.Equal(6, BitConverter.ToInt32(tableData, fcPlcfendRef));
        }

        [Fact]
        public void WriteDocBlocks_WithReservedAndInvalidEndnotes_FiltersUnsupportedEntries()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            var model = new DocumentModel();
            model.Content.Add(new ParagraphModel
            {
                Runs =
                {
                    new RunModel { Text = "ABCD" }
                }
            });
            model.Endnotes.Add(new EndnoteModel
            {
                Id = "0",
                Text = "Reserved",
                ReferenceCp = 1
            });
            model.Endnotes.Add(new EndnoteModel
            {
                Id = "9",
                Text = "Invalid anchor",
                ReferenceCp = -1
            });
            model.Endnotes.Add(new EndnoteModel
            {
                Id = "2",
                Text = "Live",
                ReferenceCp = 2
            });

            var writer = new DocWriter();
            using var ms = new MemoryStream();

            writer.WriteDocBlocks(model, ms);
            ms.Position = 0;

            using var compoundFile = new OpenMcdf.CompoundFile(ms);
            Assert.True(compoundFile.RootStorage.TryGetStream("WordDocument", out var wordDocStream));
            Assert.True(compoundFile.RootStorage.TryGetStream("1Table", out var tableStream));

            var wordDocData = wordDocStream.GetData();
            var tableData = tableStream.GetData();
            string expectedText = "AB\x0002CD\r\x0002Live\r";
            var textBytes = new byte[expectedText.Length];
            System.Array.Copy(wordDocData, 1536, textBytes, 0, expectedText.Length);
            var extractedText = Encoding.GetEncoding(1252).GetString(textBytes);

            Assert.Equal(expectedText, extractedText);
            Assert.Equal(6, BitConverter.ToInt32(wordDocData, 64));
            Assert.Equal(6, BitConverter.ToInt32(wordDocData, 80));

            int fcPlcfendRef = BitConverter.ToInt32(wordDocData, 154 + (Fib.EndnoteReferencePairIndex * 8));
            Assert.Equal(2, BitConverter.ToInt32(tableData, fcPlcfendRef));
            Assert.Equal(6, BitConverter.ToInt32(tableData, fcPlcfendRef + 4));
        }

        [Fact]
        public void WriteDocBlocks_WithEndnoteSeparatorStories_WritesHeaderDocumentStories()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            var model = new DocumentModel();
            model.Content.Add(new ParagraphModel
            {
                Runs =
                {
                    new RunModel { Text = "ABCD" }
                }
            });
            model.Endnotes.Add(new EndnoteModel
            {
                Id = "2",
                Text = "Note",
                ReferenceCp = 2
            });
            model.EndnoteSeparatorText = "End Sep";
            model.EndnoteContinuationSeparatorText = "End Cont";

            var writer = new DocWriter();
            using var ms = new MemoryStream();

            writer.WriteDocBlocks(model, ms);
            ms.Position = 0;

            using var compoundFile = new OpenMcdf.CompoundFile(ms);
            Assert.True(compoundFile.RootStorage.TryGetStream("WordDocument", out var wordDocStream));
            Assert.True(compoundFile.RootStorage.TryGetStream("1Table", out var tableStream));

            var wordDocData = wordDocStream.GetData();
            var tableData = tableStream.GetData();
            string expectedText = "AB\x0002CD\rEnd Sep\rEnd Cont\r\x0002Note\r";
            var textBytes = new byte[expectedText.Length];
            System.Array.Copy(wordDocData, 1536, textBytes, 0, expectedText.Length);
            var extractedText = Encoding.GetEncoding(1252).GetString(textBytes);

            Assert.Equal(expectedText, extractedText);
            Assert.Equal(6, BitConverter.ToInt32(wordDocData, 64));
            Assert.Equal(17, BitConverter.ToInt32(wordDocData, 72));
            Assert.Equal(6, BitConverter.ToInt32(wordDocData, 80));

            int fcPlcfHdd = BitConverter.ToInt32(wordDocData, 154 + (Fib.HeaderStoryPairIndex * 8));
            int lcbPlcfHdd = BitConverter.ToInt32(wordDocData, 154 + (Fib.HeaderStoryPairIndex * 8) + 4);
            int fcClx = BitConverter.ToInt32(wordDocData, 154 + (Fib.ClxPairIndex * 8));
            int lcbClx = BitConverter.ToInt32(wordDocData, 154 + (Fib.ClxPairIndex * 8) + 4);

            Assert.NotEqual(0, fcPlcfHdd);
            Assert.True(lcbPlcfHdd > 0);
            Assert.Equal(0, fcClx);
            Assert.True(lcbClx > 0);
            Assert.Equal(0, BitConverter.ToInt32(tableData, fcPlcfHdd));
            Assert.Equal(0, BitConverter.ToInt32(tableData, fcPlcfHdd + 4));
            Assert.Equal(0, BitConverter.ToInt32(tableData, fcPlcfHdd + 8));
            Assert.Equal(0, BitConverter.ToInt32(tableData, fcPlcfHdd + 12));
            Assert.Equal(8, BitConverter.ToInt32(tableData, fcPlcfHdd + 16));
            Assert.Equal(17, BitConverter.ToInt32(tableData, fcPlcfHdd + 20));
            Assert.Equal(16, BitConverter.ToInt32(tableData, fcPlcfHdd + 24));
            Assert.Equal(17, BitConverter.ToInt32(tableData, fcPlcfHdd + 28));
        }

        [Fact]
        public void WriteDocBlocks_WithEmptyEndnoteSeparatorStory_WritesGuardParagraphOnly()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            var model = new DocumentModel();
            model.Content.Add(new ParagraphModel
            {
                Runs =
                {
                    new RunModel { Text = "ABCD" }
                }
            });
            model.Endnotes.Add(new EndnoteModel
            {
                Id = "2",
                Text = "Note",
                ReferenceCp = 2
            });
            model.EndnoteSeparatorText = string.Empty;

            var writer = new DocWriter();
            using var ms = new MemoryStream();

            writer.WriteDocBlocks(model, ms);
            ms.Position = 0;

            using var compoundFile = new OpenMcdf.CompoundFile(ms);
            Assert.True(compoundFile.RootStorage.TryGetStream("WordDocument", out var wordDocStream));
            Assert.True(compoundFile.RootStorage.TryGetStream("1Table", out var tableStream));

            var wordDocData = wordDocStream.GetData();
            var tableData = tableStream.GetData();
            string expectedText = "AB\x0002CD\r\r\x0002Note\r";
            var textBytes = new byte[expectedText.Length];
            System.Array.Copy(wordDocData, 1536, textBytes, 0, expectedText.Length);
            var extractedText = Encoding.GetEncoding(1252).GetString(textBytes);

            Assert.Equal(expectedText, extractedText);
            Assert.Equal(6, BitConverter.ToInt32(wordDocData, 64));
            Assert.Equal(1, BitConverter.ToInt32(wordDocData, 72));
            Assert.Equal(6, BitConverter.ToInt32(wordDocData, 80));

            int fcPlcfHdd = BitConverter.ToInt32(wordDocData, 154 + (Fib.HeaderStoryPairIndex * 8));
            Assert.Equal(0, BitConverter.ToInt32(tableData, fcPlcfHdd));
            Assert.Equal(0, BitConverter.ToInt32(tableData, fcPlcfHdd + 4));
            Assert.Equal(0, BitConverter.ToInt32(tableData, fcPlcfHdd + 8));
            Assert.Equal(0, BitConverter.ToInt32(tableData, fcPlcfHdd + 12));
            Assert.Equal(1, BitConverter.ToInt32(tableData, fcPlcfHdd + 16));
            Assert.Equal(1, BitConverter.ToInt32(tableData, fcPlcfHdd + 20));
            Assert.Equal(0, BitConverter.ToInt32(tableData, fcPlcfHdd + 24));
            Assert.Equal(1, BitConverter.ToInt32(tableData, fcPlcfHdd + 28));
        }

        [Fact]
        public void WriteDocBlocks_WithEndnoteContinuationNotice_WritesHeaderDocumentStories()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            var model = new DocumentModel();
            model.Content.Add(new ParagraphModel
            {
                Runs =
                {
                    new RunModel { Text = "ABCD" }
                }
            });
            model.Endnotes.Add(new EndnoteModel
            {
                Id = "2",
                Text = "Live",
                ReferenceCp = 2
            });
            model.EndnoteSeparatorText = "End Sep";
            model.EndnoteContinuationSeparatorText = "End Cont";
            model.EndnoteContinuationNoticeText = "End Note";

            var writer = new DocWriter();
            using var ms = new MemoryStream();

            writer.WriteDocBlocks(model, ms);
            ms.Position = 0;

            using var compoundFile = new OpenMcdf.CompoundFile(ms);
            Assert.True(compoundFile.RootStorage.TryGetStream("WordDocument", out var wordDocStream));
            Assert.True(compoundFile.RootStorage.TryGetStream("1Table", out var tableStream));

            var wordDocData = wordDocStream.GetData();
            var tableData = tableStream.GetData();
            string expectedText = "AB\x0002CD\rEnd Sep\rEnd Cont\rEnd Note\r\x0002Live\r";
            var textBytes = new byte[expectedText.Length];
            System.Array.Copy(wordDocData, 1536, textBytes, 0, expectedText.Length);
            var extractedText = Encoding.GetEncoding(1252).GetString(textBytes);

            Assert.Equal(expectedText, extractedText);
            Assert.Equal(6, BitConverter.ToInt32(wordDocData, 64));
            Assert.Equal(26, BitConverter.ToInt32(wordDocData, 72));
            Assert.Equal(6, BitConverter.ToInt32(wordDocData, 80));

            int fcPlcfHdd = BitConverter.ToInt32(wordDocData, 154 + (Fib.HeaderStoryPairIndex * 8));
            Assert.Equal(0, BitConverter.ToInt32(tableData, fcPlcfHdd));
            Assert.Equal(0, BitConverter.ToInt32(tableData, fcPlcfHdd + 4));
            Assert.Equal(0, BitConverter.ToInt32(tableData, fcPlcfHdd + 8));
            Assert.Equal(0, BitConverter.ToInt32(tableData, fcPlcfHdd + 12));
            Assert.Equal(8, BitConverter.ToInt32(tableData, fcPlcfHdd + 16));
            Assert.Equal(17, BitConverter.ToInt32(tableData, fcPlcfHdd + 20));
            Assert.Equal(25, BitConverter.ToInt32(tableData, fcPlcfHdd + 24));
            Assert.Equal(26, BitConverter.ToInt32(tableData, fcPlcfHdd + 28));
        }

        [Fact]
        public void WriteDocBlocks_WithoutEndnoteContinuationNotice_KeepsNoticeSlotCollapsed()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            var model = new DocumentModel();
            model.Content.Add(new ParagraphModel
            {
                Runs =
                {
                    new RunModel { Text = "ABCD" }
                }
            });
            model.Endnotes.Add(new EndnoteModel
            {
                Id = "2",
                Text = "Note",
                ReferenceCp = 2
            });
            model.EndnoteSeparatorText = "End Sep";
            model.EndnoteContinuationSeparatorText = "End Cont";

            var writer = new DocWriter();
            using var ms = new MemoryStream();

            writer.WriteDocBlocks(model, ms);
            ms.Position = 0;

            using var compoundFile = new OpenMcdf.CompoundFile(ms);
            Assert.True(compoundFile.RootStorage.TryGetStream("WordDocument", out var wordDocStream));
            Assert.True(compoundFile.RootStorage.TryGetStream("1Table", out var tableStream));

            var wordDocData = wordDocStream.GetData();
            var tableData = tableStream.GetData();
            Assert.Equal(17, BitConverter.ToInt32(wordDocData, 72));

            int fcPlcfHdd = BitConverter.ToInt32(wordDocData, 154 + (Fib.HeaderStoryPairIndex * 8));
            Assert.Equal(0, BitConverter.ToInt32(tableData, fcPlcfHdd));
            Assert.Equal(0, BitConverter.ToInt32(tableData, fcPlcfHdd + 4));
            Assert.Equal(0, BitConverter.ToInt32(tableData, fcPlcfHdd + 8));
            Assert.Equal(0, BitConverter.ToInt32(tableData, fcPlcfHdd + 12));
            Assert.Equal(8, BitConverter.ToInt32(tableData, fcPlcfHdd + 16));
            Assert.Equal(17, BitConverter.ToInt32(tableData, fcPlcfHdd + 20));
            Assert.Equal(16, BitConverter.ToInt32(tableData, fcPlcfHdd + 24));
            Assert.Equal(17, BitConverter.ToInt32(tableData, fcPlcfHdd + 28));
        }

        [Fact]
        public void WriteDocBlocks_WithEndnoteMultiCharacterCustomMark_WritesVisibleReferenceMarker()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            var model = new DocumentModel();
            model.Content.Add(new ParagraphModel
            {
                Runs =
                {
                    new RunModel { Text = "ABCD" }
                }
            });
            model.Endnotes.Add(new EndnoteModel
            {
                Id = "2",
                Text = "Note",
                ReferenceCp = 2,
                CustomMarkText = "(a)"
            });

            var writer = new DocWriter();
            using var ms = new MemoryStream();

            writer.WriteDocBlocks(model, ms);
            ms.Position = 0;

            using var compoundFile = new OpenMcdf.CompoundFile(ms);
            Assert.True(compoundFile.RootStorage.TryGetStream("WordDocument", out var wordDocStream));
            Assert.True(compoundFile.RootStorage.TryGetStream("1Table", out var tableStream));

            var wordDocData = wordDocStream.GetData();
            var tableData = tableStream.GetData();
            string expectedText = "AB(a)CD\r(a)Note\r";
            var textBytes = new byte[expectedText.Length];
            System.Array.Copy(wordDocData, 1536, textBytes, 0, expectedText.Length);
            var extractedText = Encoding.GetEncoding(1252).GetString(textBytes);

            Assert.Equal(expectedText, extractedText);
            Assert.Equal(8, BitConverter.ToInt32(wordDocData, 64));
            Assert.Equal(8, BitConverter.ToInt32(wordDocData, 80));

            int fcPlcfendRef = BitConverter.ToInt32(wordDocData, 154 + (Fib.EndnoteReferencePairIndex * 8));
            int fcPlcfendTxt = BitConverter.ToInt32(wordDocData, 154 + (Fib.EndnoteTextPairIndex * 8));
            Assert.Equal(2, BitConverter.ToInt32(tableData, fcPlcfendRef));
            Assert.Equal(8, BitConverter.ToInt32(tableData, fcPlcfendRef + 4));
            Assert.Equal(0, BitConverter.ToInt32(tableData, fcPlcfendTxt));
            Assert.Equal(7, BitConverter.ToInt32(tableData, fcPlcfendTxt + 4));
            Assert.Equal(8, BitConverter.ToInt32(tableData, fcPlcfendTxt + 8));
        }

        [Fact]
        public void WriteDocBlocks_WithEmptyEndnoteCustomMark_FallsBackToDefaultMarker()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            var model = new DocumentModel();
            model.Content.Add(new ParagraphModel
            {
                Runs =
                {
                    new RunModel { Text = "ABCD" }
                }
            });
            model.Endnotes.Add(new EndnoteModel
            {
                Id = "2",
                Text = "Note",
                ReferenceCp = 2,
                CustomMarkText = string.Empty
            });

            var writer = new DocWriter();
            using var ms = new MemoryStream();

            writer.WriteDocBlocks(model, ms);
            ms.Position = 0;

            using var compoundFile = new OpenMcdf.CompoundFile(ms);
            Assert.True(compoundFile.RootStorage.TryGetStream("WordDocument", out var wordDocStream));

            var wordDocData = wordDocStream.GetData();
            string expectedText = "AB\x0002CD\r\x0002Note\r";
            var textBytes = new byte[expectedText.Length];
            System.Array.Copy(wordDocData, 1536, textBytes, 0, expectedText.Length);
            var extractedText = Encoding.GetEncoding(1252).GetString(textBytes);

            Assert.Equal(expectedText, extractedText);
            Assert.Equal(6, BitConverter.ToInt32(wordDocData, 64));
            Assert.Equal(6, BitConverter.ToInt32(wordDocData, 80));
        }

        [Fact]
        public void WriteDocBlocks_WithControlCharacterEndnoteCustomMark_FallsBackToDefaultMarker()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            var model = new DocumentModel();
            model.Content.Add(new ParagraphModel
            {
                Runs =
                {
                    new RunModel { Text = "ABCD" }
                }
            });
            model.Endnotes.Add(new EndnoteModel
            {
                Id = "2",
                Text = "Note",
                ReferenceCp = 2,
                CustomMarkText = "(a)\r"
            });

            var writer = new DocWriter();
            using var ms = new MemoryStream();

            writer.WriteDocBlocks(model, ms);
            ms.Position = 0;

            using var compoundFile = new OpenMcdf.CompoundFile(ms);
            Assert.True(compoundFile.RootStorage.TryGetStream("WordDocument", out var wordDocStream));

            var wordDocData = wordDocStream.GetData();
            string expectedText = "AB\x0002CD\r\x0002Note\r";
            var textBytes = new byte[expectedText.Length];
            System.Array.Copy(wordDocData, 1536, textBytes, 0, expectedText.Length);
            var extractedText = Encoding.GetEncoding(1252).GetString(textBytes);

            Assert.Equal(expectedText, extractedText);
            Assert.Equal(6, BitConverter.ToInt32(wordDocData, 64));
            Assert.Equal(6, BitConverter.ToInt32(wordDocData, 80));
        }
    }
}
