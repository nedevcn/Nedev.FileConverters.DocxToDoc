using System.IO;
using System.Text;
using Nedev.FileConverters.DocxToDoc.Format;
using Nedev.FileConverters.DocxToDoc.Model;
using Xunit;

namespace Nedev.FileConverters.DocxToDoc.Tests.Format
{
    public class DocWriterFootnoteTests
    {
        [Fact]
        public void WriteDocBlocks_WithFootnote_WritesReferenceAndFootnoteStory()
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
            Assert.Equal(6, BitConverter.ToInt32(wordDocData, 68));

            int fcPlcffndRef = BitConverter.ToInt32(wordDocData, 154 + (Fib.FootnoteReferencePairIndex * 8));
            int lcbPlcffndRef = BitConverter.ToInt32(wordDocData, 154 + (Fib.FootnoteReferencePairIndex * 8) + 4);
            int fcPlcffndTxt = BitConverter.ToInt32(wordDocData, 154 + (Fib.FootnoteTextPairIndex * 8));
            int lcbPlcffndTxt = BitConverter.ToInt32(wordDocData, 154 + (Fib.FootnoteTextPairIndex * 8) + 4);

            Assert.NotEqual(0, fcPlcffndRef);
            Assert.True(lcbPlcffndRef > 0);
            Assert.NotEqual(0, fcPlcffndTxt);
            Assert.True(lcbPlcffndTxt > 0);
            Assert.Equal(2, BitConverter.ToInt32(tableData, fcPlcffndRef));
            Assert.Equal(6, BitConverter.ToInt32(tableData, fcPlcffndRef + 4));
            Assert.Equal((ushort)1, BitConverter.ToUInt16(tableData, fcPlcffndRef + 8));
            Assert.Equal(0, BitConverter.ToInt32(tableData, fcPlcffndTxt));
            Assert.Equal(5, BitConverter.ToInt32(tableData, fcPlcffndTxt + 4));
            Assert.Equal(6, BitConverter.ToInt32(tableData, fcPlcffndTxt + 8));
        }

        [Fact]
        public void WriteDocBlocks_WithMultiParagraphFootnote_PreservesInternalParagraphMarksInStory()
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
            Assert.Equal(14, BitConverter.ToInt32(wordDocData, 68));

            int fcPlcffndTxt = BitConverter.ToInt32(wordDocData, 154 + (Fib.FootnoteTextPairIndex * 8));
            Assert.Equal(0, BitConverter.ToInt32(tableData, fcPlcffndTxt));
            Assert.Equal(13, BitConverter.ToInt32(tableData, fcPlcffndTxt + 4));
            Assert.Equal(14, BitConverter.ToInt32(tableData, fcPlcffndTxt + 8));
        }

        [Fact]
        public void WriteDocBlocks_WithTrailingEmptyFootnoteParagraph_PreservesFinalEmptyParagraph()
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
            Assert.Equal(8, BitConverter.ToInt32(wordDocData, 68));

            int fcPlcffndTxt = BitConverter.ToInt32(wordDocData, 154 + (Fib.FootnoteTextPairIndex * 8));
            Assert.Equal(0, BitConverter.ToInt32(tableData, fcPlcffndTxt));
            Assert.Equal(7, BitConverter.ToInt32(tableData, fcPlcffndTxt + 4));
            Assert.Equal(8, BitConverter.ToInt32(tableData, fcPlcffndTxt + 8));
        }

        [Fact]
        public void WriteDocBlocks_WithFootnoteAndComment_PreservesIndependentStoryCpAccounting()
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

            var writer = new DocWriter();
            using var ms = new MemoryStream();

            writer.WriteDocBlocks(model, ms);
            ms.Position = 0;

            using var compoundFile = new OpenMcdf.CompoundFile(ms);
            Assert.True(compoundFile.RootStorage.TryGetStream("WordDocument", out var wordDocStream));
            Assert.True(compoundFile.RootStorage.TryGetStream("1Table", out var tableStream));

            var wordDocData = wordDocStream.GetData();
            var tableData = tableStream.GetData();
            string expectedText = "A\x0002BC\x0005D\r\x0002Fn\r\x0005Cm\r";
            var textBytes = new byte[expectedText.Length];
            System.Array.Copy(wordDocData, 1536, textBytes, 0, expectedText.Length);
            var extractedText = Encoding.GetEncoding(1252).GetString(textBytes);

            Assert.Equal(expectedText, extractedText);
            Assert.Equal(7, BitConverter.ToInt32(wordDocData, 64));
            Assert.Equal(4, BitConverter.ToInt32(wordDocData, 68));
            Assert.Equal(4, BitConverter.ToInt32(wordDocData, 76));

            int fcPlcffndRef = BitConverter.ToInt32(wordDocData, 154 + (Fib.FootnoteReferencePairIndex * 8));
            int fcPlcfandRef = BitConverter.ToInt32(wordDocData, 154 + (Fib.CommentReferencePairIndex * 8));

            Assert.Equal(1, BitConverter.ToInt32(tableData, fcPlcffndRef));
            Assert.Equal(4, BitConverter.ToInt32(tableData, fcPlcfandRef));
        }

        [Fact]
        public void WriteDocBlocks_WithReservedAndInvalidFootnotes_FiltersUnsupportedEntries()
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
                Id = "0",
                Text = "Reserved",
                ReferenceCp = 1
            });
            model.Footnotes.Add(new FootnoteModel
            {
                Id = "9",
                Text = "Invalid anchor",
                ReferenceCp = -1
            });
            model.Footnotes.Add(new FootnoteModel
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
            Assert.Equal(6, BitConverter.ToInt32(wordDocData, 68));

            int fcPlcffndRef = BitConverter.ToInt32(wordDocData, 154 + (Fib.FootnoteReferencePairIndex * 8));
            Assert.Equal(2, BitConverter.ToInt32(tableData, fcPlcffndRef));
            Assert.Equal(6, BitConverter.ToInt32(tableData, fcPlcffndRef + 4));
        }

        [Fact]
        public void WriteDocBlocks_WithFootnoteSeparatorStories_WritesHeaderDocumentStories()
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
                Text = "Note",
                ReferenceCp = 2
            });
            model.FootnoteSeparatorText = "Sep";
            model.FootnoteContinuationSeparatorText = "Cont";

            var writer = new DocWriter();
            using var ms = new MemoryStream();

            writer.WriteDocBlocks(model, ms);
            ms.Position = 0;

            using var compoundFile = new OpenMcdf.CompoundFile(ms);
            Assert.True(compoundFile.RootStorage.TryGetStream("WordDocument", out var wordDocStream));
            Assert.True(compoundFile.RootStorage.TryGetStream("1Table", out var tableStream));

            var wordDocData = wordDocStream.GetData();
            var tableData = tableStream.GetData();
            string expectedText = "AB\x0002CD\r\x0002Note\rSep\rCont\r";
            var textBytes = new byte[expectedText.Length];
            System.Array.Copy(wordDocData, 1536, textBytes, 0, expectedText.Length);
            var extractedText = Encoding.GetEncoding(1252).GetString(textBytes);

            Assert.Equal(expectedText, extractedText);
            Assert.Equal(6, BitConverter.ToInt32(wordDocData, 64));
            Assert.Equal(6, BitConverter.ToInt32(wordDocData, 68));
            Assert.Equal(9, BitConverter.ToInt32(wordDocData, 72));

            int fcPlcfHdd = BitConverter.ToInt32(wordDocData, 154 + (Fib.HeaderStoryPairIndex * 8));
            int lcbPlcfHdd = BitConverter.ToInt32(wordDocData, 154 + (Fib.HeaderStoryPairIndex * 8) + 4);
            int fcClx = BitConverter.ToInt32(wordDocData, 154 + (Fib.ClxPairIndex * 8));
            int lcbClx = BitConverter.ToInt32(wordDocData, 154 + (Fib.ClxPairIndex * 8) + 4);

            Assert.NotEqual(0, fcPlcfHdd);
            Assert.True(lcbPlcfHdd > 0);
            Assert.Equal(0, fcClx);
            Assert.True(lcbClx > 0);
            Assert.Equal(0, BitConverter.ToInt32(tableData, fcPlcfHdd));
            Assert.Equal(4, BitConverter.ToInt32(tableData, fcPlcfHdd + 4));
            Assert.Equal(9, BitConverter.ToInt32(tableData, fcPlcfHdd + 8));
            Assert.Equal(9, BitConverter.ToInt32(tableData, fcPlcfHdd + 12));
            Assert.Equal(9, BitConverter.ToInt32(tableData, fcPlcfHdd + 16));
            Assert.Equal(9, BitConverter.ToInt32(tableData, fcPlcfHdd + 20));
            Assert.Equal(8, BitConverter.ToInt32(tableData, fcPlcfHdd + 24));
            Assert.Equal(9, BitConverter.ToInt32(tableData, fcPlcfHdd + 28));
        }

        [Fact]
        public void WriteDocBlocks_WithoutNoteSeparatorStories_KeepsHeaderDocumentEmpty()
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
                Text = "Note",
                ReferenceCp = 2
            });

            var writer = new DocWriter();
            using var ms = new MemoryStream();

            writer.WriteDocBlocks(model, ms);
            ms.Position = 0;

            using var compoundFile = new OpenMcdf.CompoundFile(ms);
            Assert.True(compoundFile.RootStorage.TryGetStream("WordDocument", out var wordDocStream));

            var wordDocData = wordDocStream.GetData();
            Assert.Equal(0, BitConverter.ToInt32(wordDocData, 72));
            Assert.Equal(0, BitConverter.ToInt32(wordDocData, 154 + (Fib.HeaderStoryPairIndex * 8)));
            Assert.Equal(0, BitConverter.ToInt32(wordDocData, 154 + (Fib.HeaderStoryPairIndex * 8) + 4));
        }

        [Fact]
        public void WriteDocBlocks_WithFootnoteContinuationNotice_WritesHeaderDocumentStories()
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
                Text = "Live",
                ReferenceCp = 2
            });
            model.FootnoteSeparatorText = "Sep";
            model.FootnoteContinuationSeparatorText = "Cont";
            model.FootnoteContinuationNoticeText = "Note";

            var writer = new DocWriter();
            using var ms = new MemoryStream();

            writer.WriteDocBlocks(model, ms);
            ms.Position = 0;

            using var compoundFile = new OpenMcdf.CompoundFile(ms);
            Assert.True(compoundFile.RootStorage.TryGetStream("WordDocument", out var wordDocStream));
            Assert.True(compoundFile.RootStorage.TryGetStream("1Table", out var tableStream));

            var wordDocData = wordDocStream.GetData();
            var tableData = tableStream.GetData();
            string expectedText = "AB\x0002CD\r\x0002Live\rSep\rCont\rNote\r";
            var textBytes = new byte[expectedText.Length];
            System.Array.Copy(wordDocData, 1536, textBytes, 0, expectedText.Length);
            var extractedText = Encoding.GetEncoding(1252).GetString(textBytes);

            Assert.Equal(expectedText, extractedText);
            Assert.Equal(6, BitConverter.ToInt32(wordDocData, 64));
            Assert.Equal(6, BitConverter.ToInt32(wordDocData, 68));
            Assert.Equal(14, BitConverter.ToInt32(wordDocData, 72));

            int fcPlcfHdd = BitConverter.ToInt32(wordDocData, 154 + (Fib.HeaderStoryPairIndex * 8));
            Assert.Equal(0, BitConverter.ToInt32(tableData, fcPlcfHdd));
            Assert.Equal(4, BitConverter.ToInt32(tableData, fcPlcfHdd + 4));
            Assert.Equal(9, BitConverter.ToInt32(tableData, fcPlcfHdd + 8));
            Assert.Equal(14, BitConverter.ToInt32(tableData, fcPlcfHdd + 12));
            Assert.Equal(14, BitConverter.ToInt32(tableData, fcPlcfHdd + 16));
            Assert.Equal(14, BitConverter.ToInt32(tableData, fcPlcfHdd + 20));
            Assert.Equal(13, BitConverter.ToInt32(tableData, fcPlcfHdd + 24));
            Assert.Equal(14, BitConverter.ToInt32(tableData, fcPlcfHdd + 28));
        }

        [Fact]
        public void WriteDocBlocks_WithEmptyFootnoteContinuationNotice_WritesGuardParagraphOnly()
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
                Text = "Note",
                ReferenceCp = 2
            });
            model.FootnoteContinuationNoticeText = string.Empty;

            var writer = new DocWriter();
            using var ms = new MemoryStream();

            writer.WriteDocBlocks(model, ms);
            ms.Position = 0;

            using var compoundFile = new OpenMcdf.CompoundFile(ms);
            Assert.True(compoundFile.RootStorage.TryGetStream("WordDocument", out var wordDocStream));
            Assert.True(compoundFile.RootStorage.TryGetStream("1Table", out var tableStream));

            var wordDocData = wordDocStream.GetData();
            var tableData = tableStream.GetData();
            string expectedText = "AB\x0002CD\r\x0002Note\r\r";
            var textBytes = new byte[expectedText.Length];
            System.Array.Copy(wordDocData, 1536, textBytes, 0, expectedText.Length);
            var extractedText = Encoding.GetEncoding(1252).GetString(textBytes);

            Assert.Equal(expectedText, extractedText);
            Assert.Equal(6, BitConverter.ToInt32(wordDocData, 64));
            Assert.Equal(6, BitConverter.ToInt32(wordDocData, 68));
            Assert.Equal(1, BitConverter.ToInt32(wordDocData, 72));

            int fcPlcfHdd = BitConverter.ToInt32(wordDocData, 154 + (Fib.HeaderStoryPairIndex * 8));
            Assert.Equal(0, BitConverter.ToInt32(tableData, fcPlcfHdd));
            Assert.Equal(0, BitConverter.ToInt32(tableData, fcPlcfHdd + 4));
            Assert.Equal(0, BitConverter.ToInt32(tableData, fcPlcfHdd + 8));
            Assert.Equal(1, BitConverter.ToInt32(tableData, fcPlcfHdd + 12));
            Assert.Equal(1, BitConverter.ToInt32(tableData, fcPlcfHdd + 16));
            Assert.Equal(1, BitConverter.ToInt32(tableData, fcPlcfHdd + 20));
            Assert.Equal(0, BitConverter.ToInt32(tableData, fcPlcfHdd + 24));
            Assert.Equal(1, BitConverter.ToInt32(tableData, fcPlcfHdd + 28));
        }

        [Fact]
        public void WriteDocBlocks_WithFootnoteMultiCharacterCustomMark_WritesVisibleReferenceMarker()
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
                Text = "Note",
                ReferenceCp = 2,
                CustomMarkText = "[12]"
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
            string expectedText = "AB[12]CD\r[12]Note\r";
            var textBytes = new byte[expectedText.Length];
            System.Array.Copy(wordDocData, 1536, textBytes, 0, expectedText.Length);
            var extractedText = Encoding.GetEncoding(1252).GetString(textBytes);

            Assert.Equal(expectedText, extractedText);
            Assert.Equal(9, BitConverter.ToInt32(wordDocData, 64));
            Assert.Equal(9, BitConverter.ToInt32(wordDocData, 68));

            int fcPlcffndRef = BitConverter.ToInt32(wordDocData, 154 + (Fib.FootnoteReferencePairIndex * 8));
            int fcPlcffndTxt = BitConverter.ToInt32(wordDocData, 154 + (Fib.FootnoteTextPairIndex * 8));
            Assert.Equal(2, BitConverter.ToInt32(tableData, fcPlcffndRef));
            Assert.Equal(9, BitConverter.ToInt32(tableData, fcPlcffndRef + 4));
            Assert.Equal(0, BitConverter.ToInt32(tableData, fcPlcffndTxt));
            Assert.Equal(8, BitConverter.ToInt32(tableData, fcPlcffndTxt + 4));
            Assert.Equal(9, BitConverter.ToInt32(tableData, fcPlcffndTxt + 8));
        }

        [Fact]
        public void WriteDocBlocks_WithEmptyFootnoteCustomMark_FallsBackToDefaultMarker()
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
            Assert.Equal(6, BitConverter.ToInt32(wordDocData, 68));
        }
    }
}
