using System;
using System.IO;
using System.Text;
using Nedev.FileConverters.DocxToDoc.Format;
using Nedev.FileConverters.DocxToDoc.Model;
using Xunit;

namespace Nedev.FileConverters.DocxToDoc.Tests.Format
{
    public class DocWriterSectionTests
    {
        [Fact]
        public void WriteDocBlocks_WithSections_WritesPlcfsed()
        {
            // Register encoding provider for Windows-1252 
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Arrange
            var model = new DocumentModel();
            
            model.Content.Add(new ParagraphModel { Runs = { new RunModel { Text = "SectionTest" } } });
            
            var section = new SectionModel
            {
                PageWidth = 12240,
                PageHeight = 15840,
                MarginLeft = 1440,
                MarginRight = 1440
            };
            model.Sections.Add(section);

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

            // Check FIB offset for Plcfsed (index 6 in RgFcLcb, offset 154 + 6*8 = 202)
            int fcPlcfsed = BitConverter.ToInt32(wordDocData, 202);
            int lcbPlcfsed = BitConverter.ToInt32(wordDocData, 206);

            Assert.NotEqual(0, fcPlcfsed);
            Assert.True(lcbPlcfsed >= 4 + 4 + 12); // CP0, CP_End, 1 SED (12 bytes)

            // Verify CP boundaries in Plcfsed (0 and 12)
            Assert.Equal(0, BitConverter.ToInt32(tableData, fcPlcfsed));
            Assert.Equal(12, BitConverter.ToInt32(tableData, fcPlcfsed + 4));

            // Verify SED (at fcPlcfsed + 8)
            // fn = 0
            Assert.Equal(0, BitConverter.ToInt16(tableData, fcPlcfsed + 8));
            // fcSep (offset in WordDocument)
            int fcSep = BitConverter.ToInt32(tableData, fcPlcfsed + 10);
            Assert.True(fcSep > 1536);

            // Verify SEP sprms in WordDocument at fcSep
            // cb (short)
            short cbSep = BitConverter.ToInt16(wordDocData, fcSep);
            Assert.True(cbSep > 0);
            
            // Check first sprm (sprmSXaPage = 0xB603)
            Assert.Equal(0x03, wordDocData[fcSep + 2]);
            Assert.Equal(0xB6, wordDocData[fcSep + 3]);
            Assert.Equal(12240, BitConverter.ToInt16(wordDocData, fcSep + 4));
        }

        [Fact]
        public void WriteDocBlocks_WithDefaultHeaderFooter_WritesHeaderDocumentStoriesAfterNoteSlots()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            var model = new DocumentModel();
            model.Content.Add(new ParagraphModel
            {
                Runs =
                {
                    new RunModel { Text = "Body" }
                }
            });

            model.Sections.Add(new SectionModel
            {
                HeaderText = "Head",
                FooterText = "Foot"
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
            string expectedText = "Body\rHead\r\rFoot\r\r";
            var textBytes = new byte[expectedText.Length];
            System.Array.Copy(wordDocData, 1536, textBytes, 0, expectedText.Length);
            var extractedText = Encoding.GetEncoding(1252).GetString(textBytes);

            Assert.Equal(expectedText, extractedText);
            Assert.Equal(5, BitConverter.ToInt32(wordDocData, 64));
            Assert.Equal(12, BitConverter.ToInt32(wordDocData, 72));

            int fcPlcfHdd = BitConverter.ToInt32(wordDocData, 154 + (Fib.HeaderStoryPairIndex * 8));
            int lcbPlcfHdd = BitConverter.ToInt32(wordDocData, 154 + (Fib.HeaderStoryPairIndex * 8) + 4);

            Assert.NotEqual(0, fcPlcfHdd);
            Assert.Equal(56, lcbPlcfHdd);

            Assert.Equal(0, BitConverter.ToInt32(tableData, fcPlcfHdd));
            Assert.Equal(0, BitConverter.ToInt32(tableData, fcPlcfHdd + 4));
            Assert.Equal(0, BitConverter.ToInt32(tableData, fcPlcfHdd + 8));
            Assert.Equal(0, BitConverter.ToInt32(tableData, fcPlcfHdd + 12));
            Assert.Equal(0, BitConverter.ToInt32(tableData, fcPlcfHdd + 16));
            Assert.Equal(0, BitConverter.ToInt32(tableData, fcPlcfHdd + 20));
            Assert.Equal(0, BitConverter.ToInt32(tableData, fcPlcfHdd + 24));
            Assert.Equal(0, BitConverter.ToInt32(tableData, fcPlcfHdd + 28));
            Assert.Equal(6, BitConverter.ToInt32(tableData, fcPlcfHdd + 32));
            Assert.Equal(6, BitConverter.ToInt32(tableData, fcPlcfHdd + 36));
            Assert.Equal(12, BitConverter.ToInt32(tableData, fcPlcfHdd + 40));
            Assert.Equal(12, BitConverter.ToInt32(tableData, fcPlcfHdd + 44));
            Assert.Equal(11, BitConverter.ToInt32(tableData, fcPlcfHdd + 48));
            Assert.Equal(12, BitConverter.ToInt32(tableData, fcPlcfHdd + 52));
        }

        [Fact]
        public void WriteDocBlocks_WithNestedTableInsideHeaderCellContent_WritesNestedTableMarkers()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            var model = new DocumentModel();
            model.Content.Add(new ParagraphModel
            {
                Runs =
                {
                    new RunModel { Text = "Body" }
                }
            });

            var nestedTable = new TableModel();
            var nestedRow = new TableRowModel();
            var nestedCell = new TableCellModel { Width = 2400 };
            var nestedCellParagraph = new ParagraphModel();
            nestedCellParagraph.Runs.Add(new RunModel { Text = "Cell" });
            nestedCell.Paragraphs.Add(nestedCellParagraph);
            nestedRow.Cells.Add(nestedCell);
            nestedTable.Rows.Add(nestedRow);

            var outerTable = new TableModel();
            var outerRow = new TableRowModel();
            var outerCell = new TableCellModel { Width = 5000 };
            var innerLead = new ParagraphModel();
            innerLead.Runs.Add(new RunModel { Text = "Inner lead" });
            var innerTail = new ParagraphModel();
            innerTail.Runs.Add(new RunModel { Text = "Inner tail" });
            outerCell.Content.Add(innerLead);
            outerCell.Content.Add(nestedTable);
            outerCell.Content.Add(innerTail);
            outerCell.Paragraphs.Add(innerLead);
            outerCell.Paragraphs.Add(nestedCellParagraph);
            outerCell.Paragraphs.Add(innerTail);
            outerRow.Cells.Add(outerCell);
            outerTable.Rows.Add(outerRow);

            var story = new HeaderFooterStoryModel();
            var lead = new ParagraphModel();
            lead.Runs.Add(new RunModel { Text = "Lead" });
            var tail = new ParagraphModel();
            tail.Runs.Add(new RunModel { Text = "Tail" });
            story.Content.Add(lead);
            story.Content.Add(outerTable);
            story.Content.Add(tail);
            story.Paragraphs.Add(lead);
            story.Paragraphs.Add(tail);
            story.Text = "Lead\rInner lead\rCell\rInner tail\rTail";

            model.Sections.Add(new SectionModel
            {
                DefaultHeaderStory = story,
                DefaultHeaderText = story.Text
            });

            var writer = new DocWriter();
            using var ms = new MemoryStream();

            writer.WriteDocBlocks(model, ms);
            ms.Position = 0;

            using var compoundFile = new OpenMcdf.CompoundFile(ms);
            Assert.True(compoundFile.RootStorage.TryGetStream("WordDocument", out var wordDocStream));
            Assert.True(compoundFile.RootStorage.TryGetStream("1Table", out var tableStream));

            var wordDocData = wordDocStream.GetData();
            string expectedText = "Body\rLead\rInner lead\rCell\r\x0007\rInner tail\r\x0007\rTail\r\r";
            var textBytes = new byte[expectedText.Length];
            System.Array.Copy(wordDocData, 1536, textBytes, 0, expectedText.Length);
            var extractedText = Encoding.GetEncoding(1252).GetString(textBytes);

            Assert.Equal(expectedText, extractedText);

            int lcbPlcfbteTapx = BitConverter.ToInt32(wordDocData, 154 + (Fib.TapxPairIndex * 8) + 4);
            Assert.True(lcbPlcfbteTapx > 0);
            Assert.True(GetTapxRunCount(wordDocData, tableStream.GetData()) >= 3);
        }

        [Fact]
        public void WriteDocBlocks_WithNoteSeparatorAndNestedTableInsideHeaderCellContent_PreservesSharedHeaderDocumentOrdering()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            var model = new DocumentModel();
            model.Content.Add(new ParagraphModel
            {
                Runs =
                {
                    new RunModel { Text = "Body" }
                }
            });
            model.FootnoteSeparatorText = "Sep";

            var story = CreateStructuredHeaderFooterNestedCellStory("Lead", "Cell", "Tail");
            model.Sections.Add(new SectionModel
            {
                DefaultHeaderStory = story,
                DefaultHeaderText = story.Text
            });

            var writer = new DocWriter();
            using var ms = new MemoryStream();

            writer.WriteDocBlocks(model, ms);
            ms.Position = 0;

            using var compoundFile = new OpenMcdf.CompoundFile(ms);
            Assert.True(compoundFile.RootStorage.TryGetStream("WordDocument", out var wordDocStream));
            Assert.True(compoundFile.RootStorage.TryGetStream("1Table", out var tableStream));

            var wordDocData = wordDocStream.GetData();
            string expectedText = "Body\rSep\rLead\rInner lead\rCell\r\x0007\rInner tail\r\x0007\rTail\r\r";
            var textBytes = new byte[expectedText.Length];
            System.Array.Copy(wordDocData, 1536, textBytes, 0, expectedText.Length);
            var extractedText = Encoding.GetEncoding(1252).GetString(textBytes);

            Assert.Equal(expectedText, extractedText);

            int lcbPlcfbteTapx = BitConverter.ToInt32(wordDocData, 154 + (Fib.TapxPairIndex * 8) + 4);
            Assert.True(lcbPlcfbteTapx > 0);
            Assert.True(GetTapxRunCount(wordDocData, tableStream.GetData()) >= 3);
        }

        [Fact]
        public void WriteDocBlocks_WithNestedTableInsideFooterCellContent_WritesNestedTableMarkers()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            var model = new DocumentModel();
            model.Content.Add(new ParagraphModel
            {
                Runs =
                {
                    new RunModel { Text = "Body" }
                }
            });

            var story = CreateStructuredHeaderFooterNestedCellStory("Lead", "Cell", "Tail");
            model.Sections.Add(new SectionModel
            {
                DefaultFooterStory = story,
                DefaultFooterText = story.Text
            });

            var writer = new DocWriter();
            using var ms = new MemoryStream();

            writer.WriteDocBlocks(model, ms);
            ms.Position = 0;

            using var compoundFile = new OpenMcdf.CompoundFile(ms);
            Assert.True(compoundFile.RootStorage.TryGetStream("WordDocument", out var wordDocStream));
            Assert.True(compoundFile.RootStorage.TryGetStream("1Table", out var tableStream));

            var wordDocData = wordDocStream.GetData();
            string expectedText = "Body\rLead\rInner lead\rCell\r\x0007\rInner tail\r\x0007\rTail\r\r";
            var textBytes = new byte[expectedText.Length];
            System.Array.Copy(wordDocData, 1536, textBytes, 0, expectedText.Length);
            var extractedText = Encoding.GetEncoding(1252).GetString(textBytes);

            Assert.Equal(expectedText, extractedText);

            int lcbPlcfbteTapx = BitConverter.ToInt32(wordDocData, 154 + (Fib.TapxPairIndex * 8) + 4);
            Assert.True(lcbPlcfbteTapx > 0);
            Assert.True(GetTapxRunCount(wordDocData, tableStream.GetData()) >= 3);
        }

        [Fact]
        public void WriteDocBlocks_WithNoteSeparatorAndNestedTableInsideFooterCellContent_PreservesSharedHeaderDocumentOrdering()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            var model = new DocumentModel();
            model.Content.Add(new ParagraphModel
            {
                Runs =
                {
                    new RunModel { Text = "Body" }
                }
            });
            model.FootnoteSeparatorText = "Sep";

            var story = CreateStructuredHeaderFooterNestedCellStory("Lead", "Cell", "Tail");
            model.Sections.Add(new SectionModel
            {
                DefaultFooterStory = story,
                DefaultFooterText = story.Text
            });

            var writer = new DocWriter();
            using var ms = new MemoryStream();

            writer.WriteDocBlocks(model, ms);
            ms.Position = 0;

            using var compoundFile = new OpenMcdf.CompoundFile(ms);
            Assert.True(compoundFile.RootStorage.TryGetStream("WordDocument", out var wordDocStream));
            Assert.True(compoundFile.RootStorage.TryGetStream("1Table", out var tableStream));

            var wordDocData = wordDocStream.GetData();
            string expectedText = "Body\rSep\rLead\rInner lead\rCell\r\x0007\rInner tail\r\x0007\rTail\r\r";
            var textBytes = new byte[expectedText.Length];
            System.Array.Copy(wordDocData, 1536, textBytes, 0, expectedText.Length);
            var extractedText = Encoding.GetEncoding(1252).GetString(textBytes);

            Assert.Equal(expectedText, extractedText);

            int lcbPlcfbteTapx = BitConverter.ToInt32(wordDocData, 154 + (Fib.TapxPairIndex * 8) + 4);
            Assert.True(lcbPlcfbteTapx > 0);
            Assert.True(GetTapxRunCount(wordDocData, tableStream.GetData()) >= 3);
        }

        [Fact]
        public void WriteDocBlocks_WithNoteSeparatorAndDefaultHeaderFooter_PreservesSharedHeaderDocumentOrdering()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            var model = new DocumentModel();
            model.Content.Add(new ParagraphModel
            {
                Runs =
                {
                    new RunModel { Text = "Body" }
                }
            });

            model.FootnoteSeparatorText = "Sep";
            model.Sections.Add(new SectionModel
            {
                HeaderText = "Head",
                FooterText = "Foot"
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
            string expectedText = "Body\rSep\rHead\r\rFoot\r\r";
            var textBytes = new byte[expectedText.Length];
            System.Array.Copy(wordDocData, 1536, textBytes, 0, expectedText.Length);
            var extractedText = Encoding.GetEncoding(1252).GetString(textBytes);

            Assert.Equal(expectedText, extractedText);
            Assert.Equal(5, BitConverter.ToInt32(wordDocData, 64));
            Assert.Equal(16, BitConverter.ToInt32(wordDocData, 72));

            int fcPlcfHdd = BitConverter.ToInt32(wordDocData, 154 + (Fib.HeaderStoryPairIndex * 8));

            Assert.Equal(0, BitConverter.ToInt32(tableData, fcPlcfHdd));
            Assert.Equal(4, BitConverter.ToInt32(tableData, fcPlcfHdd + 4));
            Assert.Equal(4, BitConverter.ToInt32(tableData, fcPlcfHdd + 8));
            Assert.Equal(4, BitConverter.ToInt32(tableData, fcPlcfHdd + 12));
            Assert.Equal(4, BitConverter.ToInt32(tableData, fcPlcfHdd + 16));
            Assert.Equal(4, BitConverter.ToInt32(tableData, fcPlcfHdd + 20));
            Assert.Equal(4, BitConverter.ToInt32(tableData, fcPlcfHdd + 24));
            Assert.Equal(4, BitConverter.ToInt32(tableData, fcPlcfHdd + 28));
            Assert.Equal(10, BitConverter.ToInt32(tableData, fcPlcfHdd + 32));
            Assert.Equal(10, BitConverter.ToInt32(tableData, fcPlcfHdd + 36));
            Assert.Equal(16, BitConverter.ToInt32(tableData, fcPlcfHdd + 40));
            Assert.Equal(16, BitConverter.ToInt32(tableData, fcPlcfHdd + 44));
            Assert.Equal(15, BitConverter.ToInt32(tableData, fcPlcfHdd + 48));
            Assert.Equal(16, BitConverter.ToInt32(tableData, fcPlcfHdd + 52));
        }

        [Fact]
        public void WriteDocBlocks_WithDifferentFirstPageHeaderFooter_WritesDedicatedStoriesAndTitlePageSprm()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            var model = new DocumentModel();
            model.Content.Add(new ParagraphModel
            {
                Runs =
                {
                    new RunModel { Text = "Body" }
                }
            });

            model.Sections.Add(new SectionModel
            {
                DefaultHeaderText = "Head",
                DefaultFooterText = "Foot",
                FirstPageHeaderText = "FirstHead",
                FirstPageFooterText = "FirstFoot",
                DifferentFirstPage = true
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

            string expectedText = "Body\rHead\r\rFoot\r\rFirstHead\r\rFirstFoot\r\r";
            var textBytes = new byte[expectedText.Length];
            System.Array.Copy(wordDocData, 1536, textBytes, 0, expectedText.Length);
            var extractedText = Encoding.GetEncoding(1252).GetString(textBytes);

            Assert.Equal(expectedText, extractedText);
            Assert.Equal(5, BitConverter.ToInt32(wordDocData, 64));
            Assert.Equal(34, BitConverter.ToInt32(wordDocData, 72));

            int fcPlcfHdd = BitConverter.ToInt32(wordDocData, 154 + (Fib.HeaderStoryPairIndex * 8));
            int lcbPlcfHdd = BitConverter.ToInt32(wordDocData, 154 + (Fib.HeaderStoryPairIndex * 8) + 4);

            Assert.NotEqual(0, fcPlcfHdd);
            Assert.Equal(56, lcbPlcfHdd);

            Assert.Equal(0, BitConverter.ToInt32(tableData, fcPlcfHdd));
            Assert.Equal(0, BitConverter.ToInt32(tableData, fcPlcfHdd + 4));
            Assert.Equal(0, BitConverter.ToInt32(tableData, fcPlcfHdd + 8));
            Assert.Equal(0, BitConverter.ToInt32(tableData, fcPlcfHdd + 12));
            Assert.Equal(0, BitConverter.ToInt32(tableData, fcPlcfHdd + 16));
            Assert.Equal(0, BitConverter.ToInt32(tableData, fcPlcfHdd + 20));
            Assert.Equal(0, BitConverter.ToInt32(tableData, fcPlcfHdd + 24));
            Assert.Equal(0, BitConverter.ToInt32(tableData, fcPlcfHdd + 28));
            Assert.Equal(6, BitConverter.ToInt32(tableData, fcPlcfHdd + 32));
            Assert.Equal(6, BitConverter.ToInt32(tableData, fcPlcfHdd + 36));
            Assert.Equal(12, BitConverter.ToInt32(tableData, fcPlcfHdd + 40));
            Assert.Equal(23, BitConverter.ToInt32(tableData, fcPlcfHdd + 44));
            Assert.Equal(33, BitConverter.ToInt32(tableData, fcPlcfHdd + 48));
            Assert.Equal(34, BitConverter.ToInt32(tableData, fcPlcfHdd + 52));

            int fcPlcfsed = BitConverter.ToInt32(wordDocData, 202);
            int fcSep = BitConverter.ToInt32(tableData, fcPlcfsed + 10);
            short cbSep = BitConverter.ToInt16(wordDocData, fcSep);
            bool foundTitlePageSprm = false;
            for (int offset = fcSep + 2; offset <= fcSep + 2 + cbSep - 3; offset++)
            {
                if (wordDocData[offset] == 0x0A &&
                    wordDocData[offset + 1] == 0x30 &&
                    wordDocData[offset + 2] == 0x01)
                {
                    foundTitlePageSprm = true;
                    break;
                }
            }

            Assert.True(foundTitlePageSprm);
        }

        [Fact]
        public void WriteDocBlocks_WithFirstPageOnlyHeaderFooter_DoesNotBackfillDefaultStorySlots()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            var model = new DocumentModel();
            model.DifferentOddAndEvenPages = true;
            model.Content.Add(new ParagraphModel
            {
                Runs =
                {
                    new RunModel { Text = "Body" }
                }
            });

            model.Sections.Add(new SectionModel
            {
                HeaderText = "FirstHead",
                FooterText = "FirstFoot",
                FirstPageHeaderText = "FirstHead",
                FirstPageFooterText = "FirstFoot",
                DifferentFirstPage = true
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

            string expectedText = "Body\rFirstHead\r\rFirstFoot\r\r";
            var textBytes = new byte[expectedText.Length];
            System.Array.Copy(wordDocData, 1536, textBytes, 0, expectedText.Length);
            var extractedText = Encoding.GetEncoding(1252).GetString(textBytes);

            Assert.Equal(expectedText, extractedText);
            Assert.Equal(5, BitConverter.ToInt32(wordDocData, 64));
            Assert.Equal(22, BitConverter.ToInt32(wordDocData, 72));

            int fcPlcfHdd = BitConverter.ToInt32(wordDocData, 154 + (Fib.HeaderStoryPairIndex * 8));

            Assert.Equal(0, BitConverter.ToInt32(tableData, fcPlcfHdd));
            Assert.Equal(0, BitConverter.ToInt32(tableData, fcPlcfHdd + 4));
            Assert.Equal(0, BitConverter.ToInt32(tableData, fcPlcfHdd + 8));
            Assert.Equal(0, BitConverter.ToInt32(tableData, fcPlcfHdd + 12));
            Assert.Equal(0, BitConverter.ToInt32(tableData, fcPlcfHdd + 16));
            Assert.Equal(0, BitConverter.ToInt32(tableData, fcPlcfHdd + 20));
            Assert.Equal(0, BitConverter.ToInt32(tableData, fcPlcfHdd + 24));
            Assert.Equal(0, BitConverter.ToInt32(tableData, fcPlcfHdd + 28));
            Assert.Equal(0, BitConverter.ToInt32(tableData, fcPlcfHdd + 32));
            Assert.Equal(0, BitConverter.ToInt32(tableData, fcPlcfHdd + 36));
            Assert.Equal(0, BitConverter.ToInt32(tableData, fcPlcfHdd + 40));
            Assert.Equal(11, BitConverter.ToInt32(tableData, fcPlcfHdd + 44));
            Assert.Equal(21, BitConverter.ToInt32(tableData, fcPlcfHdd + 48));
            Assert.Equal(22, BitConverter.ToInt32(tableData, fcPlcfHdd + 52));
        }

        [Fact]
        public void WriteDocBlocks_WithDefaultAndEvenPageHeaderFooter_WritesDedicatedEvenStorySlots()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            var model = new DocumentModel();
            model.DifferentOddAndEvenPages = true;
            model.Content.Add(new ParagraphModel
            {
                Runs =
                {
                    new RunModel { Text = "Body" }
                }
            });

            model.Sections.Add(new SectionModel
            {
                DefaultHeaderText = "Head",
                DefaultFooterText = "Foot",
                EvenPagesHeaderText = "EvenHead",
                EvenPagesFooterText = "EvenFoot"
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

            string expectedText = "Body\rEvenHead\r\rHead\r\rEvenFoot\r\rFoot\r\r";
            var textBytes = new byte[expectedText.Length];
            System.Array.Copy(wordDocData, 1536, textBytes, 0, expectedText.Length);
            var extractedText = Encoding.GetEncoding(1252).GetString(textBytes);

            Assert.Equal(expectedText, extractedText);
            Assert.Equal(5, BitConverter.ToInt32(wordDocData, 64));
            Assert.Equal(32, BitConverter.ToInt32(wordDocData, 72));

            int fcPlcfHdd = BitConverter.ToInt32(wordDocData, 154 + (Fib.HeaderStoryPairIndex * 8));
            int lcbPlcfHdd = BitConverter.ToInt32(wordDocData, 154 + (Fib.HeaderStoryPairIndex * 8) + 4);

            Assert.NotEqual(0, fcPlcfHdd);
            Assert.Equal(56, lcbPlcfHdd);

            Assert.Equal(0, BitConverter.ToInt32(tableData, fcPlcfHdd));
            Assert.Equal(0, BitConverter.ToInt32(tableData, fcPlcfHdd + 4));
            Assert.Equal(0, BitConverter.ToInt32(tableData, fcPlcfHdd + 8));
            Assert.Equal(0, BitConverter.ToInt32(tableData, fcPlcfHdd + 12));
            Assert.Equal(0, BitConverter.ToInt32(tableData, fcPlcfHdd + 16));
            Assert.Equal(0, BitConverter.ToInt32(tableData, fcPlcfHdd + 20));
            Assert.Equal(0, BitConverter.ToInt32(tableData, fcPlcfHdd + 24));
            Assert.Equal(10, BitConverter.ToInt32(tableData, fcPlcfHdd + 28));
            Assert.Equal(16, BitConverter.ToInt32(tableData, fcPlcfHdd + 32));
            Assert.Equal(26, BitConverter.ToInt32(tableData, fcPlcfHdd + 36));
            Assert.Equal(32, BitConverter.ToInt32(tableData, fcPlcfHdd + 40));
            Assert.Equal(32, BitConverter.ToInt32(tableData, fcPlcfHdd + 44));
            Assert.Equal(31, BitConverter.ToInt32(tableData, fcPlcfHdd + 48));
            Assert.Equal(32, BitConverter.ToInt32(tableData, fcPlcfHdd + 52));
        }

        [Fact]
        public void WriteDocBlocks_WithEvenPageOnlyHeaderFooter_DoesNotBackfillDefaultStorySlots()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            var model = new DocumentModel();
            model.DifferentOddAndEvenPages = true;
            model.Content.Add(new ParagraphModel
            {
                Runs =
                {
                    new RunModel { Text = "Body" }
                }
            });

            model.Sections.Add(new SectionModel
            {
                HeaderText = "EvenHead",
                FooterText = "EvenFoot",
                EvenPagesHeaderText = "EvenHead",
                EvenPagesFooterText = "EvenFoot"
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

            string expectedText = "Body\rEvenHead\r\rEvenFoot\r\r";
            var textBytes = new byte[expectedText.Length];
            System.Array.Copy(wordDocData, 1536, textBytes, 0, expectedText.Length);
            var extractedText = Encoding.GetEncoding(1252).GetString(textBytes);

            Assert.Equal(expectedText, extractedText);
            Assert.Equal(5, BitConverter.ToInt32(wordDocData, 64));
            Assert.Equal(20, BitConverter.ToInt32(wordDocData, 72));

            int fcPlcfHdd = BitConverter.ToInt32(wordDocData, 154 + (Fib.HeaderStoryPairIndex * 8));

            Assert.Equal(0, BitConverter.ToInt32(tableData, fcPlcfHdd));
            Assert.Equal(0, BitConverter.ToInt32(tableData, fcPlcfHdd + 4));
            Assert.Equal(0, BitConverter.ToInt32(tableData, fcPlcfHdd + 8));
            Assert.Equal(0, BitConverter.ToInt32(tableData, fcPlcfHdd + 12));
            Assert.Equal(0, BitConverter.ToInt32(tableData, fcPlcfHdd + 16));
            Assert.Equal(0, BitConverter.ToInt32(tableData, fcPlcfHdd + 20));
            Assert.Equal(0, BitConverter.ToInt32(tableData, fcPlcfHdd + 24));
            Assert.Equal(10, BitConverter.ToInt32(tableData, fcPlcfHdd + 28));
            Assert.Equal(10, BitConverter.ToInt32(tableData, fcPlcfHdd + 32));
            Assert.Equal(20, BitConverter.ToInt32(tableData, fcPlcfHdd + 36));
            Assert.Equal(20, BitConverter.ToInt32(tableData, fcPlcfHdd + 40));
            Assert.Equal(20, BitConverter.ToInt32(tableData, fcPlcfHdd + 44));
            Assert.Equal(19, BitConverter.ToInt32(tableData, fcPlcfHdd + 48));
            Assert.Equal(20, BitConverter.ToInt32(tableData, fcPlcfHdd + 52));
        }

        [Fact]
        public void WriteDocBlocks_WithNoteSeparatorAndDefaultAndEvenHeaderFooter_PreservesSharedHeaderDocumentOrdering()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            var model = new DocumentModel();
            model.DifferentOddAndEvenPages = true;
            model.Content.Add(new ParagraphModel
            {
                Runs =
                {
                    new RunModel { Text = "Body" }
                }
            });

            model.FootnoteSeparatorText = "Sep";
            model.Sections.Add(new SectionModel
            {
                DefaultHeaderText = "Head",
                DefaultFooterText = "Foot",
                EvenPagesHeaderText = "EvenHead",
                EvenPagesFooterText = "EvenFoot"
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

            string expectedText = "Body\rSep\rEvenHead\r\rHead\r\rEvenFoot\r\rFoot\r\r";
            var textBytes = new byte[expectedText.Length];
            System.Array.Copy(wordDocData, 1536, textBytes, 0, expectedText.Length);
            var extractedText = Encoding.GetEncoding(1252).GetString(textBytes);

            Assert.Equal(expectedText, extractedText);
            Assert.Equal(5, BitConverter.ToInt32(wordDocData, 64));
            Assert.Equal(36, BitConverter.ToInt32(wordDocData, 72));

            int fcPlcfHdd = BitConverter.ToInt32(wordDocData, 154 + (Fib.HeaderStoryPairIndex * 8));

            Assert.Equal(0, BitConverter.ToInt32(tableData, fcPlcfHdd));
            Assert.Equal(4, BitConverter.ToInt32(tableData, fcPlcfHdd + 4));
            Assert.Equal(4, BitConverter.ToInt32(tableData, fcPlcfHdd + 8));
            Assert.Equal(4, BitConverter.ToInt32(tableData, fcPlcfHdd + 12));
            Assert.Equal(4, BitConverter.ToInt32(tableData, fcPlcfHdd + 16));
            Assert.Equal(4, BitConverter.ToInt32(tableData, fcPlcfHdd + 20));
            Assert.Equal(4, BitConverter.ToInt32(tableData, fcPlcfHdd + 24));
            Assert.Equal(14, BitConverter.ToInt32(tableData, fcPlcfHdd + 28));
            Assert.Equal(20, BitConverter.ToInt32(tableData, fcPlcfHdd + 32));
            Assert.Equal(30, BitConverter.ToInt32(tableData, fcPlcfHdd + 36));
            Assert.Equal(36, BitConverter.ToInt32(tableData, fcPlcfHdd + 40));
            Assert.Equal(36, BitConverter.ToInt32(tableData, fcPlcfHdd + 44));
            Assert.Equal(35, BitConverter.ToInt32(tableData, fcPlcfHdd + 48));
            Assert.Equal(36, BitConverter.ToInt32(tableData, fcPlcfHdd + 52));
        }

        [Fact]
        public void WriteDocBlocks_WithoutDifferentOddAndEvenPages_IgnoresEvenPageStories()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            var model = new DocumentModel();
            model.Content.Add(new ParagraphModel
            {
                Runs =
                {
                    new RunModel { Text = "Body" }
                }
            });

            model.Sections.Add(new SectionModel
            {
                DefaultHeaderText = "Head",
                DefaultFooterText = "Foot",
                EvenPagesHeaderText = "EvenHead",
                EvenPagesFooterText = "EvenFoot"
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

            string expectedText = "Body\rHead\r\rFoot\r\r";
            var textBytes = new byte[expectedText.Length];
            System.Array.Copy(wordDocData, 1536, textBytes, 0, expectedText.Length);
            var extractedText = Encoding.GetEncoding(1252).GetString(textBytes);

            Assert.Equal(expectedText, extractedText);
            Assert.Equal(5, BitConverter.ToInt32(wordDocData, 64));
            Assert.Equal(12, BitConverter.ToInt32(wordDocData, 72));

            int fcPlcfHdd = BitConverter.ToInt32(wordDocData, 154 + (Fib.HeaderStoryPairIndex * 8));

            Assert.Equal(0, BitConverter.ToInt32(tableData, fcPlcfHdd));
            Assert.Equal(0, BitConverter.ToInt32(tableData, fcPlcfHdd + 4));
            Assert.Equal(0, BitConverter.ToInt32(tableData, fcPlcfHdd + 8));
            Assert.Equal(0, BitConverter.ToInt32(tableData, fcPlcfHdd + 12));
            Assert.Equal(0, BitConverter.ToInt32(tableData, fcPlcfHdd + 16));
            Assert.Equal(0, BitConverter.ToInt32(tableData, fcPlcfHdd + 20));
            Assert.Equal(0, BitConverter.ToInt32(tableData, fcPlcfHdd + 24));
            Assert.Equal(0, BitConverter.ToInt32(tableData, fcPlcfHdd + 28));
            Assert.Equal(6, BitConverter.ToInt32(tableData, fcPlcfHdd + 32));
            Assert.Equal(6, BitConverter.ToInt32(tableData, fcPlcfHdd + 36));
            Assert.Equal(12, BitConverter.ToInt32(tableData, fcPlcfHdd + 40));
            Assert.Equal(12, BitConverter.ToInt32(tableData, fcPlcfHdd + 44));
            Assert.Equal(11, BitConverter.ToInt32(tableData, fcPlcfHdd + 48));
            Assert.Equal(12, BitConverter.ToInt32(tableData, fcPlcfHdd + 52));
        }

        [Fact]
        public void WriteDocBlocks_WithEnabledEmptyEvenPageStories_ReservesDedicatedEvenStorySlots()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            var model = new DocumentModel();
            model.DifferentOddAndEvenPages = true;
            model.Content.Add(new ParagraphModel
            {
                Runs =
                {
                    new RunModel { Text = "Body" }
                }
            });

            model.Sections.Add(new SectionModel
            {
                DefaultHeaderText = "Head",
                DefaultFooterText = "Foot",
                EvenPagesHeaderText = string.Empty,
                EvenPagesFooterText = string.Empty
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

            string expectedText = "Body\r\r\rHead\r\r\r\rFoot\r\r";
            var textBytes = new byte[expectedText.Length];
            System.Array.Copy(wordDocData, 1536, textBytes, 0, expectedText.Length);
            var extractedText = Encoding.GetEncoding(1252).GetString(textBytes);

            Assert.Equal(expectedText, extractedText);
            Assert.Equal(5, BitConverter.ToInt32(wordDocData, 64));
            Assert.Equal(16, BitConverter.ToInt32(wordDocData, 72));

            int fcPlcfHdd = BitConverter.ToInt32(wordDocData, 154 + (Fib.HeaderStoryPairIndex * 8));

            Assert.Equal(0, BitConverter.ToInt32(tableData, fcPlcfHdd));
            Assert.Equal(0, BitConverter.ToInt32(tableData, fcPlcfHdd + 4));
            Assert.Equal(0, BitConverter.ToInt32(tableData, fcPlcfHdd + 8));
            Assert.Equal(0, BitConverter.ToInt32(tableData, fcPlcfHdd + 12));
            Assert.Equal(0, BitConverter.ToInt32(tableData, fcPlcfHdd + 16));
            Assert.Equal(0, BitConverter.ToInt32(tableData, fcPlcfHdd + 20));
            Assert.Equal(0, BitConverter.ToInt32(tableData, fcPlcfHdd + 24));
            Assert.Equal(2, BitConverter.ToInt32(tableData, fcPlcfHdd + 28));
            Assert.Equal(8, BitConverter.ToInt32(tableData, fcPlcfHdd + 32));
            Assert.Equal(10, BitConverter.ToInt32(tableData, fcPlcfHdd + 36));
            Assert.Equal(16, BitConverter.ToInt32(tableData, fcPlcfHdd + 40));
            Assert.Equal(16, BitConverter.ToInt32(tableData, fcPlcfHdd + 44));
            Assert.Equal(15, BitConverter.ToInt32(tableData, fcPlcfHdd + 48));
            Assert.Equal(16, BitConverter.ToInt32(tableData, fcPlcfHdd + 52));
        }

        [Fact]
        public void WriteDocBlocks_WithStructuredHeaderFieldStory_WritesHeaderFieldTable()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            var headerStory = CreateStructuredHeaderFieldStory();

            var model = new DocumentModel();
            model.Content.Add(new ParagraphModel
            {
                Runs =
                {
                    new RunModel { Text = "Body" }
                }
            });

            model.Sections.Add(new SectionModel
            {
                DefaultHeaderStory = headerStory,
                DefaultHeaderText = "Page 1 of 2"
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

            string expectedText = "Body\rPage \x0013PAGE\x00141\x0015 of \x0013SECTIONPAGES\x00142\x0015\r\r";
            var textBytes = new byte[expectedText.Length];
            System.Array.Copy(wordDocData, 1536, textBytes, 0, expectedText.Length);
            var extractedText = Encoding.GetEncoding(1252).GetString(textBytes);

            Assert.Equal(expectedText, extractedText);
            Assert.Equal(5, BitConverter.ToInt32(wordDocData, 64));
            Assert.Equal(19, BitConverter.ToInt32(wordDocData, 72));

            int fcPlcfHdd = BitConverter.ToInt32(wordDocData, 154 + (Fib.HeaderStoryPairIndex * 8));
            Assert.NotEqual(0, fcPlcfHdd);
            Assert.Equal(0, BitConverter.ToInt32(tableData, fcPlcfHdd + 24));
            Assert.Equal(0, BitConverter.ToInt32(tableData, fcPlcfHdd + 28));
            Assert.Equal(19, BitConverter.ToInt32(tableData, fcPlcfHdd + 32));
            Assert.Equal(19, BitConverter.ToInt32(tableData, fcPlcfHdd + 36));
            Assert.Equal(19, BitConverter.ToInt32(tableData, fcPlcfHdd + 40));
            Assert.Equal(19, BitConverter.ToInt32(tableData, fcPlcfHdd + 44));
            Assert.Equal(18, BitConverter.ToInt32(tableData, fcPlcfHdd + 48));
            Assert.Equal(19, BitConverter.ToInt32(tableData, fcPlcfHdd + 52));

            int fcPlcffldHdr = BitConverter.ToInt32(wordDocData, 154 + (Fib.HeaderFieldPairIndex * 8));
            int lcbPlcffldHdr = BitConverter.ToInt32(wordDocData, 154 + (Fib.HeaderFieldPairIndex * 8) + 4);

            Assert.NotEqual(0, fcPlcffldHdr);
            Assert.True(lcbPlcffldHdr > 0);
            Assert.Equal(5, BitConverter.ToInt32(tableData, fcPlcffldHdr));
            Assert.Equal(6, BitConverter.ToInt32(tableData, fcPlcffldHdr + 4));
            Assert.Equal(8, BitConverter.ToInt32(tableData, fcPlcffldHdr + 8));
            Assert.Equal(13, BitConverter.ToInt32(tableData, fcPlcffldHdr + 12));
            Assert.Equal(14, BitConverter.ToInt32(tableData, fcPlcffldHdr + 16));
            Assert.Equal(16, BitConverter.ToInt32(tableData, fcPlcffldHdr + 20));
            Assert.Equal(19, BitConverter.ToInt32(tableData, fcPlcffldHdr + 24));
            Assert.Equal(0x0813, BitConverter.ToUInt16(tableData, fcPlcffldHdr + 28));
            Assert.Equal(0x0814, BitConverter.ToUInt16(tableData, fcPlcffldHdr + 30));
            Assert.Equal(0x0815, BitConverter.ToUInt16(tableData, fcPlcffldHdr + 32));
            Assert.Equal(0x0813, BitConverter.ToUInt16(tableData, fcPlcffldHdr + 34));
            Assert.Equal(0x0814, BitConverter.ToUInt16(tableData, fcPlcffldHdr + 36));
            Assert.Equal(0x0815, BitConverter.ToUInt16(tableData, fcPlcffldHdr + 38));
        }

        [Fact]
        public void WriteDocBlocks_WithStructuredEvenPageOnlyHeaderFooter_DoesNotBackfillDefaultStorySlots()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            var model = new DocumentModel();
            model.DifferentOddAndEvenPages = true;
            model.Content.Add(new ParagraphModel
            {
                Runs =
                {
                    new RunModel { Text = "Body" }
                }
            });

            model.Sections.Add(new SectionModel
            {
                HeaderText = "SummaryHead",
                FooterText = "SummaryFoot",
                EvenPagesHeaderStory = CreateTextHeaderFooterStory("EvenHead"),
                EvenPagesFooterStory = CreateTextHeaderFooterStory("EvenFoot")
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

            string expectedText = "Body\rEvenHead\r\rEvenFoot\r\r";
            var textBytes = new byte[expectedText.Length];
            System.Array.Copy(wordDocData, 1536, textBytes, 0, expectedText.Length);
            var extractedText = Encoding.GetEncoding(1252).GetString(textBytes);

            Assert.Equal(expectedText, extractedText);
            Assert.DoesNotContain("SummaryHead", extractedText);
            Assert.DoesNotContain("SummaryFoot", extractedText);
            Assert.Equal(5, BitConverter.ToInt32(wordDocData, 64));
            Assert.Equal(20, BitConverter.ToInt32(wordDocData, 72));

            int fcPlcfHdd = BitConverter.ToInt32(wordDocData, 154 + (Fib.HeaderStoryPairIndex * 8));

            Assert.Equal(0, BitConverter.ToInt32(tableData, fcPlcfHdd));
            Assert.Equal(0, BitConverter.ToInt32(tableData, fcPlcfHdd + 4));
            Assert.Equal(0, BitConverter.ToInt32(tableData, fcPlcfHdd + 8));
            Assert.Equal(0, BitConverter.ToInt32(tableData, fcPlcfHdd + 12));
            Assert.Equal(0, BitConverter.ToInt32(tableData, fcPlcfHdd + 16));
            Assert.Equal(0, BitConverter.ToInt32(tableData, fcPlcfHdd + 20));
            Assert.Equal(0, BitConverter.ToInt32(tableData, fcPlcfHdd + 24));
            Assert.Equal(10, BitConverter.ToInt32(tableData, fcPlcfHdd + 28));
            Assert.Equal(10, BitConverter.ToInt32(tableData, fcPlcfHdd + 32));
            Assert.Equal(20, BitConverter.ToInt32(tableData, fcPlcfHdd + 36));
            Assert.Equal(20, BitConverter.ToInt32(tableData, fcPlcfHdd + 40));
            Assert.Equal(20, BitConverter.ToInt32(tableData, fcPlcfHdd + 44));
            Assert.Equal(19, BitConverter.ToInt32(tableData, fcPlcfHdd + 48));
            Assert.Equal(20, BitConverter.ToInt32(tableData, fcPlcfHdd + 52));
        }

        [Fact]
        public void WriteDocBlocks_WithStructuredFirstPageOnlyHeaderFooter_DoesNotBackfillDefaultStorySlots()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            var model = new DocumentModel();
            model.Content.Add(new ParagraphModel
            {
                Runs =
                {
                    new RunModel { Text = "Body" }
                }
            });

            model.Sections.Add(new SectionModel
            {
                HeaderText = "SummaryHead",
                FooterText = "SummaryFoot",
                FirstPageHeaderStory = CreateTextHeaderFooterStory("FirstHead"),
                FirstPageFooterStory = CreateTextHeaderFooterStory("FirstFoot"),
                DifferentFirstPage = true
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

            string expectedText = "Body\rFirstHead\r\rFirstFoot\r\r";
            var textBytes = new byte[expectedText.Length];
            System.Array.Copy(wordDocData, 1536, textBytes, 0, expectedText.Length);
            var extractedText = Encoding.GetEncoding(1252).GetString(textBytes);

            Assert.Equal(expectedText, extractedText);
            Assert.DoesNotContain("SummaryHead", extractedText);
            Assert.DoesNotContain("SummaryFoot", extractedText);
            Assert.Equal(5, BitConverter.ToInt32(wordDocData, 64));
            Assert.Equal(22, BitConverter.ToInt32(wordDocData, 72));

            int fcPlcfHdd = BitConverter.ToInt32(wordDocData, 154 + (Fib.HeaderStoryPairIndex * 8));

            Assert.Equal(0, BitConverter.ToInt32(tableData, fcPlcfHdd));
            Assert.Equal(0, BitConverter.ToInt32(tableData, fcPlcfHdd + 4));
            Assert.Equal(0, BitConverter.ToInt32(tableData, fcPlcfHdd + 8));
            Assert.Equal(0, BitConverter.ToInt32(tableData, fcPlcfHdd + 12));
            Assert.Equal(0, BitConverter.ToInt32(tableData, fcPlcfHdd + 16));
            Assert.Equal(0, BitConverter.ToInt32(tableData, fcPlcfHdd + 20));
            Assert.Equal(0, BitConverter.ToInt32(tableData, fcPlcfHdd + 24));
            Assert.Equal(0, BitConverter.ToInt32(tableData, fcPlcfHdd + 28));
            Assert.Equal(0, BitConverter.ToInt32(tableData, fcPlcfHdd + 32));
            Assert.Equal(0, BitConverter.ToInt32(tableData, fcPlcfHdd + 36));
            Assert.Equal(0, BitConverter.ToInt32(tableData, fcPlcfHdd + 40));
            Assert.Equal(11, BitConverter.ToInt32(tableData, fcPlcfHdd + 44));
            Assert.Equal(21, BitConverter.ToInt32(tableData, fcPlcfHdd + 48));
            Assert.Equal(22, BitConverter.ToInt32(tableData, fcPlcfHdd + 52));
        }

        [Fact]
        public void WriteDocBlocks_WithNoteSeparatorAndStructuredHeaderFieldStory_PreservesSharedHeaderDocumentOrdering()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            var model = new DocumentModel();
            model.Content.Add(new ParagraphModel
            {
                Runs =
                {
                    new RunModel { Text = "Body" }
                }
            });
            model.FootnoteSeparatorText = "Sep";

            model.Sections.Add(new SectionModel
            {
                DefaultHeaderStory = CreateStructuredHeaderFieldStory(),
                DefaultHeaderText = "Page 1 of 2"
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

            string expectedText = "Body\rSep\rPage \x0013PAGE\x00141\x0015 of \x0013SECTIONPAGES\x00142\x0015\r\r";
            var textBytes = new byte[expectedText.Length];
            System.Array.Copy(wordDocData, 1536, textBytes, 0, expectedText.Length);
            var extractedText = Encoding.GetEncoding(1252).GetString(textBytes);

            Assert.Equal(expectedText, extractedText);
            Assert.Equal(5, BitConverter.ToInt32(wordDocData, 64));
            Assert.Equal(23, BitConverter.ToInt32(wordDocData, 72));

            int fcPlcfHdd = BitConverter.ToInt32(wordDocData, 154 + (Fib.HeaderStoryPairIndex * 8));
            Assert.Equal(0, BitConverter.ToInt32(tableData, fcPlcfHdd));
            Assert.Equal(4, BitConverter.ToInt32(tableData, fcPlcfHdd + 4));
            Assert.Equal(4, BitConverter.ToInt32(tableData, fcPlcfHdd + 8));
            Assert.Equal(4, BitConverter.ToInt32(tableData, fcPlcfHdd + 12));
            Assert.Equal(4, BitConverter.ToInt32(tableData, fcPlcfHdd + 16));
            Assert.Equal(4, BitConverter.ToInt32(tableData, fcPlcfHdd + 20));
            Assert.Equal(4, BitConverter.ToInt32(tableData, fcPlcfHdd + 24));
            Assert.Equal(4, BitConverter.ToInt32(tableData, fcPlcfHdd + 28));
            Assert.Equal(23, BitConverter.ToInt32(tableData, fcPlcfHdd + 32));
            Assert.Equal(23, BitConverter.ToInt32(tableData, fcPlcfHdd + 36));
            Assert.Equal(23, BitConverter.ToInt32(tableData, fcPlcfHdd + 40));
            Assert.Equal(23, BitConverter.ToInt32(tableData, fcPlcfHdd + 44));
            Assert.Equal(22, BitConverter.ToInt32(tableData, fcPlcfHdd + 48));
            Assert.Equal(23, BitConverter.ToInt32(tableData, fcPlcfHdd + 52));

            int fcPlcffldHdr = BitConverter.ToInt32(wordDocData, 154 + (Fib.HeaderFieldPairIndex * 8));
            int lcbPlcffldHdr = BitConverter.ToInt32(wordDocData, 154 + (Fib.HeaderFieldPairIndex * 8) + 4);

            Assert.NotEqual(0, fcPlcffldHdr);
            Assert.True(lcbPlcffldHdr > 0);
            Assert.Equal(9, BitConverter.ToInt32(tableData, fcPlcffldHdr));
            Assert.Equal(10, BitConverter.ToInt32(tableData, fcPlcffldHdr + 4));
            Assert.Equal(12, BitConverter.ToInt32(tableData, fcPlcffldHdr + 8));
            Assert.Equal(17, BitConverter.ToInt32(tableData, fcPlcffldHdr + 12));
            Assert.Equal(18, BitConverter.ToInt32(tableData, fcPlcffldHdr + 16));
            Assert.Equal(20, BitConverter.ToInt32(tableData, fcPlcffldHdr + 20));
            Assert.Equal(23, BitConverter.ToInt32(tableData, fcPlcffldHdr + 24));
        }

        [Fact]
        public void WriteDocBlocks_WithStructuredHeaderHyperlinkStory_WritesHeaderHyperlinkField()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            var model = new DocumentModel();
            model.Content.Add(new ParagraphModel
            {
                Runs =
                {
                    new RunModel { Text = "Body" }
                }
            });

            model.Sections.Add(new SectionModel
            {
                DefaultHeaderStory = CreateStructuredHeaderHyperlinkStory(),
                DefaultHeaderText = "Go Example now"
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

            string expectedText = "Body\rGo \x0013HYPERLINK \"https://example.com\"\x0014Example\x0015 now\r\r";
            var textBytes = new byte[expectedText.Length];
            System.Array.Copy(wordDocData, 1536, textBytes, 0, expectedText.Length);
            var extractedText = Encoding.GetEncoding(1252).GetString(textBytes);

            Assert.Equal(expectedText, extractedText);
            Assert.Equal(5, BitConverter.ToInt32(wordDocData, 64));
            Assert.Equal(19, BitConverter.ToInt32(wordDocData, 72));

            int fcPlcfHdd = BitConverter.ToInt32(wordDocData, 154 + (Fib.HeaderStoryPairIndex * 8));
            Assert.NotEqual(0, fcPlcfHdd);
            Assert.Equal(0, BitConverter.ToInt32(tableData, fcPlcfHdd + 24));
            Assert.Equal(0, BitConverter.ToInt32(tableData, fcPlcfHdd + 28));
            Assert.Equal(19, BitConverter.ToInt32(tableData, fcPlcfHdd + 32));
            Assert.Equal(19, BitConverter.ToInt32(tableData, fcPlcfHdd + 36));
            Assert.Equal(19, BitConverter.ToInt32(tableData, fcPlcfHdd + 40));
            Assert.Equal(19, BitConverter.ToInt32(tableData, fcPlcfHdd + 44));
            Assert.Equal(18, BitConverter.ToInt32(tableData, fcPlcfHdd + 48));
            Assert.Equal(19, BitConverter.ToInt32(tableData, fcPlcfHdd + 52));

            int fcPlcffldHdr = BitConverter.ToInt32(wordDocData, 154 + (Fib.HeaderFieldPairIndex * 8));
            int lcbPlcffldHdr = BitConverter.ToInt32(wordDocData, 154 + (Fib.HeaderFieldPairIndex * 8) + 4);

            Assert.NotEqual(0, fcPlcffldHdr);
            Assert.True(lcbPlcffldHdr > 0);
            Assert.Equal(3, BitConverter.ToInt32(tableData, fcPlcffldHdr));
            Assert.Equal(4, BitConverter.ToInt32(tableData, fcPlcffldHdr + 4));
            Assert.Equal(12, BitConverter.ToInt32(tableData, fcPlcffldHdr + 8));
            Assert.Equal(19, BitConverter.ToInt32(tableData, fcPlcffldHdr + 12));
            Assert.Equal(0x0013, BitConverter.ToUInt16(tableData, fcPlcffldHdr + 16));
            Assert.Equal(0x0014, BitConverter.ToUInt16(tableData, fcPlcffldHdr + 18));
            Assert.Equal(0x0015, BitConverter.ToUInt16(tableData, fcPlcffldHdr + 20));
        }

        [Fact]
        public void WriteDocBlocks_WithMixedHeaderFieldAndHyperlinkStoryAndNoteSeparator_PreservesOrderingAndHeaderFields()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            var model = new DocumentModel();
            model.Content.Add(new ParagraphModel
            {
                Runs =
                {
                    new RunModel { Text = "Body" }
                }
            });
            model.FootnoteSeparatorText = "Sep";

            model.Sections.Add(new SectionModel
            {
                DefaultHeaderStory = CreateStructuredHeaderFieldAndHyperlinkStory(),
                DefaultHeaderText = "Page 1 | Example"
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

            string expectedText = "Body\rSep\rPage \x0013PAGE\x00141\x0015 | \x0013HYPERLINK \"https://example.com\"\x0014Example\x0015\r\r";
            var textBytes = new byte[expectedText.Length];
            System.Array.Copy(wordDocData, 1536, textBytes, 0, expectedText.Length);
            var extractedText = Encoding.GetEncoding(1252).GetString(textBytes);

            Assert.Equal(expectedText, extractedText);
            Assert.Equal(5, BitConverter.ToInt32(wordDocData, 64));
            Assert.Equal(28, BitConverter.ToInt32(wordDocData, 72));

            int fcPlcfHdd = BitConverter.ToInt32(wordDocData, 154 + (Fib.HeaderStoryPairIndex * 8));
            Assert.Equal(0, BitConverter.ToInt32(tableData, fcPlcfHdd));
            Assert.Equal(4, BitConverter.ToInt32(tableData, fcPlcfHdd + 4));
            Assert.Equal(4, BitConverter.ToInt32(tableData, fcPlcfHdd + 8));
            Assert.Equal(4, BitConverter.ToInt32(tableData, fcPlcfHdd + 12));
            Assert.Equal(4, BitConverter.ToInt32(tableData, fcPlcfHdd + 16));
            Assert.Equal(4, BitConverter.ToInt32(tableData, fcPlcfHdd + 20));
            Assert.Equal(4, BitConverter.ToInt32(tableData, fcPlcfHdd + 24));
            Assert.Equal(4, BitConverter.ToInt32(tableData, fcPlcfHdd + 28));
            Assert.Equal(28, BitConverter.ToInt32(tableData, fcPlcfHdd + 32));
            Assert.Equal(28, BitConverter.ToInt32(tableData, fcPlcfHdd + 36));
            Assert.Equal(28, BitConverter.ToInt32(tableData, fcPlcfHdd + 40));
            Assert.Equal(28, BitConverter.ToInt32(tableData, fcPlcfHdd + 44));
            Assert.Equal(27, BitConverter.ToInt32(tableData, fcPlcfHdd + 48));
            Assert.Equal(28, BitConverter.ToInt32(tableData, fcPlcfHdd + 52));

            int fcPlcffldHdr = BitConverter.ToInt32(wordDocData, 154 + (Fib.HeaderFieldPairIndex * 8));
            int lcbPlcffldHdr = BitConverter.ToInt32(wordDocData, 154 + (Fib.HeaderFieldPairIndex * 8) + 4);

            Assert.NotEqual(0, fcPlcffldHdr);
            Assert.True(lcbPlcffldHdr > 0);
            Assert.Equal(9, BitConverter.ToInt32(tableData, fcPlcffldHdr));
            Assert.Equal(10, BitConverter.ToInt32(tableData, fcPlcffldHdr + 4));
            Assert.Equal(12, BitConverter.ToInt32(tableData, fcPlcffldHdr + 8));
            Assert.Equal(16, BitConverter.ToInt32(tableData, fcPlcffldHdr + 12));
            Assert.Equal(17, BitConverter.ToInt32(tableData, fcPlcffldHdr + 16));
            Assert.Equal(25, BitConverter.ToInt32(tableData, fcPlcffldHdr + 20));
            Assert.Equal(28, BitConverter.ToInt32(tableData, fcPlcffldHdr + 24));
            Assert.Equal(0x0813, BitConverter.ToUInt16(tableData, fcPlcffldHdr + 28));
            Assert.Equal(0x0814, BitConverter.ToUInt16(tableData, fcPlcffldHdr + 30));
            Assert.Equal(0x0815, BitConverter.ToUInt16(tableData, fcPlcffldHdr + 32));
            Assert.Equal(0x0013, BitConverter.ToUInt16(tableData, fcPlcffldHdr + 34));
            Assert.Equal(0x0014, BitConverter.ToUInt16(tableData, fcPlcffldHdr + 36));
            Assert.Equal(0x0015, BitConverter.ToUInt16(tableData, fcPlcffldHdr + 38));
        }

        [Fact]
        public void WriteDocBlocks_WithStructuredHeaderInlineImageStory_WritesHeaderImagePlaceholderAndData()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            var model = new DocumentModel();
            model.Content.Add(new ParagraphModel
            {
                Runs =
                {
                    new RunModel { Text = "Body" }
                }
            });

            model.Sections.Add(new SectionModel
            {
                DefaultHeaderStory = CreateStructuredHeaderImageStory(),
                DefaultHeaderText = "HiThere"
            });

            var writer = new DocWriter();
            using var ms = new MemoryStream();

            writer.WriteDocBlocks(model, ms);
            ms.Position = 0;

            using var compoundFile = new OpenMcdf.CompoundFile(ms);
            Assert.True(compoundFile.RootStorage.TryGetStream("WordDocument", out var wordDocStream));
            Assert.True(compoundFile.RootStorage.TryGetStream("1Table", out var tableStream));
            Assert.True(compoundFile.RootStorage.TryGetStream("Data", out var dataStream));

            var wordDocData = wordDocStream.GetData();
            var tableData = tableStream.GetData();
            var data = dataStream.GetData();

            string expectedText = "Body\rHi\x0001There\r\r";
            var textBytes = new byte[expectedText.Length];
            System.Array.Copy(wordDocData, 1536, textBytes, 0, expectedText.Length);
            var extractedText = Encoding.GetEncoding(1252).GetString(textBytes);

            Assert.Equal(expectedText, extractedText);
            Assert.Equal(5, BitConverter.ToInt32(wordDocData, 64));
            Assert.Equal(10, BitConverter.ToInt32(wordDocData, 72));
            Assert.True(data.Length > 0);

            int fcPlcfHdd = BitConverter.ToInt32(wordDocData, 154 + (Fib.HeaderStoryPairIndex * 8));
            Assert.NotEqual(0, fcPlcfHdd);
            Assert.Equal(0, BitConverter.ToInt32(tableData, fcPlcfHdd + 24));
            Assert.Equal(0, BitConverter.ToInt32(tableData, fcPlcfHdd + 28));
            Assert.Equal(10, BitConverter.ToInt32(tableData, fcPlcfHdd + 32));
            Assert.Equal(10, BitConverter.ToInt32(tableData, fcPlcfHdd + 36));
            Assert.Equal(10, BitConverter.ToInt32(tableData, fcPlcfHdd + 40));
            Assert.Equal(10, BitConverter.ToInt32(tableData, fcPlcfHdd + 44));
            Assert.Equal(9, BitConverter.ToInt32(tableData, fcPlcfHdd + 48));
            Assert.Equal(10, BitConverter.ToInt32(tableData, fcPlcfHdd + 52));

            int fcPlcSpaHdr = BitConverter.ToInt32(wordDocData, 154 + (Fib.HeaderShapePairIndex * 8));
            int lcbPlcSpaHdr = BitConverter.ToInt32(wordDocData, 154 + (Fib.HeaderShapePairIndex * 8) + 4);

            Assert.Equal(0, fcPlcSpaHdr);
            Assert.Equal(0, lcbPlcSpaHdr);
        }

        [Fact]
        public void WriteDocBlocks_WithHeaderFieldAndInlineImageStoryAndNoteSeparator_PreservesOrderingAndData()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            var model = new DocumentModel();
            model.Content.Add(new ParagraphModel
            {
                Runs =
                {
                    new RunModel { Text = "Body" }
                }
            });
            model.FootnoteSeparatorText = "Sep";

            model.Sections.Add(new SectionModel
            {
                DefaultHeaderStory = CreateStructuredHeaderFieldAndImageStory(),
                DefaultHeaderText = "Page 1 tail"
            });

            var writer = new DocWriter();
            using var ms = new MemoryStream();

            writer.WriteDocBlocks(model, ms);
            ms.Position = 0;

            using var compoundFile = new OpenMcdf.CompoundFile(ms);
            Assert.True(compoundFile.RootStorage.TryGetStream("WordDocument", out var wordDocStream));
            Assert.True(compoundFile.RootStorage.TryGetStream("1Table", out var tableStream));
            Assert.True(compoundFile.RootStorage.TryGetStream("Data", out var dataStream));

            var wordDocData = wordDocStream.GetData();
            var tableData = tableStream.GetData();
            var data = dataStream.GetData();

            string expectedText = "Body\rSep\rPage \x0013PAGE\x00141\x0015 \x0001 tail\r\r";
            var textBytes = new byte[expectedText.Length];
            System.Array.Copy(wordDocData, 1536, textBytes, 0, expectedText.Length);
            var extractedText = Encoding.GetEncoding(1252).GetString(textBytes);

            Assert.Equal(expectedText, extractedText);
            Assert.Equal(5, BitConverter.ToInt32(wordDocData, 64));
            Assert.Equal(22, BitConverter.ToInt32(wordDocData, 72));
            Assert.True(data.Length > 0);

            int fcPlcfHdd = BitConverter.ToInt32(wordDocData, 154 + (Fib.HeaderStoryPairIndex * 8));
            Assert.Equal(0, BitConverter.ToInt32(tableData, fcPlcfHdd));
            Assert.Equal(4, BitConverter.ToInt32(tableData, fcPlcfHdd + 4));
            Assert.Equal(4, BitConverter.ToInt32(tableData, fcPlcfHdd + 8));
            Assert.Equal(4, BitConverter.ToInt32(tableData, fcPlcfHdd + 12));
            Assert.Equal(4, BitConverter.ToInt32(tableData, fcPlcfHdd + 16));
            Assert.Equal(4, BitConverter.ToInt32(tableData, fcPlcfHdd + 20));
            Assert.Equal(4, BitConverter.ToInt32(tableData, fcPlcfHdd + 24));
            Assert.Equal(4, BitConverter.ToInt32(tableData, fcPlcfHdd + 28));
            Assert.Equal(22, BitConverter.ToInt32(tableData, fcPlcfHdd + 32));
            Assert.Equal(22, BitConverter.ToInt32(tableData, fcPlcfHdd + 36));
            Assert.Equal(22, BitConverter.ToInt32(tableData, fcPlcfHdd + 40));
            Assert.Equal(22, BitConverter.ToInt32(tableData, fcPlcfHdd + 44));
            Assert.Equal(21, BitConverter.ToInt32(tableData, fcPlcfHdd + 48));
            Assert.Equal(22, BitConverter.ToInt32(tableData, fcPlcfHdd + 52));

            int fcPlcffldHdr = BitConverter.ToInt32(wordDocData, 154 + (Fib.HeaderFieldPairIndex * 8));
            int lcbPlcffldHdr = BitConverter.ToInt32(wordDocData, 154 + (Fib.HeaderFieldPairIndex * 8) + 4);

            Assert.NotEqual(0, fcPlcffldHdr);
            Assert.True(lcbPlcffldHdr > 0);
            Assert.Equal(9, BitConverter.ToInt32(tableData, fcPlcffldHdr));
            Assert.Equal(10, BitConverter.ToInt32(tableData, fcPlcffldHdr + 4));
            Assert.Equal(12, BitConverter.ToInt32(tableData, fcPlcffldHdr + 8));
            Assert.Equal(22, BitConverter.ToInt32(tableData, fcPlcffldHdr + 12));
            Assert.Equal(0x0813, BitConverter.ToUInt16(tableData, fcPlcffldHdr + 16));
            Assert.Equal(0x0814, BitConverter.ToUInt16(tableData, fcPlcffldHdr + 18));
            Assert.Equal(0x0815, BitConverter.ToUInt16(tableData, fcPlcffldHdr + 20));
        }

        [Fact]
        public void WriteDocBlocks_WithHeaderFieldAndFloatingImageStoryAndNoteSeparator_PreservesOrderingHeaderFieldsAndGeometry()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            var model = new DocumentModel();
            model.Content.Add(new ParagraphModel
            {
                Runs =
                {
                    new RunModel { Text = "Body" }
                }
            });
            model.FootnoteSeparatorText = "Sep";

            model.Sections.Add(new SectionModel
            {
                DefaultHeaderStory = CreateStructuredHeaderFieldAndFloatingImageStory(),
                DefaultHeaderText = "Page 1 tail"
            });

            var writer = new DocWriter();
            using var ms = new MemoryStream();

            writer.WriteDocBlocks(model, ms);
            ms.Position = 0;

            using var compoundFile = new OpenMcdf.CompoundFile(ms);
            Assert.True(compoundFile.RootStorage.TryGetStream("WordDocument", out var wordDocStream));
            Assert.True(compoundFile.RootStorage.TryGetStream("1Table", out var tableStream));
            Assert.True(compoundFile.RootStorage.TryGetStream("Data", out var dataStream));

            var wordDocData = wordDocStream.GetData();
            var tableData = tableStream.GetData();
            var data = dataStream.GetData();

            string expectedText = "Body\rSep\rPage \x0013PAGE\x00141\x0015 \x0001 tail\r\r";
            var textBytes = new byte[expectedText.Length];
            System.Array.Copy(wordDocData, 1536, textBytes, 0, expectedText.Length);
            var extractedText = Encoding.GetEncoding(1252).GetString(textBytes);

            Assert.Equal(expectedText, extractedText);
            Assert.Equal(5, BitConverter.ToInt32(wordDocData, 64));
            Assert.Equal(22, BitConverter.ToInt32(wordDocData, 72));
            Assert.True(data.Length > 0);

            int fcPlcfHdd = BitConverter.ToInt32(wordDocData, 154 + (Fib.HeaderStoryPairIndex * 8));
            Assert.Equal(0, BitConverter.ToInt32(tableData, fcPlcfHdd));
            Assert.Equal(4, BitConverter.ToInt32(tableData, fcPlcfHdd + 4));
            Assert.Equal(4, BitConverter.ToInt32(tableData, fcPlcfHdd + 8));
            Assert.Equal(4, BitConverter.ToInt32(tableData, fcPlcfHdd + 12));
            Assert.Equal(4, BitConverter.ToInt32(tableData, fcPlcfHdd + 16));
            Assert.Equal(4, BitConverter.ToInt32(tableData, fcPlcfHdd + 20));
            Assert.Equal(4, BitConverter.ToInt32(tableData, fcPlcfHdd + 24));
            Assert.Equal(4, BitConverter.ToInt32(tableData, fcPlcfHdd + 28));
            Assert.Equal(22, BitConverter.ToInt32(tableData, fcPlcfHdd + 32));
            Assert.Equal(22, BitConverter.ToInt32(tableData, fcPlcfHdd + 36));
            Assert.Equal(22, BitConverter.ToInt32(tableData, fcPlcfHdd + 40));
            Assert.Equal(22, BitConverter.ToInt32(tableData, fcPlcfHdd + 44));
            Assert.Equal(21, BitConverter.ToInt32(tableData, fcPlcfHdd + 48));
            Assert.Equal(22, BitConverter.ToInt32(tableData, fcPlcfHdd + 52));

            int fcPlcffldHdr = BitConverter.ToInt32(wordDocData, 154 + (Fib.HeaderFieldPairIndex * 8));
            int lcbPlcffldHdr = BitConverter.ToInt32(wordDocData, 154 + (Fib.HeaderFieldPairIndex * 8) + 4);

            Assert.NotEqual(0, fcPlcffldHdr);
            Assert.True(lcbPlcffldHdr > 0);
            Assert.Equal(9, BitConverter.ToInt32(tableData, fcPlcffldHdr));
            Assert.Equal(10, BitConverter.ToInt32(tableData, fcPlcffldHdr + 4));
            Assert.Equal(12, BitConverter.ToInt32(tableData, fcPlcffldHdr + 8));
            Assert.Equal(22, BitConverter.ToInt32(tableData, fcPlcffldHdr + 12));
            Assert.Equal(0x0813, BitConverter.ToUInt16(tableData, fcPlcffldHdr + 16));
            Assert.Equal(0x0814, BitConverter.ToUInt16(tableData, fcPlcffldHdr + 18));
            Assert.Equal(0x0815, BitConverter.ToUInt16(tableData, fcPlcffldHdr + 20));

            int fcPlcSpaHdr = BitConverter.ToInt32(wordDocData, 154 + (Fib.HeaderShapePairIndex * 8));
            int lcbPlcSpaHdr = BitConverter.ToInt32(wordDocData, 154 + (Fib.HeaderShapePairIndex * 8) + 4);
            int fcDggInfo = BitConverter.ToInt32(wordDocData, 154 + (50 * 8));

            Assert.NotEqual(0, fcPlcSpaHdr);
            Assert.True(lcbPlcSpaHdr > 0);
            Assert.True(fcDggInfo > 0);
            Assert.Equal(14, BitConverter.ToInt32(tableData, fcPlcSpaHdr));
            Assert.Equal(22, BitConverter.ToInt32(tableData, fcPlcSpaHdr + 4));
            Assert.Equal(1026, BitConverter.ToInt32(tableData, fcPlcSpaHdr + 8));
            Assert.Equal(0, BitConverter.ToInt32(tableData, fcPlcSpaHdr + 12));
            Assert.Equal(0, BitConverter.ToInt32(tableData, fcPlcSpaHdr + 16));
            Assert.Equal(1440, BitConverter.ToInt32(tableData, fcPlcSpaHdr + 20));
            Assert.Equal(1440, BitConverter.ToInt32(tableData, fcPlcSpaHdr + 24));
            Assert.Equal(18, BitConverter.ToInt32(tableData, fcPlcSpaHdr + 30));
        }

        [Fact]
        public void WriteDocBlocks_WithHeaderHyperlinkAndInlineImageStoryAndNoteSeparator_PreservesOrderingAndData()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            var model = new DocumentModel();
            model.Content.Add(new ParagraphModel
            {
                Runs =
                {
                    new RunModel { Text = "Body" }
                }
            });
            model.FootnoteSeparatorText = "Sep";

            model.Sections.Add(new SectionModel
            {
                DefaultHeaderStory = CreateStructuredHeaderHyperlinkAndImageStory(),
                DefaultHeaderText = "Go Example now"
            });

            var writer = new DocWriter();
            using var ms = new MemoryStream();

            writer.WriteDocBlocks(model, ms);
            ms.Position = 0;

            using var compoundFile = new OpenMcdf.CompoundFile(ms);
            Assert.True(compoundFile.RootStorage.TryGetStream("WordDocument", out var wordDocStream));
            Assert.True(compoundFile.RootStorage.TryGetStream("1Table", out var tableStream));
            Assert.True(compoundFile.RootStorage.TryGetStream("Data", out var dataStream));

            var wordDocData = wordDocStream.GetData();
            var tableData = tableStream.GetData();
            var data = dataStream.GetData();

            string expectedText = "Body\rSep\rGo \x0013HYPERLINK \"https://example.com\"\x0014Example\x0015 \x0001 now\r\r";
            var textBytes = new byte[expectedText.Length];
            System.Array.Copy(wordDocData, 1536, textBytes, 0, expectedText.Length);
            var extractedText = Encoding.GetEncoding(1252).GetString(textBytes);

            Assert.Equal(expectedText, extractedText);
            Assert.Equal(5, BitConverter.ToInt32(wordDocData, 64));
            Assert.Equal(25, BitConverter.ToInt32(wordDocData, 72));
            Assert.True(data.Length > 0);

            int fcPlcfHdd = BitConverter.ToInt32(wordDocData, 154 + (Fib.HeaderStoryPairIndex * 8));
            Assert.Equal(0, BitConverter.ToInt32(tableData, fcPlcfHdd));
            Assert.Equal(4, BitConverter.ToInt32(tableData, fcPlcfHdd + 4));
            Assert.Equal(4, BitConverter.ToInt32(tableData, fcPlcfHdd + 8));
            Assert.Equal(4, BitConverter.ToInt32(tableData, fcPlcfHdd + 12));
            Assert.Equal(4, BitConverter.ToInt32(tableData, fcPlcfHdd + 16));
            Assert.Equal(4, BitConverter.ToInt32(tableData, fcPlcfHdd + 20));
            Assert.Equal(4, BitConverter.ToInt32(tableData, fcPlcfHdd + 24));
            Assert.Equal(4, BitConverter.ToInt32(tableData, fcPlcfHdd + 28));
            Assert.Equal(25, BitConverter.ToInt32(tableData, fcPlcfHdd + 32));
            Assert.Equal(25, BitConverter.ToInt32(tableData, fcPlcfHdd + 36));
            Assert.Equal(25, BitConverter.ToInt32(tableData, fcPlcfHdd + 40));
            Assert.Equal(25, BitConverter.ToInt32(tableData, fcPlcfHdd + 44));
            Assert.Equal(24, BitConverter.ToInt32(tableData, fcPlcfHdd + 48));
            Assert.Equal(25, BitConverter.ToInt32(tableData, fcPlcfHdd + 52));

            int fcPlcffldHdr = BitConverter.ToInt32(wordDocData, 154 + (Fib.HeaderFieldPairIndex * 8));
            int lcbPlcffldHdr = BitConverter.ToInt32(wordDocData, 154 + (Fib.HeaderFieldPairIndex * 8) + 4);

            Assert.NotEqual(0, fcPlcffldHdr);
            Assert.True(lcbPlcffldHdr > 0);
            Assert.Equal(7, BitConverter.ToInt32(tableData, fcPlcffldHdr));
            Assert.Equal(8, BitConverter.ToInt32(tableData, fcPlcffldHdr + 4));
            Assert.Equal(16, BitConverter.ToInt32(tableData, fcPlcffldHdr + 8));
            Assert.Equal(25, BitConverter.ToInt32(tableData, fcPlcffldHdr + 12));
            Assert.Equal(0x0013, BitConverter.ToUInt16(tableData, fcPlcffldHdr + 16));
            Assert.Equal(0x0014, BitConverter.ToUInt16(tableData, fcPlcffldHdr + 18));
            Assert.Equal(0x0015, BitConverter.ToUInt16(tableData, fcPlcffldHdr + 20));

            int fcPlcSpaHdr = BitConverter.ToInt32(wordDocData, 154 + (Fib.HeaderShapePairIndex * 8));
            int lcbPlcSpaHdr = BitConverter.ToInt32(wordDocData, 154 + (Fib.HeaderShapePairIndex * 8) + 4);

            Assert.Equal(0, fcPlcSpaHdr);
            Assert.Equal(0, lcbPlcSpaHdr);
        }

        [Fact]
        public void WriteDocBlocks_WithFloatingHeaderImageStory_WritesHeaderImagePlaceholderAndData()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            var model = new DocumentModel();
            model.Content.Add(new ParagraphModel
            {
                Runs =
                {
                    new RunModel { Text = "Body" }
                }
            });

            model.Sections.Add(new SectionModel
            {
                DefaultHeaderStory = CreateStructuredHeaderFloatingImageStory(),
                DefaultHeaderText = "HiThere"
            });

            var writer = new DocWriter();
            using var ms = new MemoryStream();

            writer.WriteDocBlocks(model, ms);
            ms.Position = 0;

            using var compoundFile = new OpenMcdf.CompoundFile(ms);
            Assert.True(compoundFile.RootStorage.TryGetStream("WordDocument", out var wordDocStream));
            Assert.True(compoundFile.RootStorage.TryGetStream("1Table", out var tableStream));
            Assert.True(compoundFile.RootStorage.TryGetStream("Data", out var dataStream));

            var wordDocData = wordDocStream.GetData();
            var tableData = tableStream.GetData();
            var data = dataStream.GetData();

            string expectedText = "Body\rHi\x0001There\r\r";
            var textBytes = new byte[expectedText.Length];
            System.Array.Copy(wordDocData, 1536, textBytes, 0, expectedText.Length);
            var extractedText = Encoding.GetEncoding(1252).GetString(textBytes);

            Assert.Equal(expectedText, extractedText);
            Assert.Equal(5, BitConverter.ToInt32(wordDocData, 64));
            Assert.Equal(10, BitConverter.ToInt32(wordDocData, 72));
            Assert.True(data.Length > 0);

            int fcPlcfHdd = BitConverter.ToInt32(wordDocData, 154 + (Fib.HeaderStoryPairIndex * 8));
            Assert.NotEqual(0, fcPlcfHdd);
            Assert.Equal(0, BitConverter.ToInt32(tableData, fcPlcfHdd + 24));
            Assert.Equal(0, BitConverter.ToInt32(tableData, fcPlcfHdd + 28));
            Assert.Equal(10, BitConverter.ToInt32(tableData, fcPlcfHdd + 32));
            Assert.Equal(10, BitConverter.ToInt32(tableData, fcPlcfHdd + 36));
            Assert.Equal(10, BitConverter.ToInt32(tableData, fcPlcfHdd + 40));
            Assert.Equal(10, BitConverter.ToInt32(tableData, fcPlcfHdd + 44));
            Assert.Equal(9, BitConverter.ToInt32(tableData, fcPlcfHdd + 48));
            Assert.Equal(10, BitConverter.ToInt32(tableData, fcPlcfHdd + 52));

            int fcPlcSpaHdr = BitConverter.ToInt32(wordDocData, 154 + (Fib.HeaderShapePairIndex * 8));
            int lcbPlcSpaHdr = BitConverter.ToInt32(wordDocData, 154 + (Fib.HeaderShapePairIndex * 8) + 4);
            int fcDggInfo = BitConverter.ToInt32(wordDocData, 154 + (50 * 8));

            Assert.NotEqual(0, fcPlcSpaHdr);
            Assert.True(lcbPlcSpaHdr > 0);
            Assert.True(fcDggInfo > 0);
            Assert.Equal(2, BitConverter.ToInt32(tableData, fcPlcSpaHdr));
            Assert.Equal(10, BitConverter.ToInt32(tableData, fcPlcSpaHdr + 4));
            Assert.Equal(1026, BitConverter.ToInt32(tableData, fcPlcSpaHdr + 8));
            Assert.Equal(0, BitConverter.ToInt32(tableData, fcPlcSpaHdr + 12));
            Assert.Equal(0, BitConverter.ToInt32(tableData, fcPlcSpaHdr + 16));
            Assert.Equal(1440, BitConverter.ToInt32(tableData, fcPlcSpaHdr + 20));
            Assert.Equal(1440, BitConverter.ToInt32(tableData, fcPlcSpaHdr + 24));
            Assert.Equal(18, BitConverter.ToInt32(tableData, fcPlcSpaHdr + 30));
        }

        [Fact]
        public void WriteDocBlocks_WithHeaderHyperlinkAndFloatingImageStoryAndNoteSeparator_PreservesOrderingAndData()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            var model = new DocumentModel();
            model.Content.Add(new ParagraphModel
            {
                Runs =
                {
                    new RunModel { Text = "Body" }
                }
            });
            model.FootnoteSeparatorText = "Sep";

            model.Sections.Add(new SectionModel
            {
                DefaultHeaderStory = CreateStructuredHeaderHyperlinkAndFloatingImageStory(),
                DefaultHeaderText = "Go Example now"
            });

            var writer = new DocWriter();
            using var ms = new MemoryStream();

            writer.WriteDocBlocks(model, ms);
            ms.Position = 0;

            using var compoundFile = new OpenMcdf.CompoundFile(ms);
            Assert.True(compoundFile.RootStorage.TryGetStream("WordDocument", out var wordDocStream));
            Assert.True(compoundFile.RootStorage.TryGetStream("1Table", out var tableStream));
            Assert.True(compoundFile.RootStorage.TryGetStream("Data", out var dataStream));

            var wordDocData = wordDocStream.GetData();
            var tableData = tableStream.GetData();
            var data = dataStream.GetData();

            string expectedText = "Body\rSep\rGo \x0013HYPERLINK \"https://example.com\"\x0014Example\x0015 \x0001 now\r\r";
            var textBytes = new byte[expectedText.Length];
            System.Array.Copy(wordDocData, 1536, textBytes, 0, expectedText.Length);
            var extractedText = Encoding.GetEncoding(1252).GetString(textBytes);

            Assert.Equal(expectedText, extractedText);
            Assert.Equal(5, BitConverter.ToInt32(wordDocData, 64));
            Assert.Equal(25, BitConverter.ToInt32(wordDocData, 72));
            Assert.True(data.Length > 0);

            int fcPlcfHdd = BitConverter.ToInt32(wordDocData, 154 + (Fib.HeaderStoryPairIndex * 8));
            Assert.Equal(0, BitConverter.ToInt32(tableData, fcPlcfHdd));
            Assert.Equal(4, BitConverter.ToInt32(tableData, fcPlcfHdd + 4));
            Assert.Equal(4, BitConverter.ToInt32(tableData, fcPlcfHdd + 8));
            Assert.Equal(4, BitConverter.ToInt32(tableData, fcPlcfHdd + 12));
            Assert.Equal(4, BitConverter.ToInt32(tableData, fcPlcfHdd + 16));
            Assert.Equal(4, BitConverter.ToInt32(tableData, fcPlcfHdd + 20));
            Assert.Equal(4, BitConverter.ToInt32(tableData, fcPlcfHdd + 24));
            Assert.Equal(4, BitConverter.ToInt32(tableData, fcPlcfHdd + 28));
            Assert.Equal(25, BitConverter.ToInt32(tableData, fcPlcfHdd + 32));
            Assert.Equal(25, BitConverter.ToInt32(tableData, fcPlcfHdd + 36));
            Assert.Equal(25, BitConverter.ToInt32(tableData, fcPlcfHdd + 40));
            Assert.Equal(25, BitConverter.ToInt32(tableData, fcPlcfHdd + 44));
            Assert.Equal(24, BitConverter.ToInt32(tableData, fcPlcfHdd + 48));
            Assert.Equal(25, BitConverter.ToInt32(tableData, fcPlcfHdd + 52));

            int fcPlcffldHdr = BitConverter.ToInt32(wordDocData, 154 + (Fib.HeaderFieldPairIndex * 8));
            int lcbPlcffldHdr = BitConverter.ToInt32(wordDocData, 154 + (Fib.HeaderFieldPairIndex * 8) + 4);

            Assert.NotEqual(0, fcPlcffldHdr);
            Assert.True(lcbPlcffldHdr > 0);
            Assert.Equal(7, BitConverter.ToInt32(tableData, fcPlcffldHdr));
            Assert.Equal(8, BitConverter.ToInt32(tableData, fcPlcffldHdr + 4));
            Assert.Equal(16, BitConverter.ToInt32(tableData, fcPlcffldHdr + 8));
            Assert.Equal(25, BitConverter.ToInt32(tableData, fcPlcffldHdr + 12));
            Assert.Equal(0x0013, BitConverter.ToUInt16(tableData, fcPlcffldHdr + 16));
            Assert.Equal(0x0014, BitConverter.ToUInt16(tableData, fcPlcffldHdr + 18));
            Assert.Equal(0x0015, BitConverter.ToUInt16(tableData, fcPlcffldHdr + 20));

            int fcPlcSpaHdr = BitConverter.ToInt32(wordDocData, 154 + (Fib.HeaderShapePairIndex * 8));
            int lcbPlcSpaHdr = BitConverter.ToInt32(wordDocData, 154 + (Fib.HeaderShapePairIndex * 8) + 4);
            int expectedHeaderImageCp = 18;

            Assert.NotEqual(0, fcPlcSpaHdr);
            Assert.True(lcbPlcSpaHdr > 0);
            Assert.Equal(expectedHeaderImageCp, BitConverter.ToInt32(tableData, fcPlcSpaHdr));
            Assert.Equal(25, BitConverter.ToInt32(tableData, fcPlcSpaHdr + 4));
            Assert.Equal(1026, BitConverter.ToInt32(tableData, fcPlcSpaHdr + 8));
            Assert.Equal(0, BitConverter.ToInt32(tableData, fcPlcSpaHdr + 12));
            Assert.Equal(0, BitConverter.ToInt32(tableData, fcPlcSpaHdr + 16));
            Assert.Equal(1440, BitConverter.ToInt32(tableData, fcPlcSpaHdr + 20));
            Assert.Equal(1440, BitConverter.ToInt32(tableData, fcPlcSpaHdr + 24));
            Assert.Equal(18, BitConverter.ToInt32(tableData, fcPlcSpaHdr + 30));
        }

        [Fact]
        public void WriteDocBlocks_WithHeaderParagraphAndTableStory_WritesHeaderTableMarkers()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            var model = new DocumentModel();
            model.Content.Add(new ParagraphModel
            {
                Runs =
                {
                    new RunModel { Text = "Body" }
                }
            });

            model.Sections.Add(new SectionModel
            {
                DefaultHeaderStory = CreateStructuredHeaderParagraphAndTableStory(),
                DefaultHeaderText = "Head\rCell"
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

            string expectedText = "Body\rHead\rCell\r\x0007\r\r";
            var textBytes = new byte[expectedText.Length];
            System.Array.Copy(wordDocData, 1536, textBytes, 0, expectedText.Length);
            var extractedText = Encoding.GetEncoding(1252).GetString(textBytes);

            Assert.Equal(expectedText, extractedText);
            Assert.Equal(5, BitConverter.ToInt32(wordDocData, 64));
            Assert.Equal(13, BitConverter.ToInt32(wordDocData, 72));

            int fcPlcfHdd = BitConverter.ToInt32(wordDocData, 154 + (Fib.HeaderStoryPairIndex * 8));
            int lcbPlcfHdd = BitConverter.ToInt32(wordDocData, 154 + (Fib.HeaderStoryPairIndex * 8) + 4);
            Assert.NotEqual(0, fcPlcfHdd);
            Assert.Equal(56, lcbPlcfHdd);
            Assert.Equal(0, BitConverter.ToInt32(tableData, fcPlcfHdd));
            Assert.Equal(0, BitConverter.ToInt32(tableData, fcPlcfHdd + 4));
            Assert.Equal(0, BitConverter.ToInt32(tableData, fcPlcfHdd + 8));
            Assert.Equal(0, BitConverter.ToInt32(tableData, fcPlcfHdd + 12));
            Assert.Equal(0, BitConverter.ToInt32(tableData, fcPlcfHdd + 16));
            Assert.Equal(0, BitConverter.ToInt32(tableData, fcPlcfHdd + 20));
            Assert.Equal(0, BitConverter.ToInt32(tableData, fcPlcfHdd + 24));
            Assert.Equal(0, BitConverter.ToInt32(tableData, fcPlcfHdd + 28));
            Assert.Equal(13, BitConverter.ToInt32(tableData, fcPlcfHdd + 32));
            Assert.Equal(13, BitConverter.ToInt32(tableData, fcPlcfHdd + 36));
            Assert.Equal(13, BitConverter.ToInt32(tableData, fcPlcfHdd + 40));
            Assert.Equal(13, BitConverter.ToInt32(tableData, fcPlcfHdd + 44));
            Assert.Equal(12, BitConverter.ToInt32(tableData, fcPlcfHdd + 48));
            Assert.Equal(13, BitConverter.ToInt32(tableData, fcPlcfHdd + 52));

            int lcbPlcfbteTapx = BitConverter.ToInt32(wordDocData, 154 + (Fib.TapxPairIndex * 8) + 4);
            Assert.True(lcbPlcfbteTapx > 0);
        }

        [Fact]
        public void WriteDocBlocks_WithHeaderTableRowHeaderAndCantSplit_WritesRowFlagSprmsIntoTapx()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            var model = new DocumentModel();
            model.Content.Add(new ParagraphModel
            {
                Runs =
                {
                    new RunModel { Text = "Body" }
                }
            });

            model.Sections.Add(new SectionModel
            {
                DefaultHeaderStory = CreateStructuredHeaderTableWithRowFlagsStory()
            });

            var writer = new DocWriter();
            using var ms = new MemoryStream();
            writer.WriteDocBlocks(model, ms);
            ms.Position = 0;

            using var compoundFile = new OpenMcdf.CompoundFile(ms);
            Assert.True(compoundFile.RootStorage.TryGetStream("WordDocument", out var wordDocStream));
            Assert.True(compoundFile.RootStorage.TryGetStream("1Table", out var tableStream));

            byte[] tapxPage = GetTapxPageData(wordDocStream.GetData(), tableStream.GetData());
            Assert.True(ContainsSubsequence(tapxPage, new byte[] { 0x03, 0x34, 0x01 }));
            Assert.True(ContainsSubsequence(tapxPage, new byte[] { 0x04, 0x34, 0x01 }));
        }

        [Fact]
        public void WriteDocBlocks_WithFooterTableRowHeaderAndCantSplit_WritesRowFlagSprmsIntoTapx()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            var model = new DocumentModel();
            model.Content.Add(new ParagraphModel
            {
                Runs =
                {
                    new RunModel { Text = "Body" }
                }
            });

            model.Sections.Add(new SectionModel
            {
                DefaultFooterStory = CreateStructuredHeaderTableWithRowFlagsStory()
            });

            var writer = new DocWriter();
            using var ms = new MemoryStream();
            writer.WriteDocBlocks(model, ms);
            ms.Position = 0;

            using var compoundFile = new OpenMcdf.CompoundFile(ms);
            Assert.True(compoundFile.RootStorage.TryGetStream("WordDocument", out var wordDocStream));
            Assert.True(compoundFile.RootStorage.TryGetStream("1Table", out var tableStream));

            byte[] tapxPage = GetTapxPageData(wordDocStream.GetData(), tableStream.GetData());
            Assert.True(ContainsSubsequence(tapxPage, new byte[] { 0x03, 0x34, 0x01 }));
            Assert.True(ContainsSubsequence(tapxPage, new byte[] { 0x04, 0x34, 0x01 }));
        }

        [Fact]
        public void WriteDocBlocks_WithHeaderTableRowAtLeastHeight_ShiftsFollowingFloatingImageDown()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            int topWithAutoRowHeight = GetHeaderFloatingImageTopWithTableRowHeight(0, TableRowHeightRule.Auto);
            int topWithAtLeastHeight = GetHeaderFloatingImageTopWithTableRowHeight(2400, TableRowHeightRule.AtLeast);

            Assert.True(topWithAtLeastHeight > topWithAutoRowHeight);
        }

        [Fact]
        public void WriteDocBlocks_WithHeaderTableCellVerticalPadding_ShiftsFollowingFloatingImageDown()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            int topWithoutPadding = GetHeaderFloatingImageTopWithTableCellPadding(0, 0);
            int topWithPadding = GetHeaderFloatingImageTopWithTableCellPadding(900, 900);

            Assert.True(topWithPadding > topWithoutPadding);
        }

        [Fact]
        public void WriteDocBlocks_WithHeaderTableCellBottomAlignment_ShiftsFloatingImageDownWithinRow()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            int topWithTopAlignment = GetHeaderTableCellFloatingImageTopWithAlignment(TableCellVerticalAlignment.Top);
            int topWithBottomAlignment = GetHeaderTableCellFloatingImageTopWithAlignment(TableCellVerticalAlignment.Bottom);

            Assert.True(topWithBottomAlignment > topWithTopAlignment);
        }

        [Fact]
        public void WriteDocBlocks_WithHeaderTableExactRowHeight_ClipsSecondParagraphFloatingImageTop()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            int topWithAutoHeight = GetHeaderSecondParagraphFloatingImageTopWithRowRule(0, TableRowHeightRule.Auto, null);
            int topWithExactHeight = GetHeaderSecondParagraphFloatingImageTopWithRowRule(900, TableRowHeightRule.Exact, null);

            Assert.True(topWithExactHeight < topWithAutoHeight);
        }

        [Fact]
        public void WriteDocBlocks_WithHeaderTableExactRowHeight_ClipsBottomAlignedSecondParagraphFloatingImageTop()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            int topWithAutoHeight = GetHeaderSecondParagraphFloatingImageTopWithRowRule(0, TableRowHeightRule.Auto, "bottom");
            int topWithExactHeight = GetHeaderSecondParagraphFloatingImageTopWithRowRule(900, TableRowHeightRule.Exact, "bottom");

            Assert.True(topWithExactHeight < topWithAutoHeight);
        }

        [Fact]
        public void WriteDocBlocks_WithHeaderTableExactRowHeight_ClipsSecondParagraphFloatingImageTop_WithNegativeSpacingAfter()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            int topWithAutoHeight = GetHeaderSecondParagraphFloatingImageTopWithRowRule(0, TableRowHeightRule.Auto, null, firstParagraphSpacingAfterTwips: -600);
            int topWithExactHeight = GetHeaderSecondParagraphFloatingImageTopWithRowRule(900, TableRowHeightRule.Exact, null, firstParagraphSpacingAfterTwips: -600);

            Assert.True(topWithExactHeight < topWithAutoHeight);
        }

        [Fact]
        public void WriteDocBlocks_WithHeaderTableExactRowHeight_ClipsSecondParagraphFloatingImageTop_WithLargePositiveSpacingAfter()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            int topWithAutoHeight = GetHeaderSecondParagraphFloatingImageTopWithRowRule(0, TableRowHeightRule.Auto, null, firstParagraphSpacingAfterTwips: 1200);
            int topWithExactHeight = GetHeaderSecondParagraphFloatingImageTopWithRowRule(900, TableRowHeightRule.Exact, null, firstParagraphSpacingAfterTwips: 1200);

            Assert.True(topWithExactHeight < topWithAutoHeight);
        }

        [Fact]
        public void WriteDocBlocks_WithHeaderTableExactRowHeight_ClampsSecondParagraphTop_WhenSpacingBeforeExceedsVisibleHeight()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            int topWithAutoHeight = GetHeaderSecondParagraphFloatingImageTopWithRowRule(
                0,
                TableRowHeightRule.Auto,
                null,
                firstParagraphSpacingAfterTwips: 0,
                secondParagraphSpacingBeforeTwips: 1800);
            int topWithExactHeight = GetHeaderSecondParagraphFloatingImageTopWithRowRule(
                900,
                TableRowHeightRule.Exact,
                null,
                firstParagraphSpacingAfterTwips: 0,
                secondParagraphSpacingBeforeTwips: 1800);

            Assert.True(topWithExactHeight < topWithAutoHeight);
        }

        [Fact]
        public void WriteDocBlocks_WithHeaderTableExactRowHeight_ClampsSecondParagraphTop_WhenParagraphRelativePositionYExceedsVisibleHeight()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            int topWithAutoHeight = GetHeaderSecondParagraphFloatingImageTopWithRowRule(
                0,
                TableRowHeightRule.Auto,
                null,
                firstParagraphSpacingAfterTwips: 0,
                secondParagraphSpacingBeforeTwips: 0,
                secondParagraphPositionYTwips: 1800);
            int topWithExactHeight = GetHeaderSecondParagraphFloatingImageTopWithRowRule(
                900,
                TableRowHeightRule.Exact,
                null,
                firstParagraphSpacingAfterTwips: 0,
                secondParagraphSpacingBeforeTwips: 0,
                secondParagraphPositionYTwips: 1800);

            Assert.True(topWithExactHeight < topWithAutoHeight);
        }

        [Fact]
        public void WriteDocBlocks_WithHeaderTableExactRowHeight_ClampsHyperlinkFloatingImageTop_WhenParagraphRelativePositionYExceedsVisibleHeight()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            int topWithAutoHeight = GetHeaderHyperlinkFloatingImageTopWithRowRule(0, TableRowHeightRule.Auto, 1800);
            int topWithExactHeight = GetHeaderHyperlinkFloatingImageTopWithRowRule(900, TableRowHeightRule.Exact, 1800);

            Assert.True(topWithExactHeight < topWithAutoHeight);
        }

        [Fact]
        public void WriteDocBlocks_WithHeaderFloatingImageDistance_ExpandsAnchorBounds()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            (int left, int top, int right, int bottom) withoutDistance = GetHeaderFloatingImageBoundsWithDistance(0, 0, 0, 0);
            (int left, int top, int right, int bottom) withDistance = GetHeaderFloatingImageBoundsWithDistance(120, 240, 360, 480);

            Assert.Equal(withoutDistance.left - 120, withDistance.left);
            Assert.Equal(withoutDistance.top - 240, withDistance.top);
            Assert.Equal(withoutDistance.right + 360, withDistance.right);
            Assert.Equal(withoutDistance.bottom + 480, withDistance.bottom);
        }

        [Fact]
        public void WriteDocBlocks_WithHeaderFloatingImageNegativeDistance_DoesNotShrinkBounds()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            (int left, int top, int right, int bottom) baseBounds = GetHeaderFloatingImageBoundsWithDistance(0, 0, 0, 0);
            (int left, int top, int right, int bottom) negativeBounds = GetHeaderFloatingImageBoundsWithDistance(-120, -240, -360, -480);

            Assert.Equal(baseBounds.left, negativeBounds.left);
            Assert.Equal(baseBounds.top, negativeBounds.top);
            Assert.Equal(baseBounds.right, negativeBounds.right);
            Assert.Equal(baseBounds.bottom, negativeBounds.bottom);
        }

        [Fact]
        public void WriteDocBlocks_WithHeaderFloatingImageVeryLargeDistance_KeepsBoundsOrdered()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            (int left, int top, int right, int bottom) bounds = GetHeaderFloatingImageBoundsWithDistance(1_500_000_000, 1_500_000_000, 1_500_000_000, 1_500_000_000);

            Assert.True(bounds.left <= bounds.right);
            Assert.True(bounds.top <= bounds.bottom);
        }

        [Fact]
        public void WriteDocBlocks_WithHeaderFloatingImageVeryLargePixelDimensions_KeepsBoundsOrdered()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            (int left, int top, int right, int bottom) bounds = GetHeaderFloatingImageBoundsWithDistance(
                0,
                0,
                0,
                0,
                imageWidthPixels: int.MaxValue,
                imageHeightPixels: int.MaxValue);

            Assert.True(bounds.left <= bounds.right);
            Assert.True(bounds.top <= bounds.bottom);
        }

        [Fact]
        public void WriteDocBlocks_WithHeaderLineRelativeTopAndBottomFloatingImage_DistanceExpandsOnlyVerticalBounds()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            (int left, int top, int right, int bottom) withoutDistance = GetHeaderFloatingImageBoundsWithDistance(0, 0, 0, 0, ImageWrapType.TopAndBottom, "line");
            (int left, int top, int right, int bottom) withDistance = GetHeaderFloatingImageBoundsWithDistance(120, 240, 360, 480, ImageWrapType.TopAndBottom, "line");

            Assert.Equal(withoutDistance.left, withDistance.left);
            Assert.Equal(withoutDistance.top - 240, withDistance.top);
            Assert.Equal(withoutDistance.right, withDistance.right);
            Assert.Equal(withoutDistance.bottom + 480, withDistance.bottom);
        }

        [Fact]
        public void WriteDocBlocks_WithHeaderParagraphRelativeTopAndBottomFloatingImage_DistanceExpandsOnlyVerticalBounds()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            (int left, int top, int right, int bottom) withoutDistance = GetHeaderFloatingImageBoundsWithDistance(0, 0, 0, 0, ImageWrapType.TopAndBottom, "paragraph");
            (int left, int top, int right, int bottom) withDistance = GetHeaderFloatingImageBoundsWithDistance(120, 240, 360, 480, ImageWrapType.TopAndBottom, "paragraph");

            Assert.Equal(withoutDistance.left, withDistance.left);
            Assert.Equal(withoutDistance.top - 240, withDistance.top);
            Assert.Equal(withoutDistance.right, withDistance.right);
            Assert.Equal(withoutDistance.bottom + 480, withDistance.bottom);
        }

        [Fact]
        public void WriteDocBlocks_WithHeaderWrapNoneFloatingImage_DistanceDoesNotExpandBounds()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            (int left, int top, int right, int bottom) withoutDistance = GetHeaderFloatingImageBoundsWithDistance(0, 0, 0, 0, ImageWrapType.None, "paragraph");
            (int left, int top, int right, int bottom) withDistance = GetHeaderFloatingImageBoundsWithDistance(120, 240, 360, 480, ImageWrapType.None, "paragraph");

            Assert.Equal(withoutDistance.left, withDistance.left);
            Assert.Equal(withoutDistance.top, withDistance.top);
            Assert.Equal(withoutDistance.right, withDistance.right);
            Assert.Equal(withoutDistance.bottom, withDistance.bottom);
        }

        [Fact]
        public void WriteDocBlocks_WithHeaderBehindTextFloatingImage_DistanceDoesNotExpandBounds()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            (int left, int top, int right, int bottom) withoutDistance = GetHeaderFloatingImageBoundsWithDistance(0, 0, 0, 0, ImageWrapType.Square, "paragraph", behindText: true);
            (int left, int top, int right, int bottom) withDistance = GetHeaderFloatingImageBoundsWithDistance(120, 240, 360, 480, ImageWrapType.Square, "paragraph", behindText: true);

            Assert.Equal(withoutDistance.left, withDistance.left);
            Assert.Equal(withoutDistance.top, withDistance.top);
            Assert.Equal(withoutDistance.right, withDistance.right);
            Assert.Equal(withoutDistance.bottom, withDistance.bottom);
        }

        [Fact]
        public void WriteDocBlocks_WithHeaderFloatingImageVerticalRelativeToLine_UsesSmallerAnchorThanParagraph()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            int paragraphRelativeTop = GetHeaderFloatingImageTopWithVerticalRelativeTo("paragraph");
            int lineRelativeTop = GetHeaderFloatingImageTopWithVerticalRelativeTo("line");

            Assert.True(lineRelativeTop < paragraphRelativeTop);
        }

        [Fact]
        public void WriteDocBlocks_WithDifferentOddEvenHeaderInsideAlignment_FlipsInsideDirectionByStoryParity()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            var model = new DocumentModel();
            model.DifferentOddAndEvenPages = true;
            model.Content.Add(new ParagraphModel
            {
                Runs =
                {
                    new RunModel { Text = "Body" }
                }
            });

            model.Sections.Add(new SectionModel
            {
                EvenPagesHeaderStory = CreateHeaderAlignedFloatingImageStory("EvenHeader", "inside"),
                DefaultHeaderStory = CreateHeaderAlignedFloatingImageStory("OddHeader", "inside")
            });

            (int evenHeaderLeft, int oddHeaderLeft) = GetFirstTwoHeaderStoryShapeLeftPositions(model);

            Assert.Equal(9026, evenHeaderLeft);
            Assert.Equal(1440, oddHeaderLeft);
            Assert.True(evenHeaderLeft > oddHeaderLeft);
        }

        [Fact]
        public void WriteDocBlocks_WithDifferentOddEvenHeaderOutsideAlignment_FlipsOutsideDirectionByStoryParity()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            var model = new DocumentModel();
            model.DifferentOddAndEvenPages = true;
            model.Content.Add(new ParagraphModel
            {
                Runs =
                {
                    new RunModel { Text = "Body" }
                }
            });

            model.Sections.Add(new SectionModel
            {
                EvenPagesHeaderStory = CreateHeaderAlignedFloatingImageStory("EvenHeader", "outside"),
                DefaultHeaderStory = CreateHeaderAlignedFloatingImageStory("OddHeader", "outside")
            });

            (int evenHeaderLeft, int oddHeaderLeft) = GetFirstTwoHeaderStoryShapeLeftPositions(model);

            Assert.Equal(1440, evenHeaderLeft);
            Assert.Equal(9026, oddHeaderLeft);
            Assert.True(oddHeaderLeft > evenHeaderLeft);
        }

        [Fact]
        public void WriteDocBlocks_WithDifferentOddEvenFooterOutsideAlignment_FlipsOutsideDirectionByStoryParity()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            var model = new DocumentModel();
            model.DifferentOddAndEvenPages = true;
            model.Content.Add(new ParagraphModel
            {
                Runs =
                {
                    new RunModel { Text = "Body" }
                }
            });

            model.Sections.Add(new SectionModel
            {
                EvenPagesFooterStory = CreateHeaderAlignedFloatingImageStory("EvenFooter", "outside"),
                DefaultFooterStory = CreateHeaderAlignedFloatingImageStory("OddFooter", "outside")
            });

            (int evenFooterLeft, int oddFooterLeft) = GetFirstTwoHeaderStoryShapeLeftPositions(model);

            Assert.Equal(1440, evenFooterLeft);
            Assert.Equal(9026, oddFooterLeft);
            Assert.True(oddFooterLeft > evenFooterLeft);
        }

        [Fact]
        public void WriteDocBlocks_WithDifferentOddEvenHeaderInsideMarginRelativeTo_FlipsSideByStoryParity()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            var model = new DocumentModel();
            model.DifferentOddAndEvenPages = true;
            model.Content.Add(new ParagraphModel
            {
                Runs =
                {
                    new RunModel { Text = "Body" }
                }
            });

            model.Sections.Add(new SectionModel
            {
                EvenPagesHeaderStory = CreateHeaderRelativeFloatingImageStory("EvenHeader", "insideMargin"),
                DefaultHeaderStory = CreateHeaderRelativeFloatingImageStory("OddHeader", "insideMargin")
            });

            (int evenHeaderLeft, int oddHeaderLeft) = GetFirstTwoHeaderStoryShapeLeftPositions(model);

            Assert.Equal(10466, evenHeaderLeft);
            Assert.Equal(1440, oddHeaderLeft);
            Assert.True(evenHeaderLeft > oddHeaderLeft);
        }

        [Fact]
        public void WriteDocBlocks_WithDifferentOddEvenHeaderOutsideMarginRelativeTo_FlipsSideByStoryParity()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            var model = new DocumentModel();
            model.DifferentOddAndEvenPages = true;
            model.Content.Add(new ParagraphModel
            {
                Runs =
                {
                    new RunModel { Text = "Body" }
                }
            });

            model.Sections.Add(new SectionModel
            {
                EvenPagesHeaderStory = CreateHeaderRelativeFloatingImageStory("EvenHeader", "outsideMargin"),
                DefaultHeaderStory = CreateHeaderRelativeFloatingImageStory("OddHeader", "outsideMargin")
            });

            (int evenHeaderLeft, int oddHeaderLeft) = GetFirstTwoHeaderStoryShapeLeftPositions(model);

            Assert.Equal(1440, evenHeaderLeft);
            Assert.Equal(10466, oddHeaderLeft);
            Assert.True(oddHeaderLeft > evenHeaderLeft);
        }

        [Fact]
        public void WriteDocBlocks_WithDifferentOddEvenFooterInsideMarginRelativeTo_FlipsSideByStoryParity()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            var model = new DocumentModel();
            model.DifferentOddAndEvenPages = true;
            model.Content.Add(new ParagraphModel
            {
                Runs =
                {
                    new RunModel { Text = "Body" }
                }
            });

            model.Sections.Add(new SectionModel
            {
                EvenPagesFooterStory = CreateHeaderRelativeFloatingImageStory("EvenFooter", "insideMargin"),
                DefaultFooterStory = CreateHeaderRelativeFloatingImageStory("OddFooter", "insideMargin")
            });

            (int evenFooterLeft, int oddFooterLeft) = GetFirstTwoHeaderStoryShapeLeftPositions(model);

            Assert.Equal(10466, evenFooterLeft);
            Assert.Equal(1440, oddFooterLeft);
            Assert.True(evenFooterLeft > oddFooterLeft);
        }

        [Fact]
        public void WriteDocBlocks_WithDifferentOddEvenFooterOutsideMarginRelativeTo_FlipsSideByStoryParity()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            var model = new DocumentModel();
            model.DifferentOddAndEvenPages = true;
            model.Content.Add(new ParagraphModel
            {
                Runs =
                {
                    new RunModel { Text = "Body" }
                }
            });

            model.Sections.Add(new SectionModel
            {
                EvenPagesFooterStory = CreateHeaderRelativeFloatingImageStory("EvenFooter", "outsideMargin"),
                DefaultFooterStory = CreateHeaderRelativeFloatingImageStory("OddFooter", "outsideMargin")
            });

            (int evenFooterLeft, int oddFooterLeft) = GetFirstTwoHeaderStoryShapeLeftPositions(model);

            Assert.Equal(1440, evenFooterLeft);
            Assert.Equal(10466, oddFooterLeft);
            Assert.True(oddFooterLeft > evenFooterLeft);
        }

        [Fact]
        public void WriteDocBlocks_WithFooterFloatingImageDistance_ExpandsAnchorBounds()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            (int left, int top, int right, int bottom) withoutDistance = GetFooterFloatingImageBoundsWithDistance(0, 0, 0, 0);
            (int left, int top, int right, int bottom) withDistance = GetFooterFloatingImageBoundsWithDistance(120, 240, 360, 480);

            Assert.Equal(withoutDistance.left - 120, withDistance.left);
            Assert.Equal(withoutDistance.top - 240, withDistance.top);
            Assert.Equal(withoutDistance.right + 360, withDistance.right);
            Assert.Equal(withoutDistance.bottom + 480, withDistance.bottom);
        }

        [Fact]
        public void WriteDocBlocks_WithFooterFloatingImageNegativeDistance_DoesNotShrinkBounds()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            (int left, int top, int right, int bottom) baseBounds = GetFooterFloatingImageBoundsWithDistance(0, 0, 0, 0);
            (int left, int top, int right, int bottom) negativeBounds = GetFooterFloatingImageBoundsWithDistance(-120, -240, -360, -480);

            Assert.Equal(baseBounds.left, negativeBounds.left);
            Assert.Equal(baseBounds.top, negativeBounds.top);
            Assert.Equal(baseBounds.right, negativeBounds.right);
            Assert.Equal(baseBounds.bottom, negativeBounds.bottom);
        }

        [Fact]
        public void WriteDocBlocks_WithFooterFloatingImageVeryLargeDistance_KeepsBoundsOrdered()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            (int left, int top, int right, int bottom) bounds = GetFooterFloatingImageBoundsWithDistance(1_500_000_000, 1_500_000_000, 1_500_000_000, 1_500_000_000);

            Assert.True(bounds.left <= bounds.right);
            Assert.True(bounds.top <= bounds.bottom);
        }

        [Fact]
        public void WriteDocBlocks_WithFooterFloatingImageVeryLargePixelDimensions_KeepsBoundsOrdered()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            (int left, int top, int right, int bottom) bounds = GetFooterFloatingImageBoundsWithDistance(
                0,
                0,
                0,
                0,
                imageWidthPixels: int.MaxValue,
                imageHeightPixels: int.MaxValue);

            Assert.True(bounds.left <= bounds.right);
            Assert.True(bounds.top <= bounds.bottom);
        }

        [Fact]
        public void WriteDocBlocks_WithFooterLineRelativeTopAndBottomFloatingImage_DistanceExpandsOnlyVerticalBounds()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            (int left, int top, int right, int bottom) withoutDistance = GetFooterFloatingImageBoundsWithDistance(0, 0, 0, 0, ImageWrapType.TopAndBottom, "line");
            (int left, int top, int right, int bottom) withDistance = GetFooterFloatingImageBoundsWithDistance(120, 240, 360, 480, ImageWrapType.TopAndBottom, "line");

            Assert.Equal(withoutDistance.left, withDistance.left);
            Assert.Equal(withoutDistance.top - 240, withDistance.top);
            Assert.Equal(withoutDistance.right, withDistance.right);
            Assert.Equal(withoutDistance.bottom + 480, withDistance.bottom);
        }

        [Fact]
        public void WriteDocBlocks_WithFooterParagraphRelativeTopAndBottomFloatingImage_DistanceExpandsOnlyVerticalBounds()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            (int left, int top, int right, int bottom) withoutDistance = GetFooterFloatingImageBoundsWithDistance(0, 0, 0, 0, ImageWrapType.TopAndBottom, "paragraph");
            (int left, int top, int right, int bottom) withDistance = GetFooterFloatingImageBoundsWithDistance(120, 240, 360, 480, ImageWrapType.TopAndBottom, "paragraph");

            Assert.Equal(withoutDistance.left, withDistance.left);
            Assert.Equal(withoutDistance.top - 240, withDistance.top);
            Assert.Equal(withoutDistance.right, withDistance.right);
            Assert.Equal(withoutDistance.bottom + 480, withDistance.bottom);
        }

        [Fact]
        public void WriteDocBlocks_WithFooterWrapNoneFloatingImage_DistanceDoesNotExpandBounds()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            (int left, int top, int right, int bottom) withoutDistance = GetFooterFloatingImageBoundsWithDistance(0, 0, 0, 0, ImageWrapType.None, "paragraph");
            (int left, int top, int right, int bottom) withDistance = GetFooterFloatingImageBoundsWithDistance(120, 240, 360, 480, ImageWrapType.None, "paragraph");

            Assert.Equal(withoutDistance.left, withDistance.left);
            Assert.Equal(withoutDistance.top, withDistance.top);
            Assert.Equal(withoutDistance.right, withDistance.right);
            Assert.Equal(withoutDistance.bottom, withDistance.bottom);
        }

        [Fact]
        public void WriteDocBlocks_WithFooterBehindTextFloatingImage_DistanceDoesNotExpandBounds()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            (int left, int top, int right, int bottom) withoutDistance = GetFooterFloatingImageBoundsWithDistance(0, 0, 0, 0, ImageWrapType.Square, "paragraph", behindText: true);
            (int left, int top, int right, int bottom) withDistance = GetFooterFloatingImageBoundsWithDistance(120, 240, 360, 480, ImageWrapType.Square, "paragraph", behindText: true);

            Assert.Equal(withoutDistance.left, withDistance.left);
            Assert.Equal(withoutDistance.top, withDistance.top);
            Assert.Equal(withoutDistance.right, withDistance.right);
            Assert.Equal(withoutDistance.bottom, withDistance.bottom);
        }

        [Fact]
        public void WriteDocBlocks_WithFooterFloatingImageVerticalRelativeToLine_UsesSmallerAnchorThanParagraph()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            int paragraphRelativeTop = GetFooterFloatingImageTopWithVerticalRelativeTo("paragraph");
            int lineRelativeTop = GetFooterFloatingImageTopWithVerticalRelativeTo("line");

            Assert.True(lineRelativeTop < paragraphRelativeTop);
        }

        [Fact]
        public void WriteDocBlocks_WithFooterTableExactRowHeight_ClipsSecondParagraphFloatingImageTop()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            int topWithAutoHeight = GetFooterSecondParagraphFloatingImageTopWithRowRule(0, TableRowHeightRule.Auto, null);
            int topWithExactHeight = GetFooterSecondParagraphFloatingImageTopWithRowRule(900, TableRowHeightRule.Exact, null);

            Assert.True(topWithExactHeight < topWithAutoHeight);
        }

        [Fact]
        public void WriteDocBlocks_WithFooterTableExactRowHeight_ClipsSecondParagraphFloatingImageTop_WithNegativeSpacingAfter()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            int topWithAutoHeight = GetFooterSecondParagraphFloatingImageTopWithRowRule(0, TableRowHeightRule.Auto, null, firstParagraphSpacingAfterTwips: -600);
            int topWithExactHeight = GetFooterSecondParagraphFloatingImageTopWithRowRule(900, TableRowHeightRule.Exact, null, firstParagraphSpacingAfterTwips: -600);

            Assert.True(topWithExactHeight < topWithAutoHeight);
        }

        [Fact]
        public void WriteDocBlocks_WithFooterTableExactRowHeight_ClipsSecondParagraphFloatingImageTop_WithLargePositiveSpacingAfter()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            int topWithAutoHeight = GetFooterSecondParagraphFloatingImageTopWithRowRule(0, TableRowHeightRule.Auto, null, firstParagraphSpacingAfterTwips: 1200);
            int topWithExactHeight = GetFooterSecondParagraphFloatingImageTopWithRowRule(900, TableRowHeightRule.Exact, null, firstParagraphSpacingAfterTwips: 1200);

            Assert.True(topWithExactHeight < topWithAutoHeight);
        }

        [Fact]
        public void WriteDocBlocks_WithFooterTableExactRowHeight_ClampsSecondParagraphTop_WhenSpacingBeforeExceedsVisibleHeight()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            int topWithAutoHeight = GetFooterSecondParagraphFloatingImageTopWithRowRule(
                0,
                TableRowHeightRule.Auto,
                null,
                firstParagraphSpacingAfterTwips: 0,
                secondParagraphSpacingBeforeTwips: 1800);
            int topWithExactHeight = GetFooterSecondParagraphFloatingImageTopWithRowRule(
                900,
                TableRowHeightRule.Exact,
                null,
                firstParagraphSpacingAfterTwips: 0,
                secondParagraphSpacingBeforeTwips: 1800);

            Assert.True(topWithExactHeight < topWithAutoHeight);
        }

        [Fact]
        public void WriteDocBlocks_WithFooterTableExactRowHeight_ClampsHyperlinkFloatingImageTop_WhenParagraphRelativePositionYExceedsVisibleHeight()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            int topWithAutoHeight = GetFooterHyperlinkFloatingImageTopWithRowRule(0, TableRowHeightRule.Auto, 1800);
            int topWithExactHeight = GetFooterHyperlinkFloatingImageTopWithRowRule(900, TableRowHeightRule.Exact, 1800);

            Assert.True(topWithExactHeight < topWithAutoHeight);
        }

        [Fact]
        public void WriteDocBlocks_WithFooterTableExactRowHeight_ClampsSecondParagraphTop_WhenParagraphRelativePositionYIsNegativeBeyondVisibleStart()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            int topWithAutoNegative = GetFooterSecondParagraphFloatingImageTopWithRowRule(0, TableRowHeightRule.Auto, null, secondParagraphPositionYTwips: -1800);
            int topWithAutoZero = GetFooterSecondParagraphFloatingImageTopWithRowRule(0, TableRowHeightRule.Auto, null, secondParagraphPositionYTwips: 0);
            int topWithExactNegative = GetFooterSecondParagraphFloatingImageTopWithRowRule(900, TableRowHeightRule.Exact, null, secondParagraphPositionYTwips: -1800);
            int topWithExactZero = GetFooterSecondParagraphFloatingImageTopWithRowRule(900, TableRowHeightRule.Exact, null, secondParagraphPositionYTwips: 0);

            Assert.True(topWithAutoNegative < topWithAutoZero);
            Assert.True(topWithExactNegative < topWithAutoZero);
        }

        [Fact]
        public void WriteDocBlocks_WithFooterTableExactRowHeight_ClampsHyperlinkFloatingImageTop_WhenParagraphRelativePositionYIsNegativeBeyondVisibleStart()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            int topWithAutoNegative = GetFooterHyperlinkFloatingImageTopWithRowRule(0, TableRowHeightRule.Auto, -1800);
            int topWithAutoZero = GetFooterHyperlinkFloatingImageTopWithRowRule(0, TableRowHeightRule.Auto, 0);
            int topWithExactNegative = GetFooterHyperlinkFloatingImageTopWithRowRule(900, TableRowHeightRule.Exact, -1800);
            int topWithExactZero = GetFooterHyperlinkFloatingImageTopWithRowRule(900, TableRowHeightRule.Exact, 0);

            Assert.True(topWithAutoNegative < topWithAutoZero);
            Assert.True(topWithExactNegative >= topWithExactZero);
        }

        [Fact]
        public void WriteDocBlocks_WithFooterTableExactRowHeight_ClampsNegativeParagraphRelativePositionY_ToParagraphStart()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            int topWithExactNegative = GetFooterSecondParagraphFloatingImageTopWithRowRule(
                900,
                TableRowHeightRule.Exact,
                null,
                secondParagraphPositionYTwips: -1800);
            int topWithExactZero = GetFooterSecondParagraphFloatingImageTopWithRowRule(
                900,
                TableRowHeightRule.Exact,
                null,
                secondParagraphPositionYTwips: 0);

            Assert.Equal(topWithExactZero, topWithExactNegative);
        }

        [Fact]
        public void WriteDocBlocks_WithFooterTableExactRowHeight_ClipsNestedTableFloatingImageTop()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            int topWithAutoHeight = GetFooterNestedTableFloatingImageTopWithOuterRowRule(0, TableRowHeightRule.Auto, 0);
            int topWithExactHeight = GetFooterNestedTableFloatingImageTopWithOuterRowRule(1200, TableRowHeightRule.Exact, 0);

            Assert.True(topWithExactHeight < topWithAutoHeight);
        }

        [Fact]
        public void WriteDocBlocks_WithFooterTableExactRowHeight_ClipsNestedTableFloatingImageTop_WithNegativeSpacingAfter()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            int topWithAutoHeight = GetFooterNestedTableFloatingImageTopWithOuterRowRule(0, TableRowHeightRule.Auto, -600);
            int topWithExactHeight = GetFooterNestedTableFloatingImageTopWithOuterRowRule(1200, TableRowHeightRule.Exact, -600);

            Assert.True(topWithExactHeight < topWithAutoHeight);
        }

        [Fact]
        public void WriteDocBlocks_WithFooterTableExactRowHeight_ClipsNestedTableFloatingImageTop_WithLargePositiveSpacingAfter()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            int topWithAutoHeight = GetFooterNestedTableFloatingImageTopWithOuterRowRule(0, TableRowHeightRule.Auto, 1200);
            int topWithExactHeight = GetFooterNestedTableFloatingImageTopWithOuterRowRule(1200, TableRowHeightRule.Exact, 1200);

            Assert.True(topWithExactHeight < topWithAutoHeight);
        }

        [Fact]
        public void WriteDocBlocks_WithFooterTableExactRowHeight_ClampsNestedSecondParagraphTop_WhenSpacingBeforeExceedsVisibleHeight()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            int topWithAutoHeight = GetFooterNestedTableFloatingImageTopWithOuterRowRule(
                0,
                TableRowHeightRule.Auto,
                firstNestedParagraphSpacingAfterTwips: 0,
                secondNestedParagraphSpacingBeforeTwips: 1800,
                secondNestedParagraphPositionYTwips: 0);
            int topWithExactHeight = GetFooterNestedTableFloatingImageTopWithOuterRowRule(
                1200,
                TableRowHeightRule.Exact,
                firstNestedParagraphSpacingAfterTwips: 0,
                secondNestedParagraphSpacingBeforeTwips: 1800,
                secondNestedParagraphPositionYTwips: 0);

            Assert.True(topWithExactHeight < topWithAutoHeight);
        }

        [Fact]
        public void WriteDocBlocks_WithFooterTableExactRowHeight_ClampsNestedSecondParagraphTop_WhenParagraphRelativePositionYExceedsVisibleHeight()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            int topWithAutoHeight = GetFooterNestedTableFloatingImageTopWithOuterRowRule(
                0,
                TableRowHeightRule.Auto,
                firstNestedParagraphSpacingAfterTwips: 0,
                secondNestedParagraphSpacingBeforeTwips: 0,
                secondNestedParagraphPositionYTwips: 1800);
            int topWithExactHeight = GetFooterNestedTableFloatingImageTopWithOuterRowRule(
                1200,
                TableRowHeightRule.Exact,
                firstNestedParagraphSpacingAfterTwips: 0,
                secondNestedParagraphSpacingBeforeTwips: 0,
                secondNestedParagraphPositionYTwips: 1800);

            Assert.True(topWithExactHeight < topWithAutoHeight);
        }

        [Fact]
        public void WriteDocBlocks_WithFooterTableExactRowHeight_ClampsNestedHyperlinkFloatingImageTop_WhenParagraphRelativePositionYExceedsVisibleHeight()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            int topWithAutoHeight = GetFooterNestedHyperlinkFloatingImageTopWithOuterRowRule(0, TableRowHeightRule.Auto, 1800);
            int topWithExactHeight = GetFooterNestedHyperlinkFloatingImageTopWithOuterRowRule(1200, TableRowHeightRule.Exact, 1800);

            Assert.True(topWithExactHeight < topWithAutoHeight);
        }

        [Fact]
        public void WriteDocBlocks_WithHeaderTableExactRowHeight_ClampsSecondParagraphTop_WhenParagraphRelativePositionYIsNegativeBeyondVisibleStart()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            int topWithAutoNegative = GetHeaderSecondParagraphFloatingImageTopWithRowRule(
                0,
                TableRowHeightRule.Auto,
                null,
                firstParagraphSpacingAfterTwips: 0,
                secondParagraphSpacingBeforeTwips: 0,
                secondParagraphPositionYTwips: -1800);
            int topWithAutoZero = GetHeaderSecondParagraphFloatingImageTopWithRowRule(
                0,
                TableRowHeightRule.Auto,
                null,
                firstParagraphSpacingAfterTwips: 0,
                secondParagraphSpacingBeforeTwips: 0,
                secondParagraphPositionYTwips: 0);
            int topWithExactNegative = GetHeaderSecondParagraphFloatingImageTopWithRowRule(
                900,
                TableRowHeightRule.Exact,
                null,
                firstParagraphSpacingAfterTwips: 0,
                secondParagraphSpacingBeforeTwips: 0,
                secondParagraphPositionYTwips: -1800);
            int topWithExactZero = GetHeaderSecondParagraphFloatingImageTopWithRowRule(
                900,
                TableRowHeightRule.Exact,
                null,
                firstParagraphSpacingAfterTwips: 0,
                secondParagraphSpacingBeforeTwips: 0,
                secondParagraphPositionYTwips: 0);

            Assert.True(topWithAutoNegative < topWithAutoZero);
            Assert.True(topWithExactNegative < topWithAutoZero);
        }

        [Fact]
        public void WriteDocBlocks_WithHeaderTableExactRowHeight_ClampsHyperlinkFloatingImageTop_WhenParagraphRelativePositionYIsNegativeBeyondVisibleStart()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            int topWithAutoNegative = GetHeaderHyperlinkFloatingImageTopWithRowRule(0, TableRowHeightRule.Auto, -1800);
            int topWithAutoZero = GetHeaderHyperlinkFloatingImageTopWithRowRule(0, TableRowHeightRule.Auto, 0);
            int topWithExactNegative = GetHeaderHyperlinkFloatingImageTopWithRowRule(900, TableRowHeightRule.Exact, -1800);
            int topWithExactZero = GetHeaderHyperlinkFloatingImageTopWithRowRule(900, TableRowHeightRule.Exact, 0);

            Assert.True(topWithAutoNegative < topWithAutoZero);
            Assert.True(topWithExactNegative >= topWithExactZero);
        }

        [Fact]
        public void WriteDocBlocks_WithHeaderTableExactRowHeight_ClipsNestedTableFloatingImageTop()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            int topWithAutoHeight = GetHeaderNestedTableFloatingImageTopWithOuterRowRule(0, TableRowHeightRule.Auto, 0);
            int topWithExactHeight = GetHeaderNestedTableFloatingImageTopWithOuterRowRule(1200, TableRowHeightRule.Exact, 0);

            Assert.True(topWithExactHeight < topWithAutoHeight);
        }

        [Fact]
        public void WriteDocBlocks_WithHeaderTableExactRowHeight_ClipsNestedTableFloatingImageTop_WithNegativeSpacingAfter()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            int topWithAutoHeight = GetHeaderNestedTableFloatingImageTopWithOuterRowRule(0, TableRowHeightRule.Auto, -600);
            int topWithExactHeight = GetHeaderNestedTableFloatingImageTopWithOuterRowRule(1200, TableRowHeightRule.Exact, -600);

            Assert.True(topWithExactHeight < topWithAutoHeight);
        }

        [Fact]
        public void WriteDocBlocks_WithHeaderTableExactRowHeight_ClipsNestedTableFloatingImageTop_WithLargePositiveSpacingAfter()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            int topWithAutoHeight = GetHeaderNestedTableFloatingImageTopWithOuterRowRule(0, TableRowHeightRule.Auto, 1200);
            int topWithExactHeight = GetHeaderNestedTableFloatingImageTopWithOuterRowRule(1200, TableRowHeightRule.Exact, 1200);

            Assert.True(topWithExactHeight < topWithAutoHeight);
        }

        [Fact]
        public void WriteDocBlocks_WithHeaderTableExactRowHeight_ClampsNestedSecondParagraphTop_WhenSpacingBeforeExceedsVisibleHeight()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            int topWithAutoHeight = GetHeaderNestedTableFloatingImageTopWithOuterRowRule(
                0,
                TableRowHeightRule.Auto,
                firstNestedParagraphSpacingAfterTwips: 0,
                secondNestedParagraphSpacingBeforeTwips: 1800,
                secondNestedParagraphPositionYTwips: 0);
            int topWithExactHeight = GetHeaderNestedTableFloatingImageTopWithOuterRowRule(
                1200,
                TableRowHeightRule.Exact,
                firstNestedParagraphSpacingAfterTwips: 0,
                secondNestedParagraphSpacingBeforeTwips: 1800,
                secondNestedParagraphPositionYTwips: 0);

            Assert.True(topWithExactHeight < topWithAutoHeight);
        }

        [Fact]
        public void WriteDocBlocks_WithHeaderTableExactRowHeight_ClampsNestedSecondParagraphTop_WhenParagraphRelativePositionYExceedsVisibleHeight()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            int topWithAutoHeight = GetHeaderNestedTableFloatingImageTopWithOuterRowRule(
                0,
                TableRowHeightRule.Auto,
                firstNestedParagraphSpacingAfterTwips: 0,
                secondNestedParagraphSpacingBeforeTwips: 0,
                secondNestedParagraphPositionYTwips: 1800);
            int topWithExactHeight = GetHeaderNestedTableFloatingImageTopWithOuterRowRule(
                1200,
                TableRowHeightRule.Exact,
                firstNestedParagraphSpacingAfterTwips: 0,
                secondNestedParagraphSpacingBeforeTwips: 0,
                secondNestedParagraphPositionYTwips: 1800);

            Assert.True(topWithExactHeight < topWithAutoHeight);
        }

        [Fact]
        public void WriteDocBlocks_WithHeaderTableExactRowHeight_ClampsNestedHyperlinkFloatingImageTop_WhenParagraphRelativePositionYExceedsVisibleHeight()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            int topWithAutoHeight = GetHeaderNestedHyperlinkFloatingImageTopWithOuterRowRule(0, TableRowHeightRule.Auto, 1800);
            int topWithExactHeight = GetHeaderNestedHyperlinkFloatingImageTopWithOuterRowRule(1200, TableRowHeightRule.Exact, 1800);

            Assert.True(topWithExactHeight < topWithAutoHeight);
        }

        [Fact]
        public void WriteDocBlocks_WithHeaderTableCellVerticalBorders_ShiftsFollowingFloatingImageDown()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            int topWithoutBorders = GetHeaderFloatingImageTopWithTableCellBorders(0, 0);
            int topWithBorders = GetHeaderFloatingImageTopWithTableCellBorders(300, 300);

            Assert.True(topWithBorders > topWithoutBorders);
        }

        [Fact]
        public void WriteDocBlocks_WithHeaderTableFixedFirstCellWidth_ShrinksSecondCellAndPushesSecondParagraphImageDown()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            int topWithAutoFirstCell = GetHeaderSecondCellSecondParagraphImageTop(hasFixedFirstCellWidth: false);
            int topWithFixedFirstCell = GetHeaderSecondCellSecondParagraphImageTop(hasFixedFirstCellWidth: true);

            Assert.True(topWithFixedFirstCell > topWithAutoFirstCell);
        }

        [Fact]
        public void WriteDocBlocks_WithHeaderTableInsideHorizontalBorder_ShiftsSecondRowFloatingImageDown()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            int topWithoutInsideHorizontal = GetHeaderSecondRowFloatingImageTopWithInsideHorizontalBorder(0);
            int topWithInsideHorizontal = GetHeaderSecondRowFloatingImageTopWithInsideHorizontalBorder(300);

            Assert.True(topWithInsideHorizontal > topWithoutInsideHorizontal);
        }

        [Fact]
        public void WriteDocBlocks_WithHeaderTableInsideVerticalBorder_ShrinksSecondCellAndPushesSecondParagraphImageDown()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            int topWithoutInsideVertical = GetHeaderSecondCellSecondParagraphImageTopWithInsideVerticalBorder(0);
            int topWithInsideVertical = GetHeaderSecondCellSecondParagraphImageTopWithInsideVerticalBorder(1200);

            Assert.True(topWithInsideVertical > topWithoutInsideVertical);
        }

        [Fact]
        public void WriteDocBlocks_WithHeaderTableInsideVerticalBorder_ExplicitNoneOnPreviousRightBorderSuppressesImpact()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            int topWithInsideOnly = GetHeaderSecondCellSecondParagraphImageTopWithInsideVerticalBorderConflict(1200, explicitPreviousRightBorderTwips: null, setExplicitOverride: false);
            int topWithExplicitNone = GetHeaderSecondCellSecondParagraphImageTopWithInsideVerticalBorderConflict(1200, explicitPreviousRightBorderTwips: 0, setExplicitOverride: true);

            Assert.True(topWithExplicitNone < topWithInsideOnly);
        }

        [Fact]
        public void WriteDocBlocks_WithHeaderTableInsideVerticalBorder_ExplicitPreviousRightBorderAmplifiesImpact()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            int topWithInsideOnly = GetHeaderSecondCellSecondParagraphImageTopWithInsideVerticalBorderConflict(1200, explicitPreviousRightBorderTwips: null, setExplicitOverride: false);
            int topWithExplicitRightBorder = GetHeaderSecondCellSecondParagraphImageTopWithInsideVerticalBorderConflict(1200, explicitPreviousRightBorderTwips: 1800, setExplicitOverride: true);

            Assert.True(topWithExplicitRightBorder > topWithInsideOnly);
        }

        [Fact]
        public void WriteDocBlocks_WithHeaderTableInsideVerticalBorder_ExplicitCurrentLeftNoneSuppressesImpact()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            int topWithInsideOnly = GetHeaderSecondCellSecondParagraphImageTopWithInsideVerticalBorderCurrentLeftConflict(1200, explicitCurrentLeftBorderTwips: null, setExplicitOverride: false);
            int topWithExplicitLeftNone = GetHeaderSecondCellSecondParagraphImageTopWithInsideVerticalBorderCurrentLeftConflict(1200, explicitCurrentLeftBorderTwips: 0, setExplicitOverride: true);

            Assert.True(topWithExplicitLeftNone < topWithInsideOnly);
        }

        [Fact]
        public void WriteDocBlocks_WithHeaderTableInsideVerticalBorder_ExplicitCurrentLeftBorderAmplifiesImpact()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            int topWithInsideOnly = GetHeaderSecondCellSecondParagraphImageTopWithInsideVerticalBorderCurrentLeftConflict(1200, explicitCurrentLeftBorderTwips: null, setExplicitOverride: false);
            int topWithExplicitLeftBorder = GetHeaderSecondCellSecondParagraphImageTopWithInsideVerticalBorderCurrentLeftConflict(1200, explicitCurrentLeftBorderTwips: 1800, setExplicitOverride: true);

            Assert.True(topWithExplicitLeftBorder > topWithInsideOnly);
        }

        [Fact]
        public void WriteDocBlocks_WithHeaderTableInsideHorizontalBorder_ExplicitNoneOnPreviousBottomBorderSuppressesImpact()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            int topWithInsideOnly = GetHeaderSecondRowFloatingImageTopWithInsideHorizontalBorderConflict(900, explicitPreviousBottomBorderTwips: null, setExplicitOverride: false);
            int topWithExplicitNone = GetHeaderSecondRowFloatingImageTopWithInsideHorizontalBorderConflict(900, explicitPreviousBottomBorderTwips: 0, setExplicitOverride: true);

            Assert.True(topWithExplicitNone < topWithInsideOnly);
        }

        [Fact]
        public void WriteDocBlocks_WithHeaderTableInsideHorizontalBorder_ExplicitPreviousBottomBorderAmplifiesImpact()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            int topWithInsideOnly = GetHeaderSecondRowFloatingImageTopWithInsideHorizontalBorderConflict(300, explicitPreviousBottomBorderTwips: null, setExplicitOverride: false);
            int topWithExplicitBottomBorder = GetHeaderSecondRowFloatingImageTopWithInsideHorizontalBorderConflict(300, explicitPreviousBottomBorderTwips: 1200, setExplicitOverride: true);

            Assert.True(topWithExplicitBottomBorder > topWithInsideOnly);
        }

        [Fact]
        public void WriteDocBlocks_WithHeaderTableInsideHorizontalBorder_ExplicitCurrentTopNoneSuppressesImpact()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            int topWithInsideOnly = GetHeaderSecondRowFloatingImageTopWithInsideHorizontalBorderCurrentTopConflict(900, explicitCurrentTopBorderTwips: null, setExplicitOverride: false);
            int topWithExplicitTopNone = GetHeaderSecondRowFloatingImageTopWithInsideHorizontalBorderCurrentTopConflict(900, explicitCurrentTopBorderTwips: 0, setExplicitOverride: true);

            Assert.True(topWithExplicitTopNone < topWithInsideOnly);
        }

        [Fact]
        public void WriteDocBlocks_WithHeaderTableInsideHorizontalBorder_ExplicitCurrentTopBorderAmplifiesImpact()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            int topWithInsideOnly = GetHeaderSecondRowFloatingImageTopWithInsideHorizontalBorderCurrentTopConflict(900, explicitCurrentTopBorderTwips: null, setExplicitOverride: false);
            int topWithExplicitTopBorder = GetHeaderSecondRowFloatingImageTopWithInsideHorizontalBorderCurrentTopConflict(900, explicitCurrentTopBorderTwips: 1200, setExplicitOverride: true);

            Assert.True(topWithExplicitTopBorder > topWithInsideOnly);
        }

        [Fact]
        public void WriteDocBlocks_WithEmptyStructuredHeaderStory_WritesGuardParagraphOnly()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            var model = new DocumentModel();
            model.Content.Add(new ParagraphModel
            {
                Runs =
                {
                    new RunModel { Text = "Body" }
                }
            });

            model.Sections.Add(new SectionModel
            {
                DefaultHeaderStory = new HeaderFooterStoryModel()
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

            string expectedText = "Body\r\r\r";
            var textBytes = new byte[expectedText.Length];
            System.Array.Copy(wordDocData, 1536, textBytes, 0, expectedText.Length);
            var extractedText = Encoding.GetEncoding(1252).GetString(textBytes);

            Assert.Equal(expectedText, extractedText);
            Assert.Equal(5, BitConverter.ToInt32(wordDocData, 64));
            Assert.Equal(2, BitConverter.ToInt32(wordDocData, 72));

            int fcPlcfHdd = BitConverter.ToInt32(wordDocData, 154 + (Fib.HeaderStoryPairIndex * 8));
            int lcbPlcfHdd = BitConverter.ToInt32(wordDocData, 154 + (Fib.HeaderStoryPairIndex * 8) + 4);

            Assert.NotEqual(0, fcPlcfHdd);
            Assert.Equal(56, lcbPlcfHdd);
            Assert.Equal(0, BitConverter.ToInt32(tableData, fcPlcfHdd));
            Assert.Equal(0, BitConverter.ToInt32(tableData, fcPlcfHdd + 4));
            Assert.Equal(0, BitConverter.ToInt32(tableData, fcPlcfHdd + 8));
            Assert.Equal(0, BitConverter.ToInt32(tableData, fcPlcfHdd + 12));
            Assert.Equal(0, BitConverter.ToInt32(tableData, fcPlcfHdd + 16));
            Assert.Equal(0, BitConverter.ToInt32(tableData, fcPlcfHdd + 20));
            Assert.Equal(0, BitConverter.ToInt32(tableData, fcPlcfHdd + 24));
            Assert.Equal(0, BitConverter.ToInt32(tableData, fcPlcfHdd + 28));
            Assert.Equal(2, BitConverter.ToInt32(tableData, fcPlcfHdd + 32));
            Assert.Equal(2, BitConverter.ToInt32(tableData, fcPlcfHdd + 36));
            Assert.Equal(2, BitConverter.ToInt32(tableData, fcPlcfHdd + 40));
            Assert.Equal(2, BitConverter.ToInt32(tableData, fcPlcfHdd + 44));
            Assert.Equal(1, BitConverter.ToInt32(tableData, fcPlcfHdd + 48));
            Assert.Equal(2, BitConverter.ToInt32(tableData, fcPlcfHdd + 52));
        }

        [Fact]
        public void WriteDocBlocks_WithDocumentEndCommentAndStructuredHeaderFieldStory_PreservesCommentAndHeaderStories()
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
            model.Comments.Add(new CommentModel
            {
                Id = "0",
                Text = "Tail",
                StartCp = 5,
                EndCp = 5
            });
            model.Sections.Add(new SectionModel
            {
                DefaultHeaderStory = CreateStructuredHeaderFieldStory(),
                DefaultHeaderText = "Page 1 of 2"
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

            string expectedText = "ABCD\r\x0005Page \x0013PAGE\x00141\x0015 of \x0013SECTIONPAGES\x00142\x0015\r\r\x0005Tail\r";
            var textBytes = new byte[expectedText.Length];
            System.Array.Copy(wordDocData, 1536, textBytes, 0, expectedText.Length);
            var extractedText = Encoding.GetEncoding(1252).GetString(textBytes);

            Assert.Equal(expectedText, extractedText);
            Assert.Equal(6, BitConverter.ToInt32(wordDocData, 64));
            Assert.Equal(19, BitConverter.ToInt32(wordDocData, 72));
            Assert.Equal(6, BitConverter.ToInt32(wordDocData, 76));

            int fcPlcfandRef = BitConverter.ToInt32(wordDocData, 154 + (Fib.CommentReferencePairIndex * 8));
            Assert.Equal(5, BitConverter.ToInt32(tableData, fcPlcfandRef));
            Assert.Equal(6, BitConverter.ToInt32(tableData, fcPlcfandRef + 4));

            int fcPlcffldHdr = BitConverter.ToInt32(wordDocData, 154 + (Fib.HeaderFieldPairIndex * 8));
            int lcbPlcffldHdr = BitConverter.ToInt32(wordDocData, 154 + (Fib.HeaderFieldPairIndex * 8) + 4);
            Assert.NotEqual(0, fcPlcffldHdr);
            Assert.True(lcbPlcffldHdr > 0);
            Assert.Equal(5, BitConverter.ToInt32(tableData, fcPlcffldHdr));
            Assert.Equal(6, BitConverter.ToInt32(tableData, fcPlcffldHdr + 4));
            Assert.Equal(8, BitConverter.ToInt32(tableData, fcPlcffldHdr + 8));
            Assert.Equal(13, BitConverter.ToInt32(tableData, fcPlcffldHdr + 12));
            Assert.Equal(14, BitConverter.ToInt32(tableData, fcPlcffldHdr + 16));
            Assert.Equal(16, BitConverter.ToInt32(tableData, fcPlcffldHdr + 20));
            Assert.Equal(19, BitConverter.ToInt32(tableData, fcPlcffldHdr + 24));
        }

        private static HeaderFooterStoryModel CreateStructuredHeaderFieldStory()
        {
            var pageField = new FieldModel
            {
                Type = FieldType.Page,
                Instruction = "PAGE"
            };

            var sectionPagesField = new FieldModel
            {
                Type = FieldType.SectionPages,
                Instruction = "SECTIONPAGES"
            };

            var story = new HeaderFooterStoryModel();
            var paragraph = new ParagraphModel();
            paragraph.Runs.Add(new RunModel { Text = "Page " });
            paragraph.Runs.Add(new RunModel { Field = pageField, IsFieldBegin = true });
            paragraph.Runs.Add(new RunModel { Field = pageField });
            paragraph.Runs.Add(new RunModel { Field = pageField, IsFieldSeparate = true });
            paragraph.Runs.Add(new RunModel { Text = "1" });
            paragraph.Runs.Add(new RunModel { Field = pageField, IsFieldEnd = true });
            paragraph.Runs.Add(new RunModel { Text = " of " });
            paragraph.Runs.Add(new RunModel { Field = sectionPagesField, IsFieldBegin = true });
            paragraph.Runs.Add(new RunModel { Field = sectionPagesField });
            paragraph.Runs.Add(new RunModel { Field = sectionPagesField, IsFieldSeparate = true });
            paragraph.Runs.Add(new RunModel { Text = "2" });
            paragraph.Runs.Add(new RunModel { Field = sectionPagesField, IsFieldEnd = true });
            story.Paragraphs.Add(paragraph);
            story.Text = "Page 1 of 2";
            return story;
        }

        private static HeaderFooterStoryModel CreateTextHeaderFooterStory(string text)
        {
            var story = new HeaderFooterStoryModel();
            var paragraph = new ParagraphModel();
            paragraph.Runs.Add(new RunModel { Text = text });
            story.Paragraphs.Add(paragraph);
            story.Text = text;
            return story;
        }

        private static HeaderFooterStoryModel CreateStructuredHeaderHyperlinkStory()
        {
            var hyperlink = new HyperlinkModel
            {
                TargetUrl = "https://example.com",
                DisplayText = "Example"
            };

            var story = new HeaderFooterStoryModel();
            var paragraph = new ParagraphModel();
            paragraph.Runs.Add(new RunModel { Text = "Go " });
            paragraph.Runs.Add(new RunModel { Text = "Ex", Hyperlink = hyperlink });
            paragraph.Runs.Add(new RunModel { Text = "ample", Hyperlink = hyperlink });
            paragraph.Runs.Add(new RunModel { Text = " now" });
            story.Paragraphs.Add(paragraph);
            story.Text = "Go Example now";
            return story;
        }

        private static HeaderFooterStoryModel CreateStructuredHeaderFieldAndHyperlinkStory()
        {
            var pageField = new FieldModel
            {
                Type = FieldType.Page,
                Instruction = "PAGE"
            };

            var hyperlink = new HyperlinkModel
            {
                TargetUrl = "https://example.com",
                DisplayText = "Example"
            };

            var story = new HeaderFooterStoryModel();
            var paragraph = new ParagraphModel();
            paragraph.Runs.Add(new RunModel { Text = "Page " });
            paragraph.Runs.Add(new RunModel { Field = pageField, IsFieldBegin = true });
            paragraph.Runs.Add(new RunModel { Field = pageField });
            paragraph.Runs.Add(new RunModel { Field = pageField, IsFieldSeparate = true });
            paragraph.Runs.Add(new RunModel { Text = "1" });
            paragraph.Runs.Add(new RunModel { Field = pageField, IsFieldEnd = true });
            paragraph.Runs.Add(new RunModel { Text = " | " });
            paragraph.Runs.Add(new RunModel { Text = "Ex", Hyperlink = hyperlink });
            paragraph.Runs.Add(new RunModel { Text = "ample", Hyperlink = hyperlink });
            story.Paragraphs.Add(paragraph);
            story.Text = "Page 1 | Example";
            return story;
        }

        private static HeaderFooterStoryModel CreateStructuredHeaderImageStory()
        {
            var story = new HeaderFooterStoryModel();
            var paragraph = new ParagraphModel();
            paragraph.Runs.Add(new RunModel { Text = "Hi" });
            paragraph.Runs.Add(new RunModel
            {
                Image = new ImageModel
                {
                    Data = GetTestPngBytes(),
                    ContentType = "image/png",
                    Width = 96,
                    Height = 96,
                    LayoutType = ImageLayoutType.Inline,
                    WrapType = ImageWrapType.Inline
                }
            });
            paragraph.Runs.Add(new RunModel { Text = "There" });
            story.Paragraphs.Add(paragraph);
            story.Text = "HiThere";
            return story;
        }

        private static HeaderFooterStoryModel CreateStructuredHeaderFieldAndImageStory()
        {
            var pageField = new FieldModel
            {
                Type = FieldType.Page,
                Instruction = "PAGE"
            };

            var story = new HeaderFooterStoryModel();
            var paragraph = new ParagraphModel();
            paragraph.Runs.Add(new RunModel { Text = "Page " });
            paragraph.Runs.Add(new RunModel { Field = pageField, IsFieldBegin = true });
            paragraph.Runs.Add(new RunModel { Field = pageField });
            paragraph.Runs.Add(new RunModel { Field = pageField, IsFieldSeparate = true });
            paragraph.Runs.Add(new RunModel { Text = "1" });
            paragraph.Runs.Add(new RunModel { Field = pageField, IsFieldEnd = true });
            paragraph.Runs.Add(new RunModel { Text = " " });
            paragraph.Runs.Add(new RunModel
            {
                Image = new ImageModel
                {
                    Data = GetTestPngBytes(),
                    ContentType = "image/png",
                    Width = 96,
                    Height = 96,
                    LayoutType = ImageLayoutType.Inline,
                    WrapType = ImageWrapType.Inline
                }
            });
            paragraph.Runs.Add(new RunModel { Text = " tail" });
            story.Paragraphs.Add(paragraph);
            story.Text = "Page 1 tail";
            return story;
        }

        private static HeaderFooterStoryModel CreateStructuredHeaderFieldAndFloatingImageStory()
        {
            var pageField = new FieldModel
            {
                Type = FieldType.Page,
                Instruction = "PAGE"
            };

            var story = new HeaderFooterStoryModel();
            var paragraph = new ParagraphModel();
            paragraph.Runs.Add(new RunModel { Text = "Page " });
            paragraph.Runs.Add(new RunModel { Field = pageField, IsFieldBegin = true });
            paragraph.Runs.Add(new RunModel { Field = pageField });
            paragraph.Runs.Add(new RunModel { Field = pageField, IsFieldSeparate = true });
            paragraph.Runs.Add(new RunModel { Text = "1" });
            paragraph.Runs.Add(new RunModel { Field = pageField, IsFieldEnd = true });
            paragraph.Runs.Add(new RunModel { Text = " " });
            paragraph.Runs.Add(new RunModel
            {
                Image = new ImageModel
                {
                    Data = GetTestPngBytes(),
                    ContentType = "image/png",
                    Width = 96,
                    Height = 96,
                    LayoutType = ImageLayoutType.Floating,
                    WrapType = ImageWrapType.Square
                }
            });
            paragraph.Runs.Add(new RunModel { Text = " tail" });
            story.Paragraphs.Add(paragraph);
            story.Text = "Page 1 tail";
            return story;
        }

        private static HeaderFooterStoryModel CreateStructuredHeaderHyperlinkAndImageStory()
        {
            var hyperlink = new HyperlinkModel
            {
                TargetUrl = "https://example.com",
                DisplayText = "Example"
            };

            var story = new HeaderFooterStoryModel();
            var paragraph = new ParagraphModel();
            paragraph.Runs.Add(new RunModel { Text = "Go " });
            paragraph.Runs.Add(new RunModel { Text = "Example", Hyperlink = hyperlink });
            paragraph.Runs.Add(new RunModel { Text = " " });
            paragraph.Runs.Add(new RunModel
            {
                Image = new ImageModel
                {
                    Data = GetTestPngBytes(),
                    ContentType = "image/png",
                    Width = 96,
                    Height = 96,
                    LayoutType = ImageLayoutType.Inline,
                    WrapType = ImageWrapType.Inline
                }
            });
            paragraph.Runs.Add(new RunModel { Text = " now" });
            story.Paragraphs.Add(paragraph);
            story.Text = "Go Example now";
            return story;
        }

        private static HeaderFooterStoryModel CreateStructuredHeaderHyperlinkAndFloatingImageStory()
        {
            var hyperlink = new HyperlinkModel
            {
                TargetUrl = "https://example.com",
                DisplayText = "Example"
            };

            var story = new HeaderFooterStoryModel();
            var paragraph = new ParagraphModel();
            paragraph.Runs.Add(new RunModel { Text = "Go " });
            paragraph.Runs.Add(new RunModel { Text = "Example", Hyperlink = hyperlink });
            paragraph.Runs.Add(new RunModel { Text = " " });
            paragraph.Runs.Add(new RunModel
            {
                Image = new ImageModel
                {
                    Data = GetTestPngBytes(),
                    ContentType = "image/png",
                    Width = 96,
                    Height = 96,
                    LayoutType = ImageLayoutType.Floating,
                    WrapType = ImageWrapType.Square
                }
            });
            paragraph.Runs.Add(new RunModel { Text = " now" });
            story.Paragraphs.Add(paragraph);
            story.Text = "Go Example now";
            return story;
        }

        private static HeaderFooterStoryModel CreateStructuredHeaderFloatingImageStory()
        {
            var story = new HeaderFooterStoryModel();
            var paragraph = new ParagraphModel();
            paragraph.Runs.Add(new RunModel { Text = "Hi" });
            paragraph.Runs.Add(new RunModel
            {
                Image = new ImageModel
                {
                    Data = GetTestPngBytes(),
                    ContentType = "image/png",
                    Width = 96,
                    Height = 96,
                    LayoutType = ImageLayoutType.Floating,
                    WrapType = ImageWrapType.Square
                }
            });
            paragraph.Runs.Add(new RunModel { Text = "There" });
            story.Paragraphs.Add(paragraph);
            story.Text = "HiThere";
            return story;
        }

        private static HeaderFooterStoryModel CreateHeaderAlignedFloatingImageStory(string text, string horizontalAlignment)
        {
            var story = new HeaderFooterStoryModel();
            var paragraph = new ParagraphModel();
            paragraph.Runs.Add(new RunModel { Text = text });
            paragraph.Runs.Add(new RunModel
            {
                Image = new ImageModel
                {
                    Data = GetTestPngBytes(),
                    ContentType = "image/png",
                    Width = 96,
                    Height = 96,
                    LayoutType = ImageLayoutType.Floating,
                    WrapType = ImageWrapType.Square,
                    HorizontalRelativeTo = "margin",
                    VerticalRelativeTo = "margin",
                    HorizontalAlignment = horizontalAlignment,
                    VerticalAlignment = "top"
                }
            });
            story.Paragraphs.Add(paragraph);
            story.Text = text;
            return story;
        }

        private static HeaderFooterStoryModel CreateHeaderRelativeFloatingImageStory(string text, string horizontalRelativeTo)
        {
            var story = new HeaderFooterStoryModel();
            var paragraph = new ParagraphModel();
            paragraph.Runs.Add(new RunModel { Text = text });
            paragraph.Runs.Add(new RunModel
            {
                Image = new ImageModel
                {
                    Data = GetTestPngBytes(),
                    ContentType = "image/png",
                    Width = 96,
                    Height = 96,
                    LayoutType = ImageLayoutType.Floating,
                    WrapType = ImageWrapType.Square,
                    HorizontalRelativeTo = horizontalRelativeTo,
                    VerticalRelativeTo = "margin",
                    PositionXTwips = 0,
                    PositionYTwips = 0
                }
            });
            story.Paragraphs.Add(paragraph);
            story.Text = text;
            return story;
        }

        private static (int firstLeft, int secondLeft) GetFirstTwoHeaderStoryShapeLeftPositions(DocumentModel model)
        {
            var writer = new DocWriter();
            using var ms = new MemoryStream();
            writer.WriteDocBlocks(model, ms);
            ms.Position = 0;

            using var compoundFile = new OpenMcdf.CompoundFile(ms);
            Assert.True(compoundFile.RootStorage.TryGetStream("WordDocument", out var wordDocStream));
            Assert.True(compoundFile.RootStorage.TryGetStream("1Table", out var tableStream));

            var wordDocData = wordDocStream.GetData();
            var tableData = tableStream.GetData();
            int fibPairOffset = 154 + (Fib.HeaderShapePairIndex * 8);
            int fcPlcSpaHdr = BitConverter.ToInt32(wordDocData, fibPairOffset);
            int lcbPlcSpaHdr = BitConverter.ToInt32(wordDocData, fibPairOffset + 4);
            Assert.NotEqual(0, fcPlcSpaHdr);
            Assert.True(lcbPlcSpaHdr >= 64);

            int pictureCount = (lcbPlcSpaHdr - 4) / 30;
            Assert.True(pictureCount >= 2);

            int dataStart = fcPlcSpaHdr + ((pictureCount + 1) * 4);
            int firstLeft = BitConverter.ToInt32(tableData, dataStart + 4);
            int secondLeft = BitConverter.ToInt32(tableData, dataStart + 26 + 4);
            return (firstLeft, secondLeft);
        }

        private static HeaderFooterStoryModel CreateStructuredHeaderParagraphAndTableStory()
        {
            var story = new HeaderFooterStoryModel();

            var paragraph = new ParagraphModel();
            paragraph.Runs.Add(new RunModel { Text = "Head" });
            story.Paragraphs.Add(paragraph);
            story.Content.Add(paragraph);

            var table = new TableModel();
            table.GridColumnWidths.Add(2400);

            var row = new TableRowModel();
            var cell = new TableCellModel
            {
                Width = 2400
            };

            var cellParagraph = new ParagraphModel();
            cellParagraph.Runs.Add(new RunModel { Text = "Cell" });
            cell.Paragraphs.Add(cellParagraph);
            row.Cells.Add(cell);
            table.Rows.Add(row);

            story.Content.Add(table);
            story.Text = "Head\rCell";
            return story;
        }

        private static HeaderFooterStoryModel CreateStructuredHeaderTableWithRowFlagsStory()
        {
            var story = new HeaderFooterStoryModel();
            var table = new TableModel();
            table.GridColumnWidths.Add(2400);

            var row = new TableRowModel
            {
                IsHeader = true,
                CannotSplit = true
            };
            var cell = new TableCellModel { Width = 2400 };
            var cellParagraph = new ParagraphModel();
            cellParagraph.Runs.Add(new RunModel { Text = "Cell" });
            cell.Paragraphs.Add(cellParagraph);
            row.Cells.Add(cell);
            table.Rows.Add(row);

            story.Content.Add(table);
            story.Text = "Cell";
            return story;
        }

        private static int GetHeaderFloatingImageTopWithTableRowHeight(int rowHeightTwips, TableRowHeightRule heightRule)
        {
            var story = new HeaderFooterStoryModel();

            var lead = new ParagraphModel();
            lead.Runs.Add(new RunModel { Text = "Lead" });
            story.Content.Add(lead);
            story.Paragraphs.Add(lead);

            var table = new TableModel();
            table.GridColumnWidths.Add(2400);
            var row = new TableRowModel { HeightTwips = rowHeightTwips, HeightRule = heightRule };
            var cell = new TableCellModel { Width = 2400 };
            var cellParagraph = new ParagraphModel();
            cellParagraph.Runs.Add(new RunModel { Text = "Cell" });
            cell.Paragraphs.Add(cellParagraph);
            row.Cells.Add(cell);
            table.Rows.Add(row);
            story.Content.Add(table);

            var tail = new ParagraphModel();
            tail.Runs.Add(new RunModel
            {
                Image = new ImageModel
                {
                    Data = GetTestPngBytes(),
                    ContentType = "image/png",
                    Width = 96,
                    Height = 96,
                    LayoutType = ImageLayoutType.Floating,
                    WrapType = ImageWrapType.Square,
                    VerticalRelativeTo = "paragraph",
                    PositionYTwips = 0
                }
            });
            story.Content.Add(tail);
            story.Paragraphs.Add(tail);

            var model = new DocumentModel();
            model.Content.Add(new ParagraphModel
            {
                Runs =
                {
                    new RunModel { Text = "Body" }
                }
            });
            model.Sections.Add(new SectionModel
            {
                DefaultHeaderStory = story
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
            int fcPlcSpaHdr = BitConverter.ToInt32(wordDocData, 154 + (Fib.HeaderShapePairIndex * 8));
            Assert.NotEqual(0, fcPlcSpaHdr);

            return BitConverter.ToInt32(tableData, fcPlcSpaHdr + 16);
        }

        private static int GetHeaderFloatingImageTopWithTableCellPadding(int topPaddingTwips, int bottomPaddingTwips)
        {
            var story = new HeaderFooterStoryModel();

            var lead = new ParagraphModel();
            lead.Runs.Add(new RunModel { Text = "Lead" });
            story.Content.Add(lead);
            story.Paragraphs.Add(lead);

            var table = new TableModel
            {
                DefaultCellPaddingTopTwips = topPaddingTwips,
                DefaultCellPaddingBottomTwips = bottomPaddingTwips
            };
            table.GridColumnWidths.Add(2400);
            var row = new TableRowModel();
            var cell = new TableCellModel { Width = 2400 };
            var cellParagraph = new ParagraphModel();
            cellParagraph.Runs.Add(new RunModel { Text = "Cell" });
            cell.Paragraphs.Add(cellParagraph);
            row.Cells.Add(cell);
            table.Rows.Add(row);
            story.Content.Add(table);

            var tail = new ParagraphModel();
            tail.Runs.Add(new RunModel
            {
                Image = new ImageModel
                {
                    Data = GetTestPngBytes(),
                    ContentType = "image/png",
                    Width = 96,
                    Height = 96,
                    LayoutType = ImageLayoutType.Floating,
                    WrapType = ImageWrapType.Square,
                    VerticalRelativeTo = "paragraph",
                    PositionYTwips = 0
                }
            });
            story.Content.Add(tail);
            story.Paragraphs.Add(tail);

            var model = new DocumentModel();
            model.Content.Add(new ParagraphModel
            {
                Runs =
                {
                    new RunModel { Text = "Body" }
                }
            });
            model.Sections.Add(new SectionModel
            {
                DefaultHeaderStory = story
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
            int fcPlcSpaHdr = BitConverter.ToInt32(wordDocData, 154 + (Fib.HeaderShapePairIndex * 8));
            Assert.NotEqual(0, fcPlcSpaHdr);

            return BitConverter.ToInt32(tableData, fcPlcSpaHdr + 16);
        }

        private static int GetHeaderTableCellFloatingImageTopWithAlignment(TableCellVerticalAlignment alignment)
        {
            var story = new HeaderFooterStoryModel();

            var table = new TableModel();
            table.GridColumnWidths.Add(2400);
            var row = new TableRowModel { HeightTwips = 2800, HeightRule = TableRowHeightRule.AtLeast };
            var cell = new TableCellModel
            {
                Width = 2400,
                VerticalAlignment = alignment
            };
            var cellParagraph = new ParagraphModel();
            cellParagraph.Runs.Add(new RunModel
            {
                Image = new ImageModel
                {
                    Data = GetTestPngBytes(),
                    ContentType = "image/png",
                    Width = 96,
                    Height = 96,
                    LayoutType = ImageLayoutType.Floating,
                    WrapType = ImageWrapType.Square,
                    VerticalRelativeTo = "paragraph",
                    PositionYTwips = 0
                }
            });
            cell.Paragraphs.Add(cellParagraph);
            row.Cells.Add(cell);
            table.Rows.Add(row);
            story.Content.Add(table);
            story.Paragraphs.Add(cellParagraph);

            var model = new DocumentModel();
            model.Content.Add(new ParagraphModel
            {
                Runs =
                {
                    new RunModel { Text = "Body" }
                }
            });
            model.Sections.Add(new SectionModel
            {
                DefaultHeaderStory = story
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
            int fcPlcSpaHdr = BitConverter.ToInt32(wordDocData, 154 + (Fib.HeaderShapePairIndex * 8));
            Assert.NotEqual(0, fcPlcSpaHdr);

            return BitConverter.ToInt32(tableData, fcPlcSpaHdr + 16);
        }

        private static int GetHeaderSecondParagraphFloatingImageTopWithRowRule(int rowHeightTwips, TableRowHeightRule heightRule, string? imageVerticalAlignment, int firstParagraphSpacingAfterTwips = 0, int secondParagraphSpacingBeforeTwips = 0, int secondParagraphPositionYTwips = 0)
        {
            var story = new HeaderFooterStoryModel();

            var table = new TableModel();
            table.GridColumnWidths.Add(2400);
            var row = new TableRowModel { HeightTwips = rowHeightTwips, HeightRule = heightRule };
            var cell = new TableCellModel
            {
                Width = 2400,
                VerticalAlignment = TableCellVerticalAlignment.Top
            };

            var firstParagraph = new ParagraphModel();
            firstParagraph.Properties.SpacingAfterTwips = firstParagraphSpacingAfterTwips;
            firstParagraph.Runs.Add(new RunModel
            {
                Text = "Tall intro line for clipping behavior",
                Properties =
                {
                    FontSize = 72
                }
            });

            var secondParagraph = new ParagraphModel();
            secondParagraph.Properties.SpacingBeforeTwips = secondParagraphSpacingBeforeTwips;
            secondParagraph.Runs.Add(new RunModel
            {
                Image = new ImageModel
                {
                    Data = GetTestPngBytes(),
                    ContentType = "image/png",
                    Width = 96,
                    Height = 96,
                    LayoutType = ImageLayoutType.Floating,
                    WrapType = ImageWrapType.Square,
                    VerticalRelativeTo = "paragraph",
                    VerticalAlignment = imageVerticalAlignment,
                    PositionYTwips = secondParagraphPositionYTwips
                }
            });

            cell.Content.Add(firstParagraph);
            cell.Content.Add(secondParagraph);
            cell.Paragraphs.Add(firstParagraph);
            cell.Paragraphs.Add(secondParagraph);
            row.Cells.Add(cell);
            table.Rows.Add(row);
            story.Content.Add(table);
            story.Paragraphs.Add(firstParagraph);
            story.Paragraphs.Add(secondParagraph);

            var model = new DocumentModel();
            model.Content.Add(new ParagraphModel
            {
                Runs =
                {
                    new RunModel { Text = "Body" }
                }
            });
            model.Sections.Add(new SectionModel
            {
                DefaultHeaderStory = story
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
            int fcPlcSpaHdr = BitConverter.ToInt32(wordDocData, 154 + (Fib.HeaderShapePairIndex * 8));
            Assert.NotEqual(0, fcPlcSpaHdr);

            return BitConverter.ToInt32(tableData, fcPlcSpaHdr + 16);
        }

        private static int GetHeaderHyperlinkFloatingImageTopWithRowRule(int rowHeightTwips, TableRowHeightRule heightRule, int hyperlinkImagePositionYTwips)
        {
            var story = new HeaderFooterStoryModel();

            var table = new TableModel();
            table.GridColumnWidths.Add(2400);
            var row = new TableRowModel { HeightTwips = rowHeightTwips, HeightRule = heightRule };
            var cell = new TableCellModel
            {
                Width = 2400,
                VerticalAlignment = TableCellVerticalAlignment.Top
            };

            var firstParagraph = new ParagraphModel();
            firstParagraph.Runs.Add(new RunModel
            {
                Text = "Tall intro line for hyperlink clipping behavior",
                Properties =
                {
                    FontSize = 72
                }
            });

            var hyperlink = new HyperlinkModel
            {
                TargetUrl = "https://example.com/image",
                DisplayText = "ImageLink"
            };

            var secondParagraph = new ParagraphModel();
            secondParagraph.Runs.Add(new RunModel
            {
                Hyperlink = hyperlink,
                Image = new ImageModel
                {
                    Data = GetTestPngBytes(),
                    ContentType = "image/png",
                    Width = 96,
                    Height = 96,
                    LayoutType = ImageLayoutType.Floating,
                    WrapType = ImageWrapType.Square,
                    VerticalRelativeTo = "paragraph",
                    PositionYTwips = hyperlinkImagePositionYTwips
                }
            });

            cell.Content.Add(firstParagraph);
            cell.Content.Add(secondParagraph);
            cell.Paragraphs.Add(firstParagraph);
            cell.Paragraphs.Add(secondParagraph);
            row.Cells.Add(cell);
            table.Rows.Add(row);
            story.Content.Add(table);
            story.Paragraphs.Add(firstParagraph);
            story.Paragraphs.Add(secondParagraph);

            var model = new DocumentModel();
            model.Content.Add(new ParagraphModel
            {
                Runs =
                {
                    new RunModel { Text = "Body" }
                }
            });
            model.Sections.Add(new SectionModel
            {
                DefaultHeaderStory = story
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
            int fcPlcSpaHdr = BitConverter.ToInt32(wordDocData, 154 + (Fib.HeaderShapePairIndex * 8));
            Assert.NotEqual(0, fcPlcSpaHdr);

            return BitConverter.ToInt32(tableData, fcPlcSpaHdr + 16);
        }

        private static (int left, int top, int right, int bottom) GetHeaderFloatingImageBoundsWithDistance(
            int distanceLeftTwips,
            int distanceTopTwips,
            int distanceRightTwips,
            int distanceBottomTwips,
            ImageWrapType wrapType = ImageWrapType.Square,
            string verticalRelativeTo = "paragraph",
            bool behindText = false,
            int imageWidthPixels = 96,
            int imageHeightPixels = 48)
        {
            var story = new HeaderFooterStoryModel();
            var paragraph = new ParagraphModel();
            paragraph.Runs.Add(new RunModel
            {
                Text = "Header"
            });
            paragraph.Runs.Add(new RunModel
            {
                Image = new ImageModel
                {
                    Data = GetTestPngBytes(),
                    ContentType = "image/png",
                    Width = imageWidthPixels,
                    Height = imageHeightPixels,
                    LayoutType = ImageLayoutType.Floating,
                    WrapType = wrapType,
                    BehindText = behindText,
                    HorizontalRelativeTo = "margin",
                    VerticalRelativeTo = verticalRelativeTo,
                    PositionXTwips = 0,
                    PositionYTwips = 0,
                    DistanceLeftTwips = distanceLeftTwips,
                    DistanceTopTwips = distanceTopTwips,
                    DistanceRightTwips = distanceRightTwips,
                    DistanceBottomTwips = distanceBottomTwips
                }
            });

            story.Content.Add(paragraph);
            story.Paragraphs.Add(paragraph);

            var model = new DocumentModel();
            model.Content.Add(new ParagraphModel { Runs = { new RunModel { Text = "Body" } } });
            model.Sections.Add(new SectionModel
            {
                DefaultHeaderStory = story
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
            int fcPlcSpaHdr = BitConverter.ToInt32(wordDocData, 154 + (Fib.HeaderShapePairIndex * 8));
            Assert.NotEqual(0, fcPlcSpaHdr);

            int recordOffset = fcPlcSpaHdr;
            int left = BitConverter.ToInt32(tableData, recordOffset + 12);
            int top = BitConverter.ToInt32(tableData, recordOffset + 16);
            int right = BitConverter.ToInt32(tableData, recordOffset + 20);
            int bottom = BitConverter.ToInt32(tableData, recordOffset + 24);
            return (left, top, right, bottom);
        }

        private static int GetHeaderFloatingImageTopWithVerticalRelativeTo(string verticalRelativeTo)
        {
            var story = new HeaderFooterStoryModel();
            var paragraph = new ParagraphModel();
            paragraph.Properties.LineSpacing = 1200;
            paragraph.Properties.LineSpacingRule = "exact";
            paragraph.Runs.Add(new RunModel
            {
                Text = "This is a long header paragraph that should wrap into multiple lines to make line-relative and paragraph-relative anchors diverge.",
                Properties =
                {
                    FontSize = 24
                }
            });
            paragraph.Runs.Add(new RunModel
            {
                Image = new ImageModel
                {
                    Data = GetTestPngBytes(),
                    ContentType = "image/png",
                    Width = 16,
                    Height = 8,
                    LayoutType = ImageLayoutType.Floating,
                    VerticalAlignment = "bottom",
                    VerticalRelativeTo = verticalRelativeTo,
                    PositionYTwips = 0
                }
            });

            story.Content.Add(paragraph);
            story.Paragraphs.Add(paragraph);

            var model = new DocumentModel();
            model.Content.Add(new ParagraphModel { Runs = { new RunModel { Text = "Body" } } });
            model.Sections.Add(new SectionModel
            {
                DefaultHeaderStory = story
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
            int fcPlcSpaHdr = BitConverter.ToInt32(wordDocData, 154 + (Fib.HeaderShapePairIndex * 8));
            Assert.NotEqual(0, fcPlcSpaHdr);

            return BitConverter.ToInt32(tableData, fcPlcSpaHdr + 16);
        }

        private static (int left, int top, int right, int bottom) GetFooterFloatingImageBoundsWithDistance(
            int distanceLeftTwips,
            int distanceTopTwips,
            int distanceRightTwips,
            int distanceBottomTwips,
            ImageWrapType wrapType = ImageWrapType.Square,
            string verticalRelativeTo = "paragraph",
            bool behindText = false,
            int imageWidthPixels = 96,
            int imageHeightPixels = 48)
        {
            var story = new HeaderFooterStoryModel();
            var paragraph = new ParagraphModel();
            paragraph.Runs.Add(new RunModel
            {
                Text = "Footer"
            });
            paragraph.Runs.Add(new RunModel
            {
                Image = new ImageModel
                {
                    Data = GetTestPngBytes(),
                    ContentType = "image/png",
                    Width = imageWidthPixels,
                    Height = imageHeightPixels,
                    LayoutType = ImageLayoutType.Floating,
                    WrapType = wrapType,
                    BehindText = behindText,
                    HorizontalRelativeTo = "margin",
                    VerticalRelativeTo = verticalRelativeTo,
                    PositionXTwips = 0,
                    PositionYTwips = 0,
                    DistanceLeftTwips = distanceLeftTwips,
                    DistanceTopTwips = distanceTopTwips,
                    DistanceRightTwips = distanceRightTwips,
                    DistanceBottomTwips = distanceBottomTwips
                }
            });

            story.Content.Add(paragraph);
            story.Paragraphs.Add(paragraph);

            var model = new DocumentModel();
            model.Content.Add(new ParagraphModel { Runs = { new RunModel { Text = "Body" } } });
            model.Sections.Add(new SectionModel
            {
                DefaultFooterStory = story
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
            int fcPlcSpaHdr = BitConverter.ToInt32(wordDocData, 154 + (Fib.HeaderShapePairIndex * 8));
            Assert.NotEqual(0, fcPlcSpaHdr);

            int recordOffset = fcPlcSpaHdr;
            int left = BitConverter.ToInt32(tableData, recordOffset + 12);
            int top = BitConverter.ToInt32(tableData, recordOffset + 16);
            int right = BitConverter.ToInt32(tableData, recordOffset + 20);
            int bottom = BitConverter.ToInt32(tableData, recordOffset + 24);
            return (left, top, right, bottom);
        }

        private static int GetFooterFloatingImageTopWithVerticalRelativeTo(string verticalRelativeTo)
        {
            var story = new HeaderFooterStoryModel();
            var paragraph = new ParagraphModel();
            paragraph.Properties.LineSpacing = 1200;
            paragraph.Properties.LineSpacingRule = "exact";
            paragraph.Runs.Add(new RunModel
            {
                Text = "This is a long footer paragraph that should wrap into multiple lines to make line-relative and paragraph-relative anchors diverge.",
                Properties =
                {
                    FontSize = 24
                }
            });
            paragraph.Runs.Add(new RunModel
            {
                Image = new ImageModel
                {
                    Data = GetTestPngBytes(),
                    ContentType = "image/png",
                    Width = 16,
                    Height = 8,
                    LayoutType = ImageLayoutType.Floating,
                    VerticalAlignment = "bottom",
                    VerticalRelativeTo = verticalRelativeTo,
                    PositionYTwips = 0
                }
            });

            story.Content.Add(paragraph);
            story.Paragraphs.Add(paragraph);

            var model = new DocumentModel();
            model.Content.Add(new ParagraphModel { Runs = { new RunModel { Text = "Body" } } });
            model.Sections.Add(new SectionModel
            {
                DefaultFooterStory = story
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
            int fcPlcSpaHdr = BitConverter.ToInt32(wordDocData, 154 + (Fib.HeaderShapePairIndex * 8));
            Assert.NotEqual(0, fcPlcSpaHdr);

            return BitConverter.ToInt32(tableData, fcPlcSpaHdr + 16);
        }

        private static int GetFooterSecondParagraphFloatingImageTopWithRowRule(int rowHeightTwips, TableRowHeightRule heightRule, string? imageVerticalAlignment, int firstParagraphSpacingAfterTwips = 0, int secondParagraphSpacingBeforeTwips = 0, int secondParagraphPositionYTwips = 0)
        {
            var story = new HeaderFooterStoryModel();

            var table = new TableModel();
            table.GridColumnWidths.Add(2400);
            var row = new TableRowModel { HeightTwips = rowHeightTwips, HeightRule = heightRule };
            var cell = new TableCellModel
            {
                Width = 2400,
                VerticalAlignment = TableCellVerticalAlignment.Top
            };

            var firstParagraph = new ParagraphModel();
            firstParagraph.Properties.SpacingAfterTwips = firstParagraphSpacingAfterTwips;
            firstParagraph.Runs.Add(new RunModel
            {
                Text = "Tall intro line for clipping behavior",
                Properties =
                {
                    FontSize = 72
                }
            });

            var secondParagraph = new ParagraphModel();
            secondParagraph.Properties.SpacingBeforeTwips = secondParagraphSpacingBeforeTwips;
            secondParagraph.Runs.Add(new RunModel
            {
                Image = new ImageModel
                {
                    Data = GetTestPngBytes(),
                    ContentType = "image/png",
                    Width = 96,
                    Height = 96,
                    LayoutType = ImageLayoutType.Floating,
                    WrapType = ImageWrapType.Square,
                    VerticalRelativeTo = "paragraph",
                    VerticalAlignment = imageVerticalAlignment,
                    PositionYTwips = secondParagraphPositionYTwips
                }
            });

            cell.Content.Add(firstParagraph);
            cell.Content.Add(secondParagraph);
            cell.Paragraphs.Add(firstParagraph);
            cell.Paragraphs.Add(secondParagraph);
            row.Cells.Add(cell);
            table.Rows.Add(row);
            story.Content.Add(table);
            story.Paragraphs.Add(firstParagraph);
            story.Paragraphs.Add(secondParagraph);

            var model = new DocumentModel();
            model.Content.Add(new ParagraphModel
            {
                Runs =
                {
                    new RunModel { Text = "Body" }
                }
            });
            model.Sections.Add(new SectionModel
            {
                DefaultFooterStory = story
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
            int fcPlcSpaHdr = BitConverter.ToInt32(wordDocData, 154 + (Fib.HeaderShapePairIndex * 8));
            Assert.NotEqual(0, fcPlcSpaHdr);

            return BitConverter.ToInt32(tableData, fcPlcSpaHdr + 16);
        }

        private static int GetFooterHyperlinkFloatingImageTopWithRowRule(int rowHeightTwips, TableRowHeightRule heightRule, int hyperlinkImagePositionYTwips)
        {
            var story = new HeaderFooterStoryModel();

            var table = new TableModel();
            table.GridColumnWidths.Add(2400);
            var row = new TableRowModel { HeightTwips = rowHeightTwips, HeightRule = heightRule };
            var cell = new TableCellModel
            {
                Width = 2400,
                VerticalAlignment = TableCellVerticalAlignment.Top
            };

            var firstParagraph = new ParagraphModel();
            firstParagraph.Runs.Add(new RunModel
            {
                Text = "Tall intro line for hyperlink clipping behavior",
                Properties =
                {
                    FontSize = 72
                }
            });

            var hyperlink = new HyperlinkModel
            {
                TargetUrl = "https://example.com/image",
                DisplayText = "ImageLink"
            };

            var secondParagraph = new ParagraphModel();
            secondParagraph.Runs.Add(new RunModel
            {
                Hyperlink = hyperlink,
                Image = new ImageModel
                {
                    Data = GetTestPngBytes(),
                    ContentType = "image/png",
                    Width = 96,
                    Height = 96,
                    LayoutType = ImageLayoutType.Floating,
                    WrapType = ImageWrapType.Square,
                    VerticalRelativeTo = "paragraph",
                    PositionYTwips = hyperlinkImagePositionYTwips
                }
            });

            cell.Content.Add(firstParagraph);
            cell.Content.Add(secondParagraph);
            cell.Paragraphs.Add(firstParagraph);
            cell.Paragraphs.Add(secondParagraph);
            row.Cells.Add(cell);
            table.Rows.Add(row);
            story.Content.Add(table);
            story.Paragraphs.Add(firstParagraph);
            story.Paragraphs.Add(secondParagraph);

            var model = new DocumentModel();
            model.Content.Add(new ParagraphModel
            {
                Runs =
                {
                    new RunModel { Text = "Body" }
                }
            });
            model.Sections.Add(new SectionModel
            {
                DefaultFooterStory = story
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
            int fcPlcSpaHdr = BitConverter.ToInt32(wordDocData, 154 + (Fib.HeaderShapePairIndex * 8));
            Assert.NotEqual(0, fcPlcSpaHdr);

            return BitConverter.ToInt32(tableData, fcPlcSpaHdr + 16);
        }

        private static int GetFooterNestedTableFloatingImageTopWithOuterRowRule(
            int outerRowHeightTwips,
            TableRowHeightRule outerHeightRule,
            int firstNestedParagraphSpacingAfterTwips,
            int secondNestedParagraphSpacingBeforeTwips = 0,
            int secondNestedParagraphPositionYTwips = 0)
        {
            var story = new HeaderFooterStoryModel();

            var outerTable = new TableModel();
            outerTable.GridColumnWidths.Add(2400);
            var outerRow = new TableRowModel { HeightTwips = outerRowHeightTwips, HeightRule = outerHeightRule };
            var outerCell = new TableCellModel { Width = 2400, VerticalAlignment = TableCellVerticalAlignment.Top };

            var nestedTable = new TableModel();
            nestedTable.GridColumnWidths.Add(2200);
            var nestedRow = new TableRowModel();
            var nestedCell = new TableCellModel { Width = 2200 };

            var nestedFirstParagraph = new ParagraphModel();
            nestedFirstParagraph.Properties.SpacingAfterTwips = firstNestedParagraphSpacingAfterTwips;
            nestedFirstParagraph.Runs.Add(new RunModel
            {
                Text = "Tall nested intro line for clipping behavior",
                Properties =
                {
                    FontSize = 72
                }
            });

            var nestedSecondParagraph = new ParagraphModel();
            nestedSecondParagraph.Properties.SpacingBeforeTwips = secondNestedParagraphSpacingBeforeTwips;
            nestedSecondParagraph.Runs.Add(new RunModel
            {
                Image = new ImageModel
                {
                    Data = GetTestPngBytes(),
                    ContentType = "image/png",
                    Width = 96,
                    Height = 96,
                    LayoutType = ImageLayoutType.Floating,
                    WrapType = ImageWrapType.Square,
                    VerticalRelativeTo = "paragraph",
                    PositionYTwips = secondNestedParagraphPositionYTwips
                }
            });

            nestedCell.Content.Add(nestedFirstParagraph);
            nestedCell.Content.Add(nestedSecondParagraph);
            nestedCell.Paragraphs.Add(nestedFirstParagraph);
            nestedCell.Paragraphs.Add(nestedSecondParagraph);
            nestedRow.Cells.Add(nestedCell);
            nestedTable.Rows.Add(nestedRow);

            outerCell.Content.Add(nestedTable);
            outerRow.Cells.Add(outerCell);
            outerTable.Rows.Add(outerRow);

            story.Content.Add(outerTable);
            story.Paragraphs.Add(nestedFirstParagraph);
            story.Paragraphs.Add(nestedSecondParagraph);

            var model = new DocumentModel();
            model.Content.Add(new ParagraphModel
            {
                Runs =
                {
                    new RunModel { Text = "Body" }
                }
            });
            model.Sections.Add(new SectionModel
            {
                DefaultFooterStory = story
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
            int fcPlcSpaHdr = BitConverter.ToInt32(wordDocData, 154 + (Fib.HeaderShapePairIndex * 8));
            Assert.NotEqual(0, fcPlcSpaHdr);

            return BitConverter.ToInt32(tableData, fcPlcSpaHdr + 16);
        }

        private static int GetFooterNestedHyperlinkFloatingImageTopWithOuterRowRule(int outerRowHeightTwips, TableRowHeightRule outerHeightRule, int hyperlinkImagePositionYTwips)
        {
            var story = new HeaderFooterStoryModel();

            var outerTable = new TableModel();
            outerTable.GridColumnWidths.Add(2400);
            var outerRow = new TableRowModel { HeightTwips = outerRowHeightTwips, HeightRule = outerHeightRule };
            var outerCell = new TableCellModel { Width = 2400, VerticalAlignment = TableCellVerticalAlignment.Top };

            var nestedTable = new TableModel();
            nestedTable.GridColumnWidths.Add(2200);
            var nestedRow = new TableRowModel();
            var nestedCell = new TableCellModel { Width = 2200 };

            var nestedFirstParagraph = new ParagraphModel();
            nestedFirstParagraph.Runs.Add(new RunModel
            {
                Text = "Tall nested intro line for hyperlink clipping behavior",
                Properties =
                {
                    FontSize = 72
                }
            });

            var hyperlink = new HyperlinkModel
            {
                TargetUrl = "https://example.com/nested-image",
                DisplayText = "NestedImageLink"
            };

            var nestedSecondParagraph = new ParagraphModel();
            nestedSecondParagraph.Runs.Add(new RunModel
            {
                Hyperlink = hyperlink,
                Image = new ImageModel
                {
                    Data = GetTestPngBytes(),
                    ContentType = "image/png",
                    Width = 96,
                    Height = 96,
                    LayoutType = ImageLayoutType.Floating,
                    WrapType = ImageWrapType.Square,
                    VerticalRelativeTo = "paragraph",
                    PositionYTwips = hyperlinkImagePositionYTwips
                }
            });

            nestedCell.Content.Add(nestedFirstParagraph);
            nestedCell.Content.Add(nestedSecondParagraph);
            nestedCell.Paragraphs.Add(nestedFirstParagraph);
            nestedCell.Paragraphs.Add(nestedSecondParagraph);
            nestedRow.Cells.Add(nestedCell);
            nestedTable.Rows.Add(nestedRow);

            outerCell.Content.Add(nestedTable);
            outerRow.Cells.Add(outerCell);
            outerTable.Rows.Add(outerRow);

            story.Content.Add(outerTable);
            story.Paragraphs.Add(nestedFirstParagraph);
            story.Paragraphs.Add(nestedSecondParagraph);

            var model = new DocumentModel();
            model.Content.Add(new ParagraphModel
            {
                Runs =
                {
                    new RunModel { Text = "Body" }
                }
            });
            model.Sections.Add(new SectionModel
            {
                DefaultFooterStory = story
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
            int fcPlcSpaHdr = BitConverter.ToInt32(wordDocData, 154 + (Fib.HeaderShapePairIndex * 8));
            Assert.NotEqual(0, fcPlcSpaHdr);

            return BitConverter.ToInt32(tableData, fcPlcSpaHdr + 16);
        }

        private static int GetHeaderNestedTableFloatingImageTopWithOuterRowRule(
            int outerRowHeightTwips,
            TableRowHeightRule outerHeightRule,
            int firstNestedParagraphSpacingAfterTwips,
            int secondNestedParagraphSpacingBeforeTwips = 0,
            int secondNestedParagraphPositionYTwips = 0)
        {
            var story = new HeaderFooterStoryModel();

            var outerTable = new TableModel();
            outerTable.GridColumnWidths.Add(2400);
            var outerRow = new TableRowModel { HeightTwips = outerRowHeightTwips, HeightRule = outerHeightRule };
            var outerCell = new TableCellModel { Width = 2400, VerticalAlignment = TableCellVerticalAlignment.Top };

            var nestedTable = new TableModel();
            nestedTable.GridColumnWidths.Add(2200);
            var nestedRow = new TableRowModel();
            var nestedCell = new TableCellModel { Width = 2200 };

            var nestedFirstParagraph = new ParagraphModel();
            nestedFirstParagraph.Properties.SpacingAfterTwips = firstNestedParagraphSpacingAfterTwips;
            nestedFirstParagraph.Runs.Add(new RunModel
            {
                Text = "Tall nested intro line for clipping behavior",
                Properties =
                {
                    FontSize = 72
                }
            });

            var nestedSecondParagraph = new ParagraphModel();
            nestedSecondParagraph.Properties.SpacingBeforeTwips = secondNestedParagraphSpacingBeforeTwips;
            nestedSecondParagraph.Runs.Add(new RunModel
            {
                Image = new ImageModel
                {
                    Data = GetTestPngBytes(),
                    ContentType = "image/png",
                    Width = 96,
                    Height = 96,
                    LayoutType = ImageLayoutType.Floating,
                    WrapType = ImageWrapType.Square,
                    VerticalRelativeTo = "paragraph",
                    PositionYTwips = secondNestedParagraphPositionYTwips
                }
            });

            nestedCell.Content.Add(nestedFirstParagraph);
            nestedCell.Content.Add(nestedSecondParagraph);
            nestedCell.Paragraphs.Add(nestedFirstParagraph);
            nestedCell.Paragraphs.Add(nestedSecondParagraph);
            nestedRow.Cells.Add(nestedCell);
            nestedTable.Rows.Add(nestedRow);

            outerCell.Content.Add(nestedTable);
            outerRow.Cells.Add(outerCell);
            outerTable.Rows.Add(outerRow);

            story.Content.Add(outerTable);
            story.Paragraphs.Add(nestedFirstParagraph);
            story.Paragraphs.Add(nestedSecondParagraph);

            var model = new DocumentModel();
            model.Content.Add(new ParagraphModel
            {
                Runs =
                {
                    new RunModel { Text = "Body" }
                }
            });
            model.Sections.Add(new SectionModel
            {
                DefaultHeaderStory = story
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
            int fcPlcSpaHdr = BitConverter.ToInt32(wordDocData, 154 + (Fib.HeaderShapePairIndex * 8));
            Assert.NotEqual(0, fcPlcSpaHdr);

            return BitConverter.ToInt32(tableData, fcPlcSpaHdr + 16);
        }

        private static int GetHeaderNestedHyperlinkFloatingImageTopWithOuterRowRule(int outerRowHeightTwips, TableRowHeightRule outerHeightRule, int hyperlinkImagePositionYTwips)
        {
            var story = new HeaderFooterStoryModel();

            var outerTable = new TableModel();
            outerTable.GridColumnWidths.Add(2400);
            var outerRow = new TableRowModel { HeightTwips = outerRowHeightTwips, HeightRule = outerHeightRule };
            var outerCell = new TableCellModel { Width = 2400, VerticalAlignment = TableCellVerticalAlignment.Top };

            var nestedTable = new TableModel();
            nestedTable.GridColumnWidths.Add(2200);
            var nestedRow = new TableRowModel();
            var nestedCell = new TableCellModel { Width = 2200 };

            var nestedFirstParagraph = new ParagraphModel();
            nestedFirstParagraph.Runs.Add(new RunModel
            {
                Text = "Tall nested intro line for hyperlink clipping behavior",
                Properties =
                {
                    FontSize = 72
                }
            });

            var hyperlink = new HyperlinkModel
            {
                TargetUrl = "https://example.com/nested-image",
                DisplayText = "NestedImageLink"
            };

            var nestedSecondParagraph = new ParagraphModel();
            nestedSecondParagraph.Runs.Add(new RunModel
            {
                Hyperlink = hyperlink,
                Image = new ImageModel
                {
                    Data = GetTestPngBytes(),
                    ContentType = "image/png",
                    Width = 96,
                    Height = 96,
                    LayoutType = ImageLayoutType.Floating,
                    WrapType = ImageWrapType.Square,
                    VerticalRelativeTo = "paragraph",
                    PositionYTwips = hyperlinkImagePositionYTwips
                }
            });

            nestedCell.Content.Add(nestedFirstParagraph);
            nestedCell.Content.Add(nestedSecondParagraph);
            nestedCell.Paragraphs.Add(nestedFirstParagraph);
            nestedCell.Paragraphs.Add(nestedSecondParagraph);
            nestedRow.Cells.Add(nestedCell);
            nestedTable.Rows.Add(nestedRow);

            outerCell.Content.Add(nestedTable);
            outerRow.Cells.Add(outerCell);
            outerTable.Rows.Add(outerRow);

            story.Content.Add(outerTable);
            story.Paragraphs.Add(nestedFirstParagraph);
            story.Paragraphs.Add(nestedSecondParagraph);

            var model = new DocumentModel();
            model.Content.Add(new ParagraphModel
            {
                Runs =
                {
                    new RunModel { Text = "Body" }
                }
            });
            model.Sections.Add(new SectionModel
            {
                DefaultHeaderStory = story
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
            int fcPlcSpaHdr = BitConverter.ToInt32(wordDocData, 154 + (Fib.HeaderShapePairIndex * 8));
            Assert.NotEqual(0, fcPlcSpaHdr);

            return BitConverter.ToInt32(tableData, fcPlcSpaHdr + 16);
        }

        private static int GetHeaderFloatingImageTopWithTableCellBorders(int topBorderTwips, int bottomBorderTwips)
        {
            var story = new HeaderFooterStoryModel();

            var table = new TableModel();
            table.GridColumnWidths.Add(2400);
            var row = new TableRowModel();
            var cell = new TableCellModel
            {
                Width = 2400,
                BorderTopTwips = topBorderTwips,
                BorderBottomTwips = bottomBorderTwips,
                HasTopBorderOverride = true,
                HasBottomBorderOverride = true
            };
            var cellParagraph = new ParagraphModel();
            cellParagraph.Runs.Add(new RunModel { Text = "Cell" });
            cell.Paragraphs.Add(cellParagraph);
            row.Cells.Add(cell);
            table.Rows.Add(row);
            story.Content.Add(table);

            var tail = new ParagraphModel();
            tail.Runs.Add(new RunModel
            {
                Image = new ImageModel
                {
                    Data = GetTestPngBytes(),
                    ContentType = "image/png",
                    Width = 96,
                    Height = 96,
                    LayoutType = ImageLayoutType.Floating,
                    WrapType = ImageWrapType.Square,
                    VerticalRelativeTo = "paragraph",
                    PositionYTwips = 0
                }
            });
            story.Content.Add(tail);
            story.Paragraphs.Add(tail);

            var model = new DocumentModel();
            model.Content.Add(new ParagraphModel
            {
                Runs =
                {
                    new RunModel { Text = "Body" }
                }
            });
            model.Sections.Add(new SectionModel
            {
                DefaultHeaderStory = story
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
            int fcPlcSpaHdr = BitConverter.ToInt32(wordDocData, 154 + (Fib.HeaderShapePairIndex * 8));
            Assert.NotEqual(0, fcPlcSpaHdr);

            return BitConverter.ToInt32(tableData, fcPlcSpaHdr + 16);
        }

        private static int GetHeaderSecondCellSecondParagraphImageTop(bool hasFixedFirstCellWidth)
        {
            var story = new HeaderFooterStoryModel();

            var table = new TableModel
            {
                PreferredWidthUnit = TableWidthUnit.Dxa,
                PreferredWidthValue = 7000
            };
            var row = new TableRowModel();

            var firstCell = new TableCellModel
            {
                Width = hasFixedFirstCellWidth ? 3600 : 0,
                WidthUnit = hasFixedFirstCellWidth ? TableWidthUnit.Dxa : TableWidthUnit.Auto
            };
            firstCell.Paragraphs.Add(new ParagraphModel
            {
                Runs =
                {
                    new RunModel { Text = "Lead" }
                }
            });

            var secondCell = new TableCellModel();
            secondCell.Content.Add(new ParagraphModel
            {
                Runs =
                {
                    new RunModel
                    {
                        Text = "This paragraph is intentionally long to trigger wrapping when width is reduced by a fixed first cell."
                    }
                }
            });
            secondCell.Content.Add(new ParagraphModel
            {
                Runs =
                {
                    new RunModel
                    {
                        Image = new ImageModel
                        {
                            Data = GetTestPngBytes(),
                            ContentType = "image/png",
                            Width = 96,
                            Height = 96,
                            LayoutType = ImageLayoutType.Floating,
                            WrapType = ImageWrapType.Square,
                            VerticalRelativeTo = "paragraph",
                            PositionYTwips = 0
                        }
                    }
                }
            });
            secondCell.Paragraphs.AddRange(secondCell.Content.ConvertAll(static c => (ParagraphModel)c));

            row.Cells.Add(firstCell);
            row.Cells.Add(secondCell);
            table.Rows.Add(row);
            story.Content.Add(table);
            story.Paragraphs.AddRange(firstCell.Paragraphs);
            story.Paragraphs.AddRange(secondCell.Paragraphs);

            var model = new DocumentModel();
            model.Content.Add(new ParagraphModel
            {
                Runs =
                {
                    new RunModel { Text = "Body" }
                }
            });
            model.Sections.Add(new SectionModel
            {
                DefaultHeaderStory = story
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
            int fcPlcSpaHdr = BitConverter.ToInt32(wordDocData, 154 + (Fib.HeaderShapePairIndex * 8));
            Assert.NotEqual(0, fcPlcSpaHdr);

            return BitConverter.ToInt32(tableData, fcPlcSpaHdr + 16);
        }

        private static int GetHeaderSecondRowFloatingImageTopWithInsideHorizontalBorder(int insideHorizontalBorderTwips)
        {
            var story = new HeaderFooterStoryModel();

            var table = new TableModel
            {
                DefaultInsideHorizontalBorderTwips = insideHorizontalBorderTwips
            };
            table.GridColumnWidths.Add(2400);

            var firstRow = new TableRowModel();
            var firstCell = new TableCellModel { Width = 2400 };
            firstCell.Paragraphs.Add(new ParagraphModel
            {
                Runs =
                {
                    new RunModel { Text = "Row1" }
                }
            });
            firstRow.Cells.Add(firstCell);
            table.Rows.Add(firstRow);

            var secondRow = new TableRowModel();
            var secondCell = new TableCellModel { Width = 2400 };
            var secondParagraph = new ParagraphModel();
            secondParagraph.Runs.Add(new RunModel
            {
                Image = new ImageModel
                {
                    Data = GetTestPngBytes(),
                    ContentType = "image/png",
                    Width = 96,
                    Height = 96,
                    LayoutType = ImageLayoutType.Floating,
                    WrapType = ImageWrapType.Square,
                    VerticalRelativeTo = "paragraph",
                    PositionYTwips = 0
                }
            });
            secondCell.Paragraphs.Add(secondParagraph);
            secondRow.Cells.Add(secondCell);
            table.Rows.Add(secondRow);

            story.Content.Add(table);
            story.Paragraphs.AddRange(firstCell.Paragraphs);
            story.Paragraphs.Add(secondParagraph);

            var model = new DocumentModel();
            model.Content.Add(new ParagraphModel
            {
                Runs =
                {
                    new RunModel { Text = "Body" }
                }
            });
            model.Sections.Add(new SectionModel
            {
                DefaultHeaderStory = story
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
            int fcPlcSpaHdr = BitConverter.ToInt32(wordDocData, 154 + (Fib.HeaderShapePairIndex * 8));
            Assert.NotEqual(0, fcPlcSpaHdr);

            return BitConverter.ToInt32(tableData, fcPlcSpaHdr + 16);
        }

        private static int GetHeaderSecondCellSecondParagraphImageTopWithInsideVerticalBorder(int insideVerticalBorderTwips)
        {
            var story = new HeaderFooterStoryModel();

            var table = new TableModel
            {
                PreferredWidthUnit = TableWidthUnit.Dxa,
                PreferredWidthValue = 4200,
                DefaultInsideVerticalBorderTwips = insideVerticalBorderTwips
            };
            var row = new TableRowModel();

            var firstCell = new TableCellModel { Width = 2000 };
            firstCell.Paragraphs.Add(new ParagraphModel
            {
                Runs =
                {
                    new RunModel { Text = "Left" }
                }
            });

            var secondCell = new TableCellModel();
            secondCell.Content.Add(new ParagraphModel
            {
                Runs =
                {
                    new RunModel
                    {
                        Text = "This paragraph should wrap more when inside vertical border consumes horizontal width."
                    }
                }
            });
            secondCell.Content.Add(new ParagraphModel
            {
                Runs =
                {
                    new RunModel
                    {
                        Image = new ImageModel
                        {
                            Data = GetTestPngBytes(),
                            ContentType = "image/png",
                            Width = 96,
                            Height = 96,
                            LayoutType = ImageLayoutType.Floating,
                            WrapType = ImageWrapType.Square,
                            VerticalRelativeTo = "paragraph",
                            PositionYTwips = 0
                        }
                    }
                }
            });
            secondCell.Paragraphs.AddRange(secondCell.Content.ConvertAll(static c => (ParagraphModel)c));

            row.Cells.Add(firstCell);
            row.Cells.Add(secondCell);
            table.Rows.Add(row);
            story.Content.Add(table);
            story.Paragraphs.AddRange(firstCell.Paragraphs);
            story.Paragraphs.AddRange(secondCell.Paragraphs);

            var model = new DocumentModel();
            model.Content.Add(new ParagraphModel
            {
                Runs =
                {
                    new RunModel { Text = "Body" }
                }
            });
            model.Sections.Add(new SectionModel
            {
                DefaultHeaderStory = story
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
            int fcPlcSpaHdr = BitConverter.ToInt32(wordDocData, 154 + (Fib.HeaderShapePairIndex * 8));
            Assert.NotEqual(0, fcPlcSpaHdr);

            return BitConverter.ToInt32(tableData, fcPlcSpaHdr + 16);
        }

        private static int GetHeaderSecondCellSecondParagraphImageTopWithInsideVerticalBorderConflict(int insideVerticalBorderTwips, int? explicitPreviousRightBorderTwips, bool setExplicitOverride)
        {
            var story = new HeaderFooterStoryModel();

            var table = new TableModel
            {
                PreferredWidthUnit = TableWidthUnit.Dxa,
                PreferredWidthValue = 4200,
                DefaultInsideVerticalBorderTwips = insideVerticalBorderTwips
            };
            var row = new TableRowModel();

            var firstCell = new TableCellModel { Width = 2000 };
            if (explicitPreviousRightBorderTwips.HasValue)
            {
                firstCell.BorderRightTwips = explicitPreviousRightBorderTwips.Value;
            }

            if (setExplicitOverride)
            {
                firstCell.HasRightBorderOverride = true;
            }

            firstCell.Paragraphs.Add(new ParagraphModel
            {
                Runs =
                {
                    new RunModel { Text = "Left" }
                }
            });

            var secondCell = new TableCellModel();
            secondCell.Content.Add(new ParagraphModel
            {
                Runs =
                {
                    new RunModel
                    {
                        Text = "This paragraph should wrap more when inside vertical border consumes horizontal width."
                    }
                }
            });
            secondCell.Content.Add(new ParagraphModel
            {
                Runs =
                {
                    new RunModel
                    {
                        Image = new ImageModel
                        {
                            Data = GetTestPngBytes(),
                            ContentType = "image/png",
                            Width = 96,
                            Height = 96,
                            LayoutType = ImageLayoutType.Floating,
                            WrapType = ImageWrapType.Square,
                            VerticalRelativeTo = "paragraph",
                            PositionYTwips = 0
                        }
                    }
                }
            });
            secondCell.Paragraphs.AddRange(secondCell.Content.ConvertAll(static c => (ParagraphModel)c));

            row.Cells.Add(firstCell);
            row.Cells.Add(secondCell);
            table.Rows.Add(row);
            story.Content.Add(table);
            story.Paragraphs.AddRange(firstCell.Paragraphs);
            story.Paragraphs.AddRange(secondCell.Paragraphs);

            var model = new DocumentModel();
            model.Content.Add(new ParagraphModel
            {
                Runs =
                {
                    new RunModel { Text = "Body" }
                }
            });
            model.Sections.Add(new SectionModel
            {
                DefaultHeaderStory = story
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
            int fcPlcSpaHdr = BitConverter.ToInt32(wordDocData, 154 + (Fib.HeaderShapePairIndex * 8));
            Assert.NotEqual(0, fcPlcSpaHdr);

            return BitConverter.ToInt32(tableData, fcPlcSpaHdr + 16);
        }

        private static int GetHeaderSecondCellSecondParagraphImageTopWithInsideVerticalBorderCurrentLeftConflict(int insideVerticalBorderTwips, int? explicitCurrentLeftBorderTwips, bool setExplicitOverride)
        {
            var story = new HeaderFooterStoryModel();

            var table = new TableModel
            {
                PreferredWidthUnit = TableWidthUnit.Dxa,
                PreferredWidthValue = 4200,
                DefaultInsideVerticalBorderTwips = insideVerticalBorderTwips
            };
            var row = new TableRowModel();

            var firstCell = new TableCellModel { Width = 2000 };
            firstCell.Paragraphs.Add(new ParagraphModel
            {
                Runs =
                {
                    new RunModel { Text = "Left" }
                }
            });

            var secondCell = new TableCellModel();
            if (explicitCurrentLeftBorderTwips.HasValue)
            {
                secondCell.BorderLeftTwips = explicitCurrentLeftBorderTwips.Value;
            }

            if (setExplicitOverride)
            {
                secondCell.HasLeftBorderOverride = true;
            }

            secondCell.Content.Add(new ParagraphModel
            {
                Runs =
                {
                    new RunModel
                    {
                        Text = "This paragraph should wrap more when inside vertical border consumes horizontal width."
                    }
                }
            });
            secondCell.Content.Add(new ParagraphModel
            {
                Runs =
                {
                    new RunModel
                    {
                        Image = new ImageModel
                        {
                            Data = GetTestPngBytes(),
                            ContentType = "image/png",
                            Width = 96,
                            Height = 96,
                            LayoutType = ImageLayoutType.Floating,
                            WrapType = ImageWrapType.Square,
                            VerticalRelativeTo = "paragraph",
                            PositionYTwips = 0
                        }
                    }
                }
            });
            secondCell.Paragraphs.AddRange(secondCell.Content.ConvertAll(static c => (ParagraphModel)c));

            row.Cells.Add(firstCell);
            row.Cells.Add(secondCell);
            table.Rows.Add(row);
            story.Content.Add(table);
            story.Paragraphs.AddRange(firstCell.Paragraphs);
            story.Paragraphs.AddRange(secondCell.Paragraphs);

            var model = new DocumentModel();
            model.Content.Add(new ParagraphModel
            {
                Runs =
                {
                    new RunModel { Text = "Body" }
                }
            });
            model.Sections.Add(new SectionModel
            {
                DefaultHeaderStory = story
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
            int fcPlcSpaHdr = BitConverter.ToInt32(wordDocData, 154 + (Fib.HeaderShapePairIndex * 8));
            Assert.NotEqual(0, fcPlcSpaHdr);

            return BitConverter.ToInt32(tableData, fcPlcSpaHdr + 16);
        }

        private static int GetHeaderSecondRowFloatingImageTopWithInsideHorizontalBorderConflict(int insideHorizontalBorderTwips, int? explicitPreviousBottomBorderTwips, bool setExplicitOverride)
        {
            var story = new HeaderFooterStoryModel();

            var table = new TableModel
            {
                DefaultInsideHorizontalBorderTwips = insideHorizontalBorderTwips
            };
            table.GridColumnWidths.Add(2400);

            var firstRow = new TableRowModel();
            var firstCell = new TableCellModel { Width = 2400 };
            if (explicitPreviousBottomBorderTwips.HasValue)
            {
                firstCell.BorderBottomTwips = explicitPreviousBottomBorderTwips.Value;
            }

            if (setExplicitOverride)
            {
                firstCell.HasBottomBorderOverride = true;
            }

            firstCell.Paragraphs.Add(new ParagraphModel
            {
                Runs =
                {
                    new RunModel { Text = "Row1" }
                }
            });
            firstRow.Cells.Add(firstCell);
            table.Rows.Add(firstRow);

            var secondRow = new TableRowModel();
            var secondCell = new TableCellModel { Width = 2400 };
            var secondParagraph = new ParagraphModel();
            secondParagraph.Runs.Add(new RunModel
            {
                Image = new ImageModel
                {
                    Data = GetTestPngBytes(),
                    ContentType = "image/png",
                    Width = 96,
                    Height = 96,
                    LayoutType = ImageLayoutType.Floating,
                    WrapType = ImageWrapType.Square,
                    VerticalRelativeTo = "paragraph",
                    PositionYTwips = 0
                }
            });
            secondCell.Paragraphs.Add(secondParagraph);
            secondRow.Cells.Add(secondCell);
            table.Rows.Add(secondRow);

            story.Content.Add(table);
            story.Paragraphs.AddRange(firstCell.Paragraphs);
            story.Paragraphs.Add(secondParagraph);

            var model = new DocumentModel();
            model.Content.Add(new ParagraphModel
            {
                Runs =
                {
                    new RunModel { Text = "Body" }
                }
            });
            model.Sections.Add(new SectionModel
            {
                DefaultHeaderStory = story
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
            int fcPlcSpaHdr = BitConverter.ToInt32(wordDocData, 154 + (Fib.HeaderShapePairIndex * 8));
            Assert.NotEqual(0, fcPlcSpaHdr);

            return BitConverter.ToInt32(tableData, fcPlcSpaHdr + 16);
        }

        private static int GetHeaderSecondRowFloatingImageTopWithInsideHorizontalBorderCurrentTopConflict(int insideHorizontalBorderTwips, int? explicitCurrentTopBorderTwips, bool setExplicitOverride)
        {
            var story = new HeaderFooterStoryModel();

            var table = new TableModel
            {
                DefaultInsideHorizontalBorderTwips = insideHorizontalBorderTwips
            };
            table.GridColumnWidths.Add(2400);

            var firstRow = new TableRowModel();
            var firstCell = new TableCellModel { Width = 2400 };
            firstCell.Paragraphs.Add(new ParagraphModel
            {
                Runs =
                {
                    new RunModel { Text = "Row1" }
                }
            });
            firstRow.Cells.Add(firstCell);
            table.Rows.Add(firstRow);

            var secondRow = new TableRowModel();
            var secondCell = new TableCellModel { Width = 2400 };
            if (explicitCurrentTopBorderTwips.HasValue)
            {
                secondCell.BorderTopTwips = explicitCurrentTopBorderTwips.Value;
            }

            if (setExplicitOverride)
            {
                secondCell.HasTopBorderOverride = true;
            }

            var secondParagraph = new ParagraphModel();
            secondParagraph.Runs.Add(new RunModel
            {
                Image = new ImageModel
                {
                    Data = GetTestPngBytes(),
                    ContentType = "image/png",
                    Width = 96,
                    Height = 96,
                    LayoutType = ImageLayoutType.Floating,
                    WrapType = ImageWrapType.Square,
                    VerticalRelativeTo = "paragraph",
                    PositionYTwips = 0
                }
            });
            secondCell.Paragraphs.Add(secondParagraph);
            secondRow.Cells.Add(secondCell);
            table.Rows.Add(secondRow);

            story.Content.Add(table);
            story.Paragraphs.AddRange(firstCell.Paragraphs);
            story.Paragraphs.Add(secondParagraph);

            var model = new DocumentModel();
            model.Content.Add(new ParagraphModel
            {
                Runs =
                {
                    new RunModel { Text = "Body" }
                }
            });
            model.Sections.Add(new SectionModel
            {
                DefaultHeaderStory = story
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
            int fcPlcSpaHdr = BitConverter.ToInt32(wordDocData, 154 + (Fib.HeaderShapePairIndex * 8));
            Assert.NotEqual(0, fcPlcSpaHdr);

            return BitConverter.ToInt32(tableData, fcPlcSpaHdr + 16);
        }

        private static HeaderFooterStoryModel CreateStructuredHeaderFooterNestedCellStory(string leadText, string nestedCellText, string tailText)
        {
            var nestedTable = new TableModel();
            var nestedRow = new TableRowModel();
            var nestedCell = new TableCellModel { Width = 2400 };
            var nestedCellParagraph = new ParagraphModel();
            nestedCellParagraph.Runs.Add(new RunModel { Text = nestedCellText });
            nestedCell.Paragraphs.Add(nestedCellParagraph);
            nestedRow.Cells.Add(nestedCell);
            nestedTable.Rows.Add(nestedRow);

            var outerTable = new TableModel();
            var outerRow = new TableRowModel();
            var outerCell = new TableCellModel { Width = 5000 };
            var innerLead = new ParagraphModel();
            innerLead.Runs.Add(new RunModel { Text = "Inner lead" });
            var innerTail = new ParagraphModel();
            innerTail.Runs.Add(new RunModel { Text = "Inner tail" });
            outerCell.Content.Add(innerLead);
            outerCell.Content.Add(nestedTable);
            outerCell.Content.Add(innerTail);
            outerCell.Paragraphs.Add(innerLead);
            outerCell.Paragraphs.Add(nestedCellParagraph);
            outerCell.Paragraphs.Add(innerTail);
            outerRow.Cells.Add(outerCell);
            outerTable.Rows.Add(outerRow);

            var story = new HeaderFooterStoryModel();
            var lead = new ParagraphModel();
            lead.Runs.Add(new RunModel { Text = leadText });
            var tail = new ParagraphModel();
            tail.Runs.Add(new RunModel { Text = tailText });
            story.Content.Add(lead);
            story.Content.Add(outerTable);
            story.Content.Add(tail);
            story.Paragraphs.Add(lead);
            story.Paragraphs.Add(tail);
            story.Text = $"{leadText}\rInner lead\r{nestedCellText}\rInner tail\r{tailText}";
            return story;
        }

        private static byte GetTapxRunCount(byte[] wordDocData, byte[] tableData)
        {
            int fcPlcfbteTapx = BitConverter.ToInt32(wordDocData, 154 + (Fib.TapxPairIndex * 8));
            int pnTapx = BitConverter.ToInt32(tableData, fcPlcfbteTapx + 8);
            int tapxPageOffset = pnTapx * 512;
            return wordDocData[tapxPageOffset + 511];
        }

        private static byte[] GetTapxPageData(byte[] wordDocData, byte[] tableData)
        {
            int fcPlcfbteTapx = BitConverter.ToInt32(wordDocData, 154 + (Fib.TapxPairIndex * 8));
            int pnTapx = BitConverter.ToInt32(tableData, fcPlcfbteTapx + 8);
            int tapxPageOffset = pnTapx * 512;
            var buffer = new byte[511];
            Array.Copy(wordDocData, tapxPageOffset, buffer, 0, 511);
            return buffer;
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
