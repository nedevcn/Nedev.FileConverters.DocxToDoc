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
