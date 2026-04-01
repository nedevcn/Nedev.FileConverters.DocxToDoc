using System;
using System.IO;
using System.Text;
using System.Collections.Generic;
using System.Reflection;
using Nedev.FileConverters.DocxToDoc.Format;
using Nedev.FileConverters.DocxToDoc.Model;
using Xunit;

namespace Nedev.FileConverters.DocxToDoc.Tests.Format
{
    public class DocWriterTableTests
    {
        [Fact]
        public void WriteDocBlocks_WithTable_WritesTableMarkersAndTapx()
        {
            // Register encoding provider for Windows-1252 
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Arrange
            var model = new DocumentModel();
            var table = new TableModel();
            var row = new TableRowModel();
            
            var cell1 = new TableCellModel { Width = 5000 };
            cell1.Paragraphs.Add(new ParagraphModel { Runs = { new RunModel { Text = "Cell 1" } } });
            
            var cell2 = new TableCellModel { Width = 5000 };
            cell2.Paragraphs.Add(new ParagraphModel { Runs = { new RunModel { Text = "Cell 2" } } });
            
            row.Cells.Add(cell1);
            row.Cells.Add(cell2);
            table.Rows.Add(row);
            
            model.Content.Add(table);

            var writer = new DocWriter();
            using var ms = new MemoryStream();

            // Act
            try
            {
                writer.WriteDocBlocks(model, ms);
                ms.Position = 0;

                // Assert
                using var compoundFile = new OpenMcdf.CompoundFile(ms);
                Assert.True(compoundFile.RootStorage.TryGetStream("WordDocument", out var wordDocStream));
                Assert.True(compoundFile.RootStorage.TryGetStream("1Table", out var tableStream));
                
                byte[] tableData = tableStream.GetData();
                byte[] wordDocData = wordDocStream.GetData();

                // Check if cell markers (ASCII 7) are present in the text
                string text = Encoding.GetEncoding(1252).GetString(wordDocData, 1536, (int)wordDocData.Length - 1536);
                Assert.Contains("Cell 1\r\x0007Cell 2\r\x0007\r", text);

                // FIB offsets for Table (TAPX)
                int fcPlcfbteTapx = BitConverter.ToInt32(wordDocData, 154 + (Fib.TapxPairIndex * 8));
                int lcbPlcfbteTapx = BitConverter.ToInt32(wordDocData, 154 + (Fib.TapxPairIndex * 8) + 4);

                Assert.NotEqual(0, fcPlcfbteTapx);
                Assert.True(lcbPlcfbteTapx > 0);
            }
            catch (Exception ex)
            {
                throw new Exception($"Test failed with error: {ex.Message}\nStack: {ex.StackTrace}");
            }
            
            // Verify PlcfTapx exists in 1Table at fcPlcfbteTapx
            // (Assuming it was written to 1Table correctly)
        }

        [Fact]
        public void WriteDocBlocks_WithNestedTableInsideCellContent_WritesNestedTableMarkers()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            var model = new DocumentModel();
            model.Content.Add(new ParagraphModel { Runs = { new RunModel { Text = "Before" } } });

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

            var cellLead = new ParagraphModel();
            cellLead.Runs.Add(new RunModel { Text = "Inner lead" });
            var cellTail = new ParagraphModel();
            cellTail.Runs.Add(new RunModel { Text = "Inner tail" });

            outerCell.Content.Add(cellLead);
            outerCell.Content.Add(nestedTable);
            outerCell.Content.Add(cellTail);

            outerCell.Paragraphs.Add(cellLead);
            outerCell.Paragraphs.Add(nestedCellParagraph);
            outerCell.Paragraphs.Add(cellTail);

            outerRow.Cells.Add(outerCell);
            outerTable.Rows.Add(outerRow);
            model.Content.Add(outerTable);
            model.Content.Add(new ParagraphModel { Runs = { new RunModel { Text = "After" } } });

            var writer = new DocWriter();
            using var ms = new MemoryStream();

            writer.WriteDocBlocks(model, ms);
            ms.Position = 0;

            using var compoundFile = new OpenMcdf.CompoundFile(ms);
            Assert.True(compoundFile.RootStorage.TryGetStream("WordDocument", out var wordDocStream));
            Assert.True(compoundFile.RootStorage.TryGetStream("1Table", out var tableStream));

            byte[] wordDocData = wordDocStream.GetData();
            string expectedText = "Before\rInner lead\rCell\r\x0007\rInner tail\r\x0007\rAfter\r";
            var textBytes = new byte[expectedText.Length];
            Array.Copy(wordDocData, 1536, textBytes, 0, expectedText.Length);
            string extractedText = Encoding.GetEncoding(1252).GetString(textBytes);

            Assert.Equal(expectedText, extractedText);

            int lcbPlcfbteTapx = BitConverter.ToInt32(wordDocData, 154 + (Fib.TapxPairIndex * 8) + 4);
            Assert.True(lcbPlcfbteTapx > 0);
            Assert.True(GetTapxRunCount(wordDocData, tableStream.GetData()) >= 3);
        }

        [Fact]
        public void WriteDocBlocks_WithTableRowHeaderAndCantSplit_WritesRowFlagSprmsIntoTapx()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            var model = new DocumentModel();
            var table = new TableModel();
            var row = new TableRowModel
            {
                IsHeader = true,
                CannotSplit = true
            };

            var cell = new TableCellModel { Width = 5000 };
            cell.Paragraphs.Add(new ParagraphModel
            {
                Runs =
                {
                    new RunModel { Text = "Cell" }
                }
            });
            row.Cells.Add(cell);
            table.Rows.Add(row);
            model.Content.Add(table);

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
        public void ScaleWidthsToTarget_WhenLastCellCannotAbsorbDelta_RebalancesAcrossSelectedWidths()
        {
            var widths = new List<int> { 100, 100, 1 };
            var method = typeof(DocWriter).GetMethod("ScaleWidthsToTarget", BindingFlags.NonPublic | BindingFlags.Static);
            Assert.NotNull(method);

            method!.Invoke(null, new object[] { widths, 10 });

            Assert.Equal(10, widths[0] + widths[1] + widths[2]);
            Assert.True(widths[0] >= 1);
            Assert.True(widths[1] >= 1);
            Assert.True(widths[2] >= 1);
        }

        [Fact]
        public void ScaleSelectedWidthsToTarget_WhenLastSelectedCannotAbsorbDelta_RebalancesAcrossSelectedWidths()
        {
            var widths = new List<int> { 100, 100, 1, 50 };
            var selected = new List<bool> { true, true, true, false };
            var method = typeof(DocWriter).GetMethod("ScaleSelectedWidthsToTarget", BindingFlags.NonPublic | BindingFlags.Static);
            Assert.NotNull(method);

            method!.Invoke(null, new object[] { widths, selected, 10 });

            Assert.Equal(10, widths[0] + widths[1] + widths[2]);
            Assert.True(widths[0] >= 1);
            Assert.True(widths[1] >= 1);
            Assert.True(widths[2] >= 1);
            Assert.Equal(50, widths[3]);
        }

        [Fact]
        public void ScaleWidthsToTarget_WithVeryLargeWidths_DoesNotOverflowAndReachesTarget()
        {
            var widths = new List<int> { int.MaxValue, int.MaxValue, int.MaxValue };
            var method = typeof(DocWriter).GetMethod("ScaleWidthsToTarget", BindingFlags.NonPublic | BindingFlags.Static);
            Assert.NotNull(method);

            method!.Invoke(null, new object[] { widths, 1000 });

            Assert.Equal(1000, widths[0] + widths[1] + widths[2]);
            Assert.True(widths[0] >= 1);
            Assert.True(widths[1] >= 1);
            Assert.True(widths[2] >= 1);
        }

        [Fact]
        public void ScaleSelectedWidthsToTarget_WithVeryLargeSelectedWidths_DoesNotOverflowAndReachesTarget()
        {
            var widths = new List<int> { int.MaxValue, int.MaxValue, int.MaxValue, 50 };
            var selected = new List<bool> { true, true, true, false };
            var method = typeof(DocWriter).GetMethod("ScaleSelectedWidthsToTarget", BindingFlags.NonPublic | BindingFlags.Static);
            Assert.NotNull(method);

            method!.Invoke(null, new object[] { widths, selected, 1000 });

            Assert.Equal(1000, widths[0] + widths[1] + widths[2]);
            Assert.True(widths[0] >= 1);
            Assert.True(widths[1] >= 1);
            Assert.True(widths[2] >= 1);
            Assert.Equal(50, widths[3]);
        }

        [Fact]
        public void ResolveTableCellWidth_WithVeryLargeGridWidths_ClampsToIntMax()
        {
            var table = new TableModel();
            table.GridColumnWidths.Add(int.MaxValue);
            table.GridColumnWidths.Add(int.MaxValue);

            var method = typeof(DocWriter).GetMethod("ResolveTableCellWidth", BindingFlags.NonPublic | BindingFlags.Static);
            Assert.NotNull(method);

            int width = (int)method!.Invoke(null, new object[] { table, 0, 2, 0 })!;
            Assert.Equal(int.MaxValue, width);
        }

        [Fact]
        public void ResolveBaseTableWidthTwips_WithVeryLargeGridAndCellWidths_ClampsToIntMax()
        {
            var table = new TableModel();
            table.GridColumnWidths.Add(int.MaxValue);
            table.GridColumnWidths.Add(int.MaxValue);

            var row = new TableRowModel();
            row.Cells.Add(new TableCellModel { Width = int.MaxValue, WidthUnit = TableWidthUnit.Dxa });
            row.Cells.Add(new TableCellModel { Width = int.MaxValue, WidthUnit = TableWidthUnit.Dxa });

            var method = typeof(DocWriter).GetMethod("ResolveBaseTableWidthTwips", BindingFlags.NonPublic | BindingFlags.Static);
            Assert.NotNull(method);

            int widthFromGrid = (int)method!.Invoke(null, new object[] { table, row, 2, 1000 })!;
            Assert.Equal(int.MaxValue, widthFromGrid);

            table.GridColumnWidths.Clear();
            int widthFromCells = (int)method.Invoke(null, new object[] { table, row, 2, 1000 })!;
            Assert.Equal(int.MaxValue, widthFromCells);
        }

        [Fact]
        public void ScaleExplicitWidthsToTarget_WithVeryLargeWidths_DoesNotOverflow()
        {
            var widths = new List<int> { int.MaxValue, int.MaxValue, int.MaxValue };
            var explicitFlags = new List<bool> { true, true, false };
            var units = new List<TableWidthUnit> { TableWidthUnit.Dxa, TableWidthUnit.Pct, TableWidthUnit.Auto };

            var method = typeof(DocWriter).GetMethod("ScaleExplicitWidthsToTarget", BindingFlags.NonPublic | BindingFlags.Static);
            Assert.NotNull(method);

            bool scaled = (bool)method!.Invoke(null, new object[] { widths, explicitFlags, units, 1000 })!;
            Assert.True(scaled);
            Assert.True(widths[0] >= 1);
            Assert.True(widths[1] >= 1);
            Assert.True(widths[2] >= 1);
        }

        [Fact]
        public void ResolveRowCellWidthsTwips_WithVeryLargeGridAndSmallTarget_ConvergesToTarget()
        {
            var table = new TableModel();
            table.GridColumnWidths.Add(int.MaxValue);
            table.GridColumnWidths.Add(int.MaxValue);
            table.GridColumnWidths.Add(int.MaxValue);

            var row = new TableRowModel();
            row.Cells.Add(new TableCellModel { Width = 0, WidthUnit = TableWidthUnit.Auto });
            row.Cells.Add(new TableCellModel { Width = 0, WidthUnit = TableWidthUnit.Auto });
            row.Cells.Add(new TableCellModel { Width = 0, WidthUnit = TableWidthUnit.Auto });

            var method = typeof(DocWriter).GetMethod("ResolveRowCellWidthsTwips", BindingFlags.NonPublic | BindingFlags.Static);
            Assert.NotNull(method);

            var widths = (List<int>)method!.Invoke(null, new object[] { table, row, 1000, 3 })!;
            Assert.Equal(3, widths.Count);
            Assert.Equal(1000, widths[0] + widths[1] + widths[2]);
            Assert.True(widths[0] >= 1);
            Assert.True(widths[1] >= 1);
            Assert.True(widths[2] >= 1);
        }

        private bool IsWord97Format(byte[] data)
        {
            return BitConverter.ToUInt16(data, 0) == 0xA5EC;
        }

        [Fact]
        public void WriteDocBlocks_TableCellFloatingImages_UseCellLocalVerticalPositions()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            byte[] pngBytes = new byte[]
            {
                0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A,
                0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44, 0x52
            };

            var model = new DocumentModel();
            model.Content.Add(new ParagraphModel { Runs = { new RunModel { Text = "Preamble paragraph" } } });

            var table = new TableModel();
            var row = new TableRowModel();

            var cell1 = new TableCellModel { Width = 5000 };
            cell1.Paragraphs.Add(new ParagraphModel
            {
                Runs =
                {
                    new RunModel
                    {
                        Image = new ImageModel
                        {
                            Data = pngBytes,
                            ContentType = "image/png",
                            Width = 96,
                            Height = 48,
                            LayoutType = ImageLayoutType.Floating,
                            VerticalRelativeTo = "paragraph",
                            PositionYTwips = 100
                        }
                    }
                }
            });

            var cell2 = new TableCellModel { Width = 5000 };
            cell2.Paragraphs.Add(new ParagraphModel
            {
                Runs =
                {
                    new RunModel
                    {
                        Image = new ImageModel
                        {
                            Data = pngBytes,
                            ContentType = "image/png",
                            Width = 96,
                            Height = 48,
                            LayoutType = ImageLayoutType.Floating,
                            VerticalRelativeTo = "paragraph",
                            PositionYTwips = 100
                        }
                    }
                }
            });

            row.Cells.Add(cell1);
            row.Cells.Add(cell2);
            table.Rows.Add(row);
            model.Content.Add(table);

            var writer = new DocWriter();
            using var ms = new MemoryStream();

            writer.WriteDocBlocks(model, ms);
            ms.Position = 0;

            using var compoundFile = new OpenMcdf.CompoundFile(ms);
            Assert.True(compoundFile.RootStorage.TryGetStream("WordDocument", out var wordDocStream));
            Assert.True(compoundFile.RootStorage.TryGetStream("1Table", out var tableStream));

            var wordDocData = wordDocStream.GetData();
            var tableData = tableStream.GetData();
            int fcPlcfspaMom = BitConverter.ToInt32(wordDocData, 154 + (40 * 8));

            Assert.Equal(19, BitConverter.ToInt32(tableData, fcPlcfspaMom));
            Assert.Equal(22, BitConverter.ToInt32(tableData, fcPlcfspaMom + 4));
            Assert.Equal(26, BitConverter.ToInt32(tableData, fcPlcfspaMom + 8));

            int firstRecordOffset = fcPlcfspaMom + 12;
            int secondRecordOffset = firstRecordOffset + 26;

            Assert.Equal(100, BitConverter.ToInt32(tableData, firstRecordOffset + 8));
            Assert.Equal(820, BitConverter.ToInt32(tableData, firstRecordOffset + 16));
            Assert.Equal(100, BitConverter.ToInt32(tableData, secondRecordOffset + 8));
            Assert.Equal(820, BitConverter.ToInt32(tableData, secondRecordOffset + 16));
        }

        [Fact]
        public void WriteDocBlocks_WithExactRowHeight_ClipsImagePosition()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            byte[] pngBytes = new byte[]
            {
                0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A,
                0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44, 0x52
            };

            var model = new DocumentModel();
            var table = new TableModel();
            var row = new TableRowModel { HeightTwips = 1000, HeightRule = TableRowHeightRule.Exact };

            var cell = new TableCellModel { Width = 5000 };
            var para = new ParagraphModel();
            para.Runs.Add(new RunModel
            {
                Text = "Force large height",
                Properties = { FontSize = 120 } // Very large font -> large paragraph height (approx 1500+ twips)
            });
            para.Runs.Add(new RunModel
            {
                Image = new ImageModel
                {
                    Data = pngBytes,
                    ContentType = "image/png",
                    Width = 64,
                    Height = 32,
                    LayoutType = ImageLayoutType.Floating,
                    VerticalAlignment = "center",
                    VerticalRelativeTo = "paragraph"
                }
            });
            cell.Paragraphs.Add(para);
            row.Cells.Add(cell);
            table.Rows.Add(row);
            model.Content.Add(table);

            var writer = new DocWriter();
            using var ms = new MemoryStream();

            try
            {
                writer.WriteDocBlocks(model, ms);
                ms.Position = 0;

                using var compoundFile = new OpenMcdf.CompoundFile(ms);
                Assert.True(compoundFile.RootStorage.TryGetStream("WordDocument", out var wordDocStream), "WordDocument stream missing");
                Assert.True(compoundFile.RootStorage.TryGetStream("1Table", out var tableStream), "1Table stream missing");
                
                var wordDocData = wordDocStream.GetData();
                var tableData = tableStream.GetData();

                // Find spaMom (floating positions) - offset into 1Table
                int fcPlcfspaMom = BitConverter.ToInt32(wordDocData, 154 + (40 * 8));
                int lcbPlcfspaMom = BitConverter.ToInt32(wordDocData, 154 + (40 * 8) + 4);
                
                Assert.NotEqual(0, fcPlcfspaMom);
                Assert.True(lcbPlcfspaMom > 0);

                // PLCF SPA with 1 image: 2 CPs (8 bytes) + 1 FSPA record (26 bytes) = 34 bytes
                // The record starts at fcPlcfspaMom + 8
                int recordStart = fcPlcfspaMom + 8;
                
                // yaTop is at recordStart + 8 (after spid, xaLeft)
                int topTwips = BitConverter.ToInt32(tableData, recordStart + 8);
                
                // Without clipping, a 480-twip image centered in a ~1500-twip paragraph would be at ~500.
                // With 1000-twip clipping, it should be at (1000 - 480) / 2 = 260.
                Assert.True(topTwips < 500, $"Top position ({topTwips}) should be clipped to row height context (~260-300).");
            }
            catch (Exception ex)
            {
                throw new Exception($"Clipping test failed: {ex.Message}\n{ex.StackTrace}");
            }
        }

        [Fact]
        public void WriteDocBlocks_TableRowsAdvanceDocumentCursorByTallestCell()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            byte[] pngBytes = new byte[]
            {
                0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A,
                0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44, 0x52
            };

            var model = new DocumentModel();
            var table = new TableModel();
            var row = new TableRowModel();

            var tallCell = new TableCellModel { Width = 5000 };
            tallCell.Paragraphs.Add(new ParagraphModel
            {
                Runs =
                {
                    new RunModel
                    {
                        Text = "Tall",
                        Properties =
                        {
                            FontSize = 48
                        }
                    }
                }
            });

            var shortCell = new TableCellModel { Width = 5000 };
            shortCell.Paragraphs.Add(new ParagraphModel { Runs = { new RunModel { Text = "Short" } } });

            row.Cells.Add(tallCell);
            row.Cells.Add(shortCell);
            table.Rows.Add(row);
            model.Content.Add(table);
            model.Content.Add(new ParagraphModel
            {
                Runs =
                {
                    new RunModel { Text = "After table" },
                    new RunModel
                    {
                        Image = new ImageModel
                        {
                            Data = pngBytes,
                            ContentType = "image/png",
                            Width = 96,
                            Height = 48,
                            LayoutType = ImageLayoutType.Floating,
                            VerticalRelativeTo = "paragraph",
                            PositionYTwips = 50
                        }
                    }
                }
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
            int fcPlcfspaMom = BitConverter.ToInt32(wordDocData, 154 + (40 * 8));

            int recordOffset = fcPlcfspaMom + 8;
            Assert.Equal(25, BitConverter.ToInt32(tableData, fcPlcfspaMom));
            Assert.Equal(27, BitConverter.ToInt32(tableData, fcPlcfspaMom + 4));
            Assert.Equal(2088, BitConverter.ToInt32(tableData, recordOffset + 8));
            Assert.Equal(2808, BitConverter.ToInt32(tableData, recordOffset + 16));
        }

        [Fact]
        public void WriteDocBlocks_NarrowTableCell_UsesWrappedParagraphHeightForFloatingImage()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            byte[] pngBytes = new byte[]
            {
                0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A,
                0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44, 0x52
            };

            var model = new DocumentModel();
            var table = new TableModel();
            var row = new TableRowModel();
            var cell = new TableCellModel { Width = 2000 };

            cell.Paragraphs.Add(new ParagraphModel
            {
                Runs =
                {
                    new RunModel
                    {
                        Text = "This paragraph is long enough to wrap twice.",
                        Properties =
                        {
                            FontSize = 24
                        }
                    }
                }
            });

            cell.Paragraphs.Add(new ParagraphModel
            {
                Runs =
                {
                    new RunModel
                    {
                        Image = new ImageModel
                        {
                            Data = pngBytes,
                            ContentType = "image/png",
                            Width = 96,
                            Height = 48,
                            LayoutType = ImageLayoutType.Floating,
                            VerticalRelativeTo = "paragraph",
                            PositionYTwips = 50
                        }
                    }
                }
            });

            row.Cells.Add(cell);
            table.Rows.Add(row);
            model.Content.Add(table);

            var writer = new DocWriter();
            using var ms = new MemoryStream();

            writer.WriteDocBlocks(model, ms);
            ms.Position = 0;

            using var compoundFile = new OpenMcdf.CompoundFile(ms);
            Assert.True(compoundFile.RootStorage.TryGetStream("WordDocument", out var wordDocStream));
            Assert.True(compoundFile.RootStorage.TryGetStream("1Table", out var tableStream));

            var wordDocData = wordDocStream.GetData();
            var tableData = tableStream.GetData();
            int fcPlcfspaMom = BitConverter.ToInt32(wordDocData, 154 + (40 * 8));
            int recordOffset = fcPlcfspaMom + 8;

            Assert.Equal(1016, BitConverter.ToInt32(tableData, recordOffset + 8));
            Assert.Equal(1736, BitConverter.ToInt32(tableData, recordOffset + 16));
        }

        [Fact]
        public void WriteDocBlocks_FloatingImageDistance_ExpandsBodyAnchorBounds()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            byte[] pngBytes = new byte[]
            {
                0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A,
                0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44, 0x52
            };

            (int left, int top, int right, int bottom) withoutDistance = GetBodyFloatingImageBoundsWithDistance(pngBytes, 0, 0, 0, 0);
            (int left, int top, int right, int bottom) withDistance = GetBodyFloatingImageBoundsWithDistance(pngBytes, 120, 240, 360, 480);

            Assert.Equal(withoutDistance.left - 120, withDistance.left);
            Assert.Equal(withoutDistance.top - 240, withDistance.top);
            Assert.Equal(withoutDistance.right + 360, withDistance.right);
            Assert.Equal(withoutDistance.bottom + 480, withDistance.bottom);
        }

        [Fact]
        public void WriteDocBlocks_FloatingImageVerticalRelativeToLine_UsesSmallerAnchorThanParagraph()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            byte[] pngBytes = new byte[]
            {
                0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A,
                0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44, 0x52
            };

            int paragraphRelativeTop = GetBodyFloatingImageTopWithVerticalRelativeTo(pngBytes, "paragraph");
            int lineRelativeTop = GetBodyFloatingImageTopWithVerticalRelativeTo(pngBytes, "line");

            Assert.True(lineRelativeTop < paragraphRelativeTop);
        }

        [Fact]
        public void WriteDocBlocks_FloatingImageHorizontalRelativeToMarginSides_UsesDistinctAnchors()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            byte[] pngBytes = new byte[]
            {
                0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A,
                0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44, 0x52
            };

            int leftMarginAnchor = GetBodyFloatingImagePositionWithRelativeTo(pngBytes, "leftMargin", "topMargin", 0, 0).left;
            int rightMarginAnchor = GetBodyFloatingImagePositionWithRelativeTo(pngBytes, "rightMargin", "topMargin", 0, 0).left;

            Assert.Equal(1440, leftMarginAnchor);
            Assert.Equal(10466, rightMarginAnchor);
            Assert.Equal(9026, rightMarginAnchor - leftMarginAnchor);
        }

        [Fact]
        public void WriteDocBlocks_FloatingImageVerticalRelativeToMarginSides_UsesDistinctAnchors()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            byte[] pngBytes = new byte[]
            {
                0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A,
                0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44, 0x52
            };

            int topMarginAnchor = GetBodyFloatingImagePositionWithRelativeTo(pngBytes, "leftMargin", "topMargin", 0, 0).top;
            int bottomMarginAnchor = GetBodyFloatingImagePositionWithRelativeTo(pngBytes, "leftMargin", "bottomMargin", 0, 0).top;

            Assert.Equal(1440, topMarginAnchor);
            Assert.Equal(15398, bottomMarginAnchor);
            Assert.Equal(13958, bottomMarginAnchor - topMarginAnchor);
        }

        [Fact]
        public void WriteDocBlocks_FloatingImageHorizontalRelativeToInsideOutsideMargin_UsesDistinctAnchors()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            byte[] pngBytes = new byte[]
            {
                0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A,
                0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44, 0x52
            };

            int insideMarginAnchor = GetBodyFloatingImagePositionWithRelativeTo(pngBytes, "insideMargin", "topMargin", 0, 0).left;
            int outsideMarginAnchor = GetBodyFloatingImagePositionWithRelativeTo(pngBytes, "outsideMargin", "topMargin", 0, 0).left;

            Assert.Equal(1440, insideMarginAnchor);
            Assert.Equal(10466, outsideMarginAnchor);
            Assert.Equal(9026, outsideMarginAnchor - insideMarginAnchor);
        }

        [Fact]
        public void WriteDocBlocks_FloatingImageHorizontalRelativeToInsideOutside_Alias_UsesDistinctAnchors()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            byte[] pngBytes = new byte[]
            {
                0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A,
                0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44, 0x52
            };

            int insideAnchor = GetBodyFloatingImagePositionWithRelativeTo(pngBytes, "inside", "topMargin", 0, 0).left;
            int outsideAnchor = GetBodyFloatingImagePositionWithRelativeTo(pngBytes, "outside", "topMargin", 0, 0).left;

            Assert.Equal(1440, insideAnchor);
            Assert.Equal(10466, outsideAnchor);
            Assert.Equal(9026, outsideAnchor - insideAnchor);
        }

        [Fact]
        public void WriteDocBlocks_TableGridWidthFallback_UsesGridColumnsForFloatingImageLayout()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            byte[] pngBytes = new byte[]
            {
                0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A,
                0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44, 0x52
            };

            int topFromNarrowGrid = GetGridFallbackImageTop(1800, pngBytes);
            int topFromWideGrid = GetGridFallbackImageTop(4200, pngBytes);

            Assert.True(topFromNarrowGrid > topFromWideGrid);

            static int GetGridFallbackImageTop(int gridColumnWidth, byte[] imageBytes)
            {
                var model = new DocumentModel();
                var table = new TableModel();
                table.GridColumnWidths.Add(gridColumnWidth);

                var row = new TableRowModel();
                var cell = new TableCellModel();
                cell.Paragraphs.Add(new ParagraphModel
                {
                    Runs =
                    {
                        new RunModel
                        {
                            Text = "This paragraph is long enough to wrap when the grid column is narrow.",
                            Properties =
                            {
                                FontSize = 24
                            }
                        }
                    }
                });

                cell.Paragraphs.Add(new ParagraphModel
                {
                    Runs =
                    {
                        new RunModel
                        {
                            Image = new ImageModel
                            {
                                Data = imageBytes,
                                ContentType = "image/png",
                                Width = 96,
                                Height = 48,
                                LayoutType = ImageLayoutType.Floating,
                                VerticalRelativeTo = "paragraph",
                                PositionYTwips = 50
                            }
                        }
                    }
                });

                row.Cells.Add(cell);
                table.Rows.Add(row);
                model.Content.Add(table);

                var writer = new DocWriter();
                using var ms = new MemoryStream();
                writer.WriteDocBlocks(model, ms);
                ms.Position = 0;

                using var compoundFile = new OpenMcdf.CompoundFile(ms);
                Assert.True(compoundFile.RootStorage.TryGetStream("WordDocument", out var wordDocStream));
                Assert.True(compoundFile.RootStorage.TryGetStream("1Table", out var tableStream));

                var wordDocData = wordDocStream.GetData();
                var tableData = tableStream.GetData();
                int fcPlcfspaMom = BitConverter.ToInt32(wordDocData, 154 + (40 * 8));
                int recordOffset = fcPlcfspaMom + 8;

                return BitConverter.ToInt32(tableData, recordOffset + 8);
            }
        }

        [Fact]
        public void WriteDocBlocks_MixedGridSpanFallback_UsesRemainingGridColumnsForLaterCells()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            byte[] pngBytes = new byte[]
            {
                0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A,
                0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44, 0x52
            };

            int gridFallbackTop = GetSecondCellImageTop(useExplicitWidth: false, addPadding: false, pngBytes);
            int explicitWidthTop = GetSecondCellImageTop(useExplicitWidth: true, addPadding: false, pngBytes);

            Assert.Equal(explicitWidthTop, gridFallbackTop);
        }

        [Fact]
        public void WriteDocBlocks_TableCellMargins_ReduceAvailableWidthForFloatingImageLayout()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            byte[] pngBytes = new byte[]
            {
                0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A,
                0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44, 0x52
            };

            int topWithoutMargins = GetSecondCellImageTop(useExplicitWidth: true, addPadding: false, pngBytes);
            int topWithMargins = GetSecondCellImageTop(useExplicitWidth: true, addPadding: true, pngBytes);

            Assert.True(topWithMargins > topWithoutMargins);
        }

        [Fact]
        public void WriteDocBlocks_TableCellVerticalPadding_ShiftsFloatingImageTopAndRowAdvance()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            byte[] pngBytes = new byte[]
            {
                0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A,
                0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44, 0x52
            };

            int topWithoutPadding = GetSingleCellImageTop(0, 0, pngBytes);
            int topWithTopPadding = GetSingleCellImageTop(240, 0, pngBytes);
            int afterTableWithoutBottomPadding = GetAfterTableImageTop(0, pngBytes);
            int afterTableWithBottomPadding = GetAfterTableImageTop(360, pngBytes);

            Assert.True(topWithTopPadding > topWithoutPadding);
            Assert.True(afterTableWithBottomPadding > afterTableWithoutBottomPadding);
        }

        [Fact]
        public void WriteDocBlocks_TableCellVerticalAlignment_ShiftsFloatingImageWithinTallRow()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            byte[] pngBytes = new byte[]
            {
                0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A,
                0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44, 0x52
            };

            int topWithTopAlignment = GetAlignedCellImageTop(TableCellVerticalAlignment.Top, pngBytes);
            int topWithBottomAlignment = GetAlignedCellImageTop(TableCellVerticalAlignment.Bottom, pngBytes);

            Assert.True(topWithBottomAlignment > topWithTopAlignment);
        }

        [Fact]
        public void WriteDocBlocks_TableCellSpacing_ReducesAvailableWidthAndAdvancesFollowingContent()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            byte[] pngBytes = new byte[]
            {
                0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A,
                0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44, 0x52
            };

            int topWithoutSpacing = GetCellSpacingImageTop(0, pngBytes);
            int topWithSpacing = GetCellSpacingImageTop(360, pngBytes);
            int afterTableWithoutSpacing = GetAfterTableImageTopWithCellSpacing(0, pngBytes);
            int afterTableWithSpacing = GetAfterTableImageTopWithCellSpacing(360, pngBytes);

            Assert.True(topWithSpacing > topWithoutSpacing);
            Assert.True(afterTableWithSpacing > afterTableWithoutSpacing);
        }

        [Fact]
        public void WriteDocBlocks_TableCellBorders_ReduceAvailableWidthAndAdvanceFollowingContent()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            byte[] pngBytes = new byte[]
            {
                0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A,
                0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44, 0x52
            };

            int topWithoutBorders = GetCellBorderImageTop(0, 0, 0, 0, pngBytes);
            int topWithBorders = GetCellBorderImageTop(360, 360, 240, 240, pngBytes);
            int afterTableWithoutBorders = GetAfterTableImageTopWithCellBorder(0, 0, pngBytes);
            int afterTableWithBorders = GetAfterTableImageTopWithCellBorder(240, 240, pngBytes);

            Assert.True(topWithBorders > topWithoutBorders);
            Assert.True(afterTableWithBorders > afterTableWithoutBorders);
        }

        [Fact]
        public void WriteDocBlocks_TableRowHeightAtLeast_AdvancesFollowingContent()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            byte[] pngBytes = new byte[]
            {
                0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A,
                0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44, 0x52
            };

            int afterTableAutoHeight = GetAfterTableImageTopWithRowHeight(0, TableRowHeightRule.Auto, pngBytes);
            int afterTableAtLeastHeight = GetAfterTableImageTopWithRowHeight(1800, TableRowHeightRule.AtLeast, pngBytes);

            Assert.True(afterTableAtLeastHeight > afterTableAutoHeight);
        }

        [Fact]
        public void WriteDocBlocks_TableRowAuto_WithEmptyCell_StillAdvancesFollowingContent()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            byte[] pngBytes = new byte[]
            {
                0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A,
                0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44, 0x52
            };

            int topWithoutTable = GetAfterTableImageTopWithOptionalEmptyRow(false, pngBytes);
            int topWithEmptyTableRow = GetAfterTableImageTopWithOptionalEmptyRow(true, pngBytes);

            Assert.True(topWithEmptyTableRow > topWithoutTable);
        }

        [Fact]
        public void WriteDocBlocks_TableRowHeightExact_ShiftsBottomAlignedFloatingImage()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            byte[] pngBytes = new byte[]
            {
                0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A,
                0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44, 0x52
            };

            int topWithoutExactHeight = GetAlignedCellImageTopWithRowHeight(0, TableRowHeightRule.Auto, TableCellVerticalAlignment.Bottom, pngBytes);
            int topWithExactHeight = GetAlignedCellImageTopWithRowHeight(2200, TableRowHeightRule.Exact, TableCellVerticalAlignment.Bottom, pngBytes);

            Assert.True(topWithExactHeight > topWithoutExactHeight);
        }

        [Fact]
        public void WriteDocBlocks_TableRowHeightExact_WithNestedTableInsideCellContent_ShiftsBottomAlignedFloatingImage()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            byte[] pngBytes = new byte[]
            {
                0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A,
                0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44, 0x52
            };

            int topWithoutExactHeight = GetAlignedCellImageTopWithRowHeightAndNestedTableContent(0, TableRowHeightRule.Auto, TableCellVerticalAlignment.Bottom, pngBytes);
            int topWithExactHeight = GetAlignedCellImageTopWithRowHeightAndNestedTableContent(4200, TableRowHeightRule.Exact, TableCellVerticalAlignment.Bottom, pngBytes);

            Assert.True(topWithExactHeight > topWithoutExactHeight);
        }

        [Fact]
        public void WriteDocBlocks_TableCellVerticalAlignment_WithNestedTableInsideCellContent_ShiftsFloatingImageWithinTallRow()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            byte[] pngBytes = new byte[]
            {
                0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A,
                0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44, 0x52
            };

            int topWithTopAlignment = GetAlignedCellImageTopWithNestedTableContent(TableCellVerticalAlignment.Top, pngBytes);
            int topWithBottomAlignment = GetAlignedCellImageTopWithNestedTableContent(TableCellVerticalAlignment.Bottom, pngBytes);

            Assert.True(topWithBottomAlignment > topWithTopAlignment);
        }

        [Fact]
        public void WriteDocBlocks_TableInsideVerticalBorder_ReducesLaterCellWidth()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            byte[] pngBytes = new byte[]
            {
                0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A,
                0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44, 0x52
            };

            int topWithoutInsideV = GetSecondCellImageTopWithInsideVerticalBorder(0, pngBytes);
            int topWithInsideV = GetSecondCellImageTopWithInsideVerticalBorder(360, pngBytes);

            Assert.True(topWithInsideV > topWithoutInsideV);
        }

        [Fact]
        public void WriteDocBlocks_TableInsideHorizontalBorder_ShiftsLaterRowAndFollowingContent()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            byte[] pngBytes = new byte[]
            {
                0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A,
                0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44, 0x52
            };

            int topWithoutInsideH = GetSecondRowImageTopWithInsideHorizontalBorder(0, pngBytes);
            int topWithInsideH = GetSecondRowImageTopWithInsideHorizontalBorder(240, pngBytes);
            int afterTableWithoutInsideH = GetAfterTableImageTopWithInsideHorizontalBorder(0, pngBytes);
            int afterTableWithInsideH = GetAfterTableImageTopWithInsideHorizontalBorder(240, pngBytes);

            Assert.True(topWithInsideH > topWithoutInsideH);
            Assert.True(afterTableWithInsideH > afterTableWithoutInsideH);
        }

        [Fact]
        public void WriteDocBlocks_TableInsideHorizontalBorder_WithVerticalMergeContinue_DoesNotShiftMergedContinuationRow()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            byte[] pngBytes = new byte[]
            {
                0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A,
                0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44, 0x52
            };

            int topWithoutInsideH = GetSecondRowImageTopWithInsideHorizontalBorder(0, pngBytes, useVerticalMergeContinue: true);
            int topWithInsideH = GetSecondRowImageTopWithInsideHorizontalBorder(240, pngBytes, useVerticalMergeContinue: true);

            Assert.Equal(topWithoutInsideH, topWithInsideH);
        }

        [Fact]
        public void WriteDocBlocks_TableBorderConflict_UsesExplicitCellBorderOverInsideBorder()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            byte[] pngBytes = new byte[]
            {
                0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A,
                0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44, 0x52
            };

            int topWithInsideOnly = GetSecondCellImageTopWithBorderConflict(60, null, false, pngBytes);
            int topWithExplicitRightBorder = GetSecondCellImageTopWithBorderConflict(60, 720, true, pngBytes);

            Assert.True(topWithExplicitRightBorder > topWithInsideOnly);
        }

        [Fact]
        public void WriteDocBlocks_TableBorderConflict_ExplicitNoneSuppressesInsideBorder()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            byte[] pngBytes = new byte[]
            {
                0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A,
                0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44, 0x52
            };

            int topWithInsideOnly = GetSecondCellImageTopWithBorderConflict(360, null, false, pngBytes);
            int topWithExplicitNone = GetSecondCellImageTopWithBorderConflict(360, 0, true, pngBytes);

            Assert.True(topWithExplicitNone < topWithInsideOnly);
        }

        [Fact]
        public void WriteDocBlocks_TableCellZeroHorizontalPadding_SuppressesDefaultPadding()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            byte[] pngBytes = new byte[]
            {
                0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A,
                0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44, 0x52
            };

            int topWithDefaultPadding = GetSecondCellImageTopWithZeroPaddingOverride(false, pngBytes);
            int topWithZeroOverride = GetSecondCellImageTopWithZeroPaddingOverride(true, pngBytes);

            Assert.True(topWithZeroOverride < topWithDefaultPadding);
        }

        [Fact]
        public void WriteDocBlocks_TableCellZeroVerticalPadding_SuppressesDefaultPadding()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            byte[] pngBytes = new byte[]
            {
                0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A,
                0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44, 0x52
            };

            int topWithDefaultPadding = GetSingleCellImageTopWithZeroPaddingOverride(false, pngBytes);
            int topWithZeroOverride = GetSingleCellImageTopWithZeroPaddingOverride(true, pngBytes);

            Assert.True(topWithZeroOverride < topWithDefaultPadding);
        }

        [Fact]
        public void WriteDocBlocks_TablePreferredWidthPct_ScalesGridWidthsForLayout()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            byte[] pngBytes = new byte[]
            {
                0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A,
                0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44, 0x52
            };

            int topWithFullWidth = GetPreferredWidthImageTop(5000, TableWidthUnit.Pct, pngBytes);
            int topWithHalfWidth = GetPreferredWidthImageTop(2500, TableWidthUnit.Pct, pngBytes);

            Assert.True(topWithHalfWidth > topWithFullWidth);
        }

        [Fact]
        public void WriteDocBlocks_TablePreferredWidthDxa_ScalesExplicitCellWidthsForLayout()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            byte[] pngBytes = new byte[]
            {
                0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A,
                0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44, 0x52
            };

            int topWithWidePreferredWidth = GetPreferredExplicitWidthImageTop(5200, pngBytes);
            int topWithNarrowPreferredWidth = GetPreferredExplicitWidthImageTop(2600, pngBytes);

            Assert.True(topWithNarrowPreferredWidth > topWithWidePreferredWidth);
        }

        [Fact]
        public void WriteDocBlocks_TableCellWidthPct_ScalesExplicitCellWidthForLayout()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            byte[] pngBytes = new byte[]
            {
                0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A,
                0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44, 0x52
            };

            int topWithFullCellWidth = GetPctCellWidthImageTop(5000, pngBytes);
            int topWithHalfCellWidth = GetPctCellWidthImageTop(2500, pngBytes);

            Assert.True(topWithHalfCellWidth > topWithFullCellWidth);
        }

        [Fact]
        public void WriteDocBlocks_TablePreferredWidthPct_ScalesMixedExplicitAndGridWidths()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            byte[] pngBytes = new byte[]
            {
                0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A,
                0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44, 0x52
            };

            int topWithWidePreferredWidth = GetMixedPreferredWidthImageTop(5000, pngBytes);
            int topWithNarrowPreferredWidth = GetMixedPreferredWidthImageTop(2500, pngBytes);

            Assert.True(topWithNarrowPreferredWidth > topWithWidePreferredWidth);
        }

        [Fact]
        public void WriteDocBlocks_TablePreferredWidthDxa_AssignsRemainingWidthToAutoCells()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            byte[] pngBytes = new byte[]
            {
                0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A,
                0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44, 0x52
            };

            int topWithWideRemainingWidth = GetAutoRemainingWidthImageTop(5200, pngBytes);
            int topWithNarrowRemainingWidth = GetAutoRemainingWidthImageTop(3600, pngBytes);

            Assert.True(topWithNarrowRemainingWidth > topWithWideRemainingWidth);
        }

        [Fact]
        public void WriteDocBlocks_TablePreferredWidthDxa_AutoCellPaddingReservesHorizontalOverhead()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            byte[] pngBytes = new byte[]
            {
                0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A,
                0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44, 0x52
            };

            var paddedTops = GetAutoCellOverheadAwareImageTops(5200, addPadding: true, addBorders: false, pngBytes);

            Assert.True(paddedTops.firstCellImageTop <= paddedTops.secondCellImageTop, $"firstCellImageTop={paddedTops.firstCellImageTop}, secondCellImageTop={paddedTops.secondCellImageTop}");
        }

        [Fact]
        public void WriteDocBlocks_TablePreferredWidthDxa_AutoCellBordersReserveHorizontalOverhead()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            byte[] pngBytes = new byte[]
            {
                0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A,
                0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44, 0x52
            };

            var borderedTops = GetAutoCellOverheadAwareImageTops(5200, addPadding: false, addBorders: true, pngBytes);

            Assert.True(borderedTops.firstCellImageTop <= borderedTops.secondCellImageTop, $"firstCellImageTop={borderedTops.firstCellImageTop}, secondCellImageTop={borderedTops.secondCellImageTop}");
        }

        [Fact]
        public void WriteDocBlocks_TablePreferredWidthDxa_ShrinksExplicitCellsBeforeOvergrowingAutoCells()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            byte[] pngBytes = new byte[]
            {
                0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A,
                0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44, 0x52
            };

            int topWithWideTarget = GetOvercommittedAutoCellImageTop(5200, pngBytes);
            int topWithOvercommittedTarget = GetOvercommittedAutoCellImageTop(3000, pngBytes);

            Assert.True(topWithOvercommittedTarget > topWithWideTarget);
        }

        [Fact]
        public void WriteDocBlocks_TablePreferredWidthDxa_PreservesGridFallbackWidthBeforeShrinkingExplicitNeighbors()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            byte[] pngBytes = new byte[]
            {
                0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A,
                0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44, 0x52
            };

            int topWithGridFallbackSecondCell = GetOverflowingMixedGridFallbackImageTop(secondCellUsesGridFallback: true, preferredWidthTwips: 4500, pngBytes);
            int topWithExplicitSecondCell = GetOverflowingMixedGridFallbackImageTop(secondCellUsesGridFallback: false, preferredWidthTwips: 4500, pngBytes);

            Assert.True(topWithGridFallbackSecondCell < topWithExplicitSecondCell);
        }

        [Fact]
        public void WriteDocBlocks_TablePreferredWidthDxa_PreservesPctCellBeforeShrinkingDxaNeighbors()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            byte[] pngBytes = new byte[]
            {
                0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A,
                0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44, 0x52
            };

            int topWithPctFirstCell = GetPctVsDxaOverflowImageTop(firstCellUsesPct: true, preferredWidthTwips: 5000, pngBytes);
            int topWithDxaFirstCell = GetPctVsDxaOverflowImageTop(firstCellUsesPct: false, preferredWidthTwips: 5000, pngBytes);

            Assert.True(topWithPctFirstCell < topWithDxaFirstCell);
        }

        [Fact]
        public void WriteDocBlocks_TableRowHeightExact_ClipsOverflowingCellContentForLaterParagraphs()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            byte[] pngBytes = new byte[]
            {
                0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A,
                0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44, 0x52
            };

            int topWithAutoHeight = GetOverflowClippedImageTop(0, TableRowHeightRule.Auto, pngBytes);
            int topWithExactHeight = GetOverflowClippedImageTop(1200, TableRowHeightRule.Exact, pngBytes);

            Assert.True(topWithExactHeight < topWithAutoHeight);
        }

        [Fact]
        public void WriteDocBlocks_TableRowHeightExact_ClipsNestedTableOverflowingContentForLaterParagraphs()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            byte[] pngBytes = new byte[]
            {
                0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A,
                0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44, 0x52
            };

            int topWithAutoHeight = GetNestedOverflowClippedImageTop(0, TableRowHeightRule.Auto, pngBytes);
            int topWithExactHeight = GetNestedOverflowClippedImageTop(1200, TableRowHeightRule.Exact, pngBytes);

            Assert.True(topWithExactHeight < topWithAutoHeight);
        }

        [Fact]
        public void WriteDocBlocks_TableRowHeightExact_ClipsNestedTableOverflowingContentForLaterParagraphs_WithNegativeSpacingAfter()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            byte[] pngBytes = new byte[]
            {
                0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A,
                0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44, 0x52
            };

            int topWithAutoHeight = GetNestedOverflowClippedImageTop(0, TableRowHeightRule.Auto, pngBytes, firstNestedParagraphSpacingAfterTwips: -600);
            int topWithExactHeight = GetNestedOverflowClippedImageTop(1200, TableRowHeightRule.Exact, pngBytes, firstNestedParagraphSpacingAfterTwips: -600);

            Assert.True(topWithExactHeight < topWithAutoHeight);
        }

        [Fact]
        public void WriteDocBlocks_TableRowHeightExact_ClipsNestedTableOverflowingContentForLaterParagraphs_WithLargePositiveSpacingAfter()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            byte[] pngBytes = new byte[]
            {
                0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A,
                0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44, 0x52
            };

            int topWithAutoHeight = GetNestedOverflowClippedImageTop(0, TableRowHeightRule.Auto, pngBytes, firstNestedParagraphSpacingAfterTwips: 1200);
            int topWithExactHeight = GetNestedOverflowClippedImageTop(1200, TableRowHeightRule.Exact, pngBytes, firstNestedParagraphSpacingAfterTwips: 1200);

            Assert.True(topWithExactHeight < topWithAutoHeight);
        }

        [Fact]
        public void WriteDocBlocks_TableRowHeightExact_ClampsNestedLaterParagraphTop_WhenSpacingBeforeExceedsVisibleHeight()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            byte[] pngBytes = new byte[]
            {
                0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A,
                0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44, 0x52
            };

            int topWithAutoHeight = GetNestedOverflowClippedImageTop(
                0,
                TableRowHeightRule.Auto,
                pngBytes,
                firstNestedParagraphSpacingAfterTwips: 0,
                secondNestedParagraphSpacingBeforeTwips: 1800,
                secondNestedParagraphPositionYTwips: 0);
            int topWithExactHeight = GetNestedOverflowClippedImageTop(
                1200,
                TableRowHeightRule.Exact,
                pngBytes,
                firstNestedParagraphSpacingAfterTwips: 0,
                secondNestedParagraphSpacingBeforeTwips: 1800,
                secondNestedParagraphPositionYTwips: 0);

            Assert.True(topWithExactHeight < topWithAutoHeight);
        }

        [Fact]
        public void WriteDocBlocks_TableRowHeightExact_ClampsNestedLaterParagraphTop_WhenParagraphRelativePositionYExceedsVisibleHeight()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            byte[] pngBytes = new byte[]
            {
                0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A,
                0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44, 0x52
            };

            int topWithAutoHeight = GetNestedOverflowClippedImageTop(
                0,
                TableRowHeightRule.Auto,
                pngBytes,
                firstNestedParagraphSpacingAfterTwips: 0,
                secondNestedParagraphSpacingBeforeTwips: 0,
                secondNestedParagraphPositionYTwips: 1800);
            int topWithExactHeight = GetNestedOverflowClippedImageTop(
                1200,
                TableRowHeightRule.Exact,
                pngBytes,
                firstNestedParagraphSpacingAfterTwips: 0,
                secondNestedParagraphSpacingBeforeTwips: 0,
                secondNestedParagraphPositionYTwips: 1800);

            Assert.True(topWithExactHeight < topWithAutoHeight);
        }

        [Fact]
        public void WriteDocBlocks_TableRowHeightExact_ClampsNestedHyperlinkFloatingImageTop_WhenParagraphRelativePositionYExceedsVisibleHeight()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            byte[] pngBytes = new byte[]
            {
                0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A,
                0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44, 0x52
            };

            int topWithAutoHeight = GetNestedOverflowClippedHyperlinkImageTop(0, TableRowHeightRule.Auto, pngBytes, hyperlinkImagePositionYTwips: 1800);
            int topWithExactHeight = GetNestedOverflowClippedHyperlinkImageTop(1200, TableRowHeightRule.Exact, pngBytes, hyperlinkImagePositionYTwips: 1800);

            Assert.True(topWithExactHeight < topWithAutoHeight);
        }

        [Fact]
        public void WriteDocBlocks_TableRowHeightExact_ClipsOverflowingCellContentForLaterParagraphs_WithNegativeSpacingBefore()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            byte[] pngBytes = new byte[]
            {
                0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A,
                0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44, 0x52
            };

            int topWithAutoHeight = GetOverflowClippedImageTop(0, TableRowHeightRule.Auto, pngBytes, -600);
            int topWithExactHeight = GetOverflowClippedImageTop(1200, TableRowHeightRule.Exact, pngBytes, -600);

            Assert.True(topWithExactHeight < topWithAutoHeight);
        }

        [Fact]
        public void WriteDocBlocks_TableRowHeightExact_ClipsOverflowingCellContentForLaterParagraphs_WithNegativeSpacingAfter()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            byte[] pngBytes = new byte[]
            {
                0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A,
                0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44, 0x52
            };

            int topWithAutoHeight = GetOverflowClippedImageTop(0, TableRowHeightRule.Auto, pngBytes, secondParagraphSpacingBeforeTwips: 0, firstParagraphSpacingAfterTwips: -600);
            int topWithExactHeight = GetOverflowClippedImageTop(1200, TableRowHeightRule.Exact, pngBytes, secondParagraphSpacingBeforeTwips: 0, firstParagraphSpacingAfterTwips: -600);

            Assert.True(topWithExactHeight < topWithAutoHeight);
        }

        [Fact]
        public void WriteDocBlocks_TableRowHeightExact_ClipsOverflowingCellContentForLaterParagraphs_WithLargePositiveSpacingAfter()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            byte[] pngBytes = new byte[]
            {
                0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A,
                0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44, 0x52
            };

            int topWithAutoHeight = GetOverflowClippedImageTop(0, TableRowHeightRule.Auto, pngBytes, secondParagraphSpacingBeforeTwips: 0, firstParagraphSpacingAfterTwips: 1200);
            int topWithExactHeight = GetOverflowClippedImageTop(1200, TableRowHeightRule.Exact, pngBytes, secondParagraphSpacingBeforeTwips: 0, firstParagraphSpacingAfterTwips: 1200);

            Assert.True(topWithExactHeight < topWithAutoHeight);
        }

        [Fact]
        public void WriteDocBlocks_TableRowHeightExact_ClampsLaterParagraphTop_WhenSpacingBeforeExceedsVisibleHeight()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            byte[] pngBytes = new byte[]
            {
                0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A,
                0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44, 0x52
            };

            int topWithAutoHeight = GetOverflowClippedImageTop(0, TableRowHeightRule.Auto, pngBytes, secondParagraphSpacingBeforeTwips: 1800, firstParagraphSpacingAfterTwips: 0, secondParagraphPositionYTwips: 0);
            int topWithExactHeight = GetOverflowClippedImageTop(1200, TableRowHeightRule.Exact, pngBytes, secondParagraphSpacingBeforeTwips: 1800, firstParagraphSpacingAfterTwips: 0, secondParagraphPositionYTwips: 0);

            Assert.True(topWithAutoHeight > 1200);
            Assert.True(topWithExactHeight <= 1200);
        }

        [Fact]
        public void WriteDocBlocks_TableRowHeightExact_ClampsLaterParagraphTop_WhenParagraphRelativePositionYExceedsVisibleHeight()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            byte[] pngBytes = new byte[]
            {
                0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A,
                0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44, 0x52
            };

            int topWithAutoHeight = GetOverflowClippedImageTop(
                0,
                TableRowHeightRule.Auto,
                pngBytes,
                secondParagraphSpacingBeforeTwips: 0,
                firstParagraphSpacingAfterTwips: 0,
                secondParagraphPositionYTwips: 1800);
            int topWithExactHeight = GetOverflowClippedImageTop(
                1200,
                TableRowHeightRule.Exact,
                pngBytes,
                secondParagraphSpacingBeforeTwips: 0,
                firstParagraphSpacingAfterTwips: 0,
                secondParagraphPositionYTwips: 1800);

            Assert.True(topWithExactHeight < topWithAutoHeight);
        }

        [Fact]
        public void WriteDocBlocks_TableRowHeightExact_ClampsHyperlinkFloatingImageTop_WhenParagraphRelativePositionYExceedsVisibleHeight()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            byte[] pngBytes = new byte[]
            {
                0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A,
                0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44, 0x52
            };

            int topWithAutoHeight = GetOverflowClippedHyperlinkImageTop(0, TableRowHeightRule.Auto, pngBytes, 1800);
            int topWithExactHeight = GetOverflowClippedHyperlinkImageTop(1200, TableRowHeightRule.Exact, pngBytes, 1800);

            Assert.True(topWithExactHeight < topWithAutoHeight);
        }

        [Fact]
        public void WriteDocBlocks_TableRowHeightExact_ClampsLaterParagraphTop_WhenParagraphRelativePositionYIsNegativeBeyondVisibleStart()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            byte[] pngBytes = new byte[]
            {
                0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A,
                0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44, 0x52
            };

            int topWithAutoHeight = GetOverflowClippedImageTop(
                0,
                TableRowHeightRule.Auto,
                pngBytes,
                secondParagraphSpacingBeforeTwips: 0,
                firstParagraphSpacingAfterTwips: 0,
                secondParagraphPositionYTwips: -1800);
            int topWithExactHeight = GetOverflowClippedImageTop(
                1200,
                TableRowHeightRule.Exact,
                pngBytes,
                secondParagraphSpacingBeforeTwips: 0,
                firstParagraphSpacingAfterTwips: 0,
                secondParagraphPositionYTwips: -1800);

            Assert.True(topWithExactHeight > topWithAutoHeight);
        }

        [Fact]
        public void WriteDocBlocks_TableRowHeightExact_ClampsHyperlinkFloatingImageTop_WhenParagraphRelativePositionYIsNegativeBeyondVisibleStart()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            byte[] pngBytes = new byte[]
            {
                0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A,
                0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44, 0x52
            };

            int topWithAutoHeight = GetOverflowClippedHyperlinkImageTop(0, TableRowHeightRule.Auto, pngBytes, -1800);
            int topWithExactHeight = GetOverflowClippedHyperlinkImageTop(1200, TableRowHeightRule.Exact, pngBytes, -1800);

            Assert.True(topWithExactHeight > topWithAutoHeight);
        }

        [Fact]
        public void WriteDocBlocks_TableRowHeightExact_ClampsNegativeParagraphRelativePositionY_ToParagraphStart()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            byte[] pngBytes = new byte[]
            {
                0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A,
                0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44, 0x52
            };

            int topWithExactNegative = GetOverflowClippedImageTop(
                1200,
                TableRowHeightRule.Exact,
                pngBytes,
                secondParagraphSpacingBeforeTwips: 0,
                firstParagraphSpacingAfterTwips: 0,
                secondParagraphPositionYTwips: -1800);
            int topWithExactZero = GetOverflowClippedImageTop(
                1200,
                TableRowHeightRule.Exact,
                pngBytes,
                secondParagraphSpacingBeforeTwips: 0,
                firstParagraphSpacingAfterTwips: 0,
                secondParagraphPositionYTwips: 0);

            Assert.Equal(topWithExactZero, topWithExactNegative);
        }

        private static (int left, int top, int right, int bottom) GetBodyFloatingImageBoundsWithDistance(
            byte[] pngBytes,
            int distanceLeftTwips,
            int distanceTopTwips,
            int distanceRightTwips,
            int distanceBottomTwips)
        {
            var model = new DocumentModel();
            model.Content.Add(new ParagraphModel
            {
                Runs =
                {
                    new RunModel { Text = "Body" },
                    new RunModel
                    {
                        Image = new ImageModel
                        {
                            Data = pngBytes,
                            ContentType = "image/png",
                            Width = 96,
                            Height = 48,
                            LayoutType = ImageLayoutType.Floating,
                            HorizontalRelativeTo = "margin",
                            VerticalRelativeTo = "paragraph",
                            PositionXTwips = 0,
                            PositionYTwips = 0,
                            DistanceLeftTwips = distanceLeftTwips,
                            DistanceTopTwips = distanceTopTwips,
                            DistanceRightTwips = distanceRightTwips,
                            DistanceBottomTwips = distanceBottomTwips
                        }
                    }
                }
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
            int fcPlcfspaMom = BitConverter.ToInt32(wordDocData, 154 + (40 * 8));
            int recordOffset = fcPlcfspaMom + 8;

            int left = BitConverter.ToInt32(tableData, recordOffset + 4);
            int top = BitConverter.ToInt32(tableData, recordOffset + 8);
            int right = BitConverter.ToInt32(tableData, recordOffset + 12);
            int bottom = BitConverter.ToInt32(tableData, recordOffset + 16);
            return (left, top, right, bottom);
        }

        private static (int left, int top) GetBodyFloatingImagePositionWithRelativeTo(
            byte[] pngBytes,
            string horizontalRelativeTo,
            string verticalRelativeTo,
            int positionXTwips,
            int positionYTwips)
        {
            var model = new DocumentModel();
            model.Content.Add(new ParagraphModel
            {
                Runs =
                {
                    new RunModel { Text = "Body" },
                    new RunModel
                    {
                        Image = new ImageModel
                        {
                            Data = pngBytes,
                            ContentType = "image/png",
                            Width = 96,
                            Height = 48,
                            LayoutType = ImageLayoutType.Floating,
                            HorizontalRelativeTo = horizontalRelativeTo,
                            VerticalRelativeTo = verticalRelativeTo,
                            PositionXTwips = positionXTwips,
                            PositionYTwips = positionYTwips
                        }
                    }
                }
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
            int fcPlcfspaMom = BitConverter.ToInt32(wordDocData, 154 + (40 * 8));
            int recordOffset = fcPlcfspaMom + 8;

            int left = BitConverter.ToInt32(tableData, recordOffset + 4);
            int top = BitConverter.ToInt32(tableData, recordOffset + 8);
            return (left, top);
        }

        private static int GetBodyFloatingImageTopWithVerticalRelativeTo(byte[] pngBytes, string verticalRelativeTo)
        {
            var model = new DocumentModel();
            var paragraph = new ParagraphModel
            {
                Properties =
                {
                    LineSpacing = 1200,
                    LineSpacingRule = "exact"
                },
                Runs =
                {
                    new RunModel
                    {
                        Text = "This is a long paragraph that should wrap into multiple lines to make line-relative and paragraph-relative anchors diverge.",
                        Properties = { FontSize = 24 }
                    },
                    new RunModel
                    {
                        Image = new ImageModel
                        {
                            Data = pngBytes,
                            ContentType = "image/png",
                            Width = 16,
                            Height = 8,
                            LayoutType = ImageLayoutType.Floating,
                            VerticalAlignment = "bottom",
                            VerticalRelativeTo = verticalRelativeTo,
                            PositionYTwips = 0
                        }
                    }
                }
            };
            model.Content.Add(paragraph);

            var writer = new DocWriter();
            using var ms = new MemoryStream();
            writer.WriteDocBlocks(model, ms);
            ms.Position = 0;

            using var compoundFile = new OpenMcdf.CompoundFile(ms);
            Assert.True(compoundFile.RootStorage.TryGetStream("WordDocument", out var wordDocStream));
            Assert.True(compoundFile.RootStorage.TryGetStream("1Table", out var tableStream));

            var wordDocData = wordDocStream.GetData();
            var tableData = tableStream.GetData();
            int fcPlcfspaMom = BitConverter.ToInt32(wordDocData, 154 + (40 * 8));
            int recordOffset = fcPlcfspaMom + 8;

            return BitConverter.ToInt32(tableData, recordOffset + 8);
        }

        private static int GetSecondCellImageTop(bool useExplicitWidth, bool addPadding, byte[] pngBytes)
        {
            var model = new DocumentModel();
            var table = new TableModel();
            table.GridColumnWidths.Add(1200);
            table.GridColumnWidths.Add(1600);
            table.GridColumnWidths.Add(3400);

            if (addPadding)
            {
                table.DefaultCellPaddingLeftTwips = 360;
                table.DefaultCellPaddingRightTwips = 360;
            }

            var row = new TableRowModel();

            var firstCell = new TableCellModel { GridSpan = 2 };
            firstCell.Paragraphs.Add(new ParagraphModel
            {
                Runs = { new RunModel { Text = "Leading cell" } }
            });

            var secondCell = new TableCellModel();
            if (useExplicitWidth)
            {
                secondCell.Width = 3400;
            }

            secondCell.Paragraphs.Add(new ParagraphModel
            {
                Runs =
                {
                    new RunModel
                    {
                        Text = "This paragraph is long enough to wrap differently when the effective cell width changes.",
                        Properties =
                        {
                            FontSize = 24
                        }
                    }
                }
            });

            secondCell.Paragraphs.Add(new ParagraphModel
            {
                Runs =
                {
                    new RunModel
                    {
                        Image = new ImageModel
                        {
                            Data = pngBytes,
                            ContentType = "image/png",
                            Width = 96,
                            Height = 48,
                            LayoutType = ImageLayoutType.Floating,
                            VerticalRelativeTo = "paragraph",
                            PositionYTwips = 50
                        }
                    }
                }
            });

            row.Cells.Add(firstCell);
            row.Cells.Add(secondCell);
            table.Rows.Add(row);
            model.Content.Add(table);

            var writer = new DocWriter();
            using var ms = new MemoryStream();
            writer.WriteDocBlocks(model, ms);
            ms.Position = 0;

            using var compoundFile = new OpenMcdf.CompoundFile(ms);
            Assert.True(compoundFile.RootStorage.TryGetStream("WordDocument", out var wordDocStream));
            Assert.True(compoundFile.RootStorage.TryGetStream("1Table", out var tableStream));

            var wordDocData = wordDocStream.GetData();
            var tableData = tableStream.GetData();
            int fcPlcfspaMom = BitConverter.ToInt32(wordDocData, 154 + (40 * 8));
            int recordOffset = fcPlcfspaMom + 8;

            return BitConverter.ToInt32(tableData, recordOffset + 8);
        }

        private static int GetSecondCellImageTopWithZeroPaddingOverride(bool useZeroPaddingOverride, byte[] pngBytes)
        {
            var model = new DocumentModel();
            var table = new TableModel
            {
                DefaultCellPaddingLeftTwips = 360,
                DefaultCellPaddingRightTwips = 360
            };
            table.GridColumnWidths.Add(1800);
            table.GridColumnWidths.Add(1800);

            var row = new TableRowModel();
            var firstCell = new TableCellModel { Width = 1800 };
            firstCell.Paragraphs.Add(new ParagraphModel { Runs = { new RunModel { Text = "Leading cell" } } });

            var secondCell = new TableCellModel { Width = 1800 };
            if (useZeroPaddingOverride)
            {
                secondCell.HasLeftPaddingOverride = true;
                secondCell.HasRightPaddingOverride = true;
                secondCell.PaddingLeftTwips = 0;
                secondCell.PaddingRightTwips = 0;
            }

            secondCell.Paragraphs.Add(new ParagraphModel
            {
                Runs =
                {
                    new RunModel
                    {
                        Text = "This paragraph is intentionally long enough to wrap differently when explicit zero left and right cell padding suppresses the table default margins.",
                        Properties =
                        {
                            FontSize = 24
                        }
                    }
                }
            });
            secondCell.Paragraphs.Add(new ParagraphModel
            {
                Runs =
                {
                    new RunModel
                    {
                        Image = new ImageModel
                        {
                            Data = pngBytes,
                            ContentType = "image/png",
                            Width = 96,
                            Height = 48,
                            LayoutType = ImageLayoutType.Floating,
                            VerticalRelativeTo = "paragraph",
                            PositionYTwips = 50
                        }
                    }
                }
            });

            row.Cells.Add(firstCell);
            row.Cells.Add(secondCell);
            table.Rows.Add(row);
            model.Content.Add(table);

            var writer = new DocWriter();
            using var ms = new MemoryStream();
            writer.WriteDocBlocks(model, ms);
            ms.Position = 0;

            using var compoundFile = new OpenMcdf.CompoundFile(ms);
            Assert.True(compoundFile.RootStorage.TryGetStream("WordDocument", out var wordDocStream));
            Assert.True(compoundFile.RootStorage.TryGetStream("1Table", out var tableStream));

            var wordDocData = wordDocStream.GetData();
            var tableData = tableStream.GetData();
            int fcPlcfspaMom = BitConverter.ToInt32(wordDocData, 154 + (40 * 8));
            int recordOffset = fcPlcfspaMom + 8;

            return BitConverter.ToInt32(tableData, recordOffset + 8);
        }

        private static int GetSingleCellImageTopWithZeroPaddingOverride(bool useZeroPaddingOverride, byte[] pngBytes)
        {
            var model = new DocumentModel();
            var table = new TableModel
            {
                DefaultCellPaddingTopTwips = 240,
                DefaultCellPaddingBottomTwips = 240
            };
            var row = new TableRowModel();
            var cell = new TableCellModel { Width = 2600 };
            if (useZeroPaddingOverride)
            {
                cell.HasTopPaddingOverride = true;
                cell.HasBottomPaddingOverride = true;
                cell.PaddingTopTwips = 0;
                cell.PaddingBottomTwips = 0;
            }

            cell.Paragraphs.Add(new ParagraphModel
            {
                Runs =
                {
                    new RunModel
                    {
                        Image = new ImageModel
                        {
                            Data = pngBytes,
                            ContentType = "image/png",
                            Width = 96,
                            Height = 48,
                            LayoutType = ImageLayoutType.Floating,
                            VerticalRelativeTo = "paragraph",
                            PositionYTwips = 50
                        }
                    }
                }
            });

            row.Cells.Add(cell);
            table.Rows.Add(row);
            model.Content.Add(table);

            var writer = new DocWriter();
            using var ms = new MemoryStream();
            writer.WriteDocBlocks(model, ms);
            ms.Position = 0;

            using var compoundFile = new OpenMcdf.CompoundFile(ms);
            Assert.True(compoundFile.RootStorage.TryGetStream("WordDocument", out var wordDocStream));
            Assert.True(compoundFile.RootStorage.TryGetStream("1Table", out var tableStream));

            var wordDocData = wordDocStream.GetData();
            var tableData = tableStream.GetData();
            int fcPlcfspaMom = BitConverter.ToInt32(wordDocData, 154 + (40 * 8));
            int recordOffset = fcPlcfspaMom + 8;

            return BitConverter.ToInt32(tableData, recordOffset + 8);
        }

        private static int GetSingleCellImageTop(int topPaddingTwips, int bottomPaddingTwips, byte[] pngBytes)
        {
            var model = new DocumentModel();
            var table = new TableModel();
            var row = new TableRowModel();
            var cell = new TableCellModel { Width = 2600, PaddingTopTwips = topPaddingTwips, PaddingBottomTwips = bottomPaddingTwips };

            cell.Paragraphs.Add(new ParagraphModel
            {
                Runs =
                {
                    new RunModel
                    {
                        Image = new ImageModel
                        {
                            Data = pngBytes,
                            ContentType = "image/png",
                            Width = 96,
                            Height = 48,
                            LayoutType = ImageLayoutType.Floating,
                            VerticalRelativeTo = "paragraph",
                            PositionYTwips = 50
                        }
                    }
                }
            });

            row.Cells.Add(cell);
            table.Rows.Add(row);
            model.Content.Add(table);

            var writer = new DocWriter();
            using var ms = new MemoryStream();
            writer.WriteDocBlocks(model, ms);
            ms.Position = 0;

            using var compoundFile = new OpenMcdf.CompoundFile(ms);
            Assert.True(compoundFile.RootStorage.TryGetStream("WordDocument", out var wordDocStream));
            Assert.True(compoundFile.RootStorage.TryGetStream("1Table", out var tableStream));

            var wordDocData = wordDocStream.GetData();
            var tableData = tableStream.GetData();
            int fcPlcfspaMom = BitConverter.ToInt32(wordDocData, 154 + (40 * 8));
            int recordOffset = fcPlcfspaMom + 8;

            return BitConverter.ToInt32(tableData, recordOffset + 8);
        }

        private static int GetAfterTableImageTop(int bottomPaddingTwips, byte[] pngBytes)
        {
            var model = new DocumentModel();
            var table = new TableModel();
            var row = new TableRowModel();
            var cell = new TableCellModel { Width = 2600, PaddingBottomTwips = bottomPaddingTwips };

            cell.Paragraphs.Add(new ParagraphModel
            {
                Runs =
                {
                    new RunModel
                    {
                        Text = "Cell",
                        Properties =
                        {
                            FontSize = 24
                        }
                    }
                }
            });

            row.Cells.Add(cell);
            table.Rows.Add(row);
            model.Content.Add(table);
            model.Content.Add(new ParagraphModel
            {
                Runs =
                {
                    new RunModel { Text = "After table" },
                    new RunModel
                    {
                        Image = new ImageModel
                        {
                            Data = pngBytes,
                            ContentType = "image/png",
                            Width = 96,
                            Height = 48,
                            LayoutType = ImageLayoutType.Floating,
                            VerticalRelativeTo = "paragraph",
                            PositionYTwips = 50
                        }
                    }
                }
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
            int fcPlcfspaMom = BitConverter.ToInt32(wordDocData, 154 + (40 * 8));
            int recordOffset = fcPlcfspaMom + 8;

            return BitConverter.ToInt32(tableData, recordOffset + 8);
        }

        private static int GetAlignedCellImageTop(TableCellVerticalAlignment alignment, byte[] pngBytes)
        {
            var model = new DocumentModel();
            var table = new TableModel();
            var row = new TableRowModel();

            var tallCell = new TableCellModel { Width = 2600 };
            tallCell.Paragraphs.Add(new ParagraphModel
            {
                Runs =
                {
                    new RunModel
                    {
                        Text = "This tall cell contains enough text to produce a clearly taller row height than its neighbor.",
                        Properties =
                        {
                            FontSize = 28
                        }
                    }
                }
            });

            var alignedCell = new TableCellModel { Width = 2600, VerticalAlignment = alignment };
            alignedCell.Paragraphs.Add(new ParagraphModel
            {
                Runs =
                {
                    new RunModel
                    {
                        Image = new ImageModel
                        {
                            Data = pngBytes,
                            ContentType = "image/png",
                            Width = 96,
                            Height = 48,
                            LayoutType = ImageLayoutType.Floating,
                            VerticalRelativeTo = "paragraph",
                            PositionYTwips = 50
                        }
                    }
                }
            });

            row.Cells.Add(tallCell);
            row.Cells.Add(alignedCell);
            table.Rows.Add(row);
            model.Content.Add(table);

            var writer = new DocWriter();
            using var ms = new MemoryStream();
            writer.WriteDocBlocks(model, ms);
            ms.Position = 0;

            using var compoundFile = new OpenMcdf.CompoundFile(ms);
            Assert.True(compoundFile.RootStorage.TryGetStream("WordDocument", out var wordDocStream));
            Assert.True(compoundFile.RootStorage.TryGetStream("1Table", out var tableStream));

            var wordDocData = wordDocStream.GetData();
            var tableData = tableStream.GetData();
            int fcPlcfspaMom = BitConverter.ToInt32(wordDocData, 154 + (40 * 8));
            int recordOffset = fcPlcfspaMom + 8;

            return BitConverter.ToInt32(tableData, recordOffset + 8);
        }

        private static int GetSecondCellImageTopWithInsideVerticalBorder(int insideVerticalBorderTwips, byte[] pngBytes)
        {
            var model = new DocumentModel();
            var table = new TableModel { DefaultInsideVerticalBorderTwips = insideVerticalBorderTwips };
            table.GridColumnWidths.Add(2600);
            table.GridColumnWidths.Add(2600);

            var row = new TableRowModel();
            var firstCell = new TableCellModel { Width = 2600 };
            firstCell.Paragraphs.Add(new ParagraphModel { Runs = { new RunModel { Text = "Leading cell" } } });

            var secondCell = new TableCellModel { Width = 2600 };
            secondCell.Paragraphs.Add(new ParagraphModel
            {
                Runs =
                {
                    new RunModel
                    {
                        Text = "This paragraph is long enough to wrap differently when an inside vertical border reduces the second cell width. Adding extra words to hit the wrapper boundary reliably.",
                        Properties =
                        {
                            FontSize = 24
                        }
                    }
                }
            });
            secondCell.Paragraphs.Add(new ParagraphModel
            {
                Runs =
                {
                    new RunModel
                    {
                        Image = new ImageModel
                        {
                            Data = pngBytes,
                            ContentType = "image/png",
                            Width = 96,
                            Height = 48,
                            LayoutType = ImageLayoutType.Floating,
                            VerticalRelativeTo = "paragraph",
                            PositionYTwips = 50
                        }
                    }
                }
            });

            row.Cells.Add(firstCell);
            row.Cells.Add(secondCell);
            table.Rows.Add(row);
            model.Content.Add(table);

            var writer = new DocWriter();
            using var ms = new MemoryStream();
            writer.WriteDocBlocks(model, ms);
            ms.Position = 0;

            using var compoundFile = new OpenMcdf.CompoundFile(ms);
            Assert.True(compoundFile.RootStorage.TryGetStream("WordDocument", out var wordDocStream));
            Assert.True(compoundFile.RootStorage.TryGetStream("1Table", out var tableStream));

            var wordDocData = wordDocStream.GetData();
            var tableData = tableStream.GetData();
            int fcPlcfspaMom = BitConverter.ToInt32(wordDocData, 154 + (40 * 8));
            int recordOffset = fcPlcfspaMom + 8;

            return BitConverter.ToInt32(tableData, recordOffset + 8);
        }

        private static int GetSecondCellImageTopWithBorderConflict(int insideVerticalBorderTwips, int? explicitPreviousRightBorderTwips, bool setExplicitOverride, byte[] pngBytes)
        {
            var model = new DocumentModel();
            var table = new TableModel { DefaultInsideVerticalBorderTwips = insideVerticalBorderTwips };
            table.GridColumnWidths.Add(1800);
            table.GridColumnWidths.Add(1800);

            var row = new TableRowModel();
            var firstCell = new TableCellModel { Width = 1800 };
            if (explicitPreviousRightBorderTwips.HasValue)
            {
                firstCell.BorderRightTwips = explicitPreviousRightBorderTwips.Value;
            }
            if (setExplicitOverride)
            {
                firstCell.HasRightBorderOverride = true;
            }
            firstCell.Paragraphs.Add(new ParagraphModel { Runs = { new RunModel { Text = "Leading cell" } } });

            var secondCell = new TableCellModel { Width = 1800 };
            secondCell.Paragraphs.Add(new ParagraphModel
            {
                Runs =
                {
                    new RunModel
                    {
                        Text = "This paragraph is intentionally long enough to produce an extra wrapped line when border conflict resolution picks the larger internal boundary thickness.",
                        Properties =
                        {
                            FontSize = 24
                        }
                    }
                }
            });
            secondCell.Paragraphs.Add(new ParagraphModel
            {
                Runs =
                {
                    new RunModel
                    {
                        Image = new ImageModel
                        {
                            Data = pngBytes,
                            ContentType = "image/png",
                            Width = 96,
                            Height = 48,
                            LayoutType = ImageLayoutType.Floating,
                            VerticalRelativeTo = "paragraph",
                            PositionYTwips = 50
                        }
                    }
                }
            });

            row.Cells.Add(firstCell);
            row.Cells.Add(secondCell);
            table.Rows.Add(row);
            model.Content.Add(table);

            var writer = new DocWriter();
            using var ms = new MemoryStream();
            writer.WriteDocBlocks(model, ms);
            ms.Position = 0;

            using var compoundFile = new OpenMcdf.CompoundFile(ms);
            Assert.True(compoundFile.RootStorage.TryGetStream("WordDocument", out var wordDocStream));
            Assert.True(compoundFile.RootStorage.TryGetStream("1Table", out var tableStream));

            var wordDocData = wordDocStream.GetData();
            var tableData = tableStream.GetData();
            int fcPlcfspaMom = BitConverter.ToInt32(wordDocData, 154 + (40 * 8));
            int recordOffset = fcPlcfspaMom + 8;

            return BitConverter.ToInt32(tableData, recordOffset + 8);
        }

        private static int GetCellSpacingImageTop(int cellSpacingTwips, byte[] pngBytes)
        {
            var model = new DocumentModel();
            var table = new TableModel { CellSpacingTwips = cellSpacingTwips };
            var row = new TableRowModel();
            var cell = new TableCellModel { Width = 2600 };

            cell.Paragraphs.Add(new ParagraphModel
            {
                Runs =
                {
                    new RunModel
                    {
                        Text = "This paragraph is long enough to wrap differently when cell spacing reduces the effective width.",
                        Properties =
                        {
                            FontSize = 24
                        }
                    }
                }
            });

            cell.Paragraphs.Add(new ParagraphModel
            {
                Runs =
                {
                    new RunModel
                    {
                        Image = new ImageModel
                        {
                            Data = pngBytes,
                            ContentType = "image/png",
                            Width = 96,
                            Height = 48,
                            LayoutType = ImageLayoutType.Floating,
                            VerticalRelativeTo = "paragraph",
                            PositionYTwips = 50
                        }
                    }
                }
            });

            row.Cells.Add(cell);
            table.Rows.Add(row);
            model.Content.Add(table);

            var writer = new DocWriter();
            using var ms = new MemoryStream();
            writer.WriteDocBlocks(model, ms);
            ms.Position = 0;

            using var compoundFile = new OpenMcdf.CompoundFile(ms);
            Assert.True(compoundFile.RootStorage.TryGetStream("WordDocument", out var wordDocStream));
            Assert.True(compoundFile.RootStorage.TryGetStream("1Table", out var tableStream));

            var wordDocData = wordDocStream.GetData();
            var tableData = tableStream.GetData();
            int fcPlcfspaMom = BitConverter.ToInt32(wordDocData, 154 + (40 * 8));
            int recordOffset = fcPlcfspaMom + 8;

            return BitConverter.ToInt32(tableData, recordOffset + 8);
        }

        private static int GetAfterTableImageTopWithCellSpacing(int cellSpacingTwips, byte[] pngBytes)
        {
            var model = new DocumentModel();
            var table = new TableModel { CellSpacingTwips = cellSpacingTwips };
            var row = new TableRowModel();
            var cell = new TableCellModel { Width = 2600 };

            cell.Paragraphs.Add(new ParagraphModel
            {
                Runs =
                {
                    new RunModel
                    {
                        Text = "Cell",
                        Properties =
                        {
                            FontSize = 24
                        }
                    }
                }
            });

            row.Cells.Add(cell);
            table.Rows.Add(row);
            model.Content.Add(table);
            model.Content.Add(new ParagraphModel
            {
                Runs =
                {
                    new RunModel { Text = "After table" },
                    new RunModel
                    {
                        Image = new ImageModel
                        {
                            Data = pngBytes,
                            ContentType = "image/png",
                            Width = 96,
                            Height = 48,
                            LayoutType = ImageLayoutType.Floating,
                            VerticalRelativeTo = "paragraph",
                            PositionYTwips = 50
                        }
                    }
                }
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
            int fcPlcfspaMom = BitConverter.ToInt32(wordDocData, 154 + (40 * 8));
            int recordOffset = fcPlcfspaMom + 8;

            return BitConverter.ToInt32(tableData, recordOffset + 8);
        }

        private static int GetSecondRowImageTopWithInsideHorizontalBorder(int insideHorizontalBorderTwips, byte[] pngBytes, bool useVerticalMergeContinue = false)
        {
            var model = new DocumentModel();
            var table = new TableModel { DefaultInsideHorizontalBorderTwips = insideHorizontalBorderTwips };

            var firstRow = new TableRowModel();
            var firstCell = new TableCellModel { Width = 2600 };
            firstCell.Paragraphs.Add(new ParagraphModel
            {
                Runs = { new RunModel { Text = "Row 1" } }
            });
            firstRow.Cells.Add(firstCell);

            var secondRow = new TableRowModel();
            var secondCell = new TableCellModel { Width = 2600 };
            if (useVerticalMergeContinue)
            {
                secondCell.VerticalMerge = TableCellVerticalMerge.Continue;
            }
            secondCell.Paragraphs.Add(new ParagraphModel
            {
                Runs =
                {
                    new RunModel
                    {
                        Image = new ImageModel
                        {
                            Data = pngBytes,
                            ContentType = "image/png",
                            Width = 96,
                            Height = 48,
                            LayoutType = ImageLayoutType.Floating,
                            VerticalRelativeTo = "paragraph",
                            PositionYTwips = 50
                        }
                    }
                }
            });
            secondRow.Cells.Add(secondCell);

            table.Rows.Add(firstRow);
            table.Rows.Add(secondRow);
            model.Content.Add(table);

            var writer = new DocWriter();
            using var ms = new MemoryStream();
            writer.WriteDocBlocks(model, ms);
            ms.Position = 0;

            using var compoundFile = new OpenMcdf.CompoundFile(ms);
            Assert.True(compoundFile.RootStorage.TryGetStream("WordDocument", out var wordDocStream));
            Assert.True(compoundFile.RootStorage.TryGetStream("1Table", out var tableStream));

            var wordDocData = wordDocStream.GetData();
            var tableData = tableStream.GetData();
            int fcPlcfspaMom = BitConverter.ToInt32(wordDocData, 154 + (40 * 8));
            int recordOffset = fcPlcfspaMom + 8;

            return BitConverter.ToInt32(tableData, recordOffset + 8);
        }

        private static int GetAfterTableImageTopWithInsideHorizontalBorder(int insideHorizontalBorderTwips, byte[] pngBytes)
        {
            var model = new DocumentModel();
            var table = new TableModel { DefaultInsideHorizontalBorderTwips = insideHorizontalBorderTwips };

            var firstRow = new TableRowModel();
            var firstCell = new TableCellModel { Width = 2600 };
            firstCell.Paragraphs.Add(new ParagraphModel
            {
                Runs =
                {
                    new RunModel
                    {
                        Text = "This first row produces visible row advance before the second row.",
                        Properties =
                        {
                            FontSize = 24
                        }
                    }
                }
            });
            firstRow.Cells.Add(firstCell);

            var secondRow = new TableRowModel();
            var secondCell = new TableCellModel { Width = 2600 };
            secondCell.Paragraphs.Add(new ParagraphModel
            {
                Runs = { new RunModel { Text = "Row 2" } }
            });
            secondRow.Cells.Add(secondCell);

            table.Rows.Add(firstRow);
            table.Rows.Add(secondRow);
            model.Content.Add(table);
            model.Content.Add(new ParagraphModel
            {
                Runs =
                {
                    new RunModel { Text = "After table" },
                    new RunModel
                    {
                        Image = new ImageModel
                        {
                            Data = pngBytes,
                            ContentType = "image/png",
                            Width = 96,
                            Height = 48,
                            LayoutType = ImageLayoutType.Floating,
                            VerticalRelativeTo = "paragraph",
                            PositionYTwips = 50
                        }
                    }
                }
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
            int fcPlcfspaMom = BitConverter.ToInt32(wordDocData, 154 + (40 * 8));
            int recordOffset = fcPlcfspaMom + 8;

            return BitConverter.ToInt32(tableData, recordOffset + 8);
        }

        private static int GetPreferredWidthImageTop(int preferredWidthValue, TableWidthUnit widthUnit, byte[] pngBytes)
        {
            var model = new DocumentModel();
            var table = new TableModel { PreferredWidthValue = preferredWidthValue, PreferredWidthUnit = widthUnit };
            table.GridColumnWidths.Add(5200);

            var row = new TableRowModel();
            var cell = new TableCellModel();
            cell.Paragraphs.Add(new ParagraphModel
            {
                Runs =
                {
                    new RunModel
                    {
                        Text = "This paragraph should wrap more when the preferred table width becomes narrower.",
                        Properties =
                        {
                            FontSize = 24
                        }
                    }
                }
            });
            cell.Paragraphs.Add(new ParagraphModel
            {
                Runs =
                {
                    new RunModel
                    {
                        Image = new ImageModel
                        {
                            Data = pngBytes,
                            ContentType = "image/png",
                            Width = 96,
                            Height = 48,
                            LayoutType = ImageLayoutType.Floating,
                            VerticalRelativeTo = "paragraph",
                            PositionYTwips = 50
                        }
                    }
                }
            });

            row.Cells.Add(cell);
            table.Rows.Add(row);
            model.Content.Add(table);

            var writer = new DocWriter();
            using var ms = new MemoryStream();
            writer.WriteDocBlocks(model, ms);
            ms.Position = 0;

            using var compoundFile = new OpenMcdf.CompoundFile(ms);
            Assert.True(compoundFile.RootStorage.TryGetStream("WordDocument", out var wordDocStream));
            Assert.True(compoundFile.RootStorage.TryGetStream("1Table", out var tableStream));

            var wordDocData = wordDocStream.GetData();
            var tableData = tableStream.GetData();
            int fcPlcfspaMom = BitConverter.ToInt32(wordDocData, 154 + (40 * 8));
            int recordOffset = fcPlcfspaMom + 8;

            return BitConverter.ToInt32(tableData, recordOffset + 8);
        }

        private static int GetPreferredExplicitWidthImageTop(int preferredWidthTwips, byte[] pngBytes)
        {
            var model = new DocumentModel();
            var table = new TableModel { PreferredWidthValue = preferredWidthTwips, PreferredWidthUnit = TableWidthUnit.Dxa };

            var row = new TableRowModel();
            var cell = new TableCellModel { Width = 5200 };
            cell.Paragraphs.Add(new ParagraphModel
            {
                Runs =
                {
                    new RunModel
                    {
                        Text = "This paragraph should wrap more when preferred dxa width scales explicit cell widths down.",
                        Properties =
                        {
                            FontSize = 24
                        }
                    }
                }
            });
            cell.Paragraphs.Add(new ParagraphModel
            {
                Runs =
                {
                    new RunModel
                    {
                        Image = new ImageModel
                        {
                            Data = pngBytes,
                            ContentType = "image/png",
                            Width = 96,
                            Height = 48,
                            LayoutType = ImageLayoutType.Floating,
                            VerticalRelativeTo = "paragraph",
                            PositionYTwips = 50
                        }
                    }
                }
            });

            row.Cells.Add(cell);
            table.Rows.Add(row);
            model.Content.Add(table);

            var writer = new DocWriter();
            using var ms = new MemoryStream();
            writer.WriteDocBlocks(model, ms);
            ms.Position = 0;

            using var compoundFile = new OpenMcdf.CompoundFile(ms);
            Assert.True(compoundFile.RootStorage.TryGetStream("WordDocument", out var wordDocStream));
            Assert.True(compoundFile.RootStorage.TryGetStream("1Table", out var tableStream));

            var wordDocData = wordDocStream.GetData();
            var tableData = tableStream.GetData();
            int fcPlcfspaMom = BitConverter.ToInt32(wordDocData, 154 + (40 * 8));
            int recordOffset = fcPlcfspaMom + 8;

            return BitConverter.ToInt32(tableData, recordOffset + 8);
        }

        private static int GetPctCellWidthImageTop(int cellWidthValue, byte[] pngBytes)
        {
            var model = new DocumentModel();
            var table = new TableModel();

            var row = new TableRowModel();
            var cell = new TableCellModel { Width = cellWidthValue, WidthUnit = TableWidthUnit.Pct };
            cell.Paragraphs.Add(new ParagraphModel
            {
                Runs =
                {
                    new RunModel
                    {
                        Text = "This paragraph should wrap more when a pct tcW produces a narrower effective cell width.",
                        Properties =
                        {
                            FontSize = 24
                        }
                    }
                }
            });
            cell.Paragraphs.Add(new ParagraphModel
            {
                Runs =
                {
                    new RunModel
                    {
                        Image = new ImageModel
                        {
                            Data = pngBytes,
                            ContentType = "image/png",
                            Width = 96,
                            Height = 48,
                            LayoutType = ImageLayoutType.Floating,
                            VerticalRelativeTo = "paragraph",
                            PositionYTwips = 50
                        }
                    }
                }
            });

            row.Cells.Add(cell);
            table.Rows.Add(row);
            model.Content.Add(table);

            var writer = new DocWriter();
            using var ms = new MemoryStream();
            writer.WriteDocBlocks(model, ms);
            ms.Position = 0;

            using var compoundFile = new OpenMcdf.CompoundFile(ms);
            Assert.True(compoundFile.RootStorage.TryGetStream("WordDocument", out var wordDocStream));
            Assert.True(compoundFile.RootStorage.TryGetStream("1Table", out var tableStream));

            var wordDocData = wordDocStream.GetData();
            var tableData = tableStream.GetData();
            int fcPlcfspaMom = BitConverter.ToInt32(wordDocData, 154 + (40 * 8));
            int recordOffset = fcPlcfspaMom + 8;

            return BitConverter.ToInt32(tableData, recordOffset + 8);
        }

        private static int GetMixedPreferredWidthImageTop(int preferredWidthValue, byte[] pngBytes)
        {
            var model = new DocumentModel();
            var table = new TableModel { PreferredWidthValue = preferredWidthValue, PreferredWidthUnit = TableWidthUnit.Pct };
            table.GridColumnWidths.Add(1800);
            table.GridColumnWidths.Add(3400);

            var row = new TableRowModel();
            var firstCell = new TableCellModel { Width = 1800 };
            firstCell.Paragraphs.Add(new ParagraphModel { Runs = { new RunModel { Text = "Fixed" } } });

            var secondCell = new TableCellModel();
            secondCell.Paragraphs.Add(new ParagraphModel
            {
                Runs =
                {
                    new RunModel
                    {
                        Text = "This paragraph should react when mixed explicit and grid-derived widths are scaled together by tblW.",
                        Properties =
                        {
                            FontSize = 24
                        }
                    }
                }
            });
            secondCell.Paragraphs.Add(new ParagraphModel
            {
                Runs =
                {
                    new RunModel
                    {
                        Image = new ImageModel
                        {
                            Data = pngBytes,
                            ContentType = "image/png",
                            Width = 96,
                            Height = 48,
                            LayoutType = ImageLayoutType.Floating,
                            VerticalRelativeTo = "paragraph",
                            PositionYTwips = 50
                        }
                    }
                }
            });

            row.Cells.Add(firstCell);
            row.Cells.Add(secondCell);
            table.Rows.Add(row);
            model.Content.Add(table);

            var writer = new DocWriter();
            using var ms = new MemoryStream();
            writer.WriteDocBlocks(model, ms);
            ms.Position = 0;

            using var compoundFile = new OpenMcdf.CompoundFile(ms);
            Assert.True(compoundFile.RootStorage.TryGetStream("WordDocument", out var wordDocStream));
            Assert.True(compoundFile.RootStorage.TryGetStream("1Table", out var tableStream));

            var wordDocData = wordDocStream.GetData();
            var tableData = tableStream.GetData();
            int fcPlcfspaMom = BitConverter.ToInt32(wordDocData, 154 + (40 * 8));
            int recordOffset = fcPlcfspaMom + 8;

            return BitConverter.ToInt32(tableData, recordOffset + 8);
        }

        private static int GetAutoRemainingWidthImageTop(int preferredWidthTwips, byte[] pngBytes)
        {
            var model = new DocumentModel();
            var table = new TableModel { PreferredWidthValue = preferredWidthTwips, PreferredWidthUnit = TableWidthUnit.Dxa };

            var row = new TableRowModel();
            var firstCell = new TableCellModel { Width = 2600 };
            firstCell.Paragraphs.Add(new ParagraphModel { Runs = { new RunModel { Text = "Fixed" } } });

            var secondCell = new TableCellModel();
            secondCell.Paragraphs.Add(new ParagraphModel
            {
                Runs =
                {
                    new RunModel
                    {
                        Text = "This paragraph should use the remaining preferred table width when the cell has no tcW or grid width.",
                        Properties =
                        {
                            FontSize = 24
                        }
                    }
                }
            });
            secondCell.Paragraphs.Add(new ParagraphModel
            {
                Runs =
                {
                    new RunModel
                    {
                        Image = new ImageModel
                        {
                            Data = pngBytes,
                            ContentType = "image/png",
                            Width = 96,
                            Height = 48,
                            LayoutType = ImageLayoutType.Floating,
                            VerticalRelativeTo = "paragraph",
                            PositionYTwips = 50
                        }
                    }
                }
            });

            row.Cells.Add(firstCell);
            row.Cells.Add(secondCell);
            table.Rows.Add(row);
            model.Content.Add(table);

            var writer = new DocWriter();
            using var ms = new MemoryStream();
            writer.WriteDocBlocks(model, ms);
            ms.Position = 0;

            using var compoundFile = new OpenMcdf.CompoundFile(ms);
            Assert.True(compoundFile.RootStorage.TryGetStream("WordDocument", out var wordDocStream));
            Assert.True(compoundFile.RootStorage.TryGetStream("1Table", out var tableStream));

            var wordDocData = wordDocStream.GetData();
            var tableData = tableStream.GetData();
            int fcPlcfspaMom = BitConverter.ToInt32(wordDocData, 154 + (40 * 8));
            int recordOffset = fcPlcfspaMom + 8;

            return BitConverter.ToInt32(tableData, recordOffset + 8);
        }

        private static (int firstCellImageTop, int secondCellImageTop) GetAutoCellOverheadAwareImageTops(int preferredWidthTwips, bool addPadding, bool addBorders, byte[] pngBytes)
        {
            var model = new DocumentModel();
            var table = new TableModel { PreferredWidthValue = preferredWidthTwips, PreferredWidthUnit = TableWidthUnit.Dxa };

            var row = new TableRowModel();
            var firstCell = new TableCellModel();
            if (addPadding)
            {
                firstCell.PaddingLeftTwips = 360;
                firstCell.PaddingRightTwips = 360;
                firstCell.HasLeftPaddingOverride = true;
                firstCell.HasRightPaddingOverride = true;
            }

            if (addBorders)
            {
                firstCell.BorderLeftTwips = 120;
                firstCell.BorderRightTwips = 120;
                firstCell.BorderLeftStyle = BorderStyle.Single;
                firstCell.BorderRightStyle = BorderStyle.Single;
                firstCell.HasLeftBorderOverride = true;
                firstCell.HasRightBorderOverride = true;
            }

            firstCell.Paragraphs.Add(new ParagraphModel
            {
                Runs =
                {
                    new RunModel
                    {
                        Text = "This paragraph should keep roughly the same usable width after horizontal overhead is reserved for an auto-width cell.",
                        Properties =
                        {
                            FontSize = 24
                        }
                    }
                }
            });
            firstCell.Paragraphs.Add(new ParagraphModel
            {
                Runs =
                {
                    new RunModel
                    {
                        Image = new ImageModel
                        {
                            Data = pngBytes,
                            ContentType = "image/png",
                            Width = 96,
                            Height = 48,
                            LayoutType = ImageLayoutType.Floating,
                            VerticalRelativeTo = "paragraph",
                            PositionYTwips = 50
                        }
                    }
                }
            });

            var secondCell = new TableCellModel();
            secondCell.Paragraphs.Add(new ParagraphModel
            {
                Runs =
                {
                    new RunModel
                    {
                        Text = "This paragraph should keep roughly the same usable width after horizontal overhead is reserved for an auto-width cell.",
                        Properties =
                        {
                            FontSize = 24
                        }
                    }
                }
            });
            secondCell.Paragraphs.Add(new ParagraphModel
            {
                Runs =
                {
                    new RunModel
                    {
                        Image = new ImageModel
                        {
                            Data = pngBytes,
                            ContentType = "image/png",
                            Width = 96,
                            Height = 48,
                            LayoutType = ImageLayoutType.Floating,
                            VerticalRelativeTo = "paragraph",
                            PositionYTwips = 50
                        }
                    }
                }
            });

            row.Cells.Add(firstCell);
            row.Cells.Add(secondCell);
            table.Rows.Add(row);
            model.Content.Add(table);

            var writer = new DocWriter();
            using var ms = new MemoryStream();
            writer.WriteDocBlocks(model, ms);
            ms.Position = 0;

            using var compoundFile = new OpenMcdf.CompoundFile(ms);
            Assert.True(compoundFile.RootStorage.TryGetStream("WordDocument", out var wordDocStream));
            Assert.True(compoundFile.RootStorage.TryGetStream("1Table", out var tableStream));

            var wordDocData = wordDocStream.GetData();
            var tableData = tableStream.GetData();
            int fcPlcfspaMom = BitConverter.ToInt32(wordDocData, 154 + (40 * 8));
            int firstRecordOffset = fcPlcfspaMom + 8;
            int secondRecordOffset = firstRecordOffset + 26;

            return (
                BitConverter.ToInt32(tableData, firstRecordOffset + 8),
                BitConverter.ToInt32(tableData, secondRecordOffset + 8));
        }

        private static int GetOvercommittedAutoCellImageTop(int preferredWidthTwips, byte[] pngBytes)
        {
            var model = new DocumentModel();
            var table = new TableModel { PreferredWidthValue = preferredWidthTwips, PreferredWidthUnit = TableWidthUnit.Dxa };

            var row = new TableRowModel();
            var firstCell = new TableCellModel { Width = 2200 };
            firstCell.Paragraphs.Add(new ParagraphModel { Runs = { new RunModel { Text = "Fixed 1" } } });

            var secondCell = new TableCellModel { Width = 2200 };
            secondCell.Paragraphs.Add(new ParagraphModel { Runs = { new RunModel { Text = "Fixed 2" } } });

            var thirdCell = new TableCellModel();
            thirdCell.Paragraphs.Add(new ParagraphModel
            {
                Runs =
                {
                    new RunModel
                    {
                        Text = "This auto-width cell should still keep a bounded minimum width when explicit neighbors already exceed the preferred table width.",
                        Properties =
                        {
                            FontSize = 24
                        }
                    }
                }
            });
            thirdCell.Paragraphs.Add(new ParagraphModel
            {
                Runs =
                {
                    new RunModel
                    {
                        Image = new ImageModel
                        {
                            Data = pngBytes,
                            ContentType = "image/png",
                            Width = 96,
                            Height = 48,
                            LayoutType = ImageLayoutType.Floating,
                            VerticalRelativeTo = "paragraph",
                            PositionYTwips = 50
                        }
                    }
                }
            });

            row.Cells.Add(firstCell);
            row.Cells.Add(secondCell);
            row.Cells.Add(thirdCell);
            table.Rows.Add(row);
            model.Content.Add(table);

            var writer = new DocWriter();
            using var ms = new MemoryStream();
            writer.WriteDocBlocks(model, ms);
            ms.Position = 0;

            using var compoundFile = new OpenMcdf.CompoundFile(ms);
            Assert.True(compoundFile.RootStorage.TryGetStream("WordDocument", out var wordDocStream));
            Assert.True(compoundFile.RootStorage.TryGetStream("1Table", out var tableStream));

            var wordDocData = wordDocStream.GetData();
            var tableData = tableStream.GetData();
            int fcPlcfspaMom = BitConverter.ToInt32(wordDocData, 154 + (40 * 8));
            int recordOffset = fcPlcfspaMom + 8;

            return BitConverter.ToInt32(tableData, recordOffset + 8);
        }

        private static int GetOverflowingMixedGridFallbackImageTop(bool secondCellUsesGridFallback, int preferredWidthTwips, byte[] pngBytes)
        {
            var model = new DocumentModel();
            var table = new TableModel { PreferredWidthValue = preferredWidthTwips, PreferredWidthUnit = TableWidthUnit.Dxa };
            table.GridColumnWidths.Add(2500);
            table.GridColumnWidths.Add(2000);
            table.GridColumnWidths.Add(0);

            var row = new TableRowModel();

            var firstCell = new TableCellModel { Width = 2500 };
            firstCell.Paragraphs.Add(new ParagraphModel { Runs = { new RunModel { Text = "Fixed 1" } } });

            var secondCell = secondCellUsesGridFallback
                ? new TableCellModel()
                : new TableCellModel { Width = 2000 };

            secondCell.Paragraphs.Add(new ParagraphModel
            {
                Runs =
                {
                    new RunModel
                    {
                        Text = "This cell should keep the grid fallback width when a later auto cell reserves minimum width and an explicit neighbor must shrink first.",
                        Properties =
                        {
                            FontSize = 24
                        }
                    }
                }
            });
            secondCell.Paragraphs.Add(new ParagraphModel
            {
                Runs =
                {
                    new RunModel
                    {
                        Image = new ImageModel
                        {
                            Data = pngBytes,
                            ContentType = "image/png",
                            Width = 96,
                            Height = 48,
                            LayoutType = ImageLayoutType.Floating,
                            VerticalRelativeTo = "paragraph",
                            PositionYTwips = 50
                        }
                    }
                }
            });

            var thirdCell = new TableCellModel();
            thirdCell.Paragraphs.Add(new ParagraphModel { Runs = { new RunModel { Text = "Auto tail" } } });

            row.Cells.Add(firstCell);
            row.Cells.Add(secondCell);
            row.Cells.Add(thirdCell);
            table.Rows.Add(row);
            model.Content.Add(table);

            var writer = new DocWriter();
            using var ms = new MemoryStream();
            writer.WriteDocBlocks(model, ms);
            ms.Position = 0;

            using var compoundFile = new OpenMcdf.CompoundFile(ms);
            Assert.True(compoundFile.RootStorage.TryGetStream("WordDocument", out var wordDocStream));
            Assert.True(compoundFile.RootStorage.TryGetStream("1Table", out var tableStream));

            var wordDocData = wordDocStream.GetData();
            var tableData = tableStream.GetData();
            int fcPlcfspaMom = BitConverter.ToInt32(wordDocData, 154 + (40 * 8));

            int imageIndex = secondCellUsesGridFallback ? 0 : 0;
            int recordOffset = fcPlcfspaMom + 8 + (imageIndex * 26);

            return BitConverter.ToInt32(tableData, recordOffset + 8);
        }

        private static int GetPctVsDxaOverflowImageTop(bool firstCellUsesPct, int preferredWidthTwips, byte[] pngBytes)
        {
            var model = new DocumentModel();
            var table = new TableModel { PreferredWidthValue = preferredWidthTwips, PreferredWidthUnit = TableWidthUnit.Dxa };

            var row = new TableRowModel();
            var firstCell = new TableCellModel
            {
                Width = 2500,
                WidthUnit = firstCellUsesPct ? TableWidthUnit.Pct : TableWidthUnit.Dxa
            };
            firstCell.Paragraphs.Add(new ParagraphModel
            {
                Runs =
                {
                    new RunModel
                    {
                        Text = "This leading cell should keep more width when it is percentage-based and the row overflows because its dxa neighbor should shrink first.",
                        Properties =
                        {
                            FontSize = 24
                        }
                    }
                }
            });
            firstCell.Paragraphs.Add(new ParagraphModel
            {
                Runs =
                {
                    new RunModel
                    {
                        Image = new ImageModel
                        {
                            Data = pngBytes,
                            ContentType = "image/png",
                            Width = 96,
                            Height = 48,
                            LayoutType = ImageLayoutType.Floating,
                            VerticalRelativeTo = "paragraph",
                            PositionYTwips = 50
                        }
                    }
                }
            });

            var secondCell = new TableCellModel { Width = 2500, WidthUnit = TableWidthUnit.Dxa };
            secondCell.Paragraphs.Add(new ParagraphModel { Runs = { new RunModel { Text = "Fixed neighbor" } } });

            var thirdCell = new TableCellModel();
            thirdCell.Paragraphs.Add(new ParagraphModel { Runs = { new RunModel { Text = "Auto tail" } } });

            row.Cells.Add(firstCell);
            row.Cells.Add(secondCell);
            row.Cells.Add(thirdCell);
            table.Rows.Add(row);
            model.Content.Add(table);

            var writer = new DocWriter();
            using var ms = new MemoryStream();
            writer.WriteDocBlocks(model, ms);
            ms.Position = 0;

            using var compoundFile = new OpenMcdf.CompoundFile(ms);
            Assert.True(compoundFile.RootStorage.TryGetStream("WordDocument", out var wordDocStream));
            Assert.True(compoundFile.RootStorage.TryGetStream("1Table", out var tableStream));

            var wordDocData = wordDocStream.GetData();
            var tableData = tableStream.GetData();
            int fcPlcfspaMom = BitConverter.ToInt32(wordDocData, 154 + (40 * 8));
            int recordOffset = fcPlcfspaMom + 8;

            return BitConverter.ToInt32(tableData, recordOffset + 8);
        }

        private static int GetCellBorderImageTop(int leftBorderTwips, int rightBorderTwips, int topBorderTwips, int bottomBorderTwips, byte[] pngBytes)
        {
            var model = new DocumentModel();
            var table = new TableModel();
            var row = new TableRowModel();
            var cell = new TableCellModel
            {
                Width = 2600,
                BorderLeftTwips = leftBorderTwips,
                BorderRightTwips = rightBorderTwips,
                BorderTopTwips = topBorderTwips,
                BorderBottomTwips = bottomBorderTwips
            };

            cell.Paragraphs.Add(new ParagraphModel
            {
                Runs =
                {
                    new RunModel
                    {
                        Text = "This paragraph is long enough to wrap differently when borders reduce the effective width.",
                        Properties =
                        {
                            FontSize = 24
                        }
                    }
                }
            });

            cell.Paragraphs.Add(new ParagraphModel
            {
                Runs =
                {
                    new RunModel
                    {
                        Image = new ImageModel
                        {
                            Data = pngBytes,
                            ContentType = "image/png",
                            Width = 96,
                            Height = 48,
                            LayoutType = ImageLayoutType.Floating,
                            VerticalRelativeTo = "paragraph",
                            PositionYTwips = 50
                        }
                    }
                }
            });

            row.Cells.Add(cell);
            table.Rows.Add(row);
            model.Content.Add(table);

            var writer = new DocWriter();
            using var ms = new MemoryStream();
            writer.WriteDocBlocks(model, ms);
            ms.Position = 0;

            using var compoundFile = new OpenMcdf.CompoundFile(ms);
            Assert.True(compoundFile.RootStorage.TryGetStream("WordDocument", out var wordDocStream));
            Assert.True(compoundFile.RootStorage.TryGetStream("1Table", out var tableStream));

            var wordDocData = wordDocStream.GetData();
            var tableData = tableStream.GetData();
            int fcPlcfspaMom = BitConverter.ToInt32(wordDocData, 154 + (40 * 8));
            int recordOffset = fcPlcfspaMom + 8;

            return BitConverter.ToInt32(tableData, recordOffset + 8);
        }

        private static int GetAfterTableImageTopWithCellBorder(int topBorderTwips, int bottomBorderTwips, byte[] pngBytes)
        {
            var model = new DocumentModel();
            var table = new TableModel();
            var row = new TableRowModel();
            var cell = new TableCellModel { Width = 2600, BorderTopTwips = topBorderTwips, BorderBottomTwips = bottomBorderTwips };

            cell.Paragraphs.Add(new ParagraphModel
            {
                Runs =
                {
                    new RunModel
                    {
                        Text = "Cell",
                        Properties =
                        {
                            FontSize = 24
                        }
                    }
                }
            });

            row.Cells.Add(cell);
            table.Rows.Add(row);
            model.Content.Add(table);
            model.Content.Add(new ParagraphModel
            {
                Runs =
                {
                    new RunModel { Text = "After table" },
                    new RunModel
                    {
                        Image = new ImageModel
                        {
                            Data = pngBytes,
                            ContentType = "image/png",
                            Width = 96,
                            Height = 48,
                            LayoutType = ImageLayoutType.Floating,
                            VerticalRelativeTo = "paragraph",
                            PositionYTwips = 50
                        }
                    }
                }
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
            int fcPlcfspaMom = BitConverter.ToInt32(wordDocData, 154 + (40 * 8));
            int recordOffset = fcPlcfspaMom + 8;

            return BitConverter.ToInt32(tableData, recordOffset + 8);
        }

        private static int GetAfterTableImageTopWithRowHeight(int rowHeightTwips, TableRowHeightRule heightRule, byte[] pngBytes)
        {
            var model = new DocumentModel();
            var table = new TableModel();
            var row = new TableRowModel { HeightTwips = rowHeightTwips, HeightRule = heightRule };
            var cell = new TableCellModel { Width = 2600 };

            cell.Paragraphs.Add(new ParagraphModel
            {
                Runs =
                {
                    new RunModel
                    {
                        Text = "Cell",
                        Properties =
                        {
                            FontSize = 24
                        }
                    }
                }
            });

            row.Cells.Add(cell);
            table.Rows.Add(row);
            model.Content.Add(table);
            model.Content.Add(new ParagraphModel
            {
                Runs =
                {
                    new RunModel { Text = "After table" },
                    new RunModel
                    {
                        Image = new ImageModel
                        {
                            Data = pngBytes,
                            ContentType = "image/png",
                            Width = 96,
                            Height = 48,
                            LayoutType = ImageLayoutType.Floating,
                            VerticalRelativeTo = "paragraph",
                            PositionYTwips = 50
                        }
                    }
                }
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
            int fcPlcfspaMom = BitConverter.ToInt32(wordDocData, 154 + (40 * 8));
            int recordOffset = fcPlcfspaMom + 8;

            return BitConverter.ToInt32(tableData, recordOffset + 8);
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

        private static int GetAlignedCellImageTopWithRowHeight(int rowHeightTwips, TableRowHeightRule heightRule, TableCellVerticalAlignment alignment, byte[] pngBytes)
        {
            var model = new DocumentModel();
            var table = new TableModel();
            var row = new TableRowModel { HeightTwips = rowHeightTwips, HeightRule = heightRule };

            var contentCell = new TableCellModel { Width = 2600 };
            contentCell.Paragraphs.Add(new ParagraphModel
            {
                Runs =
                {
                    new RunModel
                    {
                        Text = "This cell stays intentionally short so explicit row height creates remaining vertical space.",
                        Properties =
                        {
                            FontSize = 24
                        }
                    }
                }
            });

            var alignedCell = new TableCellModel { Width = 2600, VerticalAlignment = alignment };
            alignedCell.Paragraphs.Add(new ParagraphModel
            {
                Runs =
                {
                    new RunModel
                    {
                        Image = new ImageModel
                        {
                            Data = pngBytes,
                            ContentType = "image/png",
                            Width = 96,
                            Height = 48,
                            LayoutType = ImageLayoutType.Floating,
                            VerticalRelativeTo = "paragraph",
                            PositionYTwips = 50
                        }
                    }
                }
            });

            row.Cells.Add(contentCell);
            row.Cells.Add(alignedCell);
            table.Rows.Add(row);
            model.Content.Add(table);

            var writer = new DocWriter();
            using var ms = new MemoryStream();
            writer.WriteDocBlocks(model, ms);
            ms.Position = 0;

            using var compoundFile = new OpenMcdf.CompoundFile(ms);
            Assert.True(compoundFile.RootStorage.TryGetStream("WordDocument", out var wordDocStream));
            Assert.True(compoundFile.RootStorage.TryGetStream("1Table", out var tableStream));

            var wordDocData = wordDocStream.GetData();
            var tableData = tableStream.GetData();
            int fcPlcfspaMom = BitConverter.ToInt32(wordDocData, 154 + (40 * 8));
            int recordOffset = fcPlcfspaMom + 8;

            return BitConverter.ToInt32(tableData, recordOffset + 8);
        }

        private static int GetAfterTableImageTopWithOptionalEmptyRow(bool includeEmptyRow, byte[] pngBytes)
        {
            var model = new DocumentModel();
            if (includeEmptyRow)
            {
                var table = new TableModel();
                var row = new TableRowModel();
                row.Cells.Add(new TableCellModel { Width = 2600 });
                table.Rows.Add(row);
                model.Content.Add(table);
            }

            model.Content.Add(new ParagraphModel
            {
                Runs =
                {
                    new RunModel { Text = "After table" },
                    new RunModel
                    {
                        Image = new ImageModel
                        {
                            Data = pngBytes,
                            ContentType = "image/png",
                            Width = 96,
                            Height = 48,
                            LayoutType = ImageLayoutType.Floating,
                            VerticalRelativeTo = "paragraph",
                            PositionYTwips = 50
                        }
                    }
                }
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
            int fcPlcfspaMom = BitConverter.ToInt32(wordDocData, 154 + (40 * 8));
            int recordOffset = fcPlcfspaMom + 8;

            return BitConverter.ToInt32(tableData, recordOffset + 8);
        }

        private static int GetAlignedCellImageTopWithRowHeightAndNestedTableContent(int rowHeightTwips, TableRowHeightRule heightRule, TableCellVerticalAlignment alignment, byte[] pngBytes)
        {
            var model = new DocumentModel();
            var table = new TableModel();
            var row = new TableRowModel { HeightTwips = rowHeightTwips, HeightRule = heightRule };

            var contentCell = CreateNestedTableMixedContentCell();

            var alignedCell = new TableCellModel { Width = 2600, VerticalAlignment = alignment };
            alignedCell.Paragraphs.Add(new ParagraphModel
            {
                Runs =
                {
                    new RunModel
                    {
                        Image = new ImageModel
                        {
                            Data = pngBytes,
                            ContentType = "image/png",
                            Width = 96,
                            Height = 48,
                            LayoutType = ImageLayoutType.Floating,
                            VerticalRelativeTo = "paragraph",
                            PositionYTwips = 50
                        }
                    }
                }
            });

            row.Cells.Add(contentCell);
            row.Cells.Add(alignedCell);
            table.Rows.Add(row);
            model.Content.Add(table);

            var writer = new DocWriter();
            using var ms = new MemoryStream();
            writer.WriteDocBlocks(model, ms);
            ms.Position = 0;

            using var compoundFile = new OpenMcdf.CompoundFile(ms);
            Assert.True(compoundFile.RootStorage.TryGetStream("WordDocument", out var wordDocStream));
            Assert.True(compoundFile.RootStorage.TryGetStream("1Table", out var tableStream));

            var wordDocData = wordDocStream.GetData();
            var tableData = tableStream.GetData();
            int fcPlcfspaMom = BitConverter.ToInt32(wordDocData, 154 + (40 * 8));
            int recordOffset = fcPlcfspaMom + 8;

            return BitConverter.ToInt32(tableData, recordOffset + 8);
        }

        private static int GetOverflowClippedImageTop(int rowHeightTwips, TableRowHeightRule heightRule, byte[] pngBytes, int secondParagraphSpacingBeforeTwips = 0, int firstParagraphSpacingAfterTwips = 0, int secondParagraphPositionYTwips = 50)
        {
            var model = new DocumentModel();
            var table = new TableModel();
            var row = new TableRowModel { HeightTwips = rowHeightTwips, HeightRule = heightRule };
            var cell = new TableCellModel { Width = 2000 };

            cell.Paragraphs.Add(new ParagraphModel
            {
                Properties =
                {
                    SpacingAfterTwips = firstParagraphSpacingAfterTwips
                },
                Runs =
                {
                    new RunModel
                    {
                        Text = "This first paragraph is intentionally long enough to create overflow before the next paragraph is laid out.",
                        Properties =
                        {
                            FontSize = 24
                        }
                    }
                }
            });
            cell.Paragraphs.Add(new ParagraphModel
            {
                Properties =
                {
                    SpacingBeforeTwips = secondParagraphSpacingBeforeTwips
                },
                Runs =
                {
                    new RunModel
                    {
                        Image = new ImageModel
                        {
                            Data = pngBytes,
                            ContentType = "image/png",
                            Width = 96,
                            Height = 48,
                            LayoutType = ImageLayoutType.Floating,
                            VerticalRelativeTo = "paragraph",
                            PositionYTwips = secondParagraphPositionYTwips
                        }
                    }
                }
            });

            row.Cells.Add(cell);
            table.Rows.Add(row);
            model.Content.Add(table);

            var writer = new DocWriter();
            using var ms = new MemoryStream();
            writer.WriteDocBlocks(model, ms);
            ms.Position = 0;

            using var compoundFile = new OpenMcdf.CompoundFile(ms);
            Assert.True(compoundFile.RootStorage.TryGetStream("WordDocument", out var wordDocStream));
            Assert.True(compoundFile.RootStorage.TryGetStream("1Table", out var tableStream));

            var wordDocData = wordDocStream.GetData();
            var tableData = tableStream.GetData();
            int fcPlcfspaMom = BitConverter.ToInt32(wordDocData, 154 + (40 * 8));
            int recordOffset = fcPlcfspaMom + 8;

            return BitConverter.ToInt32(tableData, recordOffset + 8);
        }

        private static int GetOverflowClippedHyperlinkImageTop(int rowHeightTwips, TableRowHeightRule heightRule, byte[] pngBytes, int hyperlinkImagePositionYTwips)
        {
            var model = new DocumentModel();
            var table = new TableModel();
            var row = new TableRowModel { HeightTwips = rowHeightTwips, HeightRule = heightRule };
            var cell = new TableCellModel { Width = 2000 };

            cell.Paragraphs.Add(new ParagraphModel
            {
                Runs =
                {
                    new RunModel
                    {
                        Text = "This first paragraph is intentionally long enough to create overflow before the hyperlink image paragraph is laid out.",
                        Properties = { FontSize = 24 }
                    }
                }
            });

            var hyperlink = new HyperlinkModel
            {
                TargetUrl = "https://example.com/image",
                DisplayText = "ImageLink"
            };

            cell.Paragraphs.Add(new ParagraphModel
            {
                Runs =
                {
                    new RunModel
                    {
                        Hyperlink = hyperlink,
                        Image = new ImageModel
                        {
                            Data = pngBytes,
                            ContentType = "image/png",
                            Width = 96,
                            Height = 48,
                            LayoutType = ImageLayoutType.Floating,
                            VerticalRelativeTo = "paragraph",
                            PositionYTwips = hyperlinkImagePositionYTwips
                        }
                    }
                }
            });

            row.Cells.Add(cell);
            table.Rows.Add(row);
            model.Content.Add(table);

            var writer = new DocWriter();
            using var ms = new MemoryStream();
            writer.WriteDocBlocks(model, ms);
            ms.Position = 0;

            using var compoundFile = new OpenMcdf.CompoundFile(ms);
            Assert.True(compoundFile.RootStorage.TryGetStream("WordDocument", out var wordDocStream));
            Assert.True(compoundFile.RootStorage.TryGetStream("1Table", out var tableStream));

            var wordDocData = wordDocStream.GetData();
            var tableData = tableStream.GetData();
            int fcPlcfspaMom = BitConverter.ToInt32(wordDocData, 154 + (40 * 8));
            int recordOffset = fcPlcfspaMom + 8;

            return BitConverter.ToInt32(tableData, recordOffset + 8);
        }

        private static int GetNestedOverflowClippedImageTop(
            int rowHeightTwips,
            TableRowHeightRule heightRule,
            byte[] pngBytes,
            int firstNestedParagraphSpacingAfterTwips = 0,
            int secondNestedParagraphSpacingBeforeTwips = 0,
            int secondNestedParagraphPositionYTwips = 50)
        {
            var model = new DocumentModel();
            var outerTable = new TableModel();
            var outerRow = new TableRowModel { HeightTwips = rowHeightTwips, HeightRule = heightRule };
            var outerCell = new TableCellModel { Width = 2000 };

            var nestedTable = new TableModel();
            var nestedRow = new TableRowModel();
            var nestedCell = new TableCellModel { Width = 1800 };

            nestedCell.Content.Add(new ParagraphModel
            {
                Properties =
                {
                    SpacingAfterTwips = firstNestedParagraphSpacingAfterTwips
                },
                Runs =
                {
                    new RunModel
                    {
                        Text = "This nested paragraph is intentionally long enough to create overflow before the next nested paragraph is laid out.",
                        Properties =
                        {
                            FontSize = 24
                        }
                    }
                }
            });
            nestedCell.Content.Add(new ParagraphModel
            {
                Properties =
                {
                    SpacingBeforeTwips = secondNestedParagraphSpacingBeforeTwips
                },
                Runs =
                {
                    new RunModel
                    {
                        Image = new ImageModel
                        {
                            Data = pngBytes,
                            ContentType = "image/png",
                            Width = 96,
                            Height = 48,
                            LayoutType = ImageLayoutType.Floating,
                            VerticalRelativeTo = "paragraph",
                            PositionYTwips = secondNestedParagraphPositionYTwips
                        }
                    }
                }
            });
            nestedCell.Paragraphs.AddRange(nestedCell.Content.ConvertAll(static c => (ParagraphModel)c));

            nestedRow.Cells.Add(nestedCell);
            nestedTable.Rows.Add(nestedRow);

            outerCell.Content.Add(nestedTable);
            outerRow.Cells.Add(outerCell);
            outerTable.Rows.Add(outerRow);
            model.Content.Add(outerTable);

            var writer = new DocWriter();
            using var ms = new MemoryStream();
            writer.WriteDocBlocks(model, ms);
            ms.Position = 0;

            using var compoundFile = new OpenMcdf.CompoundFile(ms);
            Assert.True(compoundFile.RootStorage.TryGetStream("WordDocument", out var wordDocStream));
            Assert.True(compoundFile.RootStorage.TryGetStream("1Table", out var tableStream));

            var wordDocData = wordDocStream.GetData();
            var tableData = tableStream.GetData();
            int fcPlcfspaMom = BitConverter.ToInt32(wordDocData, 154 + (40 * 8));
            int recordOffset = fcPlcfspaMom + 8;

            return BitConverter.ToInt32(tableData, recordOffset + 8);
        }

        private static int GetNestedOverflowClippedHyperlinkImageTop(int rowHeightTwips, TableRowHeightRule heightRule, byte[] pngBytes, int hyperlinkImagePositionYTwips)
        {
            var model = new DocumentModel();
            var outerTable = new TableModel();
            var outerRow = new TableRowModel { HeightTwips = rowHeightTwips, HeightRule = heightRule };
            var outerCell = new TableCellModel { Width = 2000 };

            var nestedTable = new TableModel();
            var nestedRow = new TableRowModel();
            var nestedCell = new TableCellModel { Width = 1800 };

            nestedCell.Content.Add(new ParagraphModel
            {
                Runs =
                {
                    new RunModel
                    {
                        Text = "This nested paragraph is intentionally long enough to create overflow before the hyperlink image paragraph is laid out.",
                        Properties = { FontSize = 24 }
                    }
                }
            });

            var hyperlink = new HyperlinkModel
            {
                TargetUrl = "https://example.com/nested-image",
                DisplayText = "NestedImageLink"
            };

            nestedCell.Content.Add(new ParagraphModel
            {
                Runs =
                {
                    new RunModel
                    {
                        Hyperlink = hyperlink,
                        Image = new ImageModel
                        {
                            Data = pngBytes,
                            ContentType = "image/png",
                            Width = 96,
                            Height = 48,
                            LayoutType = ImageLayoutType.Floating,
                            VerticalRelativeTo = "paragraph",
                            PositionYTwips = hyperlinkImagePositionYTwips
                        }
                    }
                }
            });
            nestedCell.Paragraphs.AddRange(nestedCell.Content.ConvertAll(static c => (ParagraphModel)c));

            nestedRow.Cells.Add(nestedCell);
            nestedTable.Rows.Add(nestedRow);

            outerCell.Content.Add(nestedTable);
            outerRow.Cells.Add(outerCell);
            outerTable.Rows.Add(outerRow);
            model.Content.Add(outerTable);

            var writer = new DocWriter();
            using var ms = new MemoryStream();
            writer.WriteDocBlocks(model, ms);
            ms.Position = 0;

            using var compoundFile = new OpenMcdf.CompoundFile(ms);
            Assert.True(compoundFile.RootStorage.TryGetStream("WordDocument", out var wordDocStream));
            Assert.True(compoundFile.RootStorage.TryGetStream("1Table", out var tableStream));

            var wordDocData = wordDocStream.GetData();
            var tableData = tableStream.GetData();
            int fcPlcfspaMom = BitConverter.ToInt32(wordDocData, 154 + (40 * 8));
            int recordOffset = fcPlcfspaMom + 8;

            return BitConverter.ToInt32(tableData, recordOffset + 8);
        }

        private static int GetAlignedCellImageTopWithNestedTableContent(TableCellVerticalAlignment alignment, byte[] pngBytes)
        {
            var model = new DocumentModel();
            var table = new TableModel();
            var row = new TableRowModel();

            var contentCell = CreateNestedTableMixedContentCell();

            var alignedCell = new TableCellModel { Width = 2600, VerticalAlignment = alignment };
            alignedCell.Paragraphs.Add(new ParagraphModel
            {
                Runs =
                {
                    new RunModel
                    {
                        Image = new ImageModel
                        {
                            Data = pngBytes,
                            ContentType = "image/png",
                            Width = 96,
                            Height = 48,
                            LayoutType = ImageLayoutType.Floating,
                            VerticalRelativeTo = "paragraph",
                            PositionYTwips = 50
                        }
                    }
                }
            });

            row.Cells.Add(contentCell);
            row.Cells.Add(alignedCell);
            table.Rows.Add(row);
            model.Content.Add(table);

            var writer = new DocWriter();
            using var ms = new MemoryStream();
            writer.WriteDocBlocks(model, ms);
            ms.Position = 0;

            using var compoundFile = new OpenMcdf.CompoundFile(ms);
            Assert.True(compoundFile.RootStorage.TryGetStream("WordDocument", out var wordDocStream));
            Assert.True(compoundFile.RootStorage.TryGetStream("1Table", out var tableStream));

            var wordDocData = wordDocStream.GetData();
            var tableData = tableStream.GetData();
            int fcPlcfspaMom = BitConverter.ToInt32(wordDocData, 154 + (40 * 8));
            int recordOffset = fcPlcfspaMom + 8;

            return BitConverter.ToInt32(tableData, recordOffset + 8);
        }

        private static TableCellModel CreateNestedTableMixedContentCell()
        {
            var contentCell = new TableCellModel { Width = 2600 };

            var leadParagraph = new ParagraphModel();
            leadParagraph.Runs.Add(new RunModel
            {
                Text = "Lead"
            });

            var nestedTable = new TableModel();
            var nestedRow = new TableRowModel();
            var nestedCell = new TableCellModel { Width = 1800 };
            var nestedCellParagraph = new ParagraphModel();
            nestedCellParagraph.Runs.Add(new RunModel
            {
                Text = "Nested cell text"
            });
            nestedCell.Paragraphs.Add(nestedCellParagraph);
            nestedRow.Cells.Add(nestedCell);
            nestedTable.Rows.Add(nestedRow);

            var tailParagraph = new ParagraphModel();
            tailParagraph.Runs.Add(new RunModel
            {
                Text = "Tail"
            });

            contentCell.Content.Add(leadParagraph);
            contentCell.Content.Add(nestedTable);
            contentCell.Content.Add(tailParagraph);

            contentCell.Paragraphs.Add(leadParagraph);
            contentCell.Paragraphs.Add(nestedCellParagraph);
            contentCell.Paragraphs.Add(tailParagraph);

            return contentCell;
        }
    }
}
