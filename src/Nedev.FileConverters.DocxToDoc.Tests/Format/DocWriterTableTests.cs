using System;
using System.IO;
using System.Text;
using System.Collections.Generic;
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
                int fcPlcfbteTapx = BitConverter.ToInt32(wordDocData, 186);
                int lcbPlcfbteTapx = BitConverter.ToInt32(wordDocData, 190);

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
    }
}
