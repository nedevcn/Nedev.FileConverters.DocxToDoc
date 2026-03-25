using System.IO;
using System.IO.Compression;
using System.Text;
using Nedev.FileConverters.DocxToDoc.Format;
using Nedev.FileConverters.DocxToDoc.Model;
using Xunit;

namespace Nedev.FileConverters.DocxToDoc.Tests.Format
{
    public class DocWriterTests
    {
        [Fact]
        public void WriteDocBlocks_ValidModel_CreatesValidCFB()
        {
            // Register encoding provider for Windows-1252 
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Arrange
            var model = new DocumentModel();
            model.Content.Add(new ParagraphModel { Runs = { new RunModel { Text = "Hello MS-DOC World!" } } });

            var writer = new DocWriter();
            using var ms = new MemoryStream();

            // Act
            writer.WriteDocBlocks(model, ms);
            ms.Position = 0;

            // Assert
            // We use OpenMcdf just to verify the CFB structure is valid
            using var compoundFile = new OpenMcdf.CompoundFile(ms);
            
            // Should contain WordDocument, 1Table, and Data streams
            Assert.True(compoundFile.RootStorage.TryGetStream("WordDocument", out var wordDocStream));
            Assert.True(compoundFile.RootStorage.TryGetStream("1Table", out var tableStream));
            Assert.True(compoundFile.RootStorage.TryGetStream("Data", out var dataStream));

            // Verify Text was written to WordDocument stream at offset 1536
            var wordDocData = wordDocStream.GetData();
            string expectedText = "Hello MS-DOC World!\r";
            Assert.True(wordDocData.Length >= 1536 + expectedText.Length);

            var textBytes = new byte[expectedText.Length];
            System.Array.Copy(wordDocData, 1536, textBytes, 0, expectedText.Length);
            var extractedText = Encoding.GetEncoding(1252).GetString(textBytes);
            Assert.Equal(expectedText, extractedText);
        }

        [Fact]
        public void WriteDocBlocks_WithFieldRuns_WritesFieldMarkersAndInstruction()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            var field = new FieldModel
            {
                Type = FieldType.Page,
                Instruction = "PAGE",
                IsLocked = true,
                IsDirty = true
            };

            var model = new DocumentModel();
            model.Content.Add(new ParagraphModel
            {
                Runs =
                {
                    new RunModel { Text = "Page " },
                    new RunModel { Field = field, IsFieldBegin = true },
                    new RunModel { Field = field },
                    new RunModel { Field = field, IsFieldSeparate = true },
                    new RunModel { Text = "1" },
                    new RunModel { Field = field, IsFieldEnd = true }
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
            string expectedText = "Page \x0013PAGE\x00141\x0015\r";
            var textBytes = new byte[expectedText.Length];
            System.Array.Copy(wordDocData, 1536, textBytes, 0, expectedText.Length);
            var extractedText = Encoding.GetEncoding(1252).GetString(textBytes);

            Assert.Equal(expectedText, extractedText);

            int fcPlcffldMom = BitConverter.ToInt32(wordDocData, 154 + (15 * 8));
            int lcbPlcffldMom = BitConverter.ToInt32(wordDocData, 154 + (15 * 8) + 4);

            Assert.NotEqual(0, fcPlcffldMom);
            Assert.True(lcbPlcffldMom > 0);
            Assert.Equal(5, BitConverter.ToInt32(tableData, fcPlcffldMom));
            Assert.Equal(10, BitConverter.ToInt32(tableData, fcPlcffldMom + 4));
            Assert.Equal(12, BitConverter.ToInt32(tableData, fcPlcffldMom + 8));
            Assert.Equal(14, BitConverter.ToInt32(tableData, fcPlcffldMom + 12));
            Assert.Equal(0x0B13, BitConverter.ToUInt16(tableData, fcPlcffldMom + 16));
            Assert.Equal(0x0814, BitConverter.ToUInt16(tableData, fcPlcffldMom + 18));
            Assert.Equal(0x0815, BitConverter.ToUInt16(tableData, fcPlcffldMom + 20));
        }

        [Fact]
        public void WriteDocBlocks_WithHyperlinkRuns_WritesHyperlinkFieldSequence()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            var hyperlink = new HyperlinkModel
            {
                TargetUrl = "https://example.com"
            };

            var model = new DocumentModel();
            model.Content.Add(new ParagraphModel
            {
                Runs =
                {
                    new RunModel { Text = "Visit " },
                    new RunModel { Text = "Example Website", Hyperlink = hyperlink },
                    new RunModel { Text = " now." }
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
            string expectedText = "Visit \x0013HYPERLINK \"https://example.com\"\x0014Example Website\x0015 now.\r";
            var textBytes = new byte[expectedText.Length];
            System.Array.Copy(wordDocData, 1536, textBytes, 0, expectedText.Length);
            var extractedText = Encoding.GetEncoding(1252).GetString(textBytes);

            Assert.Equal(expectedText, extractedText);

            int fcPlcffldMom = BitConverter.ToInt32(wordDocData, 154 + (15 * 8));
            int lcbPlcffldMom = BitConverter.ToInt32(wordDocData, 154 + (15 * 8) + 4);

            Assert.NotEqual(0, fcPlcffldMom);
            Assert.True(lcbPlcffldMom > 0);
            Assert.Equal(6, BitConverter.ToInt32(tableData, fcPlcffldMom));
            Assert.Equal(38, BitConverter.ToInt32(tableData, fcPlcffldMom + 4));
            Assert.Equal(54, BitConverter.ToInt32(tableData, fcPlcffldMom + 8));
            Assert.Equal(61, BitConverter.ToInt32(tableData, fcPlcffldMom + 12));
            Assert.Equal(0x0013, BitConverter.ToUInt16(tableData, fcPlcffldMom + 16));
            Assert.Equal(0x0014, BitConverter.ToUInt16(tableData, fcPlcffldMom + 18));
            Assert.Equal(0x0015, BitConverter.ToUInt16(tableData, fcPlcffldMom + 20));
        }

        [Fact]
        public void WriteDocBlocks_WithAnchorHyperlink_WritesInternalHyperlinkInstruction()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            var hyperlink = new HyperlinkModel
            {
                Anchor = "Section1"
            };

            var model = new DocumentModel();
            model.Content.Add(new ParagraphModel
            {
                Runs =
                {
                    new RunModel { Text = "Jump to " },
                    new RunModel { Text = "Section 1", Hyperlink = hyperlink }
                }
            });

            var writer = new DocWriter();
            using var ms = new MemoryStream();

            writer.WriteDocBlocks(model, ms);
            ms.Position = 0;

            using var compoundFile = new OpenMcdf.CompoundFile(ms);
            Assert.True(compoundFile.RootStorage.TryGetStream("WordDocument", out var wordDocStream));

            var wordDocData = wordDocStream.GetData();
            string expectedText = "Jump to \x0013HYPERLINK \\l \"Section1\"\x0014Section 1\x0015\r";
            var textBytes = new byte[expectedText.Length];
            System.Array.Copy(wordDocData, 1536, textBytes, 0, expectedText.Length);
            var extractedText = Encoding.GetEncoding(1252).GetString(textBytes);

            Assert.Equal(expectedText, extractedText);
        }

        [Fact]
        public void WriteDocBlocks_WithImplicitFieldResult_AutoCompletesFieldSequence()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            var field = new FieldModel
            {
                Type = FieldType.Page,
                Result = "7"
            };

            var model = new DocumentModel();
            model.Content.Add(new ParagraphModel
            {
                Runs =
                {
                    new RunModel { Text = "Page " },
                    new RunModel { Field = field, IsFieldBegin = true }
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
            string expectedText = "Page \x0013PAGE\x00147\x0015\r";
            var textBytes = new byte[expectedText.Length];
            System.Array.Copy(wordDocData, 1536, textBytes, 0, expectedText.Length);
            var extractedText = Encoding.GetEncoding(1252).GetString(textBytes);

            Assert.Equal(expectedText, extractedText);

            int fcPlcffldMom = BitConverter.ToInt32(wordDocData, 154 + (15 * 8));

            Assert.Equal(5, BitConverter.ToInt32(tableData, fcPlcffldMom));
            Assert.Equal(10, BitConverter.ToInt32(tableData, fcPlcffldMom + 4));
            Assert.Equal(12, BitConverter.ToInt32(tableData, fcPlcffldMom + 8));
            Assert.Equal(14, BitConverter.ToInt32(tableData, fcPlcffldMom + 12));
            Assert.Equal(0x0C13, BitConverter.ToUInt16(tableData, fcPlcffldMom + 16));
            Assert.Equal(0x0C14, BitConverter.ToUInt16(tableData, fcPlcffldMom + 18));
            Assert.Equal(0x0C15, BitConverter.ToUInt16(tableData, fcPlcffldMom + 20));
        }

        [Fact]
        public void WriteDocBlocks_WithImplicitFieldInstruction_DoesNotDuplicateInstructionRuns()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            var field = new FieldModel
            {
                Instruction = "DATE",
                Result = "2026-03-25"
            };

            var model = new DocumentModel();
            model.Content.Add(new ParagraphModel
            {
                Runs =
                {
                    new RunModel { Field = field, IsFieldBegin = true },
                    new RunModel { Field = field },
                    new RunModel { Text = " suffix" }
                }
            });

            var writer = new DocWriter();
            using var ms = new MemoryStream();

            writer.WriteDocBlocks(model, ms);
            ms.Position = 0;

            using var compoundFile = new OpenMcdf.CompoundFile(ms);
            Assert.True(compoundFile.RootStorage.TryGetStream("WordDocument", out var wordDocStream));

            var wordDocData = wordDocStream.GetData();
            string expectedText = "\x0013DATE\x00142026-03-25\x0015 suffix\r";
            var textBytes = new byte[expectedText.Length];
            System.Array.Copy(wordDocData, 1536, textBytes, 0, expectedText.Length);
            var extractedText = Encoding.GetEncoding(1252).GetString(textBytes);

            Assert.Equal(expectedText, extractedText);
        }

        [Fact]
        public void WriteDocBlocks_WithExplicitEndAndMissingSeparate_WritesFieldResultBeforeEnd()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            var field = new FieldModel
            {
                Instruction = "DATE",
                Result = "2026-03-25"
            };

            var model = new DocumentModel();
            model.Content.Add(new ParagraphModel
            {
                Runs =
                {
                    new RunModel { Field = field, IsFieldBegin = true },
                    new RunModel { Field = field },
                    new RunModel { Field = field, IsFieldEnd = true }
                }
            });

            var writer = new DocWriter();
            using var ms = new MemoryStream();

            writer.WriteDocBlocks(model, ms);
            ms.Position = 0;

            using var compoundFile = new OpenMcdf.CompoundFile(ms);
            Assert.True(compoundFile.RootStorage.TryGetStream("WordDocument", out var wordDocStream));

            var wordDocData = wordDocStream.GetData();
            string expectedText = "\x0013DATE\x00142026-03-25\x0015\r";
            var textBytes = new byte[expectedText.Length];
            System.Array.Copy(wordDocData, 1536, textBytes, 0, expectedText.Length);
            var extractedText = Encoding.GetEncoding(1252).GetString(textBytes);

            Assert.Equal(expectedText, extractedText);
        }

        [Fact]
        public void WriteDocBlocks_WithTwoImplicitFields_WritesSeparateFieldEntriesForEach()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            var pageField = new FieldModel
            {
                Type = FieldType.Page,
                Result = "2"
            };

            var numPagesField = new FieldModel
            {
                Type = FieldType.NumPages,
                Result = "10"
            };

            var model = new DocumentModel();
            model.Content.Add(new ParagraphModel
            {
                Runs =
                {
                    new RunModel { Text = "Page " },
                    new RunModel { Field = pageField, IsFieldBegin = true },
                    new RunModel { Text = " of " },
                    new RunModel { Field = numPagesField, IsFieldBegin = true }
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
            string expectedText = "Page \x0013PAGE\x00142\x0015 of \x0013NUMPAGES\x001410\x0015\r";
            var textBytes = new byte[expectedText.Length];
            System.Array.Copy(wordDocData, 1536, textBytes, 0, expectedText.Length);
            var extractedText = Encoding.GetEncoding(1252).GetString(textBytes);

            Assert.Equal(expectedText, extractedText);

            int fcPlcffldMom = BitConverter.ToInt32(wordDocData, 154 + (15 * 8));
            Assert.Equal(5, BitConverter.ToInt32(tableData, fcPlcffldMom));
            Assert.Equal(10, BitConverter.ToInt32(tableData, fcPlcffldMom + 4));
            Assert.Equal(12, BitConverter.ToInt32(tableData, fcPlcffldMom + 8));
            Assert.Equal(17, BitConverter.ToInt32(tableData, fcPlcffldMom + 12));
            Assert.Equal(26, BitConverter.ToInt32(tableData, fcPlcffldMom + 16));
            Assert.Equal(29, BitConverter.ToInt32(tableData, fcPlcffldMom + 20));
            Assert.Equal(31, BitConverter.ToInt32(tableData, fcPlcffldMom + 24));
            Assert.Equal(0x0C13, BitConverter.ToUInt16(tableData, fcPlcffldMom + 28));
            Assert.Equal(0x0C14, BitConverter.ToUInt16(tableData, fcPlcffldMom + 30));
            Assert.Equal(0x0C15, BitConverter.ToUInt16(tableData, fcPlcffldMom + 32));
            Assert.Equal(0x0C13, BitConverter.ToUInt16(tableData, fcPlcffldMom + 34));
            Assert.Equal(0x0C14, BitConverter.ToUInt16(tableData, fcPlcffldMom + 36));
            Assert.Equal(0x0C15, BitConverter.ToUInt16(tableData, fcPlcffldMom + 38));
        }

        [Fact]
        public void WriteDocBlocks_WithMissingFieldEnd_AutoClosesFieldAtParagraphEnd()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            var field = new FieldModel
            {
                Instruction = "PAGE"
            };

            var model = new DocumentModel();
            model.Content.Add(new ParagraphModel
            {
                Runs =
                {
                    new RunModel { Field = field, IsFieldBegin = true },
                    new RunModel { Field = field },
                    new RunModel { Field = field, IsFieldSeparate = true },
                    new RunModel { Text = "1" }
                }
            });

            var writer = new DocWriter();
            using var ms = new MemoryStream();

            writer.WriteDocBlocks(model, ms);
            ms.Position = 0;

            using var compoundFile = new OpenMcdf.CompoundFile(ms);
            Assert.True(compoundFile.RootStorage.TryGetStream("WordDocument", out var wordDocStream));

            var wordDocData = wordDocStream.GetData();
            string expectedText = "\x0013PAGE\x00141\x0015\r";
            var textBytes = new byte[expectedText.Length];
            System.Array.Copy(wordDocData, 1536, textBytes, 0, expectedText.Length);
            var extractedText = Encoding.GetEncoding(1252).GetString(textBytes);

            Assert.Equal(expectedText, extractedText);
        }

        [Fact]
        public void WriteDocBlocks_WithMissingFieldEndAndStoredResult_AppendsResultBeforeAutoClose()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            var field = new FieldModel
            {
                Instruction = "DATE",
                Result = "2026-03-25"
            };

            var model = new DocumentModel();
            model.Content.Add(new ParagraphModel
            {
                Runs =
                {
                    new RunModel { Field = field, IsFieldBegin = true },
                    new RunModel { Field = field }
                }
            });

            var writer = new DocWriter();
            using var ms = new MemoryStream();

            writer.WriteDocBlocks(model, ms);
            ms.Position = 0;

            using var compoundFile = new OpenMcdf.CompoundFile(ms);
            Assert.True(compoundFile.RootStorage.TryGetStream("WordDocument", out var wordDocStream));

            var wordDocData = wordDocStream.GetData();
            string expectedText = "\x0013DATE\x00142026-03-25\x0015\r";
            var textBytes = new byte[expectedText.Length];
            System.Array.Copy(wordDocData, 1536, textBytes, 0, expectedText.Length);
            var extractedText = Encoding.GetEncoding(1252).GetString(textBytes);

            Assert.Equal(expectedText, extractedText);
        }

        [Fact]
        public void WriteDocBlocks_WithNestedExplicitFields_PreservesNestedFieldOrder()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            var outerField = new FieldModel
            {
                Instruction = "IF"
            };

            var innerField = new FieldModel
            {
                Instruction = "PAGE",
                Result = "3"
            };

            var model = new DocumentModel();
            model.Content.Add(new ParagraphModel
            {
                Runs =
                {
                    new RunModel { Field = outerField, IsFieldBegin = true },
                    new RunModel { Field = outerField },
                    new RunModel { Field = outerField, IsFieldSeparate = true },
                    new RunModel { Text = "Page " },
                    new RunModel { Field = innerField, IsFieldBegin = true },
                    new RunModel { Field = innerField },
                    new RunModel { Field = innerField, IsFieldSeparate = true },
                    new RunModel { Text = "3" },
                    new RunModel { Field = innerField, IsFieldEnd = true },
                    new RunModel { Field = outerField, IsFieldEnd = true }
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
            string expectedText = "\x0013IF\x0014Page \x0013PAGE\x00143\x0015\x0015\r";
            var textBytes = new byte[expectedText.Length];
            System.Array.Copy(wordDocData, 1536, textBytes, 0, expectedText.Length);
            var extractedText = Encoding.GetEncoding(1252).GetString(textBytes);

            Assert.Equal(expectedText, extractedText);

            int fcPlcffldMom = BitConverter.ToInt32(wordDocData, 154 + (15 * 8));
            Assert.Equal(0x0013, BitConverter.ToUInt16(tableData, fcPlcffldMom + 28));
            Assert.Equal(0x0014, BitConverter.ToUInt16(tableData, fcPlcffldMom + 30));
            Assert.Equal(0x1413, BitConverter.ToUInt16(tableData, fcPlcffldMom + 32));
            Assert.Equal(0x1414, BitConverter.ToUInt16(tableData, fcPlcffldMom + 34));
            Assert.Equal(0x1415, BitConverter.ToUInt16(tableData, fcPlcffldMom + 36));
            Assert.Equal(0x0015, BitConverter.ToUInt16(tableData, fcPlcffldMom + 38));
        }

        [Fact]
        public void WriteDocBlocks_WithReusedFieldInstanceAcrossParagraphs_DoesNotLeakCompletionState()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            var reusedField = new FieldModel
            {
                Instruction = "DATE",
                Result = "2026-03-25"
            };

            var model = new DocumentModel();
            model.Content.Add(new ParagraphModel
            {
                Runs =
                {
                    new RunModel { Field = reusedField, IsFieldBegin = true }
                }
            });
            model.Content.Add(new ParagraphModel
            {
                Runs =
                {
                    new RunModel { Text = "Again " },
                    new RunModel { Field = reusedField, IsFieldBegin = true }
                }
            });

            var writer = new DocWriter();
            using var ms = new MemoryStream();

            writer.WriteDocBlocks(model, ms);
            ms.Position = 0;

            using var compoundFile = new OpenMcdf.CompoundFile(ms);
            Assert.True(compoundFile.RootStorage.TryGetStream("WordDocument", out var wordDocStream));

            var wordDocData = wordDocStream.GetData();
            string expectedText = "\x0013DATE\x00142026-03-25\x0015\rAgain \x0013DATE\x00142026-03-25\x0015\r";
            var textBytes = new byte[expectedText.Length];
            System.Array.Copy(wordDocData, 1536, textBytes, 0, expectedText.Length);
            var extractedText = Encoding.GetEncoding(1252).GetString(textBytes);

            Assert.Equal(expectedText, extractedText);
        }

        [Fact]
        public void WriteDocBlocks_WithImageRun_WritesStructuredImageRecordToDataStream()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            byte[] imageBytes = new byte[] { 0x01, 0x23, 0x45, 0x67, 0x89 };
            var model = new DocumentModel();
            model.Content.Add(new ParagraphModel
            {
                Runs =
                {
                    new RunModel { Text = "Before " },
                    new RunModel
                    {
                        Image = new ImageModel
                        {
                            Data = imageBytes,
                            ContentType = "image/png",
                            FileName = "image1.png",
                            Width = 320,
                            Height = 240
                        }
                    },
                    new RunModel { Text = "After" }
                }
            });

            var writer = new DocWriter();
            using var ms = new MemoryStream();

            writer.WriteDocBlocks(model, ms);
            ms.Position = 0;

            using var compoundFile = new OpenMcdf.CompoundFile(ms);
            Assert.True(compoundFile.RootStorage.TryGetStream("WordDocument", out var wordDocStream));
            Assert.True(compoundFile.RootStorage.TryGetStream("Data", out var dataStream));

            var wordDocData = wordDocStream.GetData();
            string expectedText = "Before \x0001After\r";
            var textBytes = new byte[expectedText.Length];
            System.Array.Copy(wordDocData, 1536, textBytes, 0, expectedText.Length);
            var extractedText = Encoding.GetEncoding(1252).GetString(textBytes);
            Assert.Equal(expectedText, extractedText);

            var dataBytes = dataStream.GetData();
            int pictureBlockLength = BitConverter.ToInt32(dataBytes, 0);
            ushort pictureHeaderLength = BitConverter.ToUInt16(dataBytes, 4);
            short mappingMode = BitConverter.ToInt16(dataBytes, 6);
            short blockType = BitConverter.ToInt16(dataBytes, 0x0E);
            short widthGoal = BitConverter.ToInt16(dataBytes, 0x1C);
            short heightGoal = BitConverter.ToInt16(dataBytes, 0x1E);
            ushort scaleX = BitConverter.ToUInt16(dataBytes, 0x20);
            ushort scaleY = BitConverter.ToUInt16(dataBytes, 0x22);

            Assert.Equal(dataBytes.Length, pictureBlockLength);
            Assert.Equal((ushort)0x44, pictureHeaderLength);
            Assert.Equal((short)0x64, mappingMode);
            Assert.Equal((short)0x00, blockType);
            Assert.Equal((short)4800, widthGoal);
            Assert.Equal((short)3600, heightGoal);
            Assert.Equal((ushort)1000, scaleX);
            Assert.Equal((ushort)1000, scaleY);

            var extractedImageBytes = new byte[pictureBlockLength - 0x44];
            System.Array.Copy(dataBytes, 0x44, extractedImageBytes, 0, extractedImageBytes.Length);
            Assert.Equal(imageBytes, extractedImageBytes);

            ushort flags = BitConverter.ToUInt16(wordDocData, 10);
            Assert.True((flags & 0x0008) != 0);
        }

        [Fact]
        public void WriteDocBlocks_WithMultipleImageRuns_WritesDirectoryEntriesInCpOrder()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            var model = new DocumentModel();
            model.Content.Add(new ParagraphModel
            {
                Runs =
                {
                    new RunModel
                    {
                        Image = new ImageModel
                        {
                            Data = new byte[] { 0x01, 0x02 },
                            ContentType = "image/png",
                            FileName = "first.png",
                            Width = 100,
                            Height = 80
                        }
                    },
                    new RunModel { Text = " gap " },
                    new RunModel
                    {
                        Image = new ImageModel
                        {
                            Data = new byte[] { 0x0A, 0x0B, 0x0C },
                            ContentType = "image/jpeg",
                            FileName = "second.jpg",
                            Width = 200,
                            Height = 160
                        }
                    }
                }
            });

            var writer = new DocWriter();
            using var ms = new MemoryStream();

            writer.WriteDocBlocks(model, ms);
            ms.Position = 0;

            using var compoundFile = new OpenMcdf.CompoundFile(ms);
            Assert.True(compoundFile.RootStorage.TryGetStream("Data", out var dataStream));

            var dataBytes = dataStream.GetData();
            int firstLength = BitConverter.ToInt32(dataBytes, 0);
            int secondOffset = firstLength;
            int secondLength = BitConverter.ToInt32(dataBytes, secondOffset);

            Assert.True(firstLength > 0);
            Assert.True(secondLength > 0);
            Assert.Equal(firstLength + secondLength, dataBytes.Length);
            Assert.Equal((short)1500, BitConverter.ToInt16(dataBytes, 0x1C));
            Assert.Equal((short)1200, BitConverter.ToInt16(dataBytes, 0x1E));
            Assert.Equal((short)3000, BitConverter.ToInt16(dataBytes, secondOffset + 0x1C));
            Assert.Equal((short)2400, BitConverter.ToInt16(dataBytes, secondOffset + 0x1E));
        }

        [Fact]
        public void WriteDocBlocks_WithMultipleImageRuns_WritesPictureOffsetsIntoChpx()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            var model = new DocumentModel();
            model.Content.Add(new ParagraphModel
            {
                Runs =
                {
                    new RunModel
                    {
                        Image = new ImageModel
                        {
                            Data = new byte[] { 0x01 },
                            ContentType = "image/png",
                            FileName = "first.png"
                        }
                    },
                    new RunModel { Text = "x" },
                    new RunModel
                    {
                        Image = new ImageModel
                        {
                            Data = new byte[] { 0x02 },
                            ContentType = "image/jpeg",
                            FileName = "second.jpg"
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
            Assert.True(compoundFile.RootStorage.TryGetStream("Data", out var dataStream));

            var wordDocData = wordDocStream.GetData();
            var dataBytes = dataStream.GetData();
            int firstPictureOffset = 0;
            int secondPictureOffset = BitConverter.ToInt32(dataBytes, 0);

            Assert.True(ContainsSubsequence(wordDocData, BuildPictureSprmSequence(firstPictureOffset)));
            Assert.True(ContainsSubsequence(wordDocData, BuildPictureSprmSequence(secondPictureOffset)));
        }

        [Fact]
        public void WriteDocBlocks_WithImageRunAndMissingContentType_DetectsRasterPayload()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            byte[] pngBytes = new byte[]
            {
                0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A,
                0x00, 0x00, 0x00, 0x00
            };

            var model = new DocumentModel();
            model.Content.Add(new ParagraphModel
            {
                Runs =
                {
                    new RunModel
                    {
                        Image = new ImageModel
                        {
                            Data = pngBytes,
                            Width = 32,
                            Height = 16
                        }
                    }
                }
            });

            var writer = new DocWriter();
            using var ms = new MemoryStream();

            writer.WriteDocBlocks(model, ms);
            ms.Position = 0;

            using var compoundFile = new OpenMcdf.CompoundFile(ms);
            Assert.True(compoundFile.RootStorage.TryGetStream("Data", out var dataStream));

            var dataBytes = dataStream.GetData();
            Assert.Equal((short)0x64, BitConverter.ToInt16(dataBytes, 0x06));
            Assert.Equal((short)0x00, BitConverter.ToInt16(dataBytes, 0x0E));
        }

        [Fact]
        public void WriteDocBlocks_WithPngImageRun_WritesOfficeArtDggInfoToTableStream()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            byte[] pngBytes = new byte[]
            {
                0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A,
                0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44, 0x52
            };

            var model = new DocumentModel();
            model.Content.Add(new ParagraphModel
            {
                Runs =
                {
                    new RunModel
                    {
                        Image = new ImageModel
                        {
                            Data = pngBytes,
                            ContentType = "image/png",
                            Width = 32,
                            Height = 16
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
            int fcDggInfo = BitConverter.ToInt32(wordDocData, 154 + (50 * 8));
            int lcbDggInfo = BitConverter.ToInt32(wordDocData, 154 + (50 * 8) + 4);
            int fcPlcfspaMom = BitConverter.ToInt32(wordDocData, 154 + (40 * 8));
            int lcbPlcfspaMom = BitConverter.ToInt32(wordDocData, 154 + (40 * 8) + 4);

            Assert.True(fcDggInfo > 0);
            Assert.True(lcbDggInfo > 0);
            Assert.True(fcPlcfspaMom > 0);
            Assert.True(lcbPlcfspaMom > 0);
            Assert.Equal((ushort)0x000F, BitConverter.ToUInt16(tableData, fcDggInfo));
            Assert.Equal((ushort)0xF000, BitConverter.ToUInt16(tableData, fcDggInfo + 2));
            Assert.True(ContainsSubsequence(tableData, new byte[] { 0x00, 0x0F, 0x00, 0x02, 0xF0 }));
            Assert.True(ContainsSubsequence(tableData, BitConverter.GetBytes((ushort)0xF001)));
            Assert.True(ContainsSubsequence(tableData, BitConverter.GetBytes((ushort)0xF002)));
            Assert.True(ContainsSubsequence(tableData, BitConverter.GetBytes((ushort)0xF004)));
            Assert.True(ContainsSubsequence(tableData, new byte[] { 0x04, 0x41, 0x01, 0x00, 0x00, 0x00 }));
            Assert.Equal(0, BitConverter.ToInt32(tableData, fcPlcfspaMom));
            Assert.Equal(2, BitConverter.ToInt32(tableData, fcPlcfspaMom + 4));
            Assert.Equal(1026, BitConverter.ToInt32(tableData, fcPlcfspaMom + 8));
            Assert.True(ContainsSubsequence(tableData, pngBytes));
        }

        [Fact]
        public void WriteDocBlocks_WithFloatingImageRun_WritesAnchoredBoundsToPlcfspaMom()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            byte[] pngBytes = new byte[]
            {
                0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A,
                0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44, 0x52
            };

            var model = new DocumentModel();
            model.Content.Add(new ParagraphModel
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
                            WrapType = ImageWrapType.Square,
                            PositionXTwips = 1440,
                            PositionYTwips = 720
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

            Assert.Equal(0, BitConverter.ToInt32(tableData, fcPlcfspaMom));
            Assert.Equal(2, BitConverter.ToInt32(tableData, fcPlcfspaMom + 4));
            Assert.Equal(1026, BitConverter.ToInt32(tableData, fcPlcfspaMom + 8));
            Assert.Equal(1440, BitConverter.ToInt32(tableData, fcPlcfspaMom + 12));
            Assert.Equal(720, BitConverter.ToInt32(tableData, fcPlcfspaMom + 16));
            Assert.Equal(2880, BitConverter.ToInt32(tableData, fcPlcfspaMom + 20));
            Assert.Equal(1440, BitConverter.ToInt32(tableData, fcPlcfspaMom + 24));
        }

        [Fact]
        public void WriteDocBlocks_WithAlignedFloatingImage_WritesMarginRelativeBoundsAndFlags()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            byte[] pngBytes = new byte[]
            {
                0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A,
                0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44, 0x52
            };

            var model = new DocumentModel();
            model.Sections.Add(new SectionModel
            {
                PageWidth = 10000,
                PageHeight = 15000,
                MarginLeft = 1000,
                MarginRight = 1000,
                MarginTop = 1200,
                MarginBottom = 1800
            });

            model.Content.Add(new ParagraphModel
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
                            WrapType = ImageWrapType.TopAndBottom,
                            HorizontalRelativeTo = "margin",
                            VerticalRelativeTo = "page",
                            HorizontalAlignment = "center",
                            VerticalAlignment = "bottom",
                            BehindText = true,
                            AllowOverlap = false
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

            Assert.Equal(0, BitConverter.ToInt32(tableData, fcPlcfspaMom));
            Assert.Equal(2, BitConverter.ToInt32(tableData, fcPlcfspaMom + 4));
            Assert.Equal(1026, BitConverter.ToInt32(tableData, fcPlcfspaMom + 8));
            Assert.Equal(4280, BitConverter.ToInt32(tableData, fcPlcfspaMom + 12));
            Assert.Equal(14280, BitConverter.ToInt32(tableData, fcPlcfspaMom + 16));
            Assert.Equal(5720, BitConverter.ToInt32(tableData, fcPlcfspaMom + 20));
            Assert.Equal(15000, BitConverter.ToInt32(tableData, fcPlcfspaMom + 24));
            Assert.Equal(141, BitConverter.ToInt32(tableData, fcPlcfspaMom + 30));
        }

        [Fact]
        public void WriteDocBlocks_WithParagraphRelativeFloatingImage_PreservesOffsetsAndRelativeFlags()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            byte[] pngBytes = new byte[]
            {
                0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A,
                0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44, 0x52
            };

            var model = new DocumentModel();
            model.Sections.Add(new SectionModel
            {
                PageWidth = 10000,
                PageHeight = 15000,
                MarginLeft = 1000,
                MarginRight = 1000,
                MarginTop = 1200,
                MarginBottom = 1800
            });

            model.Content.Add(new ParagraphModel
            {
                Runs =
                {
                    new RunModel { Text = "Anchor paragraph" },
                    new RunModel
                    {
                        Image = new ImageModel
                        {
                            Data = pngBytes,
                            ContentType = "image/png",
                            Width = 96,
                            Height = 48,
                            LayoutType = ImageLayoutType.Floating,
                            WrapType = ImageWrapType.Square,
                            HorizontalRelativeTo = "paragraph",
                            VerticalRelativeTo = "paragraph",
                            PositionXTwips = 360,
                            PositionYTwips = 540,
                            BehindText = false,
                            AllowOverlap = true
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

            Assert.Equal(16, BitConverter.ToInt32(tableData, fcPlcfspaMom));
            Assert.Equal(18, BitConverter.ToInt32(tableData, fcPlcfspaMom + 4));
            Assert.Equal(1026, BitConverter.ToInt32(tableData, fcPlcfspaMom + 8));
            Assert.Equal(360, BitConverter.ToInt32(tableData, fcPlcfspaMom + 12));
            Assert.Equal(1740, BitConverter.ToInt32(tableData, fcPlcfspaMom + 16));
            Assert.Equal(1800, BitConverter.ToInt32(tableData, fcPlcfspaMom + 20));
            Assert.Equal(2460, BitConverter.ToInt32(tableData, fcPlcfspaMom + 24));
            Assert.Equal(338, BitConverter.ToInt32(tableData, fcPlcfspaMom + 30));
        }

        [Fact]
        public void WriteDocBlocks_WithParagraphRelativeImageInLaterParagraph_AccumulatesEstimatedParagraphTop()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            byte[] pngBytes = new byte[]
            {
                0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A,
                0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44, 0x52
            };

            var model = new DocumentModel();
            model.Sections.Add(new SectionModel
            {
                PageWidth = 10000,
                PageHeight = 15000,
                MarginTop = 1200,
                MarginBottom = 1800
            });

            model.Content.Add(new ParagraphModel
            {
                Runs =
                {
                    new RunModel
                    {
                        Text = "Heading",
                        Properties =
                        {
                            FontSize = 48
                        }
                    }
                }
            });

            model.Content.Add(new ParagraphModel
            {
                Runs =
                {
                    new RunModel { Text = "Second paragraph" },
                    new RunModel
                    {
                        Image = new ImageModel
                        {
                            Data = pngBytes,
                            ContentType = "image/png",
                            Width = 96,
                            Height = 48,
                            LayoutType = ImageLayoutType.Floating,
                            WrapType = ImageWrapType.Square,
                            VerticalRelativeTo = "paragraph",
                            PositionYTwips = 300
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

            Assert.Equal(24, BitConverter.ToInt32(tableData, fcPlcfspaMom));
            Assert.Equal(26, BitConverter.ToInt32(tableData, fcPlcfspaMom + 4));
            Assert.Equal(2098, BitConverter.ToInt32(tableData, fcPlcfspaMom + 16));
            Assert.Equal(2818, BitConverter.ToInt32(tableData, fcPlcfspaMom + 24));
        }

        [Fact]
        public void WriteDocBlocks_WithParagraphSpacing_AppliesSpacingToParagraphRelativeImageTop()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            byte[] pngBytes = new byte[]
            {
                0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A,
                0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44, 0x52
            };

            var model = new DocumentModel();
            model.Sections.Add(new SectionModel
            {
                PageWidth = 10000,
                PageHeight = 15000,
                MarginTop = 1200,
                MarginBottom = 1800
            });

            model.Content.Add(new ParagraphModel
            {
                Properties =
                {
                    SpacingAfterTwips = 300
                },
                Runs =
                {
                    new RunModel { Text = "Lead paragraph" }
                }
            });

            model.Content.Add(new ParagraphModel
            {
                Properties =
                {
                    SpacingBeforeTwips = 200,
                    LineSpacing = 480,
                    LineSpacingRule = "auto"
                },
                Runs =
                {
                    new RunModel { Text = "Anchor paragraph" },
                    new RunModel
                    {
                        Image = new ImageModel
                        {
                            Data = pngBytes,
                            ContentType = "image/png",
                            Width = 96,
                            Height = 48,
                            LayoutType = ImageLayoutType.Floating,
                            WrapType = ImageWrapType.Square,
                            VerticalRelativeTo = "paragraph",
                            PositionYTwips = 100
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

            Assert.Equal(31, BitConverter.ToInt32(tableData, fcPlcfspaMom));
            Assert.Equal(33, BitConverter.ToInt32(tableData, fcPlcfspaMom + 4));
            Assert.Equal(2122, BitConverter.ToInt32(tableData, fcPlcfspaMom + 16));
            Assert.Equal(2842, BitConverter.ToInt32(tableData, fcPlcfspaMom + 24));
        }

        [Fact]
        public void WriteDocBlocks_WithNarrowBodyWidth_UsesWrappedParagraphHeightForFloatingImage()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            byte[] pngBytes = new byte[]
            {
                0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A,
                0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44, 0x52
            };

            var model = new DocumentModel();
            model.Sections.Add(new SectionModel
            {
                PageWidth = 4000,
                MarginLeft = 1000,
                MarginRight = 1000,
                MarginTop = 1200,
                MarginBottom = 1800
            });

            model.Content.Add(new ParagraphModel
            {
                Runs =
                {
                    new RunModel
                    {
                        Text = "This paragraph is deliberately long enough to wrap.",
                        Properties =
                        {
                            FontSize = 24
                        }
                    }
                }
            });

            model.Content.Add(new ParagraphModel
            {
                Runs =
                {
                    new RunModel { Text = "Anchor" },
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

            Assert.Equal(2216, BitConverter.ToInt32(tableData, fcPlcfspaMom + 16));
            Assert.Equal(2936, BitConverter.ToInt32(tableData, fcPlcfspaMom + 24));
        }

        [Fact]
        public void WriteDocBlocks_WithIndentedParagraph_UsesReducedWidthForFloatingImage()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            byte[] pngBytes = new byte[]
            {
                0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A,
                0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44, 0x52
            };

            int topWithoutIndent = GetParagraphRelativeImageTop(BuildFloatingParagraphModel(includeIndent: false, pngBytes));
            int topWithIndent = GetParagraphRelativeImageTop(BuildFloatingParagraphModel(includeIndent: true, pngBytes));

            Assert.True(topWithIndent > topWithoutIndent);

            static DocumentModel BuildFloatingParagraphModel(bool includeIndent, byte[] imageBytes)
            {
                var model = new DocumentModel();
                model.Sections.Add(new SectionModel
                {
                    PageWidth = 5000,
                    MarginLeft = 1000,
                    MarginRight = 1000,
                    MarginTop = 1200,
                    MarginBottom = 1800
                });

                var paragraph = new ParagraphModel();
                paragraph.Runs.Add(new RunModel
                {
                    Text = "This paragraph is deliberately long enough to wrap when indentation narrows the measure.",
                    Properties =
                    {
                        FontSize = 24
                    }
                });

                if (includeIndent)
                {
                    paragraph.Properties.LeftIndentTwips = 720;
                    paragraph.Properties.RightIndentTwips = 720;
                    paragraph.Properties.FirstLineIndentTwips = 360;
                }

                model.Content.Add(paragraph);
                model.Content.Add(new ParagraphModel
                {
                    Runs =
                    {
                        new RunModel { Text = "Anchor" },
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

                return model;
            }

        }

        [Fact]
        public void WriteDocBlocks_WithLargerRunFonts_UsesGreaterEstimatedParagraphWidthForFloatingImage()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            byte[] pngBytes = new byte[]
            {
                0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A,
                0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44, 0x52
            };

            int topWithSmallFont = GetParagraphRelativeImageTop(BuildFontSensitiveParagraphModel(20, pngBytes));
            int topWithLargeFont = GetParagraphRelativeImageTop(BuildFontSensitiveParagraphModel(36, pngBytes));

            Assert.True(topWithLargeFont > topWithSmallFont);

            static DocumentModel BuildFontSensitiveParagraphModel(int fontSizeHalfPoints, byte[] imageBytes)
            {
                var model = new DocumentModel();
                model.Sections.Add(new SectionModel
                {
                    PageWidth = 5200,
                    MarginLeft = 1000,
                    MarginRight = 1000,
                    MarginTop = 1200,
                    MarginBottom = 1800
                });

                model.Content.Add(new ParagraphModel
                {
                    Runs =
                    {
                        new RunModel
                        {
                            Text = "Wide letters WWWW mixed with narrow iiiii should still react to actual run font size.",
                            Properties =
                            {
                                FontSize = fontSizeHalfPoints
                            }
                        }
                    }
                });

                model.Content.Add(new ParagraphModel
                {
                    Runs =
                    {
                        new RunModel { Text = "Anchor" },
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

                return model;
            }
        }

        private static int GetParagraphRelativeImageTop(DocumentModel model)
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
            int fcPlcfspaMom = BitConverter.ToInt32(wordDocData, 154 + (40 * 8));

            return BitConverter.ToInt32(tableData, fcPlcfspaMom + 16);
        }

        [Fact]
        public void WriteDocBlocks_WithJpegImageRun_WritesJpegOfficeArtBlipRecord()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            byte[] jpegBytes = new byte[]
            {
                0xFF, 0xD8, 0xFF, 0xE0, 0x00, 0x10, 0x4A, 0x46,
                0x49, 0x46, 0x00, 0x01
            };

            var model = new DocumentModel();
            model.Content.Add(new ParagraphModel
            {
                Runs =
                {
                    new RunModel
                    {
                        Image = new ImageModel
                        {
                            Data = jpegBytes,
                            ContentType = "image/jpeg",
                            Width = 24,
                            Height = 12
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
            int fcDggInfo = BitConverter.ToInt32(wordDocData, 154 + (50 * 8));

            Assert.True(fcDggInfo > 0);
            Assert.True(ContainsSubsequence(tableData, BitConverter.GetBytes((ushort)0xF01D)));
            Assert.True(ContainsSubsequence(tableData, new byte[] { 0x04, 0x41, 0x01, 0x00, 0x00, 0x00 }));
            Assert.True(ContainsSubsequence(tableData, jpegBytes));
        }

        [Fact]
        public void WriteDocBlocks_WithEmfImageRun_WritesMetafileOfficeArtBlipRecord()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            byte[] emfBytes = new byte[]
            {
                0x01, 0x00, 0x00, 0x00, 0x20, 0x45, 0x4D, 0x46,
                0x00, 0x00, 0x00, 0x00
            };

            var model = new DocumentModel();
            model.Content.Add(new ParagraphModel
            {
                Runs =
                {
                    new RunModel
                    {
                        Image = new ImageModel
                        {
                            Data = emfBytes,
                            ContentType = "image/x-emf",
                            Width = 24,
                            Height = 12
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
            int fcDggInfo = BitConverter.ToInt32(wordDocData, 154 + (50 * 8));
            int blipOffset = FindSubsequenceOffset(tableData, BitConverter.GetBytes((ushort)0xF01A));

            Assert.True(fcDggInfo > 0);
            Assert.True(blipOffset >= 0);
            Assert.True(ContainsSubsequence(tableData, new byte[] { 0x04, 0x41, 0x01, 0x00, 0x00, 0x00 }));
            Assert.Equal(12, BitConverter.ToInt32(tableData, blipOffset + 24));
            Assert.Equal(360, BitConverter.ToInt32(tableData, blipOffset + 36));
            Assert.Equal(180, BitConverter.ToInt32(tableData, blipOffset + 40));
            Assert.Equal(228600, BitConverter.ToInt32(tableData, blipOffset + 44));
            Assert.Equal(114300, BitConverter.ToInt32(tableData, blipOffset + 48));
            Assert.Equal(0x00, tableData[blipOffset + 56]);
            Assert.Equal(0xFE, tableData[blipOffset + 57]);
            Assert.Equal(emfBytes, InflateMetafilePayload(tableData, blipOffset));
        }

        [Fact]
        public void WriteDocBlocks_WithWmfImageRun_WritesMetafileOfficeArtBlipRecord()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            byte[] wmfBytes = new byte[]
            {
                0xD7, 0xCD, 0xC6, 0x9A, 0x00, 0x00, 0x00, 0x00,
                0x00, 0x00, 0x00, 0x00, 0x10, 0x00, 0x10, 0x00,
                0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
                0x03, 0x00, 0x09, 0x00
            };

            var model = new DocumentModel();
            model.Content.Add(new ParagraphModel
            {
                Runs =
                {
                    new RunModel
                    {
                        Image = new ImageModel
                        {
                            Data = wmfBytes,
                            ContentType = "image/x-wmf",
                            Width = 24,
                            Height = 12
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
            int fcDggInfo = BitConverter.ToInt32(wordDocData, 154 + (50 * 8));
            int blipOffset = FindSubsequenceOffset(tableData, BitConverter.GetBytes((ushort)0xF01B));
            byte[] expectedPayload = new byte[] { 0x03, 0x00, 0x09, 0x00 };

            Assert.True(fcDggInfo > 0);
            Assert.True(blipOffset >= 0);
            Assert.True(ContainsSubsequence(tableData, new byte[] { 0x04, 0x41, 0x01, 0x00, 0x00, 0x00 }));
            Assert.Equal(expectedPayload.Length, BitConverter.ToInt32(tableData, blipOffset + 24));
            Assert.Equal(360, BitConverter.ToInt32(tableData, blipOffset + 36));
            Assert.Equal(180, BitConverter.ToInt32(tableData, blipOffset + 40));
            Assert.Equal(228600, BitConverter.ToInt32(tableData, blipOffset + 44));
            Assert.Equal(114300, BitConverter.ToInt32(tableData, blipOffset + 48));
            Assert.Equal(0x00, tableData[blipOffset + 56]);
            Assert.Equal(0xFE, tableData[blipOffset + 57]);
            Assert.Equal(expectedPayload, InflateMetafilePayload(tableData, blipOffset));
        }

        [Fact]
        public void WriteDocBlocks_WithWmfImageRunAndMissingDimensions_InferTwipsFromPlaceableHeader()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            byte[] wmfBytes = new byte[]
            {
                0xD7, 0xCD, 0xC6, 0x9A,
                0x00, 0x00,
                0x00, 0x00,
                0x00, 0x00,
                0xA0, 0x05,
                0xD0, 0x02,
                0xA0, 0x05,
                0x00, 0x00, 0x00, 0x00,
                0x00, 0x00,
                0x03, 0x00, 0x09, 0x00
            };

            var model = new DocumentModel();
            model.Content.Add(new ParagraphModel
            {
                Runs =
                {
                    new RunModel
                    {
                        Image = new ImageModel
                        {
                            Data = wmfBytes,
                            ContentType = "image/x-wmf"
                        }
                    }
                }
            });

            var writer = new DocWriter();
            using var ms = new MemoryStream();

            writer.WriteDocBlocks(model, ms);
            ms.Position = 0;

            using var compoundFile = new OpenMcdf.CompoundFile(ms);
            Assert.True(compoundFile.RootStorage.TryGetStream("Data", out var dataStream));
            Assert.True(compoundFile.RootStorage.TryGetStream("WordDocument", out var wordDocStream));
            Assert.True(compoundFile.RootStorage.TryGetStream("1Table", out var tableStream));

            var dataBytes = dataStream.GetData();
            var wordDocData = wordDocStream.GetData();
            var tableData = tableStream.GetData();
            int blipOffset = FindSubsequenceOffset(tableData, BitConverter.GetBytes((ushort)0xF01B));

            Assert.Equal((short)1440, BitConverter.ToInt16(dataBytes, 0x1C));
            Assert.Equal((short)720, BitConverter.ToInt16(dataBytes, 0x1E));
            Assert.Equal(1440, BitConverter.ToInt32(tableData, blipOffset + 36));
            Assert.Equal(720, BitConverter.ToInt32(tableData, blipOffset + 40));
            Assert.Equal(914400, BitConverter.ToInt32(tableData, blipOffset + 44));
            Assert.Equal(457200, BitConverter.ToInt32(tableData, blipOffset + 48));
        }

        [Fact]
        public void WriteDocBlocks_WithEmfImageRunAndMissingDimensions_InferTwipsFromFrameHeader()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            byte[] emfBytes = new byte[]
            {
                0x01, 0x00, 0x00, 0x00,
                0x2C, 0x00, 0x00, 0x00,
                0x00, 0x00, 0x00, 0x00,
                0x00, 0x00, 0x00, 0x00,
                0x00, 0x00, 0x00, 0x00,
                0x00, 0x00, 0x00, 0x00,
                0x00, 0x00, 0x00, 0x00,
                0x00, 0x00, 0x00, 0x00,
                0xEC, 0x09, 0x00, 0x00,
                0xF6, 0x04, 0x00, 0x00,
                0x20, 0x45, 0x4D, 0x46
            };

            var model = new DocumentModel();
            model.Content.Add(new ParagraphModel
            {
                Runs =
                {
                    new RunModel
                    {
                        Image = new ImageModel
                        {
                            Data = emfBytes,
                            ContentType = "image/x-emf"
                        }
                    }
                }
            });

            var writer = new DocWriter();
            using var ms = new MemoryStream();

            writer.WriteDocBlocks(model, ms);
            ms.Position = 0;

            using var compoundFile = new OpenMcdf.CompoundFile(ms);
            Assert.True(compoundFile.RootStorage.TryGetStream("Data", out var dataStream));
            Assert.True(compoundFile.RootStorage.TryGetStream("1Table", out var tableStream));

            var dataBytes = dataStream.GetData();
            var tableData = tableStream.GetData();
            int blipOffset = FindSubsequenceOffset(tableData, BitConverter.GetBytes((ushort)0xF01A));

            Assert.Equal((short)1440, BitConverter.ToInt16(dataBytes, 0x1C));
            Assert.Equal((short)720, BitConverter.ToInt16(dataBytes, 0x1E));
            Assert.Equal(1440, BitConverter.ToInt32(tableData, blipOffset + 36));
            Assert.Equal(720, BitConverter.ToInt32(tableData, blipOffset + 40));
            Assert.Equal(914400, BitConverter.ToInt32(tableData, blipOffset + 44));
            Assert.Equal(457200, BitConverter.ToInt32(tableData, blipOffset + 48));
        }

        [Fact]
        public void WriteDocBlocks_WithMetafileSignatureAndMissingContentType_DetectsMetafilePayload()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            byte[] emfBytes = new byte[] { 0x01, 0x00, 0x00, 0x00, 0x20, 0x45, 0x4D, 0x46 };
            var model = new DocumentModel();
            model.Content.Add(new ParagraphModel
            {
                Runs =
                {
                    new RunModel
                    {
                        Image = new ImageModel
                        {
                            Data = emfBytes,
                            Width = 40,
                            Height = 20
                        }
                    }
                }
            });

            var writer = new DocWriter();
            using var ms = new MemoryStream();

            writer.WriteDocBlocks(model, ms);
            ms.Position = 0;

            using var compoundFile = new OpenMcdf.CompoundFile(ms);
            Assert.True(compoundFile.RootStorage.TryGetStream("Data", out var dataStream));

            var dataBytes = dataStream.GetData();
            Assert.Equal((short)8, BitConverter.ToInt16(dataBytes, 0x06));
            Assert.Equal((short)0x08, BitConverter.ToInt16(dataBytes, 0x0E));
        }

        private static bool ContainsSubsequence(byte[] buffer, byte[] subsequence)
        {
            if (subsequence.Length == 0 || subsequence.Length > buffer.Length)
            {
                return false;
            }

            for (int start = 0; start <= buffer.Length - subsequence.Length; start++)
            {
                bool matches = true;
                for (int index = 0; index < subsequence.Length; index++)
                {
                    if (buffer[start + index] != subsequence[index])
                    {
                        matches = false;
                        break;
                    }
                }

                if (matches)
                {
                    return true;
                }
            }

            return false;
        }

        private static int FindSubsequenceOffset(byte[] buffer, byte[] subsequence)
        {
            if (subsequence.Length == 0 || subsequence.Length > buffer.Length)
            {
                return -1;
            }

            for (int start = 0; start <= buffer.Length - subsequence.Length; start++)
            {
                bool matches = true;
                for (int index = 0; index < subsequence.Length; index++)
                {
                    if (buffer[start + index] != subsequence[index])
                    {
                        matches = false;
                        break;
                    }
                }

                if (matches)
                {
                    return start - 2;
                }
            }

            return -1;
        }

        private static byte[] InflateMetafilePayload(byte[] tableData, int blipOffset)
        {
            int compressedSize = BitConverter.ToInt32(tableData, blipOffset + 52);
            using var input = new MemoryStream(tableData, blipOffset + 58, compressedSize);
            using var inflater = new DeflateStream(input, CompressionMode.Decompress);
            using var output = new MemoryStream();
            inflater.CopyTo(output);
            return output.ToArray();
        }

        private static byte[] BuildPictureSprmSequence(int pictureOffset)
        {
            var sequence = new byte[9];
            sequence[0] = 0x55;
            sequence[1] = 0x08;
            sequence[2] = 0x01;
            sequence[3] = 0x03;
            sequence[4] = 0x6A;
            System.Array.Copy(BitConverter.GetBytes(pictureOffset), 0, sequence, 5, 4);
            return sequence;
        }
    }
}
