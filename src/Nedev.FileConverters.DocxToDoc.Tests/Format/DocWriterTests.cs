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
        public void WriteDocBlocks_WithHyperlinkImageRun_WritesImagePlaceholderInsideHyperlinkField()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            var hyperlink = new HyperlinkModel
            {
                TargetUrl = "https://example.com/image"
            };

            var imageData = new byte[]
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

            var model = new DocumentModel();
            model.Content.Add(new ParagraphModel
            {
                Runs =
                {
                    new RunModel { Text = "Visit " },
                    new RunModel
                    {
                        Hyperlink = hyperlink,
                        Image = new ImageModel
                        {
                            Data = imageData,
                            ContentType = "image/png",
                            Width = 1,
                            Height = 1
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
            Assert.True(compoundFile.RootStorage.TryGetStream("Data", out var dataStream));

            var wordDocData = wordDocStream.GetData();
            var tableData = tableStream.GetData();
            string expectedText = "Visit \x0013HYPERLINK \"https://example.com/image\"\x0014\x0001\x0015\r";
            var textBytes = new byte[expectedText.Length];
            System.Array.Copy(wordDocData, 1536, textBytes, 0, expectedText.Length);
            var extractedText = Encoding.GetEncoding(1252).GetString(textBytes);

            Assert.Equal(expectedText, extractedText);
            Assert.True(dataStream.GetData().Length > 0);

            int fcPlcffldMom = BitConverter.ToInt32(wordDocData, 154 + (15 * 8));
            int beginCp = expectedText.IndexOf('\x0013');
            int separateCp = expectedText.IndexOf('\x0014');
            int endCp = expectedText.IndexOf('\x0015');
            int paragraphEndCp = expectedText.Length;

            Assert.Equal(beginCp, BitConverter.ToInt32(tableData, fcPlcffldMom));
            Assert.Equal(separateCp, BitConverter.ToInt32(tableData, fcPlcffldMom + 4));
            Assert.Equal(endCp, BitConverter.ToInt32(tableData, fcPlcffldMom + 8));
            Assert.Equal(paragraphEndCp, BitConverter.ToInt32(tableData, fcPlcffldMom + 12));
            Assert.Equal(0x0013, BitConverter.ToUInt16(tableData, fcPlcffldMom + 16));
            Assert.Equal(0x0014, BitConverter.ToUInt16(tableData, fcPlcffldMom + 18));
            Assert.Equal(0x0015, BitConverter.ToUInt16(tableData, fcPlcffldMom + 20));
        }

        [Fact]
        public void WriteDocBlocks_WithPlainComment_WritesCommentReferenceAndAnnotationStory()
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
                Initials = "JD",
                Text = "Note",
                StartCp = 2,
                EndCp = 2
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
            string expectedText = "AB\x0005CD\r\x0005Note\r";
            var textBytes = new byte[expectedText.Length];
            System.Array.Copy(wordDocData, 1536, textBytes, 0, expectedText.Length);
            var extractedText = Encoding.GetEncoding(1252).GetString(textBytes);

            Assert.Equal(expectedText, extractedText);
            Assert.Equal(4, Fib.CommentReferencePairIndex);
            Assert.Equal(5, Fib.CommentTextPairIndex);
            Assert.Equal(6, BitConverter.ToInt32(wordDocData, 64));
            Assert.Equal(6, BitConverter.ToInt32(wordDocData, 76));

            int fcPlcfandRef = BitConverter.ToInt32(wordDocData, 154 + (Fib.CommentReferencePairIndex * 8));
            int lcbPlcfandRef = BitConverter.ToInt32(wordDocData, 154 + (Fib.CommentReferencePairIndex * 8) + 4);
            int fcPlcfandTxt = BitConverter.ToInt32(wordDocData, 154 + (Fib.CommentTextPairIndex * 8));
            int lcbPlcfandTxt = BitConverter.ToInt32(wordDocData, 154 + (Fib.CommentTextPairIndex * 8) + 4);

            Assert.NotEqual(0, fcPlcfandRef);
            Assert.True(lcbPlcfandRef > 0);
            Assert.NotEqual(0, fcPlcfandTxt);
            Assert.True(lcbPlcfandTxt > 0);
            Assert.Equal(2, BitConverter.ToInt32(tableData, fcPlcfandRef));
            Assert.Equal(6, BitConverter.ToInt32(tableData, fcPlcfandRef + 4));
            Assert.Equal(0, BitConverter.ToInt32(tableData, fcPlcfandTxt));
            Assert.Equal(5, BitConverter.ToInt32(tableData, fcPlcfandTxt + 4));
            Assert.Equal(6, BitConverter.ToInt32(tableData, fcPlcfandTxt + 8));
        }

        [Fact]
        public void WriteDocBlocks_WithMultiParagraphComment_PreservesInternalParagraphMarksInAnnotationStory()
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
                Initials = "JD",
                Text = "First\rSecond",
                StartCp = 2,
                EndCp = 2
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
            string expectedText = "AB\x0005CD\r\x0005First\rSecond\r";
            var textBytes = new byte[expectedText.Length];
            System.Array.Copy(wordDocData, 1536, textBytes, 0, expectedText.Length);
            var extractedText = Encoding.GetEncoding(1252).GetString(textBytes);

            Assert.Equal(expectedText, extractedText);
            Assert.Equal(6, BitConverter.ToInt32(wordDocData, 64));
            Assert.Equal(14, BitConverter.ToInt32(wordDocData, 76));

            int fcPlcfandTxt = BitConverter.ToInt32(wordDocData, 154 + (Fib.CommentTextPairIndex * 8));
            Assert.Equal(0, BitConverter.ToInt32(tableData, fcPlcfandTxt));
            Assert.Equal(13, BitConverter.ToInt32(tableData, fcPlcfandTxt + 4));
            Assert.Equal(14, BitConverter.ToInt32(tableData, fcPlcfandTxt + 8));
        }

        [Fact]
        public void WriteDocBlocks_WithTrailingEmptyCommentParagraph_PreservesFinalEmptyParagraph()
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
                Initials = "JD",
                Text = "First\r",
                StartCp = 2,
                EndCp = 2
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
            string expectedText = "AB\x0005CD\r\x0005First\r\r";
            var textBytes = new byte[expectedText.Length];
            System.Array.Copy(wordDocData, 1536, textBytes, 0, expectedText.Length);
            var extractedText = Encoding.GetEncoding(1252).GetString(textBytes);

            Assert.Equal(expectedText, extractedText);
            Assert.Equal(6, BitConverter.ToInt32(wordDocData, 64));
            Assert.Equal(8, BitConverter.ToInt32(wordDocData, 76));

            int fcPlcfandTxt = BitConverter.ToInt32(wordDocData, 154 + (Fib.CommentTextPairIndex * 8));
            Assert.Equal(0, BitConverter.ToInt32(tableData, fcPlcfandTxt));
            Assert.Equal(7, BitConverter.ToInt32(tableData, fcPlcfandTxt + 4));
            Assert.Equal(8, BitConverter.ToInt32(tableData, fcPlcfandTxt + 8));
        }

        [Fact]
        public void WriteDocBlocks_WithCommentAnchoredAtDocumentEnd_EmitsReferenceAtFinalCp()
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

            var writer = new DocWriter();
            using var ms = new MemoryStream();

            writer.WriteDocBlocks(model, ms);
            ms.Position = 0;

            using var compoundFile = new OpenMcdf.CompoundFile(ms);
            Assert.True(compoundFile.RootStorage.TryGetStream("WordDocument", out var wordDocStream));
            Assert.True(compoundFile.RootStorage.TryGetStream("1Table", out var tableStream));

            var wordDocData = wordDocStream.GetData();
            var tableData = tableStream.GetData();
            string expectedText = "ABCD\r\x0005\x0005Tail\r";
            var textBytes = new byte[expectedText.Length];
            System.Array.Copy(wordDocData, 1536, textBytes, 0, expectedText.Length);
            var extractedText = Encoding.GetEncoding(1252).GetString(textBytes);

            Assert.Equal(expectedText, extractedText);
            Assert.Equal(6, BitConverter.ToInt32(wordDocData, 64));
            Assert.Equal(6, BitConverter.ToInt32(wordDocData, 76));

            int fcPlcfandRef = BitConverter.ToInt32(wordDocData, 154 + (Fib.CommentReferencePairIndex * 8));
            Assert.Equal(5, BitConverter.ToInt32(tableData, fcPlcfandRef));
            Assert.Equal(6, BitConverter.ToInt32(tableData, fcPlcfandRef + 4));
        }

        [Fact]
        public void WriteDocBlocks_WithOutOfRangeCommentAnchor_ClampsReferenceToDocumentEnd()
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
                Text = "Late",
                StartCp = 99,
                EndCp = 99
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
            string expectedText = "ABCD\r\x0005\x0005Late\r";
            var textBytes = new byte[expectedText.Length];
            System.Array.Copy(wordDocData, 1536, textBytes, 0, expectedText.Length);
            var extractedText = Encoding.GetEncoding(1252).GetString(textBytes);

            Assert.Equal(expectedText, extractedText);

            int fcPlcfandRef = BitConverter.ToInt32(wordDocData, 154 + (Fib.CommentReferencePairIndex * 8));
            Assert.Equal(5, BitConverter.ToInt32(tableData, fcPlcfandRef));
            Assert.Equal(6, BitConverter.ToInt32(tableData, fcPlcfandRef + 4));
        }

        [Fact]
        public void WriteDocBlocks_WithAuthorOnlyComment_DerivesAtrdInitialsFromAuthor()
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
                Author = "Jane Smith",
                Text = "Note",
                StartCp = 2,
                EndCp = 2
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
            int fcPlcfandRef = BitConverter.ToInt32(wordDocData, 154 + (Fib.CommentReferencePairIndex * 8));
            string initials = Encoding.Unicode.GetString(tableData, fcPlcfandRef + 8, 20).TrimEnd('\0');

            Assert.Equal("JS", initials);
        }

        [Fact]
        public void WriteDocBlocks_WithExplicitInitialsAndAuthor_PreservesExplicitInitials()
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
                Author = "Jane Smith",
                Initials = "QA",
                Text = "Note",
                StartCp = 2,
                EndCp = 2
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
            int fcPlcfandRef = BitConverter.ToInt32(wordDocData, 154 + (Fib.CommentReferencePairIndex * 8));
            string initials = Encoding.Unicode.GetString(tableData, fcPlcfandRef + 8, 20).TrimEnd('\0');

            Assert.Equal("QA", initials);
        }

        [Fact]
        public void WriteDocBlocks_WithInvalidAnchorComments_SkipsUnsupportedEmission()
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
                Text = "Negative",
                StartCp = -1,
                EndCp = 0
            });
            model.Comments.Add(new CommentModel
            {
                Id = "1",
                Text = "Inverted",
                StartCp = 3,
                EndCp = 2
            });

            var writer = new DocWriter();
            using var ms = new MemoryStream();

            writer.WriteDocBlocks(model, ms);
            ms.Position = 0;

            using var compoundFile = new OpenMcdf.CompoundFile(ms);
            Assert.True(compoundFile.RootStorage.TryGetStream("WordDocument", out var wordDocStream));
            Assert.True(compoundFile.RootStorage.TryGetStream("1Table", out var tableStream));

            var wordDocData = wordDocStream.GetData();
            string expectedText = "ABCD\r";
            var textBytes = new byte[expectedText.Length];
            System.Array.Copy(wordDocData, 1536, textBytes, 0, expectedText.Length);
            var extractedText = Encoding.GetEncoding(1252).GetString(textBytes);

            Assert.Equal(expectedText, extractedText);
            Assert.Equal(0, BitConverter.ToInt32(wordDocData, 76));
            Assert.Equal(0, BitConverter.ToInt32(wordDocData, 154 + (Fib.CommentReferencePairIndex * 8)));
            Assert.Equal(0, BitConverter.ToInt32(wordDocData, 154 + (Fib.CommentReferencePairIndex * 8) + 4));
            Assert.Equal(0, BitConverter.ToInt32(wordDocData, 154 + (Fib.CommentTextPairIndex * 8)));
            Assert.Equal(0, BitConverter.ToInt32(wordDocData, 154 + (Fib.CommentTextPairIndex * 8) + 4));
        }

        [Fact]
        public void WriteDocBlocks_WithTableAndComment_WritesTapxAndAnnotationPairsWithoutCollision()
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

            var table = new TableModel();
            var row = new TableRowModel();
            var cell1 = new TableCellModel { Width = 5000 };
            var cell2 = new TableCellModel { Width = 5000 };
            cell1.Paragraphs.Add(new ParagraphModel { Runs = { new RunModel { Text = "Cell 1" } } });
            cell2.Paragraphs.Add(new ParagraphModel { Runs = { new RunModel { Text = "Cell 2" } } });
            row.Cells.Add(cell1);
            row.Cells.Add(cell2);
            table.Rows.Add(row);
            model.Content.Add(table);

            model.Comments.Add(new CommentModel
            {
                Id = "0",
                Text = "Note",
                StartCp = 2,
                EndCp = 2
            });

            var writer = new DocWriter();
            using var ms = new MemoryStream();

            writer.WriteDocBlocks(model, ms);
            ms.Position = 0;

            using var compoundFile = new OpenMcdf.CompoundFile(ms);
            Assert.True(compoundFile.RootStorage.TryGetStream("WordDocument", out var wordDocStream));
            Assert.True(compoundFile.RootStorage.TryGetStream("1Table", out var tableStream));

            var wordDocData = wordDocStream.GetData();
            int fcPlcfandRef = BitConverter.ToInt32(wordDocData, 154 + (4 * 8));
            int lcbPlcfandRef = BitConverter.ToInt32(wordDocData, 154 + (4 * 8) + 4);
            int fcPlcfandTxt = BitConverter.ToInt32(wordDocData, 154 + (5 * 8));
            int lcbPlcfandTxt = BitConverter.ToInt32(wordDocData, 154 + (5 * 8) + 4);
            int fcPlcfbteTapx = BitConverter.ToInt32(wordDocData, 154 + (Fib.TapxPairIndex * 8));
            int lcbPlcfbteTapx = BitConverter.ToInt32(wordDocData, 154 + (Fib.TapxPairIndex * 8) + 4);

            Assert.True(lcbPlcfandRef > 0);
            Assert.True(lcbPlcfandTxt > 0);
            Assert.True(lcbPlcfbteTapx > 0);
            Assert.NotEqual(fcPlcfandRef, fcPlcfbteTapx);
            Assert.NotEqual(fcPlcfandTxt, fcPlcfbteTapx);
        }

        [Fact]
        public void WriteDocBlocks_WithCommentAfterHyperlink_AnchorsReferenceAfterVisibleText()
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
                    new RunModel { Text = "Link", Hyperlink = hyperlink }
                }
            });
            model.Comments.Add(new CommentModel
            {
                Id = "0",
                Text = "Hyperlink comment",
                StartCp = 10,
                EndCp = 10
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
            string expectedMainText = "Visit \x0013HYPERLINK \"https://example.com\"\x0014Link\x0015\x0005\r";
            string expectedCommentText = "\x0005Hyperlink comment\r";
            string expectedText = expectedMainText + expectedCommentText;
            var textBytes = new byte[expectedText.Length];
            System.Array.Copy(wordDocData, 1536, textBytes, 0, expectedText.Length);
            var extractedText = Encoding.GetEncoding(1252).GetString(textBytes);

            Assert.Equal(expectedText, extractedText);

            int fcPlcfandRef = BitConverter.ToInt32(wordDocData, 154 + (Fib.CommentReferencePairIndex * 8));
            int commentReferenceCp = expectedMainText.IndexOf('\x0005');

            Assert.Equal(commentReferenceCp, BitConverter.ToInt32(tableData, fcPlcfandRef));
            Assert.Equal(expectedMainText.Length, BitConverter.ToInt32(tableData, fcPlcfandRef + 4));
        }

        [Fact]
        public void WriteDocBlocks_WithReplyComment_FoldsReplyIntoParentCommentStory()
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
                Text = "Parent",
                StartCp = 2,
                EndCp = 2
            });
            model.Comments.Add(new CommentModel
            {
                Id = "1",
                Text = "Reply",
                StartCp = 2,
                EndCp = 2,
                IsReply = true,
                ParentId = "0"
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
            string expectedCommentText = "\x0005Parent\r[Reply]\rReply\r";
            string expectedText = "AB\x0005CD\r" + expectedCommentText;
            var textBytes = new byte[expectedText.Length];
            System.Array.Copy(wordDocData, 1536, textBytes, 0, expectedText.Length);
            var extractedText = Encoding.GetEncoding(1252).GetString(textBytes);

            Assert.Equal(expectedText, extractedText);
            Assert.Equal(expectedCommentText.Length, BitConverter.ToInt32(wordDocData, 76));

            int fcPlcfandRef = BitConverter.ToInt32(wordDocData, 154 + (Fib.CommentReferencePairIndex * 8));
            int lcbPlcfandRef = BitConverter.ToInt32(wordDocData, 154 + (Fib.CommentReferencePairIndex * 8) + 4);
            int fcPlcfandTxt = BitConverter.ToInt32(wordDocData, 154 + (Fib.CommentTextPairIndex * 8));

            Assert.Equal(38, lcbPlcfandRef);
            Assert.Equal(2, BitConverter.ToInt32(tableData, fcPlcfandRef));
            Assert.Equal(6, BitConverter.ToInt32(tableData, fcPlcfandRef + 4));
            Assert.Equal(0, BitConverter.ToInt32(tableData, fcPlcfandTxt));
            Assert.Equal(expectedCommentText.Length - 1, BitConverter.ToInt32(tableData, fcPlcfandTxt + 4));
            Assert.Equal(expectedCommentText.Length, BitConverter.ToInt32(tableData, fcPlcfandTxt + 8));
        }

        [Fact]
        public void WriteDocBlocks_WithReplyChain_FoldsRepliesIntoSingleCommentStoryWithMetadata()
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
                Text = "Parent",
                StartCp = 2,
                EndCp = 2
            });
            model.Comments.Add(new CommentModel
            {
                Id = "1",
                Text = "First reply",
                StartCp = 2,
                EndCp = 2,
                IsReply = true,
                ParentId = "0",
                Author = "Alice Example",
                Date = new DateTime(2024, 1, 2, 3, 4, 5)
            });
            model.Comments.Add(new CommentModel
            {
                Id = "2",
                Text = "Nested reply",
                StartCp = 2,
                EndCp = 2,
                IsReply = true,
                ParentId = "1",
                Initials = "BX"
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
            string expectedCommentText = "\x0005Parent\r[Reply by Alice Example at 2024-01-02 03:04:05]\rFirst reply\r[Reply by BX]\rNested reply\r";
            string expectedText = "AB\x0005CD\r" + expectedCommentText;
            var textBytes = new byte[expectedText.Length];
            System.Array.Copy(wordDocData, 1536, textBytes, 0, expectedText.Length);
            var extractedText = Encoding.GetEncoding(1252).GetString(textBytes);

            Assert.Equal(expectedText, extractedText);
            Assert.Equal(expectedCommentText.Length, BitConverter.ToInt32(wordDocData, 76));

            int fcPlcfandRef = BitConverter.ToInt32(wordDocData, 154 + (Fib.CommentReferencePairIndex * 8));
            int fcPlcfandTxt = BitConverter.ToInt32(wordDocData, 154 + (Fib.CommentTextPairIndex * 8));
            Assert.Equal(2, BitConverter.ToInt32(tableData, fcPlcfandRef));
            Assert.Equal(6, BitConverter.ToInt32(tableData, fcPlcfandRef + 4));
            Assert.Equal(0, BitConverter.ToInt32(tableData, fcPlcfandTxt));
            Assert.Equal(expectedCommentText.Length - 1, BitConverter.ToInt32(tableData, fcPlcfandTxt + 4));
            Assert.Equal(expectedCommentText.Length, BitConverter.ToInt32(tableData, fcPlcfandTxt + 8));
        }

        [Fact]
        public void WriteDocBlocks_WithOrphanReplyComment_EmitsStandaloneDowngradedComment()
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
                Id = "1",
                Text = "Reply",
                StartCp = 2,
                EndCp = 2,
                IsReply = true,
                ParentId = "missing"
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
            string expectedCommentText = "\x0005[Reply]\rReply\r";
            string expectedText = "AB\x0005CD\r" + expectedCommentText;
            var textBytes = new byte[expectedText.Length];
            System.Array.Copy(wordDocData, 1536, textBytes, 0, expectedText.Length);
            var extractedText = Encoding.GetEncoding(1252).GetString(textBytes);

            Assert.Equal(expectedText, extractedText);
            Assert.Equal(expectedCommentText.Length, BitConverter.ToInt32(wordDocData, 76));

            int fcPlcfandRef = BitConverter.ToInt32(wordDocData, 154 + (Fib.CommentReferencePairIndex * 8));
            int fcPlcfandTxt = BitConverter.ToInt32(wordDocData, 154 + (Fib.CommentTextPairIndex * 8));
            Assert.Equal(2, BitConverter.ToInt32(tableData, fcPlcfandRef));
            Assert.Equal(6, BitConverter.ToInt32(tableData, fcPlcfandRef + 4));
            Assert.Equal(0, BitConverter.ToInt32(tableData, fcPlcfandTxt));
            Assert.Equal(expectedCommentText.Length - 1, BitConverter.ToInt32(tableData, fcPlcfandTxt + 4));
            Assert.Equal(expectedCommentText.Length, BitConverter.ToInt32(tableData, fcPlcfandTxt + 8));
        }

        [Fact]
        public void WriteDocBlocks_WithReplyWhoseParentCannotAnchor_EmitsReplyAsStandaloneDowngradedComment()
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
                Text = "Parent",
                StartCp = -1,
                EndCp = -1
            });
            model.Comments.Add(new CommentModel
            {
                Id = "1",
                Text = "Reply",
                StartCp = 2,
                EndCp = 2,
                IsReply = true,
                ParentId = "0",
                Author = "Alice Example"
            });

            var writer = new DocWriter();
            using var ms = new MemoryStream();

            writer.WriteDocBlocks(model, ms);
            ms.Position = 0;

            using var compoundFile = new OpenMcdf.CompoundFile(ms);
            Assert.True(compoundFile.RootStorage.TryGetStream("WordDocument", out var wordDocStream));

            var wordDocData = wordDocStream.GetData();
            string expectedText = "AB\x0005CD\r\x0005[Reply by Alice Example]\rReply\r";
            var textBytes = new byte[expectedText.Length];
            System.Array.Copy(wordDocData, 1536, textBytes, 0, expectedText.Length);
            var extractedText = Encoding.GetEncoding(1252).GetString(textBytes);

            Assert.Equal(expectedText, extractedText);
            Assert.Equal("\x0005[Reply by Alice Example]\rReply\r".Length, BitConverter.ToInt32(wordDocData, 76));
        }

        [Fact]
        public void WriteDocBlocks_WithReplyWithoutParentId_EmitsStandaloneDowngradedComment()
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
                Id = "1",
                Text = "Reply",
                StartCp = 2,
                EndCp = 2,
                IsReply = true
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
            string expectedCommentText = "\x0005[Reply]\rReply\r";
            string expectedText = "AB\x0005CD\r" + expectedCommentText;
            var textBytes = new byte[expectedText.Length];
            System.Array.Copy(wordDocData, 1536, textBytes, 0, expectedText.Length);
            var extractedText = Encoding.GetEncoding(1252).GetString(textBytes);

            Assert.Equal(expectedText, extractedText);
            Assert.Equal(expectedCommentText.Length, BitConverter.ToInt32(wordDocData, 76));

            int fcPlcfandRef = BitConverter.ToInt32(wordDocData, 154 + (Fib.CommentReferencePairIndex * 8));
            int fcPlcfandTxt = BitConverter.ToInt32(wordDocData, 154 + (Fib.CommentTextPairIndex * 8));
            Assert.Equal(2, BitConverter.ToInt32(tableData, fcPlcfandRef));
            Assert.Equal(6, BitConverter.ToInt32(tableData, fcPlcfandRef + 4));
            Assert.Equal(0, BitConverter.ToInt32(tableData, fcPlcfandTxt));
            Assert.Equal(expectedCommentText.Length - 1, BitConverter.ToInt32(tableData, fcPlcfandTxt + 4));
            Assert.Equal(expectedCommentText.Length, BitConverter.ToInt32(tableData, fcPlcfandTxt + 8));
        }

        [Fact]
        public void WriteDocBlocks_WithNestedAndSiblingReplies_EmitsDepthFirstReplyOrder()
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
                Text = "Parent",
                StartCp = 2,
                EndCp = 2
            });
            model.Comments.Add(new CommentModel
            {
                Id = "1",
                Text = "First reply",
                StartCp = 2,
                EndCp = 2,
                IsReply = true,
                ParentId = "0",
                Author = "Alice Example"
            });
            model.Comments.Add(new CommentModel
            {
                Id = "2",
                Text = "Second reply",
                StartCp = 2,
                EndCp = 2,
                IsReply = true,
                ParentId = "0",
                Initials = "BX"
            });
            model.Comments.Add(new CommentModel
            {
                Id = "3",
                Text = "Nested reply",
                StartCp = 2,
                EndCp = 2,
                IsReply = true,
                ParentId = "1"
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
            string expectedCommentText = "\x0005Parent\r[Reply by Alice Example]\rFirst reply\r[Reply]\rNested reply\r[Reply by BX]\rSecond reply\r";
            string expectedText = "AB\x0005CD\r" + expectedCommentText;
            var textBytes = new byte[expectedText.Length];
            System.Array.Copy(wordDocData, 1536, textBytes, 0, expectedText.Length);
            var extractedText = Encoding.GetEncoding(1252).GetString(textBytes);

            Assert.Equal(expectedText, extractedText);
            Assert.Equal(expectedCommentText.Length, BitConverter.ToInt32(wordDocData, 76));

            int fcPlcfandRef = BitConverter.ToInt32(wordDocData, 154 + (Fib.CommentReferencePairIndex * 8));
            int fcPlcfandTxt = BitConverter.ToInt32(wordDocData, 154 + (Fib.CommentTextPairIndex * 8));
            Assert.Equal(2, BitConverter.ToInt32(tableData, fcPlcfandRef));
            Assert.Equal(6, BitConverter.ToInt32(tableData, fcPlcfandRef + 4));
            Assert.Equal(0, BitConverter.ToInt32(tableData, fcPlcfandTxt));
            Assert.Equal(expectedCommentText.Length - 1, BitConverter.ToInt32(tableData, fcPlcfandTxt + 4));
            Assert.Equal(expectedCommentText.Length, BitConverter.ToInt32(tableData, fcPlcfandTxt + 8));
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
        public void WriteDocBlocks_WithMarginRelativeFloatingImageOffsets_AnchorsOffsetsToMargins()
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
                            WrapType = ImageWrapType.Square,
                            HorizontalRelativeTo = "margin",
                            VerticalRelativeTo = "margin",
                            PositionXTwips = 360,
                            PositionYTwips = 540
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

            Assert.Equal(1360, BitConverter.ToInt32(tableData, fcPlcfspaMom + 12));
            Assert.Equal(1740, BitConverter.ToInt32(tableData, fcPlcfspaMom + 16));
            Assert.Equal(2800, BitConverter.ToInt32(tableData, fcPlcfspaMom + 20));
            Assert.Equal(2460, BitConverter.ToInt32(tableData, fcPlcfspaMom + 24));

            int flags = BitConverter.ToInt32(tableData, fcPlcfspaMom + 30);
            Assert.Equal(1, (flags >> 7) & 0x3);
            Assert.Equal(1, (flags >> 5) & 0x3);
        }

        [Fact]
        public void WriteDocBlocks_WithMarginAliasRelativeValues_UsesMarginAnchorsAndFlags()
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
                            WrapType = ImageWrapType.Square,
                            HorizontalRelativeTo = "column",
                            VerticalRelativeTo = "topMargin",
                            PositionXTwips = 360,
                            PositionYTwips = 540
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

            Assert.Equal(1360, BitConverter.ToInt32(tableData, fcPlcfspaMom + 12));
            Assert.Equal(1740, BitConverter.ToInt32(tableData, fcPlcfspaMom + 16));

            int flags = BitConverter.ToInt32(tableData, fcPlcfspaMom + 30);
            Assert.Equal(1, (flags >> 7) & 0x3);
            Assert.Equal(1, (flags >> 5) & 0x3);
        }

        [Fact]
        public void WriteDocBlocks_WithParagraphAliasRelativeValues_UsesParagraphFlags()
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
                    new RunModel { Text = "Alias paragraph anchor" },
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
                            HorizontalRelativeTo = "character",
                            VerticalRelativeTo = "line",
                            PositionXTwips = 360,
                            PositionYTwips = 540
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

            Assert.Equal(360, BitConverter.ToInt32(tableData, fcPlcfspaMom + 12));
            Assert.Equal(1740, BitConverter.ToInt32(tableData, fcPlcfspaMom + 16));

            int flags = BitConverter.ToInt32(tableData, fcPlcfspaMom + 30);
            Assert.Equal(2, (flags >> 7) & 0x3);
            Assert.Equal(2, (flags >> 5) & 0x3);
        }

        [Fact]
        public void WriteDocBlocks_WithSeparatedMarginAliasValues_UsesMarginAnchorsAndFlags()
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
                            WrapType = ImageWrapType.Square,
                            HorizontalRelativeTo = "left-margin",
                            VerticalRelativeTo = "top_margin",
                            PositionXTwips = 360,
                            PositionYTwips = 540
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

            Assert.Equal(1360, BitConverter.ToInt32(tableData, fcPlcfspaMom + 12));
            Assert.Equal(1740, BitConverter.ToInt32(tableData, fcPlcfspaMom + 16));

            int flags = BitConverter.ToInt32(tableData, fcPlcfspaMom + 30);
            Assert.Equal(1, (flags >> 7) & 0x3);
            Assert.Equal(1, (flags >> 5) & 0x3);
        }

        [Fact]
        public void WriteDocBlocks_WithSpacedMarginAliasValues_UsesMarginAnchorsAndFlags()
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
                            WrapType = ImageWrapType.Square,
                            HorizontalRelativeTo = "left margin",
                            VerticalRelativeTo = "top margin",
                            PositionXTwips = 360,
                            PositionYTwips = 540
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

            Assert.Equal(1360, BitConverter.ToInt32(tableData, fcPlcfspaMom + 12));
            Assert.Equal(1740, BitConverter.ToInt32(tableData, fcPlcfspaMom + 16));

            int flags = BitConverter.ToInt32(tableData, fcPlcfspaMom + 30);
            Assert.Equal(1, (flags >> 7) & 0x3);
            Assert.Equal(1, (flags >> 5) & 0x3);
        }

        [Fact]
        public void WriteDocBlocks_WithWhitespaceMarginAliasValues_UsesMarginAnchorsAndFlags()
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
                            WrapType = ImageWrapType.Square,
                            HorizontalRelativeTo = "left\tmargin",
                            VerticalRelativeTo = "top\nmargin",
                            PositionXTwips = 360,
                            PositionYTwips = 540
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

            Assert.Equal(1360, BitConverter.ToInt32(tableData, fcPlcfspaMom + 12));
            Assert.Equal(1740, BitConverter.ToInt32(tableData, fcPlcfspaMom + 16));

            int flags = BitConverter.ToInt32(tableData, fcPlcfspaMom + 30);
            Assert.Equal(1, (flags >> 7) & 0x3);
            Assert.Equal(1, (flags >> 5) & 0x3);
        }

        [Fact]
        public void WriteDocBlocks_WithInsideMarginAliasContainingWhitespaceAndPunctuation_UsesMarginAnchorsAndFlags()
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
                            WrapType = ImageWrapType.Square,
                            HorizontalRelativeTo = "inside margin",
                            VerticalRelativeTo = "inside_margin",
                            PositionXTwips = 360,
                            PositionYTwips = 540
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

            Assert.Equal(1360, BitConverter.ToInt32(tableData, fcPlcfspaMom + 12));
            Assert.Equal(1740, BitConverter.ToInt32(tableData, fcPlcfspaMom + 16));

            int flags = BitConverter.ToInt32(tableData, fcPlcfspaMom + 30);
            Assert.Equal(1, (flags >> 7) & 0x3);
            Assert.Equal(1, (flags >> 5) & 0x3);
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
