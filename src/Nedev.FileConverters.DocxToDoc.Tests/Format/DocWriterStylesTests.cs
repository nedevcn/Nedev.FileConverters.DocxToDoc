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
    public class DocWriterStylesTests
    {
        [Fact]
        public void WriteDocBlocks_WithStylesAndFonts_WritesSttbAndStsh()
        {
            // Register encoding provider for Windows-1252 
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Arrange
            var model = new DocumentModel
            {
                TextBuffer = "StyleTest\r"
            };
            
            model.Paragraphs.Add(new ParagraphModel { Runs = { new RunModel { Text = "StyleTest" } } });
            
            model.Styles.Add(new StyleModel { Id = "Normal", Name = "Normal", IsParagraphStyle = true });
            model.Fonts.Add(new FontModel { Name = "Arial" });
            model.Fonts.Add(new FontModel { Name = "Times New Roman" });

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

            // Check FIB offsets for STSH and Font Table
            // fcStshf is index 0 in RgFcLcb (offset 154)
            // fcSttbfffn is index 14 in RgFcLcb (offset 154 + 14 * 8 = 266)
            
            int fcStshf = BitConverter.ToInt32(wordDocData, 154);
            int lcbStshf = BitConverter.ToInt32(wordDocData, 158);
            
            int fcSttbfffn = BitConverter.ToInt32(wordDocData, 266);
            int lcbSttbfffn = BitConverter.ToInt32(wordDocData, 270);

            Assert.NotEqual(0, fcStshf);
            Assert.True(lcbStshf > 0);
            Assert.NotEqual(0, fcSttbfffn);
            Assert.True(lcbSttbfffn > 0);

            // Basic check on Font Table (should start with fExtend = 0xFFFF)
            Assert.Equal(0xFF, tableData[fcSttbfffn]);
            Assert.Equal(0xFF, tableData[fcSttbfffn + 1]);
            
            // Count of fonts (2)
            Assert.Equal(2, BitConverter.ToUInt16(tableData, fcSttbfffn + 2));
        }

        [Fact]
        public void WriteDocBlocks_WithStyleParagraphNumbering_WritesNumberingSprmsIntoStylePapx()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            var model = new DocumentModel();
            model.Content.Add(new ParagraphModel
            {
                Runs =
                {
                    new RunModel { Text = "StyleNumberingTest" }
                }
            });

            model.NumberingInstances.Add(new NumberingInstanceModel
            {
                Id = 10,
                AbstractNumberId = 0
            });

            model.Styles.Add(new StyleModel
            {
                Id = "HeadingList",
                Name = "HeadingList",
                IsParagraphStyle = true,
                StyleId = 1,
                ParagraphProps = new ParagraphModel.ParagraphProperties
                {
                    NumberingId = 10,
                    NumberingIdSpecified = true,
                    NumberingLevel = 2,
                    NumberingLevelSpecified = true
                }
            });

            var writer = new DocWriter();
            using var ms = new MemoryStream();

            writer.WriteDocBlocks(model, ms);
            ms.Position = 0;

            using var compoundFile = new OpenMcdf.CompoundFile(ms);
            Assert.True(compoundFile.RootStorage.TryGetStream("WordDocument", out var wordDocStream));
            Assert.True(compoundFile.RootStorage.TryGetStream("1Table", out var tableStream));

            byte[] tableData = tableStream.GetData();
            byte[] wordDocData = wordDocStream.GetData();

            int fcStshf = BitConverter.ToInt32(wordDocData, 154);
            int lcbStshf = BitConverter.ToInt32(wordDocData, 158);
            Assert.True(fcStshf >= 0);
            Assert.True(lcbStshf > 0);

            var stshData = new byte[lcbStshf];
            Array.Copy(tableData, fcStshf, stshData, 0, lcbStshf);

            Assert.True(ContainsSubsequence(stshData, new byte[] { 0x0B, 0x46, 0x01, 0x00 }));
            Assert.True(ContainsSubsequence(stshData, new byte[] { 0x11, 0x26, 0x02 }));
        }

        [Fact]
        public void WriteDocBlocks_WithStyleParagraphNumberingLevelOnly_WritesLevelSprmIntoStylePapx()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            var model = new DocumentModel();
            model.Content.Add(new ParagraphModel
            {
                Runs =
                {
                    new RunModel { Text = "StyleNumberingLevelOnlyTest" }
                }
            });

            model.Styles.Add(new StyleModel
            {
                Id = "HeadingLevelOnly",
                Name = "HeadingLevelOnly",
                IsParagraphStyle = true,
                StyleId = 2,
                ParagraphProps = new ParagraphModel.ParagraphProperties
                {
                    NumberingLevel = 4,
                    NumberingLevelSpecified = true
                }
            });

            var writer = new DocWriter();
            using var ms = new MemoryStream();

            writer.WriteDocBlocks(model, ms);
            ms.Position = 0;

            using var compoundFile = new OpenMcdf.CompoundFile(ms);
            Assert.True(compoundFile.RootStorage.TryGetStream("WordDocument", out var wordDocStream));
            Assert.True(compoundFile.RootStorage.TryGetStream("1Table", out var tableStream));

            byte[] tableData = tableStream.GetData();
            byte[] wordDocData = wordDocStream.GetData();

            int fcStshf = BitConverter.ToInt32(wordDocData, 154);
            int lcbStshf = BitConverter.ToInt32(wordDocData, 158);
            Assert.True(fcStshf >= 0);
            Assert.True(lcbStshf > 0);

            var stshData = new byte[lcbStshf];
            Array.Copy(tableData, fcStshf, stshData, 0, lcbStshf);

            Assert.True(ContainsSubsequence(stshData, new byte[] { 0x11, 0x26, 0x04 }));
        }

        [Fact]
        public void WriteDocBlocks_WithDuplicateDefaultStyleIds_WritesAllStyleNamesIntoStsh()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            var model = new DocumentModel();
            model.Content.Add(new ParagraphModel
            {
                Runs =
                {
                    new RunModel { Text = "StyleSlotsTest" }
                }
            });

            model.Styles.Add(new StyleModel
            {
                Id = "StyleA",
                Name = "StyleA",
                IsParagraphStyle = true
            });
            model.Styles.Add(new StyleModel
            {
                Id = "StyleB",
                Name = "StyleB",
                IsParagraphStyle = true
            });

            var writer = new DocWriter();
            using var ms = new MemoryStream();

            writer.WriteDocBlocks(model, ms);
            ms.Position = 0;

            using var compoundFile = new OpenMcdf.CompoundFile(ms);
            Assert.True(compoundFile.RootStorage.TryGetStream("WordDocument", out var wordDocStream));
            Assert.True(compoundFile.RootStorage.TryGetStream("1Table", out var tableStream));

            byte[] tableData = tableStream.GetData();
            byte[] wordDocData = wordDocStream.GetData();

            int fcStshf = BitConverter.ToInt32(wordDocData, 154);
            int lcbStshf = BitConverter.ToInt32(wordDocData, 158);
            Assert.True(fcStshf >= 0);
            Assert.True(lcbStshf > 0);

            var stshData = new byte[lcbStshf];
            Array.Copy(tableData, fcStshf, stshData, 0, lcbStshf);

            byte[] styleAName = Encoding.Unicode.GetBytes("StyleA\0");
            byte[] styleBName = Encoding.Unicode.GetBytes("StyleB\0");

            Assert.True(ContainsSubsequence(stshData, styleAName));
            Assert.True(ContainsSubsequence(stshData, styleBName));
        }

        [Fact]
        public void WriteDocBlocks_WithOutOfRangeStyleIds_ResolvesBasedOnAndNextStyleToAssignedSlots()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            var model = new DocumentModel();
            model.Content.Add(new ParagraphModel
            {
                Runs =
                {
                    new RunModel { Text = "StyleRefsTest" }
                }
            });

            model.Styles.Add(new StyleModel
            {
                Id = "StyleA",
                Name = "StyleA",
                IsParagraphStyle = true,
                StyleId = 100,
                NextStyle = 200
            });
            model.Styles.Add(new StyleModel
            {
                Id = "StyleB",
                Name = "StyleB",
                IsParagraphStyle = true,
                StyleId = 200,
                BasedOn = 100
            });

            var writer = new DocWriter();
            using var ms = new MemoryStream();

            writer.WriteDocBlocks(model, ms);
            ms.Position = 0;

            using var compoundFile = new OpenMcdf.CompoundFile(ms);
            Assert.True(compoundFile.RootStorage.TryGetStream("WordDocument", out var wordDocStream));
            Assert.True(compoundFile.RootStorage.TryGetStream("1Table", out var tableStream));

            byte[] tableData = tableStream.GetData();
            byte[] wordDocData = wordDocStream.GetData();

            int fcStshf = BitConverter.ToInt32(wordDocData, 154);
            int lcbStshf = BitConverter.ToInt32(wordDocData, 158);
            var stshData = new byte[lcbStshf];
            Array.Copy(tableData, fcStshf, stshData, 0, lcbStshf);

            byte[] styleAName = Encoding.Unicode.GetBytes("StyleA\0");
            byte[] styleBName = Encoding.Unicode.GetBytes("StyleB\0");

            byte[] styleAPrefix =
            {
                0x01,       // sgc paragraph
                0x00,       // istdBase -> self slot 0
                0x01, 0x00, // istdNext -> slot 1 (resolved from styleId 200)
                0x00, 0x00, // bchUpe
                0x00, 0x00, // flags
                (byte)styleAName.Length
            };

            byte[] styleBPrefix =
            {
                0x01,       // sgc paragraph
                0x00,       // istdBase -> slot 0 (resolved from styleId 100)
                0x01, 0x00, // istdNext -> self slot 1
                0x00, 0x00, // bchUpe
                0x00, 0x00, // flags
                (byte)styleBName.Length
            };

            Assert.True(ContainsSubsequence(stshData, Combine(styleAPrefix, styleAName)));
            Assert.True(ContainsSubsequence(stshData, Combine(styleBPrefix, styleBName)));
        }

        [Fact]
        public void WriteDocBlocks_WithUnknownStyleReferences_FallsBackToSelfSlot()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            var model = new DocumentModel();
            model.Content.Add(new ParagraphModel
            {
                Runs =
                {
                    new RunModel { Text = "StyleUnknownRefsTest" }
                }
            });

            model.Styles.Add(new StyleModel
            {
                Id = "StyleX",
                Name = "StyleX",
                IsParagraphStyle = true,
                StyleId = 10,
                BasedOn = 777,
                NextStyle = 888
            });

            var writer = new DocWriter();
            using var ms = new MemoryStream();

            writer.WriteDocBlocks(model, ms);
            ms.Position = 0;

            using var compoundFile = new OpenMcdf.CompoundFile(ms);
            Assert.True(compoundFile.RootStorage.TryGetStream("WordDocument", out var wordDocStream));
            Assert.True(compoundFile.RootStorage.TryGetStream("1Table", out var tableStream));

            byte[] tableData = tableStream.GetData();
            byte[] wordDocData = wordDocStream.GetData();

            int fcStshf = BitConverter.ToInt32(wordDocData, 154);
            int lcbStshf = BitConverter.ToInt32(wordDocData, 158);
            var stshData = new byte[lcbStshf];
            Array.Copy(tableData, fcStshf, stshData, 0, lcbStshf);

            byte[] styleName = Encoding.Unicode.GetBytes("StyleX\0");
            int nameOffset = IndexOfSubsequence(stshData, styleName);
            Assert.True(nameOffset >= 9);
            Assert.Equal(10, stshData[nameOffset - 8]); // istdBase
            Assert.Equal(10, BitConverter.ToUInt16(stshData, nameOffset - 7)); // istdNext
        }

        [Fact]
        public void WriteDocBlocks_WithVeryLongStyleName_TruncatesNameToStdByteLimit()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            string marker = "VeryLongStyleName_";
            string longSuffix = new string('X', 300);
            string veryLongName = marker + longSuffix;

            var model = new DocumentModel();
            model.Content.Add(new ParagraphModel
            {
                Runs =
                {
                    new RunModel { Text = "StyleLongNameTest" }
                }
            });

            model.Styles.Add(new StyleModel
            {
                Id = "LongStyle",
                Name = veryLongName,
                IsParagraphStyle = true,
                StyleId = 1
            });

            var writer = new DocWriter();
            using var ms = new MemoryStream();
            writer.WriteDocBlocks(model, ms);
            ms.Position = 0;

            using var compoundFile = new OpenMcdf.CompoundFile(ms);
            Assert.True(compoundFile.RootStorage.TryGetStream("WordDocument", out var wordDocStream));
            Assert.True(compoundFile.RootStorage.TryGetStream("1Table", out var tableStream));

            byte[] tableData = tableStream.GetData();
            byte[] wordDocData = wordDocStream.GetData();

            int fcStshf = BitConverter.ToInt32(wordDocData, 154);
            int lcbStshf = BitConverter.ToInt32(wordDocData, 158);
            var stshData = new byte[lcbStshf];
            Array.Copy(tableData, fcStshf, stshData, 0, lcbStshf);

            byte[] markerBytes = Encoding.Unicode.GetBytes(marker);
            int markerOffset = IndexOfSubsequence(stshData, markerBytes);
            Assert.True(markerOffset >= 9);

            int nameLength = stshData[markerOffset - 1];
            Assert.Equal(254, nameLength);
            Assert.Equal(0, stshData[markerOffset + nameLength - 2]);
            Assert.Equal(0, stshData[markerOffset + nameLength - 1]);
        }

        [Fact]
        public void ClampStdUpxData_WithOversizedData_TruncatesToByteMax()
        {
            var method = typeof(DocWriter).GetMethod("ClampStdUpxData", BindingFlags.NonPublic | BindingFlags.Static);
            Assert.NotNull(method);

            var input = new byte[600];
            for (int index = 0; index < input.Length; index++)
            {
                input[index] = (byte)(index % 251);
            }

            var output = (byte[])method!.Invoke(null, new object[] { input })!;

            Assert.Equal(255, output.Length);
            for (int index = 0; index < output.Length; index++)
            {
                Assert.Equal(input[index], output[index]);
            }
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

        private static byte[] Combine(byte[] first, byte[] second)
        {
            var combined = new byte[first.Length + second.Length];
            Array.Copy(first, 0, combined, 0, first.Length);
            Array.Copy(second, 0, combined, first.Length, second.Length);
            return combined;
        }

        private static int IndexOfSubsequence(byte[] buffer, byte[] subsequence)
        {
            if (subsequence.Length == 0 || buffer.Length < subsequence.Length)
            {
                return -1;
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
                    return index;
                }
            }

            return -1;
        }
    }
}
