using System;
using System.IO;
using System.IO.Compression;
using System.Text;
using System.Collections.Generic;
using System.Security.Cryptography;
using Nedev.FileConverters.DocxToDoc.Model;

namespace Nedev.FileConverters.DocxToDoc.Format
{
    /// <summary>
    /// Writes the various streams required by the MS-DOC File Format (.doc)
    /// such as the WordDocument stream, 1Table, Data, etc.
    /// </summary>
    public class DocWriter
    {
        private const short OfficeArtBlipToDisplayProperty = 0x0104;
        private const int MainDocumentDrawingId = 1;
        private const int ShapeIdBase = MainDocumentDrawingId << 10;

        private readonly struct OfficeArtPictureDescriptor
        {
            public OfficeArtPictureDescriptor(int cp, int shapeId, int blipIndex, byte[] data, string contentType, int widthTwips, int heightTwips, int leftTwips, int topTwips, ImageWrapType wrapType, bool behindText, bool allowOverlap, string? horizontalRelativeTo, string? verticalRelativeTo)
            {
                Cp = cp;
                ShapeId = shapeId;
                BlipIndex = blipIndex;
                Data = data;
                ContentType = contentType;
                WidthTwips = widthTwips;
                HeightTwips = heightTwips;
                LeftTwips = leftTwips;
                TopTwips = topTwips;
                WrapType = wrapType;
                BehindText = behindText;
                AllowOverlap = allowOverlap;
                HorizontalRelativeTo = horizontalRelativeTo;
                VerticalRelativeTo = verticalRelativeTo;
            }

            public int Cp { get; }
            public int ShapeId { get; }
            public int BlipIndex { get; }
            public byte[] Data { get; }
            public string ContentType { get; }
            public int WidthTwips { get; }
            public int HeightTwips { get; }
            public int LeftTwips { get; }
            public int TopTwips { get; }
            public ImageWrapType WrapType { get; }
            public bool BehindText { get; }
            public bool AllowOverlap { get; }
            public string? HorizontalRelativeTo { get; }
            public string? VerticalRelativeTo { get; }
            public int RightTwips => LeftTwips + WidthTwips;
            public int BottomTwips => TopTwips + HeightTwips;
        }

        public void WriteDocBlocks(DocumentModel model, Stream outputStream)
        {
            // 1. Initialize streams needed for the MS-DOC file
            using var wordDocumentStream = new MemoryStream();
            using var tableStream = new MemoryStream();
            using var dataStream = new MemoryStream();

            // Track inline picture blocks written to the Data stream.
            var embeddedObjects = new List<byte[]>();
            var officeArtBlips = new List<(int cp, byte[] data, string contentType, int widthTwips, int heightTwips, int leftTwips, int topTwips, ImageWrapType wrapType, bool behindText, bool allowOverlap, string? horizontalRelativeTo, string? verticalRelativeTo)>();
            var fieldEntries = new List<(int cp, ushort descriptor)>();
            int nextPictureOffset = 0;

            // 1. Build the text buffer and formatting structures in one pass
            var textBuilder = new StringBuilder();
            var chpxWriter = new ChpxFkpWriter();
            var papxWriter = new PapxFkpWriter();
            var tapxWriter = new TapxFkpWriter();
            var layoutSection = model.Sections.Count > 0 ? model.Sections[0] : new SectionModel();
            int documentAvailableWidthTwips = Math.Max(1440, layoutSection.PageWidth - layoutSection.MarginLeft - layoutSection.MarginRight);
            int paragraphVerticalCursorTwips = layoutSection.MarginTop;
            
            int currentCp = 0;
            var tableWriter = new BinaryWriter(tableStream);

            void ProcessParagraph(ParagraphModel para, ref int verticalCursorTwips, int availableWidthTwips, int? maxVisibleCursorTwips = null)
            {
                int paraStart = currentCp;
                int paragraphAvailableWidthTwips = ResolveParagraphAvailableWidthTwips(para, availableWidthTwips);
                int paragraphContentHeightTwips = EstimateParagraphContentHeightTwips(para, paragraphAvailableWidthTwips);
                int paragraphTopTwips = verticalCursorTwips + para.Properties.SpacingBeforeTwips;

                if (maxVisibleCursorTwips.HasValue)
                {
                    int maxAllowedHeight = Math.Max(0, maxVisibleCursorTwips.Value - paragraphTopTwips);
                    paragraphContentHeightTwips = Math.Min(paragraphContentHeightTwips, maxAllowedHeight);
                }

                int paragraphAdvanceTwips = EstimateParagraphAdvanceTwips(para, paragraphContentHeightTwips);
                var autoCompletedFields = new HashSet<FieldModel>();
                var separatedFields = new HashSet<FieldModel>();
                var openFields = new List<FieldModel>();
                for (int runIndex = 0; runIndex < para.Runs.Count; runIndex++)
                {
                    var run = para.Runs[runIndex];

                    if (run.Field != null && autoCompletedFields.Contains(run.Field))
                    {
                        continue;
                    }

                    // Handle images
                    if (run.Image != null && run.Image.Data != null)
                    {
                        string imageContentType = ResolveImageContentType(run.Image.ContentType, run.Image.Data);
                        (int imageWidthTwips, int imageHeightTwips) = ResolveImageDimensionsTwips(run.Image, imageContentType);

                        // Add a placeholder character for the image
                        // In MS-DOC, embedded objects use special characters
                        textBuilder.Append('\x0001'); // Object placeholder
                        
                        // Emit a DOC-style picture block into the Data stream and point CHPX to it.
                        byte[] pictureBlock = BuildPictureBlock(run.Image, imageContentType);
                        int pictureOffset = nextPictureOffset;
                        nextPictureOffset += pictureBlock.Length;
                        embeddedObjects.Add(pictureBlock);

                        if (SupportsOfficeArtBlip(imageContentType))
                        {
                            (int leftTwips, int topTwips, _, _) = ResolveImageBoundsTwips(run.Image, imageContentType, layoutSection, paragraphTopTwips, paragraphContentHeightTwips);
                            officeArtBlips.Add((currentCp, run.Image.Data, imageContentType, imageWidthTwips, imageHeightTwips, leftTwips, topTwips, run.Image.WrapType, run.Image.BehindText, run.Image.AllowOverlap, run.Image.HorizontalRelativeTo, run.Image.VerticalRelativeTo));
                        }
                        
                        // Add CHPX with fSpec plus a Data-stream picture location.
                        chpxWriter.AddRun(currentCp, currentCp + 1, BuildImageSprms(pictureOffset));
                        
                        currentCp += 1;
                        continue;
                    }

                    if (run.Hyperlink != null)
                    {
                        int hyperlinkStart = runIndex;
                        var hyperlink = run.Hyperlink;
                        while (runIndex + 1 < para.Runs.Count && ReferenceEquals(para.Runs[runIndex + 1].Hyperlink, hyperlink))
                        {
                            runIndex++;
                        }

                        AppendFieldCharacter('\x0013');
                        AppendPlainText(BuildHyperlinkInstruction(hyperlink));
                        AppendFieldCharacter('\x0014');

                        for (int hyperlinkRunIndex = hyperlinkStart; hyperlinkRunIndex <= runIndex; hyperlinkRunIndex++)
                        {
                            AppendFormattedRunText(para.Runs[hyperlinkRunIndex]);
                        }

                        AppendFieldCharacter('\x0015');
                        continue;
                    }

                    if (run.IsFieldBegin)
                    {
                        if (run.Field != null && !HasExplicitFieldBoundary(para.Runs, runIndex + 1, run.Field))
                        {
                            int fieldDepth = openFields.Count;
                            AppendFieldCharacter('\x0013', run.Field, fieldDepth);

                            string instruction = ResolveFieldInstruction(run.Field);
                            AppendPlainText(instruction);

                            if (!string.IsNullOrEmpty(run.Field.Result))
                            {
                                AppendFieldCharacter('\x0014', run.Field, fieldDepth);
                                AppendPlainText(run.Field.Result);
                            }

                            AppendFieldCharacter('\x0015', run.Field, fieldDepth);
                            autoCompletedFields.Add(run.Field);
                            continue;
                        }

                        AppendFieldCharacter('\x0013', run.Field, openFields.Count);
                        if (run.Field != null)
                        {
                            openFields.Add(run.Field);
                        }
                        continue;
                    }

                    if (run.Field != null && !run.IsFieldSeparate && !run.IsFieldEnd && run.Text.Length == 0 && !string.IsNullOrEmpty(run.Field.Instruction))
                    {
                        AppendPlainText(run.Field.Instruction);
                        continue;
                    }

                    if (run.IsFieldSeparate)
                    {
                        AppendFieldCharacter('\x0014', run.Field, GetFieldDepth(openFields, run.Field));
                        continue;
                    }

                    if (run.IsFieldEnd)
                    {
                        int fieldDepth = GetFieldDepth(openFields, run.Field);
                        if (run.Field != null &&
                            !separatedFields.Contains(run.Field) &&
                            !string.IsNullOrEmpty(run.Field.Result))
                        {
                            AppendFieldCharacter('\x0014', run.Field, fieldDepth);
                            AppendPlainText(run.Field.Result);
                        }

                        AppendFieldCharacter('\x0015', run.Field, fieldDepth);
                        if (run.Field != null)
                        {
                            RemoveLastOpenField(openFields, run.Field);
                        }
                        continue;
                    }

                    AppendFormattedRunText(run);
                }

                for (int index = openFields.Count - 1; index >= 0; index--)
                {
                    var openField = openFields[index];
                    int fieldDepth = index;
                    if (!separatedFields.Contains(openField) && !string.IsNullOrEmpty(openField.Result))
                    {
                        AppendFieldCharacter('\x0014', openField, fieldDepth);
                        AppendPlainText(openField.Result);
                    }

                    AppendFieldCharacter('\x0015', openField, fieldDepth);
                }
                
                // End Paragraph
                List<byte> paraSprms = new List<byte>();
                if (para.Properties.Alignment != ParagraphModel.Justification.Left)
                {
                    paraSprms.Add(0x03); paraSprms.Add(0x24); paraSprms.Add((byte)para.Properties.Alignment);
                }
                if (para.Properties.NumberingId.HasValue)
                {
                    paraSprms.Add(0x0B); paraSprms.Add(0x46);
                    int lfoIndex = model.NumberingInstances.FindIndex(n => n.Id == para.Properties.NumberingId.Value) + 1;
                    paraSprms.Add((byte)(lfoIndex & 0xFF)); paraSprms.Add((byte)((lfoIndex >> 8) & 0xFF));
                    if (para.Properties.NumberingLevel.HasValue)
                    {
                        paraSprms.Add(0x11); paraSprms.Add(0x26); paraSprms.Add((byte)para.Properties.NumberingLevel.Value);
                    }
                }
                AppendParagraphFormattingSprms(paraSprms, para.Properties);

                textBuilder.Append('\r');
                papxWriter.AddParagraph(paraStart, currentCp + 1, paraSprms.ToArray());
                currentCp += 1;
                verticalCursorTwips += paragraphAdvanceTwips;

                void AppendFieldCharacter(char marker, FieldModel? fieldModel = null, int nestingDepth = 0)
                {
                    textBuilder.Append(marker);
                    if (marker == '\x0013' || marker == '\x0014' || marker == '\x0015')
                    {
                        fieldEntries.Add((currentCp, BuildFieldDescriptor(marker, fieldModel, nestingDepth)));
                    }

                    if (marker == '\x0014' && fieldModel != null)
                    {
                        separatedFields.Add(fieldModel);
                    }

                    currentCp += 1;
                }

                void AppendPlainText(string text)
                {
                    if (string.IsNullOrEmpty(text))
                    {
                        return;
                    }

                    textBuilder.Append(text);
                    currentCp += text.Length;
                }

                void AppendFormattedRunText(RunModel runModel)
                {
                    if (runModel.Text.Length == 0)
                    {
                        return;
                    }

                    byte[] runSprms = BuildRunSprms(runModel.Properties, model.Fonts);
                    if (runSprms.Length > 0)
                    {
                        chpxWriter.AddRun(currentCp, currentCp + runModel.Text.Length, runSprms);
                    }

                    textBuilder.Append(runModel.Text);
                    currentCp += runModel.Text.Length;
                }

                static string BuildHyperlinkInstruction(HyperlinkModel hyperlinkModel)
                {
                    if (!string.IsNullOrWhiteSpace(hyperlinkModel.TargetUrl))
                    {
                        return $"HYPERLINK \"{hyperlinkModel.TargetUrl}\"";
                    }

                    if (!string.IsNullOrWhiteSpace(hyperlinkModel.Anchor))
                    {
                        return $"HYPERLINK \\l \"{hyperlinkModel.Anchor}\"";
                    }

                    return "HYPERLINK";
                }

                static ushort BuildFieldDescriptor(char marker, FieldModel? fieldModel, int nestingDepth)
                {
                    ushort descriptor = marker;
                    if (fieldModel != null)
                    {
                        byte flags = 0;
                        if (marker == '\x0013' && fieldModel.IsLocked)
                        {
                            flags |= 0x01;
                        }
                        if (marker == '\x0013' && fieldModel.IsDirty)
                        {
                            flags |= 0x02;
                        }
                        if (!string.IsNullOrEmpty(fieldModel.Result))
                        {
                            flags |= 0x04;
                        }
                        if (fieldModel.Type != FieldType.Unknown)
                        {
                            flags |= 0x08;
                        }

                        descriptor |= (ushort)(flags << 8);
                    }

                    descriptor |= (ushort)((System.Math.Min(System.Math.Max(nestingDepth, 0), 15) & 0x0F) << 12);

                    return descriptor;
                }

                static bool HasExplicitFieldBoundary(IList<RunModel> runs, int startIndex, FieldModel fieldModel)
                {
                    for (int index = startIndex; index < runs.Count; index++)
                    {
                        var candidate = runs[index];
                        if (!ReferenceEquals(candidate.Field, fieldModel))
                        {
                            continue;
                        }

                        if (candidate.IsFieldSeparate || candidate.IsFieldEnd)
                        {
                            return true;
                        }
                    }

                    return false;
                }

                static string ResolveFieldInstruction(FieldModel fieldModel)
                {
                    if (!string.IsNullOrWhiteSpace(fieldModel.Instruction))
                    {
                        return fieldModel.Instruction;
                    }

                    return fieldModel.Type switch
                    {
                        FieldType.Page => "PAGE",
                        FieldType.NumPages => "NUMPAGES",
                        FieldType.Date => "DATE",
                        FieldType.Time => "TIME",
                        FieldType.Author => "AUTHOR",
                        FieldType.Title => "TITLE",
                        FieldType.Subject => "SUBJECT",
                        FieldType.FileName => "FILENAME",
                        FieldType.Hyperlink => "HYPERLINK",
                        FieldType.Bookmark => "BOOKMARK",
                        FieldType.Index => "INDEX",
                        FieldType.Seq => "SEQ",
                        FieldType.Ref => "REF",
                        FieldType.MergeField => "MERGEFIELD",
                        _ => string.Empty
                    };
                }

                static void RemoveLastOpenField(List<FieldModel> openFieldList, FieldModel fieldModel)
                {
                    for (int index = openFieldList.Count - 1; index >= 0; index--)
                    {
                        if (ReferenceEquals(openFieldList[index], fieldModel))
                        {
                            openFieldList.RemoveAt(index);
                            return;
                        }
                    }
                }

                static int GetFieldDepth(List<FieldModel> openFieldList, FieldModel? fieldModel)
                {
                    if (fieldModel == null)
                    {
                        return 0;
                    }

                    for (int index = openFieldList.Count - 1; index >= 0; index--)
                    {
                        if (ReferenceEquals(openFieldList[index], fieldModel))
                        {
                            return index;
                        }
                    }

                    return 0;
                }
            }

            foreach (var item in model.Content)
            {
                if (item is ParagraphModel para)
                {
                    ProcessParagraph(para, ref paragraphVerticalCursorTwips, documentAvailableWidthTwips);
                }
                else if (item is TableModel table)
                {
                    for (int rowIndex = 0; rowIndex < table.Rows.Count; rowIndex++)
                    {
                        var row = table.Rows[rowIndex];
                        var previousRow = rowIndex > 0 ? table.Rows[rowIndex - 1] : null;
                        int rowStart = currentCp;
                        int rowHeightTwips = 0;
                        int rowGridColumnIndex = 0;
                        int totalColumnCount = ResolveTableTotalColumnCount(table, row);
                        int tableAvailableWidthTwips = ResolveTableAvailableWidthTwips(table, row, documentAvailableWidthTwips, totalColumnCount);
                        var resolvedRowWidthsTwips = ResolveRowCellWidthsTwips(table, row, tableAvailableWidthTwips, totalColumnCount);
                        var cellLayouts = new List<(TableCellModel cell, int availableWidthTwips, int totalHeightTwips, int topOffsetTwips, int bottomOffsetTwips, int verticalOffsetTwips)>();
                        for (int cellIndex = 0; cellIndex < row.Cells.Count; cellIndex++)
                        {
                            var cell = row.Cells[cellIndex];
                            int cellStartColumnIndex = rowGridColumnIndex;
                            int cellSpan = Math.Max(1, cell.GridSpan);
                            bool isFirstRow = rowIndex == 0;
                            bool isLastRow = rowIndex == table.Rows.Count - 1;
                            bool isFirstColumn = cellStartColumnIndex == 0;
                            bool isLastColumn = cellStartColumnIndex + cellSpan >= totalColumnCount;
                            var previousCell = cellIndex > 0 ? row.Cells[cellIndex - 1] : null;
                            int resolvedCellWidthTwips = cellIndex < resolvedRowWidthsTwips.Count ? resolvedRowWidthsTwips[cellIndex] : 0;
                            int effectiveCellWidthTwips = resolvedCellWidthTwips > 0 ? resolvedCellWidthTwips : tableAvailableWidthTwips;
                            int leftBorderTwips = ResolveTableCellLeftBorderTwips(table, cell, previousCell, isFirstColumn);
                            int rightBorderTwips = ResolveTableCellRightBorderTwips(table, cell, isLastColumn);
                            int topBorderTwips = ResolveTableCellTopBorderTwips(table, cell, previousRow, cellStartColumnIndex, cellSpan, isFirstRow);
                            int bottomBorderTwips = ResolveTableCellBottomBorderTwips(table, cell, isLastRow);
                            int topPaddingTwips = ResolveTableCellTopPaddingTwips(table, cell);
                            int bottomPaddingTwips = ResolveTableCellBottomPaddingTwips(table, cell);
                            int topCellSpacingTwips = ResolveTableCellTopSpacingTwips(table);
                            int bottomCellSpacingTwips = ResolveTableCellBottomSpacingTwips(table);
                            int horizontalCellBorderTwips = Math.Max(0, leftBorderTwips) + Math.Max(0, rightBorderTwips);
                            int horizontalCellPaddingTwips = ResolveTableCellHorizontalPaddingTwips(table, cell);
                            int horizontalCellSpacingTwips = ResolveTableCellHorizontalSpacingTwips(table);
                            int cellAvailableWidthTwips = Math.Max(720, effectiveCellWidthTwips - horizontalCellBorderTwips - horizontalCellPaddingTwips - horizontalCellSpacingTwips);
                            int cellContentHeightTwips = EstimateTableCellContentHeightTwips(cell, cellAvailableWidthTwips);
                            int cellTotalHeightTwips = topCellSpacingTwips + topBorderTwips + topPaddingTwips + cellContentHeightTwips + bottomPaddingTwips + bottomBorderTwips + bottomCellSpacingTwips;
                            rowGridColumnIndex += cellSpan;
                            rowHeightTwips = Math.Max(rowHeightTwips, cellTotalHeightTwips);
                            cellLayouts.Add((cell, cellAvailableWidthTwips, cellTotalHeightTwips, topCellSpacingTwips + topBorderTwips + topPaddingTwips, bottomPaddingTwips + bottomBorderTwips + bottomCellSpacingTwips, 0));
                        }

                        rowHeightTwips = ResolveRowHeightTwips(row, rowHeightTwips);

                        for (int cellIndex = 0; cellIndex < cellLayouts.Count; cellIndex++)
                        {
                            var cellLayout = cellLayouts[cellIndex];
                            int verticalOffsetTwips = ResolveTableCellVerticalAlignmentOffset(cellLayout.cell, rowHeightTwips, cellLayout.totalHeightTwips);
                            cellLayouts[cellIndex] = (cellLayout.cell, cellLayout.availableWidthTwips, cellLayout.totalHeightTwips, cellLayout.topOffsetTwips, cellLayout.bottomOffsetTwips, verticalOffsetTwips);
                        }

                        foreach (var cellLayout in cellLayouts)
                        {
                            int cellVerticalCursorTwips = cellLayout.topOffsetTwips + cellLayout.verticalOffsetTwips;
                            int maxVisibleCursorTwips = Math.Max(cellLayout.topOffsetTwips, rowHeightTwips - cellLayout.bottomOffsetTwips);
                            foreach (var cellPara in cellLayout.cell.Paragraphs)
                            {
                                ProcessParagraph(cellPara, ref cellVerticalCursorTwips, cellLayout.availableWidthTwips, row.HeightRule == TableRowHeightRule.Exact ? maxVisibleCursorTwips : null);
                                if (row.HeightRule == TableRowHeightRule.Exact)
                                {
                                    cellVerticalCursorTwips = Math.Min(cellVerticalCursorTwips, maxVisibleCursorTwips);
                                }
                            }

                            // Cell Mark - treated as a paragraph terminator in MS-DOC
                            int cellMarkStart = currentCp;
                            textBuilder.Append('\x0007');
                            currentCp += 1;
                            
                            // Cell marks need PAPX entries with sprmPFInTable
                            List<byte> cellMarkSprms = new List<byte>();
                            cellMarkSprms.Add(0x16); cellMarkSprms.Add(0x24); cellMarkSprms.Add(1); // sprmPFInTable = 1
                            papxWriter.AddParagraph(cellMarkStart, currentCp, cellMarkSprms.ToArray());
                        }
                        // Row Mark Paragraph
                        int rowMarkStart = currentCp;
                        textBuilder.Append('\r');
                        
                        List<byte> rowParaSprms = new List<byte>();
                        rowParaSprms.Add(0x16); rowParaSprms.Add(0x24); rowParaSprms.Add(1); // sprmPFTable = 1
                        rowParaSprms.Add(0x17); rowParaSprms.Add(0x24); rowParaSprms.Add(1); // sprmPFTermInTbl = 1
                        
                        papxWriter.AddParagraph(rowMarkStart, currentCp + 1, rowParaSprms.ToArray());
                        currentCp += 1;
                        
                        // Build TAP (Table Properties) for this row
                        List<byte> tapSprms = new List<byte>();
                        // sprmTDefTable (0xD608) - minimal definition
                        tapSprms.Add(0x08); tapSprms.Add(0xD6);
                        // Complex operand shortened for now
                        byte[] defTable = new byte[10] { 0x08, (byte)row.Cells.Count, 0, 0, 0, 0, 0, 0, 0, 0 };
                        tapSprms.AddRange(defTable);
                        
                        tapxWriter.AddRow(rowStart, currentCp, tapSprms.ToArray());
                        paragraphVerticalCursorTwips += rowHeightTwips;
                    }
                }
            }

            string finalBaseText = textBuilder.ToString();
            byte[] textBytes = System.Text.Encoding.GetEncoding(1252).GetBytes(finalBaseText);
            wordDocumentStream.Seek(1536, SeekOrigin.Begin);
            wordDocumentStream.Write(textBytes);

            // 2. Build the Piece Table (Clx)
            int fcClx = (int)tableStream.Position;
            tableWriter.Write((byte)0x02);
            int cbPlcfpcd = (2 * 4) + (1 * 8); 
            tableWriter.Write((int)cbPlcfpcd);
            tableWriter.Write((int)0);
            tableWriter.Write((int)currentCp);
            int fcBits = (1536) | 0x40000000;
            tableWriter.Write((int)fcBits);
            tableWriter.Write((short)0);
            int lcbClx = (int)tableStream.Position - fcClx;

            // 3. Process CHPX FKPs 
            int fcPlcfbteChpx = 0; int lcbPlcfbteChpx = 0;
            byte[] chpxPage = chpxWriter.GeneratePage();
            if (chpxPage.Length > 0 && chpxPage[511] > 0)
            {
                long rem = wordDocumentStream.Position % 512;
                if (rem != 0) wordDocumentStream.Write(new byte[512 - rem]);
                int pnChpx = (int)(wordDocumentStream.Position / 512);
                wordDocumentStream.Write(chpxPage);
                fcPlcfbteChpx = (int)tableStream.Position;
                tableWriter.Write((int)0); tableWriter.Write((int)currentCp); tableWriter.Write((int)pnChpx);
                lcbPlcfbteChpx = (int)tableStream.Position - fcPlcfbteChpx;
            }

            // 4. Process PAPX FKPs
            int fcPlcfbtePapx = 0; int lcbPlcfbtePapx = 0;
            byte[] papxPage = papxWriter.GeneratePage();
            if (papxPage.Length > 0 && papxPage[511] > 0)
            {
                long rem = wordDocumentStream.Position % 512;
                if (rem != 0) wordDocumentStream.Write(new byte[512 - rem]);
                int pnPapx = (int)(wordDocumentStream.Position / 512);
                wordDocumentStream.Write(papxPage);
                fcPlcfbtePapx = (int)tableStream.Position;
                tableWriter.Write((int)0); tableWriter.Write((int)currentCp); tableWriter.Write((int)pnPapx);
                lcbPlcfbtePapx = (int)tableStream.Position - fcPlcfbtePapx;
            }

            // 5. Process TAPX FKPs
            int fcPlcfbteTapx = 0; int lcbPlcfbteTapx = 0;
            byte[] tapxPage = tapxWriter.GeneratePage();
            if (tapxPage.Length > 0 && tapxPage[511] > 0)
            {
                long rem = wordDocumentStream.Position % 512;
                if (rem != 0) wordDocumentStream.Write(new byte[512 - rem]);
                int pnTapx = (int)(wordDocumentStream.Position / 512);
                wordDocumentStream.Write(tapxPage);
                fcPlcfbteTapx = (int)tableStream.Position;
                tableWriter.Write((int)0); tableWriter.Write((int)currentCp); tableWriter.Write((int)pnTapx);
                lcbPlcfbteTapx = (int)tableStream.Position - fcPlcfbteTapx;
            }

            // 6. Build Font Table (STTB FFN)
            int fcSttbfffn = (int)tableStream.Position;
            WriteFontTable(tableWriter, model.Fonts);
            int lcbSttbfffn = (int)tableStream.Position - fcSttbfffn;

            // 7. Build Style Sheet (STSH)
            int fcStshf = (int)tableStream.Position;
            WriteStyleSheet(tableWriter, model.Styles, model.Fonts);
            int lcbStshf = (int)tableStream.Position - fcStshf;

            // 8. Write Numbering (SttbLst and PlcfLfo)
            int fcSttbLst = (int)tableStream.Position;
            WriteNumbering(tableWriter, model);
            int lcbSttbLst = (int)tableStream.Position - fcSttbLst;

            int fcPlfLfo = (int)tableStream.Position;
            WriteLfo(tableWriter, model);
            int lcbPlfLfo = (int)tableStream.Position - fcPlfLfo;

            int fcPlcffldMom = 0;
            int lcbPlcffldMom = 0;
            if (fieldEntries.Count > 0)
            {
                fcPlcffldMom = (int)tableStream.Position;
                foreach (var (cp, _) in fieldEntries)
                {
                    tableWriter.Write(cp);
                }
                tableWriter.Write(currentCp);

                foreach (var (_, descriptor) in fieldEntries)
                {
                    tableWriter.Write(descriptor);
                }

                lcbPlcffldMom = (int)tableStream.Position - fcPlcffldMom;
            }

            var officeArtPictures = BuildOfficeArtPictureDescriptors(officeArtBlips);

            int fcPlcfspaMom = 0;
            int lcbPlcfspaMom = 0;
            if (officeArtPictures.Count > 0)
            {
                fcPlcfspaMom = (int)tableStream.Position;
                WritePlcfspaMom(tableWriter, officeArtPictures, currentCp);
                lcbPlcfspaMom = (int)tableStream.Position - fcPlcfspaMom;
            }

            int fcDggInfo = 0;
            int lcbDggInfo = 0;
            if (officeArtPictures.Count > 0)
            {
                byte[] officeArtContent = BuildOfficeArtContent(officeArtPictures);
                if (officeArtContent.Length > 0)
                {
                    fcDggInfo = (int)tableStream.Position;
                    tableWriter.Write(officeArtContent);
                    lcbDggInfo = officeArtContent.Length;
                }
            }

            // 8.5. Write embedded objects/images to Data stream
            if (embeddedObjects.Count > 0)
            {
                using var dataWriter = new BinaryWriter(dataStream, System.Text.Encoding.GetEncoding(1252), leaveOpen: true);
                WritePictureBlocks(dataWriter, embeddedObjects);
            }

            // 8.6. Write Bookmarks (PlcfBkmkf and PlcfBkmkl)
            int fcPlcfBkmkf = 0;
            int lcbPlcfBkmkf = 0;
            int fcPlcfBkmkl = 0;
            int lcbPlcfBkmkl = 0;
            int fcSttbfbkmk = 0;
            int lcbSttbfbkmk = 0;

            if (model.Bookmarks.Count > 0)
            {
                // Write bookmark names (STTBF)
                fcSttbfbkmk = (int)tableStream.Position;
                tableWriter.Write((ushort)0xFFFF); // fExtend
                tableWriter.Write((ushort)model.Bookmarks.Count);
                tableWriter.Write((ushort)0); // cbExtra

                foreach (var bookmark in model.Bookmarks)
                {
                    // Write bookmark name as null-terminated Unicode string
                    byte[] nameBytes = System.Text.Encoding.Unicode.GetBytes(bookmark.Name + "\0");
                    tableWriter.Write((ushort)bookmark.Name.Length);
                    tableWriter.Write(nameBytes);
                }
                lcbSttbfbkmk = (int)tableStream.Position - fcSttbfbkmk;

                // Write PlcfBkmkf (bookmark first CPs)
                fcPlcfBkmkf = (int)tableStream.Position;
                foreach (var bookmark in model.Bookmarks)
                {
                    tableWriter.Write(bookmark.StartCp);
                }
                // Add terminator
                tableWriter.Write(currentCp);
                lcbPlcfBkmkf = (int)tableStream.Position - fcPlcfBkmkf;

                // Write PlcfBkmkl (bookmark last CPs)
                fcPlcfBkmkl = (int)tableStream.Position;
                foreach (var bookmark in model.Bookmarks)
                {
                    tableWriter.Write(bookmark.EndCp);
                }
                // Add terminator
                tableWriter.Write(currentCp);
                lcbPlcfBkmkl = (int)tableStream.Position - fcPlcfBkmkl;
            }

            // 9. Process section properties: Build Plcfsed and SED/SEP
            var sections = model.Sections.Count > 0 ? model.Sections : new List<SectionModel> { new SectionModel() };
            
            // Write SEPs to WordDocument
            List<int> fcSeps = new List<int>();
            wordDocumentStream.Seek(0, SeekOrigin.End);
            foreach (var section in sections)
            {
                fcSeps.Add((int)wordDocumentStream.Position);
                using var sepBinaryWriter = new BinaryWriter(wordDocumentStream, System.Text.Encoding.GetEncoding(1252), leaveOpen: true);
                
                List<byte> sepSprms = new List<byte>();
                void AddShortSprm(ushort op, int val)
                {
                    sepSprms.Add((byte)(op & 0xFF));
                    sepSprms.Add((byte)((op >> 8) & 0xFF));
                    sepSprms.Add(BitConverter.GetBytes((short)val)[0]);
                    sepSprms.Add(BitConverter.GetBytes((short)val)[1]);
                }

                AddShortSprm(0xB603, section.PageWidth);
                AddShortSprm(0xB604, section.PageHeight);
                AddShortSprm(0xB605, section.MarginLeft);
                AddShortSprm(0xB606, section.MarginRight);
                AddShortSprm(0xB607, section.MarginTop);
                AddShortSprm(0xB608, section.MarginBottom);

                sepBinaryWriter.Write((short)sepSprms.Count);
                sepBinaryWriter.Write(sepSprms.ToArray());
            }

            // Build Plcfsed in 1Table
            int fcPlcfsed = (int)tableStream.Position;
            tableWriter.Write((int)0);
            tableWriter.Write((int)currentCp);
            foreach (var fcSep in fcSeps)
            {
                tableWriter.Write((short)0); // fn = 0 (WordDocument)
                tableWriter.Write((int)fcSep); // fcSep
                tableWriter.Write((short)0); // reserved
                tableWriter.Write(new byte[6]); // padding to 12 bytes
            }
            int lcbPlcfsed = (int)tableStream.Position - fcPlcfsed;

            // 10. Write File Information Block (FIB)
            wordDocumentStream.Seek(0, SeekOrigin.Begin);
            var fib = new Fib
            {
                fcClx = fcClx,
                lcbClx = lcbClx,
                fcPlcfbteChpx = fcPlcfbteChpx,
                lcbPlcfbteChpx = lcbPlcfbteChpx,
                fcPlcfbtePapx = fcPlcfbtePapx,
                lcbPlcfbtePapx = lcbPlcfbtePapx,
                fcPlcfsed = fcPlcfsed,
                lcbPlcfsed = lcbPlcfsed,
                fcPlcfbteTapx = fcPlcfbteTapx,
                lcbPlcfbteTapx = lcbPlcfbteTapx,
                fcStshf = fcStshf,
                lcbStshf = lcbStshf,
                fcSttbfffn = fcSttbfffn,
                lcbSttbfffn = lcbSttbfffn,
                fcPlcffldMom = fcPlcffldMom,
                lcbPlcffldMom = lcbPlcffldMom,
                fcPlcfspaMom = fcPlcfspaMom,
                lcbPlcfspaMom = lcbPlcfspaMom,
                fcDggInfo = fcDggInfo,
                lcbDggInfo = lcbDggInfo,
                fcSttbLst = fcSttbLst,
                lcbSttbLst = lcbSttbLst,
                fcPlfLfo = fcPlfLfo,
                lcbPlfLfo = lcbPlfLfo,
                fcPlcfBkmkf = fcPlcfBkmkf,
                lcbPlcfBkmkf = lcbPlcfBkmkf,
                fcPlcfBkmkl = fcPlcfBkmkl,
                lcbPlcfBkmkl = lcbPlcfBkmkl,
                fcSttbfbkmk = fcSttbfbkmk,
                lcbSttbfbkmk = lcbSttbfbkmk,
                ccpText = currentCp,
                HasPictures = embeddedObjects.Count > 0
            };
            fib.WriteTo(new BinaryWriter(wordDocumentStream, System.Text.Encoding.GetEncoding(1252), leaveOpen: true));

            // 8. Wrap the streams into OLE Compound File Binary (CFB) format
            using var cfbWriter = new CfbWriter();
            cfbWriter.AddStream("WordDocument", wordDocumentStream.ToArray());
            cfbWriter.AddStream("1Table", tableStream.ToArray());
            cfbWriter.AddStream("Data", dataStream.ToArray());
            
            if (model.VbaProjectData != null)
            {
                cfbWriter.EmbedStorage("Macros", model.VbaProjectData);
            }

            // 9. Write out the final CFB to the destination
            cfbWriter.WriteTo(outputStream);
        }

        private static void WritePictureBlocks(BinaryWriter writer, List<byte[]> pictureBlocks)
        {
            foreach (var pictureBlock in pictureBlocks)
            {
                writer.Write(pictureBlock);
            }
        }

        private static byte[] BuildImageSprms(int pictureOffset)
        {
            var sprms = new List<byte>
            {
                0x55, 0x08, 0x01,
                0x03, 0x6A
            };

            sprms.AddRange(BitConverter.GetBytes(pictureOffset));

            return sprms.ToArray();
        }

        private static byte[] BuildPictureBlock(ImageModel image, string? resolvedContentType = null)
        {
            byte[] imageData = image.Data ?? Array.Empty<byte>();
            string imageContentType = resolvedContentType ?? ResolveImageContentType(image.ContentType, imageData);
            (int widthTwips, int heightTwips) = ResolveImageDimensionsTwips(image, imageContentType);

            byte[] block = new byte[0x44 + imageData.Length];
            BitConverter.GetBytes(block.Length).CopyTo(block, 0x00);
            BitConverter.GetBytes((ushort)0x44).CopyTo(block, 0x04);
            BitConverter.GetBytes(GetPictureMappingMode(imageContentType)).CopyTo(block, 0x06);
            BitConverter.GetBytes(ClampToShort(ConvertTwipsToMm100(widthTwips))).CopyTo(block, 0x08);
            BitConverter.GetBytes(ClampToShort(ConvertTwipsToMm100(heightTwips))).CopyTo(block, 0x0A);
            BitConverter.GetBytes(GetPictureBlockType(imageContentType)).CopyTo(block, 0x0E);
            BitConverter.GetBytes(ClampToShort(widthTwips)).CopyTo(block, 0x1C);
            BitConverter.GetBytes(ClampToShort(heightTwips)).CopyTo(block, 0x1E);
            BitConverter.GetBytes((ushort)1000).CopyTo(block, 0x20);
            BitConverter.GetBytes((ushort)1000).CopyTo(block, 0x22);
            imageData.CopyTo(block, 0x44);

            return block;
        }

        private static bool SupportsOfficeArtBlip(string? contentType)
        {
            return contentType == "image/png" ||
                   contentType == "image/jpeg" ||
                   contentType == "image/x-emf" ||
                   contentType == "image/x-wmf";
        }

        private static List<OfficeArtPictureDescriptor> BuildOfficeArtPictureDescriptors(
            List<(int cp, byte[] data, string contentType, int widthTwips, int heightTwips, int leftTwips, int topTwips, ImageWrapType wrapType, bool behindText, bool allowOverlap, string? horizontalRelativeTo, string? verticalRelativeTo)> officeArtBlips)
        {
            var pictures = new List<OfficeArtPictureDescriptor>(officeArtBlips.Count);
            for (int index = 0; index < officeArtBlips.Count; index++)
            {
                var picture = officeArtBlips[index];
                pictures.Add(new OfficeArtPictureDescriptor(
                    picture.cp,
                    GetPictureShapeId(index),
                    index + 1,
                    picture.data,
                    picture.contentType,
                    picture.widthTwips,
                    picture.heightTwips,
                    picture.leftTwips,
                        picture.topTwips,
                        picture.wrapType,
                        picture.behindText,
                        picture.allowOverlap,
                        picture.horizontalRelativeTo,
                        picture.verticalRelativeTo));
            }

            return pictures;
        }

        private static void WritePlcfspaMom(
            BinaryWriter writer,
            List<OfficeArtPictureDescriptor> officeArtPictures,
            int documentEndCp)
        {
            foreach (var picture in officeArtPictures)
            {
                writer.Write(picture.Cp);
            }

            writer.Write(documentEndCp);

            foreach (var picture in officeArtPictures)
            {
                writer.Write(picture.ShapeId);
                writer.Write(picture.LeftTwips);
                writer.Write(picture.TopTwips);
                writer.Write(picture.RightTwips);
                writer.Write(picture.BottomTwips);
                writer.Write((short)0);
                writer.Write(BuildFloatingLayoutFlags(picture));
            }
        }

        private static byte[] BuildOfficeArtContent(List<OfficeArtPictureDescriptor> officeArtPictures)
        {
            if (officeArtPictures.Count == 0)
            {
                return Array.Empty<byte>();
            }

            var bseRecords = new List<byte[]>(officeArtPictures.Count);
            foreach (var picture in officeArtPictures)
            {
                bseRecords.Add(BuildBlipStoreEntry(picture.Data, picture.ContentType, picture.WidthTwips, picture.HeightTwips));
            }

            byte[] dggRecord = BuildDggRecord(officeArtPictures.Count + 1, GetNextShapeId(officeArtPictures.Count));
            byte[] bstoreContainer = BuildEscherContainer(0xF001, (ushort)officeArtPictures.Count, bseRecords);
            byte[] dggContainer = BuildEscherContainer(0xF000, 0, new[] { dggRecord, bstoreContainer });
            byte[] drawingContainer = BuildDrawingContainer(officeArtPictures);

            using var stream = new MemoryStream();
            stream.Write(dggContainer, 0, dggContainer.Length);
            stream.WriteByte(0x00);
            stream.Write(drawingContainer, 0, drawingContainer.Length);

            return stream.ToArray();
        }

        private static byte[] BuildDrawingContainer(List<OfficeArtPictureDescriptor> officeArtPictures)
        {
            var shapeContainers = new List<byte[]>(officeArtPictures.Count + 1)
            {
                BuildGroupShapeContainer()
            };

            foreach (var picture in officeArtPictures)
            {
                shapeContainers.Add(BuildPictureShapeContainer(picture.ShapeId, picture.BlipIndex));
            }

            byte[] dgRecord = BuildDgRecord(officeArtPictures.Count + 1, GetNextShapeId(officeArtPictures.Count));
            byte[] spgrContainer = BuildEscherContainer(0xF003, 0, shapeContainers);

            return BuildEscherContainer(0xF002, 0, new[] { dgRecord, spgrContainer });
        }

        private static byte[] BuildDggRecord(int savedShapeCount, int nextShapeId)
        {
            byte[] content = new byte[24];
            BitConverter.GetBytes(nextShapeId).CopyTo(content, 0x00);
            BitConverter.GetBytes(2).CopyTo(content, 0x04);
            BitConverter.GetBytes(savedShapeCount).CopyTo(content, 0x08);
            BitConverter.GetBytes(1).CopyTo(content, 0x0C);
            BitConverter.GetBytes(MainDocumentDrawingId).CopyTo(content, 0x10);
            BitConverter.GetBytes(savedShapeCount + 1).CopyTo(content, 0x14);

            return BuildEscherRecord(0xF006, 0, content);
        }

        private static byte[] BuildDgRecord(int shapeCount, int nextShapeId)
        {
            byte[] content = new byte[8];
            BitConverter.GetBytes(shapeCount).CopyTo(content, 0x00);
            BitConverter.GetBytes(nextShapeId).CopyTo(content, 0x04);

            return BuildEscherRecord(0xF008, (ushort)(MainDocumentDrawingId << 4), content);
        }

        private static byte[] BuildGroupShapeContainer()
        {
            byte[] spRecord = BuildSpRecord(GetGroupShapeId(), 0, 0x0805);
            byte[] spgrRecord = BuildSpgrRecord();

            return BuildEscherContainer(0xF004, 0, new[] { spRecord, spgrRecord });
        }

        private static byte[] BuildPictureShapeContainer(int shapeId, int blipIndex)
        {
            byte[] spRecord = BuildSpRecord(shapeId, 75, 0x0A02);
            byte[] optRecord = BuildOptRecord(new[]
            {
                BuildSimpleProperty(OfficeArtBlipToDisplayProperty, true, blipIndex)
            });

            return BuildEscherContainer(0xF004, 0, new[] { spRecord, optRecord });
        }

        private static byte[] BuildSpgrRecord()
        {
            byte[] content = new byte[16];
            BitConverter.GetBytes(0).CopyTo(content, 0x00);
            BitConverter.GetBytes(0).CopyTo(content, 0x04);
            BitConverter.GetBytes(1).CopyTo(content, 0x08);
            BitConverter.GetBytes(1).CopyTo(content, 0x0C);

            return BuildEscherRecord(0xF009, 0x0001, content);
        }

        private static byte[] BuildSpRecord(int shapeId, ushort shapeType, int flags)
        {
            byte[] content = new byte[8];
            BitConverter.GetBytes(shapeId).CopyTo(content, 0x00);
            BitConverter.GetBytes(flags).CopyTo(content, 0x04);

            return BuildEscherRecord(0xF00A, (ushort)((shapeType << 4) | 0x0002), content);
        }

        private static byte[] BuildOptRecord(IEnumerable<byte[]> properties)
        {
            using var stream = new MemoryStream();
            int propertyCount = 0;
            foreach (var property in properties)
            {
                stream.Write(property, 0, property.Length);
                propertyCount++;
            }

            return BuildEscherRecord(0xF00B, (ushort)((propertyCount << 4) | 0x0003), stream.ToArray());
        }

        private static byte[] BuildSimpleProperty(short propertyNumber, bool isBlipId, int value)
        {
            byte[] property = new byte[6];
            short propertyId = (short)(propertyNumber | (isBlipId ? 0x4000 : 0));
            BitConverter.GetBytes(propertyId).CopyTo(property, 0x00);
            BitConverter.GetBytes(value).CopyTo(property, 0x02);

            return property;
        }

        private static int GetGroupShapeId()
        {
            return ShapeIdBase + 1;
        }

        private static int GetPictureShapeId(int pictureIndex)
        {
            return ShapeIdBase + pictureIndex + 2;
        }

        private static int GetNextShapeId(int pictureCount)
        {
            return ShapeIdBase + pictureCount + 3;
        }

        private static byte[] BuildBlipStoreEntry(byte[] imageData, string contentType, int widthTwips, int heightTwips)
        {
            byte blipType = GetOfficeArtBlipType(contentType);
            byte[] uid = ComputeOfficeArtUid(imageData);
            byte[] blipRecord = BuildBlipRecord(imageData, contentType, uid, widthTwips, heightTwips);

            byte[] content = new byte[36 + blipRecord.Length];
            content[0] = blipType;
            content[1] = GetOfficeArtMacBlipType(contentType);
            uid.CopyTo(content, 2);
            BitConverter.GetBytes((short)0).CopyTo(content, 18);
            BitConverter.GetBytes(blipRecord.Length - 8).CopyTo(content, 20);
            BitConverter.GetBytes(1).CopyTo(content, 24);
            BitConverter.GetBytes(0).CopyTo(content, 28);
            content[32] = 0;
            content[33] = 0;
            content[34] = 0;
            content[35] = 0;
            blipRecord.CopyTo(content, 36);

            return BuildEscherRecord(0xF007, (ushort)((blipType << 4) | 0x0002), content);
        }

        private static byte[] BuildBlipRecord(byte[] imageData, string contentType, byte[] uid, int widthTwips, int heightTwips)
        {
            if (contentType == "image/x-emf" || contentType == "image/x-wmf")
            {
                return BuildMetafileBlipRecord(imageData, contentType, uid, widthTwips, heightTwips);
            }

            ushort recordId = contentType == "image/jpeg" ? (ushort)0xF01D : (ushort)0xF01E;
            ushort instance = contentType == "image/jpeg" ? (ushort)0x046A : (ushort)0x06E0;

            byte[] content = new byte[17 + imageData.Length];
            uid.CopyTo(content, 0);
            content[16] = 0xFF;
            imageData.CopyTo(content, 17);

            return BuildEscherRecord(recordId, instance, content);
        }

        private static byte[] BuildMetafileBlipRecord(byte[] imageData, string contentType, byte[] uid, int widthTwips, int heightTwips)
        {
            byte[] normalizedImageData = NormalizeMetafilePayload(imageData, contentType);
            byte[] compressedData = CompressMetafilePayload(normalizedImageData);
            ushort recordId = contentType == "image/x-emf" ? (ushort)0xF01A : (ushort)0xF01B;
            ushort instance = contentType == "image/x-emf" ? (ushort)0x3D40 : (ushort)0x2160;

            byte[] content = new byte[50 + compressedData.Length];
            uid.CopyTo(content, 0);
            BitConverter.GetBytes(normalizedImageData.Length).CopyTo(content, 16);
            BitConverter.GetBytes(0).CopyTo(content, 20);
            BitConverter.GetBytes(0).CopyTo(content, 24);
            BitConverter.GetBytes(widthTwips).CopyTo(content, 28);
            BitConverter.GetBytes(heightTwips).CopyTo(content, 32);
            BitConverter.GetBytes(ConvertTwipsToEmu(widthTwips)).CopyTo(content, 36);
            BitConverter.GetBytes(ConvertTwipsToEmu(heightTwips)).CopyTo(content, 40);
            BitConverter.GetBytes(compressedData.Length).CopyTo(content, 44);
            content[48] = 0x00;
            content[49] = 0xFE;
            compressedData.CopyTo(content, 50);

            return BuildEscherRecord(recordId, instance, content);
        }

        private static byte GetOfficeArtBlipType(string contentType)
        {
            return contentType switch
            {
                "image/x-emf" => 0x02,
                "image/x-wmf" => 0x03,
                "image/jpeg" => 0x05,
                "image/png" => 0x06,
                _ => 0x00
            };
        }

        private static byte GetOfficeArtMacBlipType(string contentType)
        {
            return contentType switch
            {
                "image/x-emf" => 0x04,
                "image/x-wmf" => 0x04,
                _ => GetOfficeArtBlipType(contentType)
            };
        }

        private static byte[] NormalizeMetafilePayload(byte[] imageData, string contentType)
        {
            if (contentType == "image/x-wmf" && HasWmfPlaceableHeader(imageData))
            {
                byte[] normalized = new byte[imageData.Length - 22];
                Buffer.BlockCopy(imageData, 22, normalized, 0, normalized.Length);
                return normalized;
            }

            return imageData;
        }

        private static bool HasWmfPlaceableHeader(byte[] imageData)
        {
            return imageData.Length > 22 &&
                   imageData[0] == 0xD7 &&
                   imageData[1] == 0xCD &&
                   imageData[2] == 0xC6 &&
                   imageData[3] == 0x9A;
        }

        private static byte[] CompressMetafilePayload(byte[] imageData)
        {
            using var output = new MemoryStream();
            using (var deflater = new DeflateStream(output, CompressionLevel.Optimal, leaveOpen: true))
            {
                deflater.Write(imageData, 0, imageData.Length);
            }

            return output.ToArray();
        }

        private static byte[] ComputeOfficeArtUid(byte[] imageData)
        {
            using var md5 = MD5.Create();
            return md5.ComputeHash(imageData);
        }

        private static byte[] BuildEscherContainer(ushort recordId, ushort instance, IEnumerable<byte[]> children)
        {
            using var stream = new MemoryStream();
            foreach (var child in children)
            {
                stream.Write(child, 0, child.Length);
            }

            return BuildEscherRecord(recordId, (ushort)((instance << 4) | 0x000F), stream.ToArray());
        }

        private static byte[] BuildEscherRecord(ushort recordId, ushort options, byte[] content)
        {
            byte[] record = new byte[8 + content.Length];
            BitConverter.GetBytes(options).CopyTo(record, 0x00);
            BitConverter.GetBytes(recordId).CopyTo(record, 0x02);
            BitConverter.GetBytes(content.Length).CopyTo(record, 0x04);
            content.CopyTo(record, 0x08);
            return record;
        }

        private static short GetPictureMappingMode(string? contentType)
        {
            return contentType switch
            {
                "image/x-wmf" => 8,
                "image/x-emf" => 8,
                _ => 0x64
            };
        }

        private static short GetPictureBlockType(string? contentType)
        {
            return contentType switch
            {
                "image/x-wmf" => 0x08,
                "image/x-emf" => 0x08,
                _ => 0x00
            };
        }

        private static string ResolveImageContentType(string? contentType, byte[] imageData)
        {
            if (!string.IsNullOrWhiteSpace(contentType))
            {
                return contentType;
            }

            if (imageData.Length >= 8 &&
                imageData[0] == 0x89 &&
                imageData[1] == 0x50 &&
                imageData[2] == 0x4E &&
                imageData[3] == 0x47 &&
                imageData[4] == 0x0D &&
                imageData[5] == 0x0A &&
                imageData[6] == 0x1A &&
                imageData[7] == 0x0A)
            {
                return "image/png";
            }

            if (imageData.Length >= 3 &&
                imageData[0] == 0x47 &&
                imageData[1] == 0x49 &&
                imageData[2] == 0x46)
            {
                return "image/gif";
            }

            if (imageData.Length >= 2 &&
                imageData[0] == 0xFF &&
                imageData[1] == 0xD8)
            {
                return "image/jpeg";
            }

            if (imageData.Length >= 2 &&
                imageData[0] == 0x42 &&
                imageData[1] == 0x4D)
            {
                return "image/bmp";
            }

            if (imageData.Length >= 4 &&
                ((imageData[0] == 0x49 && imageData[1] == 0x49 && imageData[2] == 0x2A && imageData[3] == 0x00) ||
                 (imageData[0] == 0x4D && imageData[1] == 0x4D && imageData[2] == 0x00 && imageData[3] == 0x2A)))
            {
                return "image/tiff";
            }

            if (imageData.Length >= 4 &&
                imageData[0] == 0x01 &&
                imageData[1] == 0x00 &&
                imageData[2] == 0x00 &&
                imageData[3] == 0x00)
            {
                return "image/x-emf";
            }

            if (imageData.Length >= 4 &&
                imageData[0] == 0xD7 &&
                imageData[1] == 0xCD &&
                imageData[2] == 0xC6 &&
                imageData[3] == 0x9A)
            {
                return "image/x-wmf";
            }

            return string.Empty;
        }

        private static int GetImageDimensionTwips(int pixels)
        {
            if (pixels <= 0)
            {
                return 1;
            }

            return pixels * 15;
        }

        private static (int widthTwips, int heightTwips) ResolveImageDimensionsTwips(ImageModel image, string? contentType)
        {
            int widthTwips = GetImageDimensionTwips(image.Width);
            int heightTwips = GetImageDimensionTwips(image.Height);

            if (image.Width > 0 && image.Height > 0)
            {
                return (widthTwips, heightTwips);
            }

            if (TryGetImageDimensionsTwipsFromPayload(image.Data, contentType, out int inferredWidthTwips, out int inferredHeightTwips))
            {
                if (image.Width <= 0)
                {
                    widthTwips = inferredWidthTwips;
                }

                if (image.Height <= 0)
                {
                    heightTwips = inferredHeightTwips;
                }
            }

            return (widthTwips, heightTwips);
        }

        private static (int leftTwips, int topTwips, int rightTwips, int bottomTwips) ResolveImageBoundsTwips(ImageModel image, string? contentType, SectionModel section, int paragraphTopTwips, int paragraphHeightTwips)
        {
            (int widthTwips, int heightTwips) = ResolveImageDimensionsTwips(image, contentType);
            int leftTwips = 0;
            int topTwips = 0;

            if (image.LayoutType == ImageLayoutType.Floating)
            {
                leftTwips = ResolveAlignedPositionTwips(
                    image.HorizontalAlignment,
                    image.HorizontalRelativeTo,
                    image.PositionXTwips,
                    widthTwips,
                    section.PageWidth,
                    section.MarginLeft,
                    section.MarginRight,
                    isHorizontal: true);

                topTwips = ResolveAlignedPositionTwips(
                    image.VerticalAlignment,
                    image.VerticalRelativeTo,
                    image.PositionYTwips,
                    heightTwips,
                    section.PageHeight,
                    section.MarginTop,
                    section.MarginBottom,
                    isHorizontal: false,
                    paragraphStartTwips: paragraphTopTwips,
                    paragraphExtentTwips: paragraphHeightTwips);
            }

            return (leftTwips, topTwips, leftTwips + widthTwips, topTwips + heightTwips);
        }

        private static int EstimateParagraphContentHeightTwips(ParagraphModel paragraph, int availableWidthTwips)
        {
            int maxFontSizeHalfPoints = 24;
            foreach (var run in paragraph.Runs)
            {
                if (run.Properties.FontSize.HasValue)
                {
                    maxFontSizeHalfPoints = Math.Max(maxFontSizeHalfPoints, run.Properties.FontSize.Value);
                }
            }

            int baseLineHeightTwips = (maxFontSizeHalfPoints * 10) + 40;
            int lineHeightTwips = Math.Max(276, (baseLineHeightTwips * 115 + 99) / 100);
            if (paragraph.Properties.LineSpacing.HasValue)
            {
                string rule = paragraph.Properties.LineSpacingRule ?? "auto";
                if (string.Equals(rule, "exact", StringComparison.OrdinalIgnoreCase))
                {
                    lineHeightTwips = Math.Max(1, paragraph.Properties.LineSpacing.Value);
                }
                else if (string.Equals(rule, "atLeast", StringComparison.OrdinalIgnoreCase))
                {
                    lineHeightTwips = Math.Max(lineHeightTwips, paragraph.Properties.LineSpacing.Value);
                }
                else
                {
                    double multiplier = Math.Max(1d / 240d, paragraph.Properties.LineSpacing.Value / 240d);
                    lineHeightTwips = Math.Max(1, (int)Math.Round(lineHeightTwips * multiplier, MidpointRounding.AwayFromZero));
                }
            }

            int estimatedLineCount = EstimateParagraphLineCount(paragraph, maxFontSizeHalfPoints, availableWidthTwips);
            return Math.Max(lineHeightTwips, lineHeightTwips * estimatedLineCount);
        }

        private static int ResolveParagraphAvailableWidthTwips(ParagraphModel paragraph, int baseAvailableWidthTwips)
        {
            int indentWidthTwips = Math.Max(0, paragraph.Properties.LeftIndentTwips)
                + Math.Max(0, paragraph.Properties.RightIndentTwips)
                + Math.Max(0, paragraph.Properties.FirstLineIndentTwips);

            return Math.Max(720, baseAvailableWidthTwips - indentWidthTwips);
        }

        private static int ResolveTableCellWidth(TableModel table, int gridColumnIndex, int gridSpan, int fallbackWidthTwips)
        {
            if (table.GridColumnWidths.Count == 0)
            {
                return fallbackWidthTwips;
            }

            int span = Math.Max(1, gridSpan);
            int widthTwips = 0;
            for (int index = 0; index < span; index++)
            {
                int currentGridIndex = gridColumnIndex + index;
                if (currentGridIndex >= 0 && currentGridIndex < table.GridColumnWidths.Count)
                {
                    widthTwips += table.GridColumnWidths[currentGridIndex];
                }
            }

            return widthTwips > 0 ? widthTwips : fallbackWidthTwips;
        }

        private static int ResolveTableCellHorizontalPaddingTwips(TableModel table, TableCellModel cell)
        {
            int leftPaddingTwips = ResolveTableCellLeftPaddingTwips(table, cell);
            int rightPaddingTwips = ResolveTableCellRightPaddingTwips(table, cell);
            return Math.Max(0, leftPaddingTwips) + Math.Max(0, rightPaddingTwips);
        }

        private static int ResolveTableTotalColumnCount(TableModel table, TableRowModel row)
        {
            if (table.GridColumnWidths.Count > 0)
            {
                return table.GridColumnWidths.Count;
            }

            int totalColumns = 0;
            foreach (var cell in row.Cells)
            {
                totalColumns += Math.Max(1, cell.GridSpan);
            }

            return Math.Max(1, totalColumns);
        }

        private static int ResolveTableAvailableWidthTwips(TableModel table, TableRowModel row, int documentAvailableWidthTwips, int totalColumnCount)
        {
            int preferredWidthTwips = ResolvePreferredTableWidthTwips(table, documentAvailableWidthTwips);
            if (preferredWidthTwips > 0)
            {
                return preferredWidthTwips;
            }

            int baseWidthTwips = ResolveBaseTableWidthTwips(table, row, totalColumnCount, documentAvailableWidthTwips);
            return baseWidthTwips > 0 ? baseWidthTwips : documentAvailableWidthTwips;
        }

        private static int ResolvePreferredTableWidthTwips(TableModel table, int documentAvailableWidthTwips)
        {
            return table.PreferredWidthUnit switch
            {
                TableWidthUnit.Dxa when table.PreferredWidthValue > 0 => Math.Max(720, table.PreferredWidthValue),
                TableWidthUnit.Pct when table.PreferredWidthValue > 0 => Math.Max(720, (int)Math.Round(documentAvailableWidthTwips * (table.PreferredWidthValue / 5000d), MidpointRounding.AwayFromZero)),
                _ => 0
            };
        }

        private static int ResolveBaseTableWidthTwips(TableModel table, TableRowModel row, int totalColumnCount, int fallbackWidthTwips)
        {
            if (table.GridColumnWidths.Count > 0)
            {
                int gridWidthTwips = 0;
                foreach (int gridWidth in table.GridColumnWidths)
                {
                    gridWidthTwips += Math.Max(0, gridWidth);
                }

                if (gridWidthTwips > 0)
                {
                    return gridWidthTwips;
                }
            }

            int explicitWidthTwips = 0;
            foreach (var cell in row.Cells)
            {
                int cellWidthTwips = ResolveExplicitCellWidthTwips(cell, fallbackWidthTwips);
                if (cellWidthTwips > 0)
                {
                    explicitWidthTwips += cellWidthTwips;
                }
            }

            if (explicitWidthTwips > 0)
            {
                return explicitWidthTwips;
            }

            return totalColumnCount > 0 ? fallbackWidthTwips : 0;
        }

        private static List<int> ResolveRowCellWidthsTwips(TableModel table, TableRowModel row, int targetTableWidthTwips, int totalColumnCount)
        {
            var widthsTwips = new List<int>(row.Cells.Count);
            var spans = new List<int>(row.Cells.Count);
            var isExplicitResolvedWidth = new List<bool>(row.Cells.Count);
            var explicitWidthUnits = new List<TableWidthUnit>(row.Cells.Count);
            var minimumAutoWidthsTwips = new List<int>(row.Cells.Count);
            int gridColumnIndex = 0;
            int resolvedWidthSumTwips = 0;
            int unresolvedSpanSum = 0;
            int unresolvedCellCount = 0;

            foreach (var cell in row.Cells)
            {
                int span = Math.Max(1, cell.GridSpan);
                int explicitWidthTwips = ResolveExplicitCellWidthTwips(cell, targetTableWidthTwips);
                int widthTwips = explicitWidthTwips;
                bool resolvedFromExplicitWidth = explicitWidthTwips > 0;
                if (widthTwips <= 0 && table.GridColumnWidths.Count > 0)
                {
                    widthTwips = ResolveTableCellWidth(table, gridColumnIndex, span, 0);
                }

                widthsTwips.Add(Math.Max(0, widthTwips));
                spans.Add(span);
                isExplicitResolvedWidth.Add(resolvedFromExplicitWidth);
                explicitWidthUnits.Add(resolvedFromExplicitWidth ? cell.WidthUnit : TableWidthUnit.Auto);
                minimumAutoWidthsTwips.Add(Math.Max(720, 720 * span));

                if (widthTwips > 0)
                {
                    resolvedWidthSumTwips += widthTwips;
                }
                else
                {
                    unresolvedSpanSum += span;
                    unresolvedCellCount += 1;
                }

                gridColumnIndex += span;
            }

            if (unresolvedSpanSum > 0)
            {
                int reservedMinimumAutoWidthTwips = 0;
                for (int index = 0; index < widthsTwips.Count; index++)
                {
                    if (widthsTwips[index] == 0)
                    {
                        reservedMinimumAutoWidthTwips += minimumAutoWidthsTwips[index];
                    }
                }

                if (resolvedWidthSumTwips > 0 && targetTableWidthTwips > 0 && resolvedWidthSumTwips + reservedMinimumAutoWidthTwips > targetTableWidthTwips)
                {
                    int availableResolvedWidthTwips = Math.Max(unresolvedCellCount, targetTableWidthTwips - reservedMinimumAutoWidthTwips);
                    if (!ScaleExplicitWidthsToTarget(widthsTwips, isExplicitResolvedWidth, explicitWidthUnits, availableResolvedWidthTwips))
                    {
                        ScaleResolvedWidthsToTarget(widthsTwips, availableResolvedWidthTwips);
                    }

                    resolvedWidthSumTwips = 0;
                    foreach (int widthTwips in widthsTwips)
                    {
                        resolvedWidthSumTwips += Math.Max(0, widthTwips);
                    }
                }

                int fallbackPoolTwips = Math.Max(0, targetTableWidthTwips - resolvedWidthSumTwips);
                var cellContentWidthsTwips = new List<int>(widthsTwips.Count);
                int totalUnresolvedContentWidthTwips = 0;
                for (int index = 0; index < widthsTwips.Count; index++)
                {
                    if (widthsTwips[index] > 0)
                    {
                        cellContentWidthsTwips.Add(0);
                    }
                    else
                    {
                        int estimatedContentWidth = EstimateTableCellContentWidthTwips(row.Cells[index]);
                        cellContentWidthsTwips.Add(Math.Max(1, estimatedContentWidth));
                        totalUnresolvedContentWidthTwips += Math.Max(1, estimatedContentWidth);
                    }
                }

                int assignedFallbackTwips = 0;
                int remainingUnresolvedSpan = unresolvedSpanSum;
                int remainingUnresolvedCells = unresolvedCellCount;
                int assignedExtraTwips = 0;

                for (int index = 0; index < widthsTwips.Count; index++)
                {
                    if (widthsTwips[index] > 0)
                    {
                        continue;
                    }

                    int span = spans[index];
                    int minimumWidthTwips = minimumAutoWidthsTwips[index];
                    int distributablePoolTwips = Math.Max(0, fallbackPoolTwips - reservedMinimumAutoWidthTwips);
                    int contentWidth = cellContentWidthsTwips[index];

                    int extraWidthTwips = remainingUnresolvedCells > 0
                        ? (remainingUnresolvedCells == 1
                            ? Math.Max(0, distributablePoolTwips - assignedExtraTwips)
                            : (int)Math.Round(distributablePoolTwips * (contentWidth / (double)totalUnresolvedContentWidthTwips), MidpointRounding.AwayFromZero))
                        : 0;

                    int widthTwips = minimumWidthTwips + Math.Max(0, extraWidthTwips);

                    widthTwips = Math.Max(1, widthTwips);
                    widthsTwips[index] = widthTwips;
                    assignedFallbackTwips += widthTwips;
                    assignedExtraTwips += Math.Max(0, extraWidthTwips);
                    remainingUnresolvedSpan -= span;
                    remainingUnresolvedCells -= 1;
                }
            }

            ScaleWidthsToTarget(widthsTwips, targetTableWidthTwips);
            return widthsTwips;
        }

        private static bool ScaleExplicitWidthsToTarget(List<int> widthsTwips, List<bool> isExplicitResolvedWidth, List<TableWidthUnit> explicitWidthUnits, int targetResolvedWidthTwips)
        {
            if (targetResolvedWidthTwips <= 0 || widthsTwips.Count != isExplicitResolvedWidth.Count || widthsTwips.Count != explicitWidthUnits.Count)
            {
                return false;
            }

            int fixedResolvedWidthTwips = 0;
            int pctExplicitWidthTwips = 0;
            int dxaExplicitWidthTwips = 0;
            int dxaExplicitCount = 0;
            for (int index = 0; index < widthsTwips.Count; index++)
            {
                int widthTwips = Math.Max(0, widthsTwips[index]);
                if (widthTwips == 0)
                {
                    continue;
                }

                if (isExplicitResolvedWidth[index])
                {
                    if (explicitWidthUnits[index] == TableWidthUnit.Pct)
                    {
                        pctExplicitWidthTwips += widthTwips;
                    }
                    else
                    {
                        dxaExplicitWidthTwips += widthTwips;
                        dxaExplicitCount += 1;
                    }
                }
                else
                {
                    fixedResolvedWidthTwips += widthTwips;
                }
            }

            int targetDxaExplicitWidthTwips = targetResolvedWidthTwips - fixedResolvedWidthTwips - pctExplicitWidthTwips;
            if (targetDxaExplicitWidthTwips > 0 && targetDxaExplicitWidthTwips < dxaExplicitWidthTwips)
            {
                // Enough room for fixed and PCT, just scale down DXA
                var shouldScale = new List<bool>(widthsTwips.Count);
                for (int index = 0; index < widthsTwips.Count; index++)
                {
                    shouldScale.Add(isExplicitResolvedWidth[index] && explicitWidthUnits[index] != TableWidthUnit.Pct);
                }
                ScaleSelectedWidthsToTarget(widthsTwips, shouldScale, Math.Max(dxaExplicitCount, targetDxaExplicitWidthTwips));
                return true;
            }
            else if (targetDxaExplicitWidthTwips <= 0)
            {
                // PCT alone overflows target (or leaves no room for DXA). Scale both PCT and DXA, protecting auto/fixed widths.
                var shouldScale = new List<bool>(widthsTwips.Count);
                for (int index = 0; index < widthsTwips.Count; index++)
                {
                    shouldScale.Add(isExplicitResolvedWidth[index]);
                }
                
                int targetExplicitWidth = Math.Max(1, targetResolvedWidthTwips - fixedResolvedWidthTwips);
                if (targetExplicitWidth < (dxaExplicitWidthTwips + pctExplicitWidthTwips))
                {
                    ScaleSelectedWidthsToTarget(widthsTwips, shouldScale, targetExplicitWidth);
                    return true;
                }
            }

            return false;
        }

        private static void ScaleResolvedWidthsToTarget(List<int> widthsTwips, int targetResolvedWidthTwips)
        {
            if (targetResolvedWidthTwips <= 0)
            {
                return;
            }

            int currentResolvedWidthTwips = 0;
            foreach (int widthTwips in widthsTwips)
            {
                if (widthTwips > 0)
                {
                    currentResolvedWidthTwips += widthTwips;
                }
            }

            if (currentResolvedWidthTwips <= 0 || currentResolvedWidthTwips <= targetResolvedWidthTwips)
            {
                return;
            }

            int scaledResolvedTotalTwips = 0;
            int lastResolvedIndex = -1;
            for (int index = 0; index < widthsTwips.Count; index++)
            {
                if (widthsTwips[index] <= 0)
                {
                    continue;
                }

                lastResolvedIndex = index;
                int scaledWidthTwips = Math.Max(1, (int)Math.Round(widthsTwips[index] * (targetResolvedWidthTwips / (double)currentResolvedWidthTwips), MidpointRounding.AwayFromZero));
                widthsTwips[index] = scaledWidthTwips;
                scaledResolvedTotalTwips += scaledWidthTwips;
            }

            if (lastResolvedIndex >= 0 && scaledResolvedTotalTwips != targetResolvedWidthTwips)
            {
                widthsTwips[lastResolvedIndex] = Math.Max(1, widthsTwips[lastResolvedIndex] + (targetResolvedWidthTwips - scaledResolvedTotalTwips));
            }
        }

        private static void ScaleSelectedWidthsToTarget(List<int> widthsTwips, List<bool> selectedWidths, int targetSelectedWidthTwips)
        {
            if (targetSelectedWidthTwips <= 0 || widthsTwips.Count != selectedWidths.Count)
            {
                return;
            }

            int currentSelectedWidthTwips = 0;
            foreach (var pair in widthsTwips.Select((widthTwips, index) => (widthTwips, index)))
            {
                if (selectedWidths[pair.index] && pair.widthTwips > 0)
                {
                    currentSelectedWidthTwips += pair.widthTwips;
                }
            }

            if (currentSelectedWidthTwips <= 0 || currentSelectedWidthTwips == targetSelectedWidthTwips)
            {
                return;
            }

            int scaledSelectedTotalTwips = 0;
            int lastSelectedIndex = -1;
            for (int index = 0; index < widthsTwips.Count; index++)
            {
                if (!selectedWidths[index] || widthsTwips[index] <= 0)
                {
                    continue;
                }

                lastSelectedIndex = index;
                int scaledWidthTwips = Math.Max(1, (int)Math.Round(widthsTwips[index] * (targetSelectedWidthTwips / (double)currentSelectedWidthTwips), MidpointRounding.AwayFromZero));
                widthsTwips[index] = scaledWidthTwips;
                scaledSelectedTotalTwips += scaledWidthTwips;
            }

            if (lastSelectedIndex >= 0 && scaledSelectedTotalTwips != targetSelectedWidthTwips)
            {
                widthsTwips[lastSelectedIndex] = Math.Max(1, widthsTwips[lastSelectedIndex] + (targetSelectedWidthTwips - scaledSelectedTotalTwips));
            }
        }

        private static void ScaleWidthsToTarget(List<int> widthsTwips, int targetTableWidthTwips)
        {
            if (widthsTwips.Count == 0 || targetTableWidthTwips <= 0)
            {
                return;
            }

            int currentTotalTwips = 0;
            foreach (int widthTwips in widthsTwips)
            {
                currentTotalTwips += Math.Max(0, widthTwips);
            }

            if (currentTotalTwips <= 0 || currentTotalTwips == targetTableWidthTwips)
            {
                return;
            }

            int scaledTotalTwips = 0;
            int lastPositiveWidthIndex = -1;
            for (int index = 0; index < widthsTwips.Count; index++)
            {
                int widthTwips = widthsTwips[index];
                if (widthTwips <= 0)
                {
                    continue;
                }

                lastPositiveWidthIndex = index;
                int scaledWidthTwips = Math.Max(1, (int)Math.Round(widthTwips * (targetTableWidthTwips / (double)currentTotalTwips), MidpointRounding.AwayFromZero));
                widthsTwips[index] = scaledWidthTwips;
                scaledTotalTwips += scaledWidthTwips;
            }

            if (lastPositiveWidthIndex >= 0 && scaledTotalTwips != targetTableWidthTwips)
            {
                widthsTwips[lastPositiveWidthIndex] = Math.Max(1, widthsTwips[lastPositiveWidthIndex] + (targetTableWidthTwips - scaledTotalTwips));
            }
        }

        private static int ResolveEffectiveTableCellWidthTwips(TableModel table, TableRowModel row, TableCellModel cell, int gridColumnIndex, int gridSpan, int tableAvailableWidthTwips, int totalColumnCount)
        {
            int baseTableWidthTwips = ResolveBaseTableWidthTwips(table, row, totalColumnCount, tableAvailableWidthTwips);
            int explicitCellWidthTwips = ResolveExplicitCellWidthTwips(cell, tableAvailableWidthTwips);
            if (explicitCellWidthTwips > 0)
            {
                return ScaleTableCellWidthTwips(explicitCellWidthTwips, tableAvailableWidthTwips, baseTableWidthTwips);
            }

            if (table.GridColumnWidths.Count > 0)
            {
                int gridWidthTwips = ResolveTableCellWidth(table, gridColumnIndex, gridSpan, tableAvailableWidthTwips);
                return ScaleTableCellWidthTwips(gridWidthTwips, tableAvailableWidthTwips, baseTableWidthTwips);
            }

            int span = Math.Max(1, gridSpan);
            int autoWidthTwips = Math.Max(720, (int)Math.Round(tableAvailableWidthTwips * (span / (double)Math.Max(1, totalColumnCount)), MidpointRounding.AwayFromZero));
            return autoWidthTwips;
        }

        private static int ResolveExplicitCellWidthTwips(TableCellModel cell, int referenceTableWidthTwips)
        {
            if (cell.Width <= 0)
            {
                return 0;
            }

            return cell.WidthUnit switch
            {
                TableWidthUnit.Pct when referenceTableWidthTwips > 0 => Math.Max(1, (int)Math.Round(referenceTableWidthTwips * (cell.Width / 5000d), MidpointRounding.AwayFromZero)),
                TableWidthUnit.Auto => 0,
                _ => Math.Max(0, cell.Width)
            };
        }

        private static int ScaleTableCellWidthTwips(int cellWidthTwips, int targetTableWidthTwips, int baseTableWidthTwips)
        {
            if (cellWidthTwips <= 0)
            {
                return 0;
            }

            if (targetTableWidthTwips <= 0 || baseTableWidthTwips <= 0 || targetTableWidthTwips == baseTableWidthTwips)
            {
                return cellWidthTwips;
            }

            return Math.Max(720, (int)Math.Round(cellWidthTwips * (targetTableWidthTwips / (double)baseTableWidthTwips), MidpointRounding.AwayFromZero));
        }

        private static int ResolveTableCellLeftBorderTwips(TableModel table, TableCellModel cell, TableCellModel? previousCell, bool isFirstColumn)
        {
            int currentLeft = HasExplicitLeftBorder(cell) ? Math.Max(0, cell.BorderLeftTwips) : -1;
            int prevRight = (previousCell != null && HasExplicitRightBorder(previousCell)) ? Math.Max(0, previousCell.BorderRightTwips) : -1;

            if (currentLeft >= 0 && prevRight >= 0)
            {
                return ShouldOverrideBorder(currentLeft, cell.BorderLeftStyle, prevRight, previousCell!.BorderRightStyle) ? currentLeft : prevRight;
            }
            else if (currentLeft >= 0)
            {
                return currentLeft;
            }
            else if (prevRight >= 0)
            {
                return prevRight;
            }

            if (isFirstColumn)
            {
                return Math.Max(0, table.DefaultBorderLeftTwips);
            }

            return Math.Max(0, table.DefaultInsideVerticalBorderTwips);
        }

        private static int ResolveTableCellRightBorderTwips(TableModel table, TableCellModel cell, bool isLastColumn)
        {
            if (HasExplicitRightBorder(cell))
            {
                return Math.Max(0, cell.BorderRightTwips);
            }

            return Math.Max(0, isLastColumn ? table.DefaultBorderRightTwips : 0);
        }

        private static int ResolveTableCellHorizontalSpacingTwips(TableModel table)
        {
            return Math.Max(0, table.CellSpacingTwips);
        }

        private static int ResolveTableCellTopSpacingTwips(TableModel table)
        {
            return Math.Max(0, table.CellSpacingTwips / 2);
        }

        private static int ResolveTableCellBottomSpacingTwips(TableModel table)
        {
            int spacingTwips = Math.Max(0, table.CellSpacingTwips);
            return spacingTwips - (spacingTwips / 2);
        }

        private static int ResolveTableCellTopPaddingTwips(TableModel table, TableCellModel cell)
        {
            return Math.Max(0, HasExplicitTopPadding(cell) ? cell.PaddingTopTwips : table.DefaultCellPaddingTopTwips);
        }

        private static int ResolveTableCellBottomPaddingTwips(TableModel table, TableCellModel cell)
        {
            return Math.Max(0, HasExplicitBottomPadding(cell) ? cell.PaddingBottomTwips : table.DefaultCellPaddingBottomTwips);
        }

        private static int ResolveTableCellLeftPaddingTwips(TableModel table, TableCellModel cell)
        {
            return HasExplicitLeftPadding(cell) ? cell.PaddingLeftTwips : table.DefaultCellPaddingLeftTwips;
        }

        private static int ResolveTableCellRightPaddingTwips(TableModel table, TableCellModel cell)
        {
            return HasExplicitRightPadding(cell) ? cell.PaddingRightTwips : table.DefaultCellPaddingRightTwips;
        }

        private static int ResolveTableCellTopBorderTwips(TableModel table, TableCellModel cell, TableRowModel? previousRow, int currentStartColumnIndex, int currentSpan, bool isFirstRow)
        {
            int currentTop = HasExplicitTopBorder(cell) ? Math.Max(0, cell.BorderTopTwips) : -1;
            var previousBottom = ResolvePreviousRowBottomBorder(previousRow, currentStartColumnIndex, currentSpan);
            int prevBottomVal = previousBottom.hasExplicitOverride ? previousBottom.widthTwips : -1;

            if (currentTop >= 0 && prevBottomVal >= 0)
            {
                return ShouldOverrideBorder(currentTop, cell.BorderTopStyle, prevBottomVal, previousBottom.style) ? currentTop : prevBottomVal;
            }
            else if (currentTop >= 0)
            {
                return currentTop;
            }
            else if (prevBottomVal >= 0)
            {
                return prevBottomVal;
            }

            if (isFirstRow)
            {
                return Math.Max(0, table.DefaultBorderTopTwips);
            }

            return Math.Max(0, table.DefaultInsideHorizontalBorderTwips);
        }

        private static int ResolveTableCellBottomBorderTwips(TableModel table, TableCellModel cell, bool isLastRow)
        {
            if (HasExplicitBottomBorder(cell))
            {
                return Math.Max(0, cell.BorderBottomTwips);
            }

            return Math.Max(0, isLastRow ? table.DefaultBorderBottomTwips : 0);
        }

        private static int ResolveTableCellVerticalAlignmentOffset(TableCellModel cell, int rowHeightTwips, int cellTotalHeightTwips)
        {
            int remainingHeightTwips = Math.Max(0, rowHeightTwips - cellTotalHeightTwips);
            return cell.VerticalAlignment switch
            {
                TableCellVerticalAlignment.Center => remainingHeightTwips / 2,
                TableCellVerticalAlignment.Bottom => remainingHeightTwips,
                _ => 0
            };
        }

        private static int ResolveRowHeightTwips(TableRowModel row, int contentHeightTwips)
        {
            if (row.HeightTwips <= 0)
            {
                return contentHeightTwips;
            }

            return row.HeightRule switch
            {
                TableRowHeightRule.Exact => Math.Max(0, row.HeightTwips),
                TableRowHeightRule.AtLeast => Math.Max(contentHeightTwips, row.HeightTwips),
                _ => contentHeightTwips
            };
        }

        private static (int widthTwips, BorderStyle style, bool hasExplicitOverride) ResolvePreviousRowBottomBorder(TableRowModel? previousRow, int startColumnIndex, int span)
        {
            if (previousRow == null)
            {
                return (0, BorderStyle.None, false);
            }

            int currentColumnIndex = 0;
            int selectedWidthTwips = -1;
            BorderStyle selectedStyle = BorderStyle.None;
            bool hasExplicitOverride = false;
            foreach (var previousCell in previousRow.Cells)
            {
                int previousSpan = Math.Max(1, previousCell.GridSpan);
                int previousEndColumnIndex = currentColumnIndex + previousSpan;
                int currentEndColumnIndex = startColumnIndex + Math.Max(1, span);
                bool overlaps = currentColumnIndex < currentEndColumnIndex && previousEndColumnIndex > startColumnIndex;
                if (overlaps)
                {
                    if (HasExplicitBottomBorder(previousCell))
                    {
                        hasExplicitOverride = true;
                        int width = Math.Max(0, previousCell.BorderBottomTwips);
                        if (selectedWidthTwips < 0 || ShouldOverrideBorder(width, previousCell.BorderBottomStyle, selectedWidthTwips, selectedStyle))
                        {
                            selectedWidthTwips = width;
                            selectedStyle = previousCell.BorderBottomStyle;
                        }
                    }
                }

                currentColumnIndex = previousEndColumnIndex;
            }

            return (Math.Max(0, selectedWidthTwips), selectedStyle, hasExplicitOverride);
        }

        private static bool ShouldOverrideBorder(int widthA, BorderStyle styleA, int widthB, BorderStyle styleB)
        {
            if (widthA > widthB) return true;
            if (widthA < widthB) return false;

            return GetBorderStylePrecedence(styleA) > GetBorderStylePrecedence(styleB);
        }

        private static int GetBorderStylePrecedence(BorderStyle style)
        {
            return style switch
            {
                BorderStyle.Double => 5,
                BorderStyle.Single => 4,
                BorderStyle.Dashed => 3,
                BorderStyle.Dotted => 2,
                BorderStyle.Other => 1,
                _ => 0
            };
        }

        private static bool HasExplicitLeftBorder(TableCellModel cell)
        {
            return cell.HasLeftBorderOverride || cell.BorderLeftTwips > 0;
        }

        private static bool HasExplicitLeftPadding(TableCellModel cell)
        {
            return cell.HasLeftPaddingOverride || cell.PaddingLeftTwips > 0;
        }

        private static bool HasExplicitRightBorder(TableCellModel cell)
        {
            return cell.HasRightBorderOverride || cell.BorderRightTwips > 0;
        }

        private static bool HasExplicitRightPadding(TableCellModel cell)
        {
            return cell.HasRightPaddingOverride || cell.PaddingRightTwips > 0;
        }

        private static bool HasExplicitTopBorder(TableCellModel cell)
        {
            return cell.HasTopBorderOverride || cell.BorderTopTwips > 0;
        }

        private static bool HasExplicitTopPadding(TableCellModel cell)
        {
            return cell.HasTopPaddingOverride || cell.PaddingTopTwips > 0;
        }

        private static bool HasExplicitBottomBorder(TableCellModel cell)
        {
            return cell.HasBottomBorderOverride || cell.BorderBottomTwips > 0;
        }

        private static bool HasExplicitBottomPadding(TableCellModel cell)
        {
            return cell.HasBottomPaddingOverride || cell.PaddingBottomTwips > 0;
        }

        private static int EstimateTableCellContentHeightTwips(TableCellModel cell, int cellAvailableWidthTwips)
        {
            int contentHeightTwips = 0;
            foreach (var paragraph in cell.Paragraphs)
            {
                int paragraphAvailableWidthTwips = ResolveParagraphAvailableWidthTwips(paragraph, cellAvailableWidthTwips);
                int paragraphContentHeightTwips = EstimateParagraphContentHeightTwips(paragraph, paragraphAvailableWidthTwips);
                contentHeightTwips += EstimateParagraphAdvanceTwips(paragraph, paragraphContentHeightTwips);
            }

            return contentHeightTwips;
        }

        private static int EstimateTableCellContentWidthTwips(TableCellModel cell)
        {
            int maxContentWidth = 0;
            foreach (var paragraph in cell.Paragraphs)
            {
                int currentWidth = 0;
                foreach (var run in paragraph.Runs)
                {
                    int runFontSizeHalfPoints = run.Properties.FontSize.GetValueOrDefault(24);
                    int runFontSizeTwips = Math.Max(120, runFontSizeHalfPoints * 10);
                    double runStyleMultiplier = 1.0;
                    if (run.Properties.IsBold) runStyleMultiplier *= 1.1;
                    if (run.Properties.IsItalic) runStyleMultiplier *= 1.02;

                    bool isMonospace = string.Equals(run.Properties.FontName, "Courier New", StringComparison.OrdinalIgnoreCase) || 
                                       string.Equals(run.Properties.FontName, "Consolas", StringComparison.OrdinalIgnoreCase);

                    string text = run.Text ?? string.Empty;
                    foreach (char c in text)
                    {
                        double baseWidth = EstimateCharacterWidthTwips(c, runFontSizeTwips, isMonospace);
                        currentWidth += (int)Math.Round(baseWidth * runStyleMultiplier, MidpointRounding.AwayFromZero);
                    }

                    if (run.Image != null && run.Image.LayoutType == ImageLayoutType.Inline)
                    {
                        currentWidth += Math.Max(960, run.Image.Width * 15);
                    }
                }

                currentWidth += Math.Max(0, paragraph.Properties.LeftIndentTwips) + Math.Max(0, paragraph.Properties.RightIndentTwips);
                maxContentWidth = Math.Max(maxContentWidth, currentWidth);
            }

            return maxContentWidth;
        }

        private static int EstimateParagraphLineCount(ParagraphModel paragraph, int maxFontSizeHalfPoints, int availableWidthTwips)
        {
            if (availableWidthTwips <= 0) return 1;

            int currentLineWidthTwips = 0;
            int lineCount = 1;

            foreach (var run in paragraph.Runs)
            {
                int runFontSizeHalfPoints = run.Properties.FontSize.GetValueOrDefault(maxFontSizeHalfPoints);
                int runFontSizeTwips = Math.Max(120, runFontSizeHalfPoints * 10);
                double runStyleMultiplier = 1.0;
                if (run.Properties.IsBold) runStyleMultiplier *= 1.1;
                if (run.Properties.IsItalic) runStyleMultiplier *= 1.02;

                bool isMonospace = string.Equals(run.Properties.FontName, "Courier New", StringComparison.OrdinalIgnoreCase) || 
                                   string.Equals(run.Properties.FontName, "Consolas", StringComparison.OrdinalIgnoreCase);

                string text = run.Text ?? string.Empty;
                int i = 0;
                while (i < text.Length)
                {
                    char c = text[i];
                    if (c == '\n' || c == '\r')
                    {
                        lineCount++;
                        currentLineWidthTwips = 0;
                        if (c == '\r' && i + 1 < text.Length && text[i + 1] == '\n') i++; 
                        i++;
                        continue;
                    }

                    int wordWidthTwips = 0;
                    bool isWhitespace = char.IsWhiteSpace(c);

                    while (i < text.Length)
                    {
                        char wc = text[i];
                        if (wc == '\n' || wc == '\r') break;

                        bool currentIsWhitespace = char.IsWhiteSpace(wc);
                        if (currentIsWhitespace != isWhitespace && wordWidthTwips > 0) break;
                        
                        bool isCjk = (wc >= 0x4E00 && wc <= 0x9FFF) || (wc >= 0x3040 && wc <= 0x30FF) || (wc >= 0xAC00 && wc <= 0xD7AF);
                        if (isCjk)
                        {
                            if (wordWidthTwips > 0) break;
                            
                            double bw = EstimateCharacterWidthTwips(wc, runFontSizeTwips, isMonospace);
                            wordWidthTwips = (int)Math.Round(bw * runStyleMultiplier, MidpointRounding.AwayFromZero);
                            i++;
                            break;
                        }

                        double baseWidth = EstimateCharacterWidthTwips(wc, runFontSizeTwips, isMonospace);
                        wordWidthTwips += (int)Math.Round(baseWidth * runStyleMultiplier, MidpointRounding.AwayFromZero);
                        i++;
                    }

                    if (wordWidthTwips > 0)
                    {
                        if (currentLineWidthTwips + wordWidthTwips > availableWidthTwips)
                        {
                            if (wordWidthTwips > availableWidthTwips)
                            {
                                if (currentLineWidthTwips > 0)
                                {
                                    lineCount++;
                                }
                                lineCount += wordWidthTwips / availableWidthTwips;
                                currentLineWidthTwips = wordWidthTwips % availableWidthTwips;
                            }
                            else if (currentLineWidthTwips > 0 && !isWhitespace)
                            {
                                lineCount++;
                                currentLineWidthTwips = wordWidthTwips;
                            }
                            else
                            {
                                currentLineWidthTwips += wordWidthTwips;
                            }
                        }
                        else
                        {
                            currentLineWidthTwips += wordWidthTwips;
                        }
                    }
                }
            }

            return lineCount;
        }

        private static double EstimateCharacterWidthTwips(char character, int fontSizeTwips, bool isMonospace)
        {
            if (isMonospace)
            {
                return fontSizeTwips * 0.60d;
            }

            double widthFactor = character switch
            {
                'i' or 'l' or 'j' or '.' or ',' or ';' or ':' or '!' or '|' or '\'' or '`' => 0.25d,
                ' ' or '\t' or '"' or 't' or 'r' or 'I' or '[' or ']' or '(' or ')' or '{' or '}' => 0.33d,
                'f' or 's' or 'c' or 'z' or 'J' => 0.42d,
                'a' or 'b' or 'd' or 'e' or 'g' or 'h' or 'k' or 'n' or 'o' or 'p' or 'q' or 'u' or 'v' or 'x' or 'y' => 0.52d,
                'm' or 'w' => 0.78d,
                'M' or 'W' or 'O' or 'Q' or 'D' or 'N' => 0.82d,
                _ when char.IsDigit(character) => 0.55d,
                _ when char.IsUpper(character) => 0.65d,
                _ when character >= 0x4E00 && character <= 0x9FFF => 1.0d,
                _ when character >= 0x3040 && character <= 0x30FF => 1.0d,
                _ when character >= 0xAC00 && character <= 0xD7AF => 1.0d,
                _ when character >= 0x2E80 && character <= 0x2EFF => 1.0d,
                _ => 0.52d
            };

            return Math.Max(40d, fontSizeTwips * widthFactor);
        }

        private static int EstimateParagraphAdvanceTwips(ParagraphModel paragraph, int paragraphContentHeightTwips)
        {
            return paragraph.Properties.SpacingBeforeTwips + paragraphContentHeightTwips + paragraph.Properties.SpacingAfterTwips;
        }

        private static int BuildFloatingLayoutFlags(OfficeArtPictureDescriptor picture)
        {
            int flags = (int)picture.WrapType & 0x07;
            if (picture.BehindText)
            {
                flags |= 1 << 3;
            }

            if (picture.AllowOverlap)
            {
                flags |= 1 << 4;
            }

            flags |= EncodeRelativeTo(picture.VerticalRelativeTo) << 5;
            flags |= EncodeRelativeTo(picture.HorizontalRelativeTo) << 7;

            return flags;
        }

        private static int EncodeRelativeTo(string? relativeTo)
        {
            if (string.Equals(relativeTo, "margin", StringComparison.OrdinalIgnoreCase))
            {
                return 1;
            }

            if (string.Equals(relativeTo, "paragraph", StringComparison.OrdinalIgnoreCase))
            {
                return 2;
            }

            return 0;
        }

        private static int ResolveAlignedPositionTwips(
            string? alignment,
            string? relativeTo,
            int offsetTwips,
            int sizeTwips,
            int pageExtentTwips,
            int leadingMarginTwips,
            int trailingMarginTwips,
            bool isHorizontal,
            int paragraphStartTwips = 0,
            int paragraphExtentTwips = 0)
        {
            if (string.Equals(relativeTo, "paragraph", StringComparison.OrdinalIgnoreCase))
            {
                if (!isHorizontal && !string.IsNullOrWhiteSpace(alignment))
                {
                    return ResolveAlignmentWithinAnchor(paragraphStartTwips, paragraphExtentTwips, sizeTwips, alignment, isHorizontal);
                }

                if (!isHorizontal)
                {
                    return paragraphStartTwips + offsetTwips;
                }

                return offsetTwips;
            }

            if (!string.IsNullOrWhiteSpace(alignment))
            {
                return ResolveAlignmentPositionTwips(alignment, relativeTo, sizeTwips, pageExtentTwips, leadingMarginTwips, trailingMarginTwips, isHorizontal);
            }

            return offsetTwips;
        }

        private static int ResolveAlignmentWithinAnchor(int anchorStart, int anchorExtent, int sizeTwips, string alignment, bool isHorizontal)
        {
            int safeExtent = Math.Max(0, anchorExtent);
            int centerPosition = anchorStart + Math.Max(0, (safeExtent - sizeTwips) / 2);
            int endPosition = anchorStart + Math.Max(0, safeExtent - sizeTwips);
            string normalizedAlignment = alignment.Trim().ToLowerInvariant();

            return normalizedAlignment switch
            {
                "left" when isHorizontal => anchorStart,
                "inside" when isHorizontal => anchorStart,
                "right" when isHorizontal => endPosition,
                "outside" when isHorizontal => endPosition,
                "center" => centerPosition,
                "top" when !isHorizontal => anchorStart,
                "bottom" when !isHorizontal => endPosition,
                _ => anchorStart
            };
        }

        private static int ResolveAlignmentPositionTwips(
            string alignment,
            string? relativeTo,
            int sizeTwips,
            int pageExtentTwips,
            int leadingMarginTwips,
            int trailingMarginTwips,
            bool isHorizontal)
        {
            int anchorStart;
            int anchorExtent;
            if (string.Equals(relativeTo, "margin", StringComparison.OrdinalIgnoreCase))
            {
                anchorStart = leadingMarginTwips;
                anchorExtent = Math.Max(0, pageExtentTwips - leadingMarginTwips - trailingMarginTwips);
            }
            else
            {
                anchorStart = 0;
                anchorExtent = Math.Max(0, pageExtentTwips);
            }

            string normalizedAlignment = alignment.Trim().ToLowerInvariant();
            int centerPosition = anchorStart + Math.Max(0, (anchorExtent - sizeTwips) / 2);
            int endPosition = anchorStart + Math.Max(0, anchorExtent - sizeTwips);

            return normalizedAlignment switch
            {
                "left" when isHorizontal => anchorStart,
                "inside" when isHorizontal => anchorStart,
                "right" when isHorizontal => endPosition,
                "outside" when isHorizontal => endPosition,
                "center" => centerPosition,
                "top" when !isHorizontal => anchorStart,
                "bottom" when !isHorizontal => endPosition,
                _ => anchorStart
            };
        }

        private static bool TryGetImageDimensionsTwipsFromPayload(byte[]? imageData, string? contentType, out int widthTwips, out int heightTwips)
        {
            widthTwips = 1;
            heightTwips = 1;

            if (imageData == null || imageData.Length == 0)
            {
                return false;
            }

            return contentType switch
            {
                "image/x-wmf" => TryGetWmfDimensionsTwips(imageData, out widthTwips, out heightTwips),
                "image/x-emf" => TryGetEmfDimensionsTwips(imageData, out widthTwips, out heightTwips),
                _ => false
            };
        }

        private static bool TryGetWmfDimensionsTwips(byte[] imageData, out int widthTwips, out int heightTwips)
        {
            widthTwips = 1;
            heightTwips = 1;

            if (!HasWmfPlaceableHeader(imageData) || imageData.Length < 22)
            {
                return false;
            }

            int left = BitConverter.ToInt16(imageData, 6);
            int top = BitConverter.ToInt16(imageData, 8);
            int right = BitConverter.ToInt16(imageData, 10);
            int bottom = BitConverter.ToInt16(imageData, 12);
            int unitsPerInch = BitConverter.ToUInt16(imageData, 14);

            if (unitsPerInch <= 0 || right <= left || bottom <= top)
            {
                return false;
            }

            widthTwips = Math.Max(1, (int)Math.Round((right - left) * 1440d / unitsPerInch));
            heightTwips = Math.Max(1, (int)Math.Round((bottom - top) * 1440d / unitsPerInch));
            return true;
        }

        private static bool TryGetEmfDimensionsTwips(byte[] imageData, out int widthTwips, out int heightTwips)
        {
            widthTwips = 1;
            heightTwips = 1;

            if (imageData.Length < 40)
            {
                return false;
            }

            int frameLeft = BitConverter.ToInt32(imageData, 24);
            int frameTop = BitConverter.ToInt32(imageData, 28);
            int frameRight = BitConverter.ToInt32(imageData, 32);
            int frameBottom = BitConverter.ToInt32(imageData, 36);

            if (frameRight <= frameLeft || frameBottom <= frameTop)
            {
                return false;
            }

            widthTwips = Math.Max(1, (int)Math.Round((frameRight - frameLeft) * 72d / 127d));
            heightTwips = Math.Max(1, (int)Math.Round((frameBottom - frameTop) * 72d / 127d));
            return true;
        }

        private static int ConvertTwipsToMm100(int twips)
        {
            return (int)Math.Round(twips * 127d / 72d);
        }

        private static int ConvertTwipsToEmu(int twips)
        {
            return twips * 635;
        }

        private static short ClampToShort(int value)
        {
            return (short)Math.Max(short.MinValue, Math.Min(short.MaxValue, value));
        }

        private void WriteNumbering(BinaryWriter writer, DocumentModel model)
        {
            // STTB (String Table) header for LST
            writer.Write((ushort)0xFFFF); // fExtend
            writer.Write((ushort)model.AbstractNumbering.Count);
            writer.Write((ushort)0); // cbExtra

            foreach (var abs in model.AbstractNumbering)
            {
                // LSTF (28 bytes)
                writer.Write(abs.Id); // lsid
                writer.Write(0); // tplc
                for (int i = 0; i < 9; i++) writer.Write((short)0); // rgwchHtml
                writer.Write((byte)1); // grf (fSimpleList = 1?)
                writer.Write((byte)0); // unused

                // LVLs (9 levels usually required)
                for (int i = 0; i < 9; i++)
                {
                    var levelModel = abs.Levels.Find(l => l.Level == i) ?? new NumberingLevelModel { Level = i };
                    
                    // LVL structure
                    writer.Write(levelModel.Start); // iStartAt
                    
                    byte nfc = levelModel.NumberFormat switch
                    {
                        "decimal" => 0,
                        "upperRoman" => 1,
                        "lowerRoman" => 2,
                        "upperLetter" => 3,
                        "lowerLetter" => 4,
                        _ => 0
                    };
                    writer.Write(nfc);
                    writer.Write((byte)0); // jc (left)
                    writer.Write(new byte[9]); // rgbxchNums
                    writer.Write((byte)0); // ixchFollow (0=tab)
                    writer.Write((int)0); // dxvIndent
                    writer.Write((int)0); // dxvSpace
                    writer.Write((byte)0); // cbGrpprlChpx
                    writer.Write((byte)0); // cbGrpprlPapx
                    writer.Write((ushort)0); // reserved
                    
                    // xst (short string for level text)
                    string text = levelModel.LevelText.Replace("%" + (i + 1), "\x0001");
                    writer.Write((ushort)text.Length);
                    foreach (char c in text) writer.Write((short)c);
                }
            }
        }

        private void WriteLfo(BinaryWriter writer, DocumentModel model)
        {
            // PlfLfo structure
            writer.Write(model.NumberingInstances.Count); // lLfo

            foreach (var instance in model.NumberingInstances)
            {
                // LFO (16 bytes)
                writer.Write(instance.AbstractNumberId); // lsid
                writer.Write(0); // reserved1
                writer.Write(0); // reserved2
                writer.Write((byte)0); // clfolvl
                writer.Write((byte)0); // ibstFltcl
                writer.Write((ushort)0); // grf
            }

            // No LFOData levels for now (simplified)
        }

        private void WriteFontTable(BinaryWriter writer, List<FontModel> fonts)
        {
            // STTB (String Table) header for FFN
            writer.Write((ushort)0xFFFF); // fExtend - Unicode strings
            writer.Write((ushort)fonts.Count);
            writer.Write((ushort)0); // cbExtra (0 for FFN)

            foreach (var font in fonts)
            {
                // FFN (Font Family Name) structure
                // Build the font name as null-terminated Unicode string
                byte[] nameBytes = System.Text.Encoding.Unicode.GetBytes(font.Name + "\0");

                // Calculate total size: prq(1) + fTrueType(1) + ff(1) + wWeight(2) + chs(1) + ixchSz(1) + name
                // Actually: prq(1) + fTrueType(1 bit) + ff(4 bits) + wWeight(2) + chs(1) + ixchSz(1) + name
                byte cbFfn = (byte)(1 + 1 + 2 + 1 + 1 + nameBytes.Length);

                // prq (bits 0-1): Pitch
                // fTrueType (bit 2): TrueType flag
                // ff (bits 3-6): Font family
                byte prqAndFlags = (byte)(((byte)font.Pitch & 0x03) | (((byte)font.Family & 0x0F) << 3));

                writer.Write(cbFfn);
                writer.Write(prqAndFlags);
                writer.Write(font.Weight);
                writer.Write(font.Charset);
                writer.Write((byte)0); // ixchSz - index to extra string (0 = none)
                writer.Write(nameBytes);
            }
        }

        private void WriteStyleSheet(BinaryWriter writer, List<Nedev.FileConverters.DocxToDoc.Model.StyleModel> styles, List<FontModel> fonts)
        {
            // STSH structure (Style Sheet)
            // STSHI header (Style Sheet Information)
            writer.Write((ushort)0); // cbStshi (placeholder)
            long startPos = writer.BaseStream.Position;

            // cstd (count of styles) - Word expects at least 15 standard styles
            ushort cstd = (ushort)Math.Max(styles.Count, 15);
            writer.Write(cstd);
            writer.Write((ushort)0x0012); // cbStd (size of STD base - 18 bytes for Word 97-2003)

            // STSHI flags
            writer.Write((ushort)0); // stshi.fStdStylenamesWord97
            writer.Write((ushort)0); // stshi.ftcStandardChpStsh
            writer.Write((ushort)0); // stshi.wSpare
            writer.Write((ushort)0); // stshi.wSpare1
            writer.Write((uint)0);   // stshi.cstdBase
            writer.Write((ushort)0); // stshi.cstdNew
            writer.Write((ushort)0); // stshi.cstdCopy

            long endPos = writer.BaseStream.Position;
            writer.BaseStream.Seek(startPos - 2, SeekOrigin.Begin);
            writer.Write((ushort)(endPos - startPos)); // Actual cbStshi
            writer.BaseStream.Seek(endPos, SeekOrigin.Begin);

            // Write STDs (Style Descriptions)
            for (int i = 0; i < cstd; i++)
            {
                var style = styles.FirstOrDefault(s => s.StyleId == i);

                if (style == null)
                {
                    // Empty slot
                    writer.Write((ushort)0); // cb (0 = empty slot)
                    continue;
                }

                // Calculate STD size
                byte[] nameBytes = System.Text.Encoding.Unicode.GetBytes(style.Name + "\0");
                int cbStd = 10 + nameBytes.Length; // Base (10) + name

                // Add PAPX if present
                byte[]? papxData = null;
                if (style.ParagraphProps != null)
                {
                    papxData = BuildPapxFromStyle(style.ParagraphProps);
                    cbStd += 1 + papxData.Length; // cbGrpprlPapx + data
                }

                // Add CHPX if present
                byte[]? chpxData = null;
                if (style.CharacterProps != null)
                {
                    chpxData = BuildChpxFromStyle(style.CharacterProps, fonts);
                    cbStd += 1 + chpxData.Length; // cbGrpprlChpx + data
                }

                writer.Write((ushort)cbStd);

                // STD base (10 bytes)
                writer.Write((byte)(style.IsParagraphStyle ? 1 : 2)); // sgc (style type)
                writer.Write((byte)style.StyleId); // istdBase (parent style)
                writer.Write((ushort)(style.NextStyle ?? style.StyleId)); // istdNext
                writer.Write((ushort)0); // bchUpe - offset to UPX
                writer.Write((ushort)0); // fHasUpe, fScratch, fHidden, etc.
                writer.Write((byte)nameBytes.Length); // stzName length
                writer.Write(nameBytes);

                // UPX (formatting)
                if (papxData != null)
                {
                    writer.Write((byte)papxData.Length);
                    writer.Write(papxData);
                }
                else
                {
                    writer.Write((byte)0); // No PAPX
                }

                if (chpxData != null)
                {
                    writer.Write((byte)chpxData.Length);
                    writer.Write(chpxData);
                }
                else
                {
                    writer.Write((byte)0); // No CHPX
                }
            }
        }

        private byte[] BuildPapxFromStyle(ParagraphModel.ParagraphProperties props)
        {
            var sprms = new List<byte>();

            if (props.Alignment != ParagraphModel.Justification.Left)
            {
                sprms.Add(0x03); sprms.Add(0x24); sprms.Add((byte)props.Alignment);
            }

            AppendParagraphFormattingSprms(sprms, props);

            return sprms.ToArray();
        }

        private static void AppendParagraphFormattingSprms(List<byte> sprms, ParagraphModel.ParagraphProperties props)
        {
            AppendParagraphIndentSprms(sprms, props);
            AppendParagraphSpacingSprms(sprms, props);
        }

        private static void AppendParagraphIndentSprms(List<byte> sprms, ParagraphModel.ParagraphProperties props)
        {
            if (props.RightIndentTwips != 0)
            {
                sprms.Add(0x0E);
                sprms.Add(0x84);
                short rightIndent = ClampToShort(props.RightIndentTwips);
                sprms.Add((byte)(rightIndent & 0xFF));
                sprms.Add((byte)((rightIndent >> 8) & 0xFF));
            }

            if (props.LeftIndentTwips != 0)
            {
                sprms.Add(0x0F);
                sprms.Add(0x84);
                short leftIndent = ClampToShort(props.LeftIndentTwips);
                sprms.Add((byte)(leftIndent & 0xFF));
                sprms.Add((byte)((leftIndent >> 8) & 0xFF));
            }

            if (props.FirstLineIndentTwips != 0)
            {
                sprms.Add(0x11);
                sprms.Add(0x84);
                short firstLineIndent = ClampToShort(props.FirstLineIndentTwips);
                sprms.Add((byte)(firstLineIndent & 0xFF));
                sprms.Add((byte)((firstLineIndent >> 8) & 0xFF));
            }
        }

        private static void AppendParagraphSpacingSprms(List<byte> sprms, ParagraphModel.ParagraphProperties props)
        {
            if (props.SpacingBeforeTwips > 0)
            {
                sprms.Add(0x22);
                sprms.Add(0x26);
                sprms.Add((byte)(props.SpacingBeforeTwips & 0xFF));
                sprms.Add((byte)((props.SpacingBeforeTwips >> 8) & 0xFF));
            }

            if (props.SpacingAfterTwips > 0)
            {
                sprms.Add(0x23);
                sprms.Add(0x26);
                sprms.Add((byte)(props.SpacingAfterTwips & 0xFF));
                sprms.Add((byte)((props.SpacingAfterTwips >> 8) & 0xFF));
            }

            if (props.LineSpacing.HasValue)
            {
                sprms.Add(0x24);
                sprms.Add(0x26);
                sprms.Add((byte)(props.LineSpacing.Value & 0xFF));
                sprms.Add((byte)((props.LineSpacing.Value >> 8) & 0xFF));
            }
        }

        private byte[] BuildChpxFromStyle(RunModel.CharacterProperties props, List<FontModel> fonts)
        {
            return BuildRunSprms(props, fonts);
        }

        private byte[] BuildRunSprms(RunModel.CharacterProperties props, List<FontModel> fonts)
        {
            var sprms = new List<byte>();

            if (props.IsBold) { sprms.Add(0x35); sprms.Add(0x08); sprms.Add(1); }
            if (props.IsItalic) { sprms.Add(0x36); sprms.Add(0x08); sprms.Add(1); }
            if (props.IsStrike) { sprms.Add(0x37); sprms.Add(0x08); sprms.Add(1); }
            if (props.FontSize.HasValue)
            {
                sprms.Add(0x43); sprms.Add(0x4A);
                sprms.Add(BitConverter.GetBytes((short)props.FontSize.Value)[0]);
                sprms.Add(BitConverter.GetBytes((short)props.FontSize.Value)[1]);
            }

            if (props.Underline != UnderlineType.None)
            {
                sprms.Add(0x3E); sprms.Add(0x2A);
                sprms.Add(props.Underline switch
                {
                    UnderlineType.Single => 1,
                    UnderlineType.Double => 3,
                    UnderlineType.Dotted => 4,
                    UnderlineType.Thick => 6,
                    UnderlineType.Dashed => 7,
                    UnderlineType.Wave => 11,
                    _ => 0
                });
            }

            if (!string.IsNullOrEmpty(props.FontName))
            {
                int fontIndex = fonts.FindIndex(f => string.Equals(f.Name, props.FontName, StringComparison.OrdinalIgnoreCase));
                if (fontIndex >= 0)
                {
                    sprms.Add(0x4F); sprms.Add(0x4A); // sprmCRgFtc0
                    sprms.Add((byte)(fontIndex & 0xFF));
                    sprms.Add((byte)((fontIndex >> 8) & 0xFF));
                }
            }

            return sprms.ToArray();
        }
    }
}
