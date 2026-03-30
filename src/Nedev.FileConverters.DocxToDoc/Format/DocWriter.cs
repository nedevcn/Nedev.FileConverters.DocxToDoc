using System;
using System.IO;
using System.IO.Compression;
using System.Text;
using System.Collections.Generic;
using System.Globalization;
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
            var headerOfficeArtBlips = new List<(int cp, byte[] data, string contentType, int widthTwips, int heightTwips, int leftTwips, int topTwips, ImageWrapType wrapType, bool behindText, bool allowOverlap, string? horizontalRelativeTo, string? verticalRelativeTo)>();
            var fieldEntries = new List<(int cp, ushort descriptor)>();
            var headerFieldEntries = new List<(int cp, ushort descriptor)>();
            int nextPictureOffset = 0;

            // 1. Build the text buffer and formatting structures in one pass
            var textBuilder = new StringBuilder();
            var chpxWriter = new ChpxFkpWriter();
            var papxWriter = new PapxFkpWriter();
            var tapxWriter = new TapxFkpWriter();
            var sections = model.Sections.Count > 0 ? model.Sections : new List<SectionModel> { new SectionModel() };
            var layoutSection = sections[0];
            int documentAvailableWidthTwips = Math.Max(1440, layoutSection.PageWidth - layoutSection.MarginLeft - layoutSection.MarginRight);
            int paragraphVerticalCursorTwips = layoutSection.MarginTop;
            
            int currentCp = 0;
            int visibleCp = 0;
            var tableWriter = new BinaryWriter(tableStream);
            var supportedFootnotes = new List<(FootnoteModel footnote, int index)>();
            for (int footnoteIndex = 0; footnoteIndex < model.Footnotes.Count; footnoteIndex++)
            {
                var footnote = model.Footnotes[footnoteIndex];
                if (!CanWriteFootnote(footnote))
                {
                    continue;
                }

                supportedFootnotes.Add((footnote, footnoteIndex));
            }

            supportedFootnotes.Sort(static (left, right) =>
            {
                int byReference = left.footnote.ReferenceCp.CompareTo(right.footnote.ReferenceCp);
                if (byReference != 0)
                {
                    return byReference;
                }

                return left.index.CompareTo(right.index);
            });

            var supportedEndnotes = new List<(EndnoteModel endnote, int index)>();
            for (int endnoteIndex = 0; endnoteIndex < model.Endnotes.Count; endnoteIndex++)
            {
                var endnote = model.Endnotes[endnoteIndex];
                if (!CanWriteEndnote(endnote))
                {
                    continue;
                }

                supportedEndnotes.Add((endnote, endnoteIndex));
            }

            supportedEndnotes.Sort(static (left, right) =>
            {
                int byReference = left.endnote.ReferenceCp.CompareTo(right.endnote.ReferenceCp);
                if (byReference != 0)
                {
                    return byReference;
                }

                return left.index.CompareTo(right.index);
            });

            var supportedComments = BuildSupportedComments(model.Comments);

            supportedComments.Sort(static (left, right) =>
            {
                int byEnd = left.comment.EndCp.CompareTo(right.comment.EndCp);
                if (byEnd != 0)
                {
                    return byEnd;
                }

                int byStart = left.comment.StartCp.CompareTo(right.comment.StartCp);
                if (byStart != 0)
                {
                    return byStart;
                }

                return left.index.CompareTo(right.index);
            });

            int nextFootnoteIndex = 0;
            var emittedFootnotes = new List<(FootnoteModel footnote, int referenceCp)>(supportedFootnotes.Count);
            int nextEndnoteIndex = 0;
            var emittedEndnotes = new List<(EndnoteModel endnote, int referenceCp)>(supportedEndnotes.Count);
            int nextCommentIndex = 0;
            var emittedComments = new List<(CommentModel comment, int referenceCp, string storyText)>(supportedComments.Count);

            void EmitFootnoteReference(FootnoteModel footnote)
            {
                emittedFootnotes.Add((footnote, currentCp));
                AppendNoteReferenceMarker(textBuilder, chpxWriter, ref currentCp, footnote.CustomMarkText);
            }

            void EmitEndnoteReference(EndnoteModel endnote)
            {
                emittedEndnotes.Add((endnote, currentCp));
                AppendNoteReferenceMarker(textBuilder, chpxWriter, ref currentCp, endnote.CustomMarkText);
            }

            void AppendPendingFootnoteReferencesAtVisibleCp()
            {
                while (nextFootnoteIndex < supportedFootnotes.Count &&
                       supportedFootnotes[nextFootnoteIndex].footnote.ReferenceCp == visibleCp)
                {
                    EmitFootnoteReference(supportedFootnotes[nextFootnoteIndex].footnote);
                    nextFootnoteIndex++;
                }
            }

            void AppendPendingEndnoteReferencesAtVisibleCp()
            {
                while (nextEndnoteIndex < supportedEndnotes.Count &&
                       supportedEndnotes[nextEndnoteIndex].endnote.ReferenceCp == visibleCp)
                {
                    EmitEndnoteReference(supportedEndnotes[nextEndnoteIndex].endnote);
                    nextEndnoteIndex++;
                }
            }

            void AppendRemainingFootnoteReferencesAtDocumentEnd()
            {
                while (nextFootnoteIndex < supportedFootnotes.Count)
                {
                    EmitFootnoteReference(supportedFootnotes[nextFootnoteIndex].footnote);
                    nextFootnoteIndex++;
                }
            }

            void AppendRemainingEndnoteReferencesAtDocumentEnd()
            {
                while (nextEndnoteIndex < supportedEndnotes.Count)
                {
                    EmitEndnoteReference(supportedEndnotes[nextEndnoteIndex].endnote);
                    nextEndnoteIndex++;
                }
            }

            void EmitCommentReference(CommentModel comment)
            {
                textBuilder.Append('\x0005');
                chpxWriter.AddRun(currentCp, currentCp + 1, BuildSpecialCharacterSprms());
                string storyText = supportedComments[nextCommentIndex].storyText;
                emittedComments.Add((comment, currentCp, storyText));
                currentCp += 1;
            }

            void AppendPendingCommentReferencesAtVisibleCp()
            {
                while (nextCommentIndex < supportedComments.Count &&
                       supportedComments[nextCommentIndex].comment.EndCp == visibleCp)
                {
                    EmitCommentReference(supportedComments[nextCommentIndex].comment);
                    nextCommentIndex++;
                }
            }

            void AppendRemainingCommentReferencesAtDocumentEnd()
            {
                while (nextCommentIndex < supportedComments.Count)
                {
                    EmitCommentReference(supportedComments[nextCommentIndex].comment);
                    nextCommentIndex++;
                }
            }

            void AppendNonVisibleText(string text)
            {
                if (string.IsNullOrEmpty(text))
                {
                    return;
                }

                textBuilder.Append(text);
                currentCp += text.Length;
            }

            void AppendVisibleText(string text, byte[]? runSprms = null)
            {
                if (string.IsNullOrEmpty(text))
                {
                    return;
                }

                int offset = 0;
                while (offset < text.Length)
                {
                    AppendPendingFootnoteReferencesAtVisibleCp();
                    AppendPendingEndnoteReferencesAtVisibleCp();
                    AppendPendingCommentReferencesAtVisibleCp();

                    int nextFootnoteAnchorVisibleCp = nextFootnoteIndex < supportedFootnotes.Count
                        ? supportedFootnotes[nextFootnoteIndex].footnote.ReferenceCp
                        : int.MaxValue;
                    int nextEndnoteAnchorVisibleCp = nextEndnoteIndex < supportedEndnotes.Count
                        ? supportedEndnotes[nextEndnoteIndex].endnote.ReferenceCp
                        : int.MaxValue;
                    int nextCommentAnchorVisibleCp = nextCommentIndex < supportedComments.Count
                        ? supportedComments[nextCommentIndex].comment.EndCp
                        : int.MaxValue;
                    int nextAnchorVisibleCp = Math.Min(Math.Min(nextFootnoteAnchorVisibleCp, nextEndnoteAnchorVisibleCp), nextCommentAnchorVisibleCp);
                    int segmentLength = Math.Min(text.Length - offset, Math.Max(0, nextAnchorVisibleCp - visibleCp));
                    if (segmentLength == 0)
                    {
                        segmentLength = text.Length - offset;
                    }

                    if (runSprms != null && runSprms.Length > 0)
                    {
                        chpxWriter.AddRun(currentCp, currentCp + segmentLength, runSprms);
                    }

                    textBuilder.Append(text, offset, segmentLength);
                    offset += segmentLength;
                    currentCp += segmentLength;
                    visibleCp += segmentLength;
                }
            }

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

                    if (run.Hyperlink != null)
                    {
                        int hyperlinkStart = runIndex;
                        var hyperlink = run.Hyperlink;
                        while (runIndex + 1 < para.Runs.Count && ReferenceEquals(para.Runs[runIndex + 1].Hyperlink, hyperlink))
                        {
                            runIndex++;
                        }

                        AppendFieldCharacter('\x0013');
                        AppendNonVisibleText(BuildHyperlinkInstructionCore(hyperlink));
                        AppendFieldCharacter('\x0014');

                        for (int hyperlinkRunIndex = hyperlinkStart; hyperlinkRunIndex <= runIndex; hyperlinkRunIndex++)
                        {
                            AppendVisibleRunContent(para.Runs[hyperlinkRunIndex]);
                        }

                        AppendFieldCharacter('\x0015');
                        continue;
                    }

                    if (run.Image != null && run.Image.Data != null)
                    {
                        AppendImageRunContent(run);
                        continue;
                    }

                    if (run.IsFieldBegin)
                    {
                        if (run.Field != null && !HasExplicitFieldBoundary(para.Runs, runIndex + 1, run.Field))
                        {
                            int fieldDepth = openFields.Count;
                            AppendFieldCharacter('\x0013', run.Field, fieldDepth);

                            string instruction = ResolveFieldInstruction(run.Field);
                            AppendNonVisibleText(instruction);

                            if (!string.IsNullOrEmpty(run.Field.Result))
                            {
                                AppendFieldCharacter('\x0014', run.Field, fieldDepth);
                                AppendVisibleText(run.Field.Result);
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
                        AppendNonVisibleText(run.Field.Instruction);
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
                            AppendVisibleText(run.Field.Result);
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
                        AppendVisibleText(openField.Result);
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

                AppendPendingFootnoteReferencesAtVisibleCp();
                AppendPendingEndnoteReferencesAtVisibleCp();
                AppendPendingCommentReferencesAtVisibleCp();
                textBuilder.Append('\r');
                papxWriter.AddParagraph(paraStart, currentCp + 1, paraSprms.ToArray());
                currentCp += 1;
                visibleCp += 1;
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

                void AppendFormattedRunText(RunModel runModel)
                {
                    if (runModel.Text.Length == 0)
                    {
                        return;
                    }

                    byte[] runSprms = BuildRunSprms(runModel.Properties, model.Fonts);
                    if (runSprms.Length > 0)
                    {
                        AppendVisibleText(runModel.Text, runSprms);
                    }
                    else
                    {
                        AppendVisibleText(runModel.Text);
                    }
                }

                void AppendImageRunContent(RunModel runModel)
                {
                    if (runModel.Image == null || runModel.Image.Data == null)
                    {
                        return;
                    }

                    string imageContentType = ResolveImageContentType(runModel.Image.ContentType, runModel.Image.Data);
                    (int imageWidthTwips, int imageHeightTwips) = ResolveImageDimensionsTwips(runModel.Image, imageContentType);

                    // In MS-DOC, embedded objects use a single placeholder character in the text stream.
                    textBuilder.Append('\x0001');

                    byte[] pictureBlock = BuildPictureBlock(runModel.Image, imageContentType);
                    int pictureOffset = nextPictureOffset;
                    nextPictureOffset += pictureBlock.Length;
                    embeddedObjects.Add(pictureBlock);

                    if (SupportsOfficeArtBlip(imageContentType))
                    {
                        (int leftTwips, int topTwips, _, _) = ResolveImageBoundsTwips(runModel.Image, imageContentType, layoutSection, paragraphTopTwips, paragraphContentHeightTwips);
                        officeArtBlips.Add((currentCp, runModel.Image.Data, imageContentType, imageWidthTwips, imageHeightTwips, leftTwips, topTwips, runModel.Image.WrapType, runModel.Image.BehindText, runModel.Image.AllowOverlap, runModel.Image.HorizontalRelativeTo, runModel.Image.VerticalRelativeTo));
                    }

                    chpxWriter.AddRun(currentCp, currentCp + 1, BuildImageSprms(pictureOffset));
                    currentCp += 1;
                }

                void AppendVisibleRunContent(RunModel runModel)
                {
                    AppendImageRunContent(runModel);
                    AppendFormattedRunText(runModel);
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
                        FieldType.SectionPages => "SECTIONPAGES",
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

            void ProcessTable(TableModel table, ref int verticalCursorTwips, int availableWidthTwips)
            {
                for (int rowIndex = 0; rowIndex < table.Rows.Count; rowIndex++)
                {
                    var row = table.Rows[rowIndex];
                    var previousRow = rowIndex > 0 ? table.Rows[rowIndex - 1] : null;
                    int rowHeightTwips = 0;
                    int rowGridColumnIndex = 0;
                    int totalColumnCount = ResolveTableTotalColumnCount(table, row);
                    int tableAvailableWidthTwips = ResolveTableAvailableWidthTwips(table, row, availableWidthTwips, totalColumnCount);
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
                        foreach (var cellBlock in EnumerateTableCellBlocks(cellLayout.cell))
                        {
                            if (cellBlock is ParagraphModel cellParagraph)
                            {
                                ProcessParagraph(cellParagraph, ref cellVerticalCursorTwips, cellLayout.availableWidthTwips, row.HeightRule == TableRowHeightRule.Exact ? maxVisibleCursorTwips : null);
                            }
                            else if (cellBlock is TableModel nestedTable)
                            {
                                ProcessTable(nestedTable, ref cellVerticalCursorTwips, cellLayout.availableWidthTwips);
                                if (row.HeightRule == TableRowHeightRule.Exact)
                                {
                                    cellVerticalCursorTwips = Math.Min(cellVerticalCursorTwips, maxVisibleCursorTwips);
                                }
                            }
                        }

                        int cellMarkStart = currentCp;
                        textBuilder.Append('\x0007');
                        currentCp += 1;

                        List<byte> cellMarkSprms = new List<byte>();
                        cellMarkSprms.Add(0x16); cellMarkSprms.Add(0x24); cellMarkSprms.Add(1);
                        papxWriter.AddParagraph(cellMarkStart, currentCp, cellMarkSprms.ToArray());
                    }

                    int rowMarkStart = currentCp;
                    textBuilder.Append('\r');

                    List<byte> rowParaSprms = new List<byte>();
                    rowParaSprms.Add(0x16); rowParaSprms.Add(0x24); rowParaSprms.Add(1);
                    rowParaSprms.Add(0x17); rowParaSprms.Add(0x24); rowParaSprms.Add(1);

                    papxWriter.AddParagraph(rowMarkStart, currentCp + 1, rowParaSprms.ToArray());
                    currentCp += 1;

                    List<byte> tapSprms = new List<byte>();
                    tapSprms.Add(0x08); tapSprms.Add(0xD6);
                    byte[] defTable = new byte[10] { 0x08, (byte)row.Cells.Count, 0, 0, 0, 0, 0, 0, 0, 0 };
                    tapSprms.AddRange(defTable);

                    tapxWriter.AddRow(rowMarkStart, currentCp, tapSprms.ToArray());
                    verticalCursorTwips += rowHeightTwips;
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
                    ProcessTable(table, ref paragraphVerticalCursorTwips, documentAvailableWidthTwips);
                }
            }

            void AppendStructuredHeaderFooterStory(HeaderFooterStoryModel story, SectionModel? storySection, ref int storyLength, List<int> storyStarts)
            {
                storyStarts.Add(storyLength);
                IReadOnlyList<object> storyBlocks = story.Content.Count > 0
                    ? story.Content
                    : story.Paragraphs.ConvertAll(static paragraph => (object)paragraph);
                int storyAvailableWidthTwips = storySection != null
                    ? Math.Max(1440, storySection.PageWidth - storySection.MarginLeft - storySection.MarginRight)
                    : 9360;
                int storyVerticalCursorTwips = storySection?.MarginTop ?? 0;

                if (storyBlocks.Count == 0)
                {
                    AppendHeaderParagraphMark(ref storyLength);
                    AppendHeaderParagraphMark(ref storyLength);
                    return;
                }

                foreach (var block in storyBlocks)
                {
                    if (block is ParagraphModel paragraph)
                    {
                        int paragraphAvailableWidthTwips = ResolveParagraphAvailableWidthTwips(paragraph, storyAvailableWidthTwips);
                        int paragraphContentHeightTwips = EstimateParagraphContentHeightTwips(paragraph, paragraphAvailableWidthTwips);
                        int paragraphTopTwips = storyVerticalCursorTwips + paragraph.Properties.SpacingBeforeTwips;
                        AppendStructuredHeaderFooterParagraph(paragraph, storySection, paragraphTopTwips, paragraphContentHeightTwips, ref storyLength, inTable: false);
                        storyVerticalCursorTwips += EstimateParagraphAdvanceTwips(paragraph, paragraphContentHeightTwips);
                    }
                    else if (block is TableModel table)
                    {
                        AppendStructuredHeaderFooterTable(table, storySection, storyAvailableWidthTwips, ref storyVerticalCursorTwips, ref storyLength);
                    }
                }

                AppendHeaderParagraphMark(ref storyLength);

                void AppendStructuredHeaderFooterTable(TableModel table, SectionModel? layoutSectionForStory, int tableAvailableWidthTwips, ref int verticalCursorTwips, ref int localStoryLength)
                {
                    foreach (var row in table.Rows)
                    {
                        int rowHeightTwips = 0;
                        int gridColumnIndex = 0;

                        foreach (var cell in row.Cells)
                        {
                            int gridSpan = Math.Max(1, cell.GridSpan);
                            int cellWidthTwips = cell.Width > 0
                                ? cell.Width
                                : ResolveTableCellWidth(table, gridColumnIndex, gridSpan, tableAvailableWidthTwips);
                            int horizontalCellPaddingTwips = ResolveTableCellHorizontalPaddingTwips(table, cell);
                            int cellAvailableWidthTwips = Math.Max(720, cellWidthTwips - horizontalCellPaddingTwips);
                            rowHeightTwips = Math.Max(rowHeightTwips, EstimateTableCellContentHeightTwips(cell, cellAvailableWidthTwips));
                            foreach (var cellBlock in EnumerateTableCellBlocks(cell))
                            {
                                if (cellBlock is ParagraphModel paragraph)
                                {
                                    int paragraphAvailableWidthTwips = ResolveParagraphAvailableWidthTwips(paragraph, cellAvailableWidthTwips);
                                    int paragraphContentHeightTwips = EstimateParagraphContentHeightTwips(paragraph, paragraphAvailableWidthTwips);
                                    int paragraphTopTwips = verticalCursorTwips + paragraph.Properties.SpacingBeforeTwips;
                                    AppendStructuredHeaderFooterParagraph(paragraph, layoutSectionForStory, paragraphTopTwips, paragraphContentHeightTwips, ref localStoryLength, inTable: true);
                                }
                                else if (cellBlock is TableModel nestedTable)
                                {
                                    AppendStructuredHeaderFooterTable(nestedTable, layoutSectionForStory, cellAvailableWidthTwips, ref verticalCursorTwips, ref localStoryLength);
                                }
                            }

                            int cellMarkStart = currentCp;
                            textBuilder.Append('\x0007');
                            currentCp += 1;
                            localStoryLength += 1;

                            List<byte> cellMarkSprms = new List<byte>();
                            cellMarkSprms.Add(0x16);
                            cellMarkSprms.Add(0x24);
                            cellMarkSprms.Add(1);
                            papxWriter.AddParagraph(cellMarkStart, currentCp, cellMarkSprms.ToArray());
                            gridColumnIndex += gridSpan;
                        }

                        int rowMarkStart = currentCp;
                        textBuilder.Append('\r');
                        localStoryLength += 1;

                        List<byte> rowParaSprms = new List<byte>();
                        rowParaSprms.Add(0x16); rowParaSprms.Add(0x24); rowParaSprms.Add(1);
                        rowParaSprms.Add(0x17); rowParaSprms.Add(0x24); rowParaSprms.Add(1);

                        papxWriter.AddParagraph(rowMarkStart, currentCp + 1, rowParaSprms.ToArray());
                        currentCp += 1;

                        List<byte> tapSprms = new List<byte>();
                        tapSprms.Add(0x08); tapSprms.Add(0xD6);
                        byte[] defTable = new byte[10] { 0x08, (byte)row.Cells.Count, 0, 0, 0, 0, 0, 0, 0, 0 };
                        tapSprms.AddRange(defTable);

                        tapxWriter.AddRow(rowMarkStart, currentCp, tapSprms.ToArray());
                        verticalCursorTwips += Math.Max(276, rowHeightTwips);
                    }
                }

                void AppendStructuredHeaderFooterParagraph(ParagraphModel paragraph, SectionModel? layoutSectionForStory, int paragraphTopTwips, int paragraphContentHeightTwips, ref int localStoryLength, bool inTable)
                {
                    int paragraphStart = currentCp;
                    var autoCompletedFields = new HashSet<FieldModel>();
                    var separatedFields = new HashSet<FieldModel>();
                    var openFields = new List<FieldModel>();

                    for (int runIndex = 0; runIndex < paragraph.Runs.Count; runIndex++)
                    {
                        var run = paragraph.Runs[runIndex];
                        if (run.Hyperlink != null)
                        {
                            int hyperlinkStart = runIndex;
                            var hyperlink = run.Hyperlink;
                            while (runIndex + 1 < paragraph.Runs.Count && ReferenceEquals(paragraph.Runs[runIndex + 1].Hyperlink, hyperlink))
                            {
                                runIndex++;
                            }

                            AppendHeaderFieldCharacter('\x0013', ref localStoryLength);
                            AppendHeaderNonVisibleText(BuildHyperlinkInstructionCore(hyperlink));
                            AppendHeaderFieldCharacter('\x0014', ref localStoryLength);

                            for (int hyperlinkRunIndex = hyperlinkStart; hyperlinkRunIndex <= runIndex; hyperlinkRunIndex++)
                            {
                                AppendHeaderVisibleRunContent(paragraph.Runs[hyperlinkRunIndex], layoutSectionForStory, paragraphTopTwips, paragraphContentHeightTwips, ref localStoryLength);
                            }

                            AppendHeaderFieldCharacter('\x0015', ref localStoryLength);
                            continue;
                        }

                        if (run.IsFieldBegin &&
                            run.Field != null &&
                            !HasExplicitFieldBoundaryCore(paragraph.Runs, runIndex + 1, run.Field))
                        {
                            int fieldDepth = openFields.Count;
                            AppendHeaderFieldCharacter('\x0013', ref localStoryLength, run.Field, fieldDepth);

                            string instruction = ResolveFieldInstructionCore(run.Field);
                            AppendHeaderNonVisibleText(instruction);

                            if (!string.IsNullOrEmpty(run.Field.Result))
                            {
                                AppendHeaderFieldCharacter('\x0014', ref localStoryLength, run.Field, fieldDepth);
                                separatedFields.Add(run.Field);
                                AppendHeaderVisibleText(run.Field.Result, ref localStoryLength);
                            }

                            AppendHeaderFieldCharacter('\x0015', ref localStoryLength, run.Field, fieldDepth);
                            autoCompletedFields.Add(run.Field);
                            continue;
                        }

                        if (run.IsFieldBegin)
                        {
                            AppendHeaderFieldCharacter('\x0013', ref localStoryLength, run.Field, openFields.Count);
                            if (run.Field != null)
                            {
                                openFields.Add(run.Field);
                            }
                            continue;
                        }

                        if (run.Field != null &&
                            !run.IsFieldSeparate &&
                            !run.IsFieldEnd &&
                            run.Text.Length == 0 &&
                            !string.IsNullOrEmpty(run.Field.Instruction))
                        {
                            AppendHeaderNonVisibleText(run.Field.Instruction);
                            continue;
                        }

                        if (run.IsFieldSeparate)
                        {
                            AppendHeaderFieldCharacter('\x0014', ref localStoryLength, run.Field, GetFieldDepthCore(openFields, run.Field));
                            if (run.Field != null)
                            {
                                separatedFields.Add(run.Field);
                            }
                            continue;
                        }

                        if (run.IsFieldEnd)
                        {
                            int fieldDepth = GetFieldDepthCore(openFields, run.Field);
                            if (run.Field != null &&
                                !autoCompletedFields.Contains(run.Field) &&
                                !separatedFields.Contains(run.Field) &&
                                !string.IsNullOrEmpty(run.Field.Result))
                            {
                                AppendHeaderFieldCharacter('\x0014', ref localStoryLength, run.Field, fieldDepth);
                                separatedFields.Add(run.Field);
                                AppendHeaderVisibleText(run.Field.Result, ref localStoryLength);
                            }

                            AppendHeaderFieldCharacter('\x0015', ref localStoryLength, run.Field, fieldDepth);
                            if (run.Field != null)
                            {
                                RemoveLastOpenFieldCore(openFields, run.Field);
                            }
                            continue;
                        }

                        AppendHeaderVisibleRunContent(run, layoutSectionForStory, paragraphTopTwips, paragraphContentHeightTwips, ref localStoryLength);
                    }

                    for (int index = openFields.Count - 1; index >= 0; index--)
                    {
                        var openField = openFields[index];
                        int fieldDepth = index;
                        if (!separatedFields.Contains(openField) && !string.IsNullOrEmpty(openField.Result))
                        {
                            AppendHeaderFieldCharacter('\x0014', ref localStoryLength, openField, fieldDepth);
                            AppendHeaderVisibleText(openField.Result, ref localStoryLength);
                        }

                        AppendHeaderFieldCharacter('\x0015', ref localStoryLength, openField, fieldDepth);
                    }

                    AppendHeaderVisibleText("\r", ref localStoryLength);

                    List<byte> paragraphSprms = new List<byte>();
                    AppendParagraphFormattingSprms(paragraphSprms, paragraph.Properties);
                    if (inTable)
                    {
                        paragraphSprms.Add(0x16);
                        paragraphSprms.Add(0x24);
                        paragraphSprms.Add(1);
                    }

                    papxWriter.AddParagraph(paragraphStart, currentCp, paragraphSprms.ToArray());
                }

                void AppendHeaderParagraphMark(ref int localStoryLength)
                {
                    int paragraphStart = currentCp;
                    textBuilder.Append('\r');
                    currentCp += 1;
                    localStoryLength += 1;
                    papxWriter.AddParagraph(paragraphStart, currentCp, Array.Empty<byte>());
                }

                void AppendHeaderNonVisibleText(string text)
                {
                    if (string.IsNullOrEmpty(text))
                    {
                        return;
                    }

                    textBuilder.Append(text);
                }

                void AppendHeaderVisibleText(string text, ref int localStoryLength, byte[]? runSprms = null)
                {
                    if (string.IsNullOrEmpty(text))
                    {
                        return;
                    }

                    int startCp = currentCp;
                    textBuilder.Append(text);
                    currentCp += text.Length;
                    localStoryLength += text.Length;

                    if (runSprms != null && runSprms.Length > 0)
                    {
                        chpxWriter.AddRun(startCp, currentCp, runSprms);
                    }
                }

                void AppendHeaderFieldCharacter(char marker, ref int localStoryLength, FieldModel? fieldModel = null, int nestingDepth = 0)
                {
                    textBuilder.Append(marker);
                    headerFieldEntries.Add((localStoryLength, BuildFieldDescriptorCore(marker, fieldModel, nestingDepth)));
                    chpxWriter.AddRun(currentCp, currentCp + 1, BuildSpecialCharacterSprms());
                    currentCp += 1;
                    localStoryLength += 1;
                }

                void AppendHeaderVisibleRunContent(RunModel runModel, SectionModel? layoutSectionForStory, int paragraphTopTwips, int paragraphContentHeightTwips, ref int localStoryLength)
                {
                    AppendHeaderImageRunContent(runModel, layoutSectionForStory, paragraphTopTwips, paragraphContentHeightTwips, ref localStoryLength);

                    if (runModel.Text.Length == 0)
                    {
                        return;
                    }

                    byte[] runSprms = BuildRunSprms(runModel.Properties, model.Fonts);
                    AppendHeaderVisibleText(runModel.Text, ref localStoryLength, runSprms.Length > 0 ? runSprms : null);
                }

                void AppendHeaderImageRunContent(RunModel runModel, SectionModel? layoutSectionForStory, int paragraphTopTwips, int paragraphContentHeightTwips, ref int localStoryLength)
                {
                    if (runModel.Image == null ||
                        runModel.Image.Data == null)
                    {
                        return;
                    }

                    string imageContentType = ResolveImageContentType(runModel.Image.ContentType, runModel.Image.Data);
                    int localPictureCp = localStoryLength;

                    textBuilder.Append('\x0001');

                    byte[] pictureBlock = BuildPictureBlock(runModel.Image, imageContentType);
                    int pictureOffset = nextPictureOffset;
                    nextPictureOffset += pictureBlock.Length;
                    embeddedObjects.Add(pictureBlock);

                    if (runModel.Image.LayoutType == ImageLayoutType.Floating &&
                        layoutSectionForStory != null &&
                        SupportsOfficeArtBlip(imageContentType))
                    {
                        (int imageWidthTwips, int imageHeightTwips) = ResolveImageDimensionsTwips(runModel.Image, imageContentType);
                        (int leftTwips, int topTwips, _, _) = ResolveImageBoundsTwips(runModel.Image, imageContentType, layoutSectionForStory, paragraphTopTwips, paragraphContentHeightTwips);
                        headerOfficeArtBlips.Add((localPictureCp, runModel.Image.Data, imageContentType, imageWidthTwips, imageHeightTwips, leftTwips, topTwips, runModel.Image.WrapType, runModel.Image.BehindText, runModel.Image.AllowOverlap, runModel.Image.HorizontalRelativeTo, runModel.Image.VerticalRelativeTo));
                    }

                    chpxWriter.AddRun(currentCp, currentCp + 1, BuildImageSprms(pictureOffset));
                    currentCp += 1;
                    localStoryLength += 1;
                }
            }

            void AppendResolvedHeaderFooterStory(
                HeaderFooterStoryModel? story,
                string? text,
                SectionModel? section,
                ref int storyLength,
                List<int> storyStarts)
            {
                if (story != null)
                {
                    AppendStructuredHeaderFooterStory(story, section, ref storyLength, storyStarts);
                    return;
                }

                AppendHeaderFooterStoryText(textBuilder, ref currentCp, ref storyLength, storyStarts, text, papxWriter);
            }

            int trailingReferenceStart = currentCp;
            AppendPendingFootnoteReferencesAtVisibleCp();
            AppendPendingEndnoteReferencesAtVisibleCp();
            AppendPendingCommentReferencesAtVisibleCp();
            AppendRemainingFootnoteReferencesAtDocumentEnd();
            AppendRemainingEndnoteReferencesAtDocumentEnd();
            AppendRemainingCommentReferencesAtDocumentEnd();
            if (currentCp > trailingReferenceStart)
            {
                papxWriter.AddParagraph(trailingReferenceStart, currentCp, Array.Empty<byte>());
            }

            int mainDocumentCp = currentCp;
            int footnoteStoryLength = 0;
            var footnoteTextStoryStarts = new List<int>(emittedFootnotes.Count + 2);
            for (int footnoteIndex = 0; footnoteIndex < emittedFootnotes.Count; footnoteIndex++)
            {
                var (footnote, _) = emittedFootnotes[footnoteIndex];
                footnoteTextStoryStarts.Add(footnoteStoryLength);
                int paragraphStart = currentCp;

                AppendNoteReferenceMarker(textBuilder, chpxWriter, ref currentCp, ref footnoteStoryLength, footnote.CustomMarkText);

                AppendSecondaryStoryText(textBuilder, ref currentCp, ref footnoteStoryLength, footnote.Text, papxWriter, paragraphStart);
            }

            int headerStoryLength = 0;
            bool hasSectionHeaderFooterStories = HasSectionHeaderFooterStories(model, sections);
            var headerStoryStarts = new List<int>(hasSectionHeaderFooterStories ? 6 + (sections.Count * 6) : 8);
            if (HasExplicitHeaderStories(model, sections))
            {
                AppendHeaderStoryText(textBuilder, ref currentCp, ref headerStoryLength, headerStoryStarts, model.FootnoteSeparatorText, papxWriter);
                AppendHeaderStoryText(textBuilder, ref currentCp, ref headerStoryLength, headerStoryStarts, model.FootnoteContinuationSeparatorText, papxWriter);
                AppendHeaderStoryText(textBuilder, ref currentCp, ref headerStoryLength, headerStoryStarts, model.FootnoteContinuationNoticeText, papxWriter);
                AppendHeaderStoryText(textBuilder, ref currentCp, ref headerStoryLength, headerStoryStarts, model.EndnoteSeparatorText, papxWriter);
                AppendHeaderStoryText(textBuilder, ref currentCp, ref headerStoryLength, headerStoryStarts, model.EndnoteContinuationSeparatorText, papxWriter);
                AppendHeaderStoryText(textBuilder, ref currentCp, ref headerStoryLength, headerStoryStarts, model.EndnoteContinuationNoticeText, papxWriter);

                if (hasSectionHeaderFooterStories)
                {
                    foreach (var section in sections)
                    {
                        AppendResolvedHeaderFooterStory(ResolveEvenPagesHeaderStory(model, section), ResolveEvenPagesHeaderText(model, section), section, ref headerStoryLength, headerStoryStarts);
                        AppendResolvedHeaderFooterStory(ResolveDefaultHeaderStory(section), ResolveDefaultHeaderText(section), section, ref headerStoryLength, headerStoryStarts);
                        AppendResolvedHeaderFooterStory(ResolveEvenPagesFooterStory(model, section), ResolveEvenPagesFooterText(model, section), section, ref headerStoryLength, headerStoryStarts);
                        AppendResolvedHeaderFooterStory(ResolveDefaultFooterStory(section), ResolveDefaultFooterText(section), section, ref headerStoryLength, headerStoryStarts);
                        AppendResolvedHeaderFooterStory(ResolveFirstPageHeaderStory(section), ResolveFirstPageHeaderText(section), section, ref headerStoryLength, headerStoryStarts);
                        AppendResolvedHeaderFooterStory(ResolveFirstPageFooterStory(section), ResolveFirstPageFooterText(section), section, ref headerStoryLength, headerStoryStarts);
                    }
                }
            }

            int commentStoryLength = 0;
            var commentTextStoryStarts = new List<int>(emittedComments.Count + 2);
            for (int commentIndex = 0; commentIndex < emittedComments.Count; commentIndex++)
            {
                var (comment, _, storyText) = emittedComments[commentIndex];
                commentTextStoryStarts.Add(commentStoryLength);
                int paragraphStart = currentCp;

                textBuilder.Append('\x0005');
                chpxWriter.AddRun(currentCp, currentCp + 1, BuildSpecialCharacterSprms());
                currentCp += 1;
                commentStoryLength += 1;

                AppendSecondaryStoryText(textBuilder, ref currentCp, ref commentStoryLength, storyText, papxWriter, paragraphStart);
            }

            int endnoteStoryLength = 0;
            var endnoteTextStoryStarts = new List<int>(emittedEndnotes.Count + 2);
            for (int endnoteIndex = 0; endnoteIndex < emittedEndnotes.Count; endnoteIndex++)
            {
                var (endnote, _) = emittedEndnotes[endnoteIndex];
                endnoteTextStoryStarts.Add(endnoteStoryLength);
                int paragraphStart = currentCp;

                AppendNoteReferenceMarker(textBuilder, chpxWriter, ref currentCp, ref endnoteStoryLength, endnote.CustomMarkText);

                AppendSecondaryStoryText(textBuilder, ref currentCp, ref endnoteStoryLength, endnote.Text, papxWriter, paragraphStart);
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
                tableWriter.Write((int)0); tableWriter.Write((int)mainDocumentCp); tableWriter.Write((int)pnPapx);
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
                tableWriter.Write((int)0); tableWriter.Write((int)mainDocumentCp); tableWriter.Write((int)pnTapx);
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
                tableWriter.Write(mainDocumentCp);

                foreach (var (_, descriptor) in fieldEntries)
                {
                    tableWriter.Write(descriptor);
                }

                lcbPlcffldMom = (int)tableStream.Position - fcPlcffldMom;
            }

            var officeArtPictures = BuildOfficeArtPictureDescriptors(officeArtBlips, 0);
            var headerOfficeArtPictures = BuildOfficeArtPictureDescriptors(headerOfficeArtBlips, officeArtPictures.Count);

            int fcPlcfspaMom = 0;
            int lcbPlcfspaMom = 0;
            if (officeArtPictures.Count > 0)
            {
                fcPlcfspaMom = (int)tableStream.Position;
                WritePlcfspa(tableWriter, officeArtPictures, mainDocumentCp);
                lcbPlcfspaMom = (int)tableStream.Position - fcPlcfspaMom;
            }

            int fcPlcSpaHdr = 0;
            int lcbPlcSpaHdr = 0;
            if (headerOfficeArtPictures.Count > 0)
            {
                fcPlcSpaHdr = (int)tableStream.Position;
                WritePlcfspa(tableWriter, headerOfficeArtPictures, headerStoryLength);
                lcbPlcSpaHdr = (int)tableStream.Position - fcPlcSpaHdr;
            }

            int fcDggInfo = 0;
            int lcbDggInfo = 0;
            if (officeArtPictures.Count > 0 || headerOfficeArtPictures.Count > 0)
            {
                if (headerOfficeArtPictures.Count > 0)
                {
                    officeArtPictures.AddRange(headerOfficeArtPictures);
                }

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
            int fcPlcffndRef = 0;
            int lcbPlcffndRef = 0;
            int fcPlcffndTxt = 0;
            int lcbPlcffndTxt = 0;
            int fcPlcfHdd = 0;
            int lcbPlcfHdd = 0;
            int fcPlcffldHdr = 0;
            int lcbPlcffldHdr = 0;
            int fcPlcfandRef = 0;
            int lcbPlcfandRef = 0;
            int fcPlcfandTxt = 0;
            int lcbPlcfandTxt = 0;
            int fcPlcfendRef = 0;
            int lcbPlcfendRef = 0;
            int fcPlcfendTxt = 0;
            int lcbPlcfendTxt = 0;

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
                tableWriter.Write(mainDocumentCp);
                lcbPlcfBkmkf = (int)tableStream.Position - fcPlcfBkmkf;

                // Write PlcfBkmkl (bookmark last CPs)
                fcPlcfBkmkl = (int)tableStream.Position;
                foreach (var bookmark in model.Bookmarks)
                {
                    tableWriter.Write(bookmark.EndCp);
                }
                // Add terminator
                tableWriter.Write(mainDocumentCp);
                lcbPlcfBkmkl = (int)tableStream.Position - fcPlcfBkmkl;
            }

            if (emittedFootnotes.Count > 0)
            {
                fcPlcffndRef = (int)tableStream.Position;
                foreach (var (_, referenceCp) in emittedFootnotes)
                {
                    tableWriter.Write(referenceCp);
                }
                tableWriter.Write(mainDocumentCp);

                foreach (var _ in emittedFootnotes)
                {
                    tableWriter.Write((ushort)1);
                }

                lcbPlcffndRef = (int)tableStream.Position - fcPlcffndRef;

                fcPlcffndTxt = (int)tableStream.Position;
                foreach (int startCp in footnoteTextStoryStarts)
                {
                    tableWriter.Write(startCp);
                }

                int closingParagraphCp = footnoteStoryLength > 0 ? footnoteStoryLength - 1 : 0;
                tableWriter.Write(closingParagraphCp);
                tableWriter.Write(footnoteStoryLength);
                lcbPlcffndTxt = (int)tableStream.Position - fcPlcffndTxt;
            }

            if (headerStoryStarts.Count > 0)
            {
                fcPlcfHdd = (int)tableStream.Position;
                foreach (int startCp in headerStoryStarts)
                {
                    tableWriter.Write(startCp);
                }

                tableWriter.Write(headerStoryLength - 1);
                tableWriter.Write(headerStoryLength);
                lcbPlcfHdd = (int)tableStream.Position - fcPlcfHdd;
            }

            if (headerFieldEntries.Count > 0)
            {
                fcPlcffldHdr = (int)tableStream.Position;
                foreach (var (cp, _) in headerFieldEntries)
                {
                    tableWriter.Write(cp);
                }
                tableWriter.Write(headerStoryLength);

                foreach (var (_, descriptor) in headerFieldEntries)
                {
                    tableWriter.Write(descriptor);
                }

                lcbPlcffldHdr = (int)tableStream.Position - fcPlcffldHdr;
            }

            if (emittedComments.Count > 0)
            {
                fcPlcfandRef = (int)tableStream.Position;
                foreach (var (_, referenceCp, _) in emittedComments)
                {
                    tableWriter.Write(referenceCp);
                }
                tableWriter.Write(mainDocumentCp);

                foreach (var (comment, _, _) in emittedComments)
                {
                    WriteAtrdPre10(tableWriter, comment);
                }

                lcbPlcfandRef = (int)tableStream.Position - fcPlcfandRef;

                fcPlcfandTxt = (int)tableStream.Position;
                foreach (int startCp in commentTextStoryStarts)
                {
                    tableWriter.Write(startCp);
                }

                int closingParagraphCp = commentStoryLength > 0 ? commentStoryLength - 1 : 0;
                tableWriter.Write(closingParagraphCp);
                tableWriter.Write(commentStoryLength);
                lcbPlcfandTxt = (int)tableStream.Position - fcPlcfandTxt;
            }

            if (emittedEndnotes.Count > 0)
            {
                fcPlcfendRef = (int)tableStream.Position;
                foreach (var (_, referenceCp) in emittedEndnotes)
                {
                    tableWriter.Write(referenceCp);
                }
                tableWriter.Write(mainDocumentCp);

                foreach (var _ in emittedEndnotes)
                {
                    tableWriter.Write((ushort)1);
                }

                lcbPlcfendRef = (int)tableStream.Position - fcPlcfendRef;

                fcPlcfendTxt = (int)tableStream.Position;
                foreach (int startCp in endnoteTextStoryStarts)
                {
                    tableWriter.Write(startCp);
                }

                int closingParagraphCp = endnoteStoryLength > 0 ? endnoteStoryLength - 1 : 0;
                tableWriter.Write(closingParagraphCp);
                tableWriter.Write(endnoteStoryLength);
                lcbPlcfendTxt = (int)tableStream.Position - fcPlcfendTxt;
            }

            // 9. Process section properties: Build Plcfsed and SED/SEP
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

                void AddByteSprm(ushort op, byte val)
                {
                    sepSprms.Add((byte)(op & 0xFF));
                    sepSprms.Add((byte)((op >> 8) & 0xFF));
                    sepSprms.Add(val);
                }

                AddShortSprm(0xB603, section.PageWidth);
                AddShortSprm(0xB604, section.PageHeight);
                AddShortSprm(0xB605, section.MarginLeft);
                AddShortSprm(0xB606, section.MarginRight);
                AddShortSprm(0xB607, section.MarginTop);
                AddShortSprm(0xB608, section.MarginBottom);
                if (HasDifferentFirstPage(section))
                {
                    AddByteSprm(0x300A, 1);
                }

                sepBinaryWriter.Write((short)sepSprms.Count);
                sepBinaryWriter.Write(sepSprms.ToArray());
            }

            // Build Plcfsed in 1Table
            int fcPlcfsed = (int)tableStream.Position;
            tableWriter.Write((int)0);
            tableWriter.Write((int)mainDocumentCp);
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
                fcPlcffndRef = fcPlcffndRef,
                lcbPlcffndRef = lcbPlcffndRef,
                fcPlcffndTxt = fcPlcffndTxt,
                lcbPlcffndTxt = lcbPlcffndTxt,
                fcPlcfHdd = fcPlcfHdd,
                lcbPlcfHdd = lcbPlcfHdd,
                fcPlcffldHdr = fcPlcffldHdr,
                lcbPlcffldHdr = lcbPlcffldHdr,
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
                fcPlcSpaHdr = fcPlcSpaHdr,
                lcbPlcSpaHdr = lcbPlcSpaHdr,
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
                fcPlcfandRef = fcPlcfandRef,
                lcbPlcfandRef = lcbPlcfandRef,
                fcPlcfandTxt = fcPlcfandTxt,
                lcbPlcfandTxt = lcbPlcfandTxt,
                fcPlcfendRef = fcPlcfendRef,
                lcbPlcfendRef = lcbPlcfendRef,
                fcPlcfendTxt = fcPlcfendTxt,
                lcbPlcfendTxt = lcbPlcfendTxt,
                ccpText = mainDocumentCp,
                ccpFtn = footnoteStoryLength,
                ccpHdd = headerStoryLength,
                ccpAtn = commentStoryLength,
                ccpEdn = endnoteStoryLength,
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

        private static byte[] BuildSpecialCharacterSprms()
        {
            var sprms = new List<byte>
            {
                0x55, 0x08, 0x01
            };

            return sprms.ToArray();
        }

        private static string NormalizeRelativeTo(string? relativeTo, bool isHorizontal)
        {
            if (string.IsNullOrWhiteSpace(relativeTo))
            {
                return "page";
            }

            string normalized = relativeTo.Trim().ToLowerInvariant();
            normalized = string.Concat(normalized.Where(char.IsLetterOrDigit));
            if (normalized == "paragraph")
            {
                return "paragraph";
            }

            if (isHorizontal && (normalized == "character" || normalized == "char"))
            {
                return "paragraph";
            }

            if (!isHorizontal && normalized == "line")
            {
                return "paragraph";
            }

            if (normalized == "margin"
                || normalized == "insidemargin"
                || normalized == "outsidemargin")
            {
                return "margin";
            }

            if (isHorizontal && (normalized == "column" || normalized == "leftmargin" || normalized == "rightmargin"))
            {
                return "margin";
            }

            if (!isHorizontal && (normalized == "topmargin" || normalized == "bottommargin"))
            {
                return "margin";
            }

            return "page";
        }

        private static void AppendSecondaryStoryText(
            StringBuilder textBuilder,
            ref int currentCp,
            ref int storyLength,
            string text,
            PapxFkpWriter? papxWriter = null,
            int? paragraphStartCp = null)
        {
            int paragraphStart = paragraphStartCp ?? currentCp;

            if (!string.IsNullOrEmpty(text))
            {
                for (int index = 0; index < text.Length; index++)
                {
                    char ch = text[index];
                    textBuilder.Append(ch);
                    currentCp += 1;
                    storyLength += 1;

                    if (ch == '\r' && papxWriter != null)
                    {
                        papxWriter.AddParagraph(paragraphStart, currentCp, Array.Empty<byte>());
                        paragraphStart = currentCp;
                    }
                }
            }

            textBuilder.Append('\r');
            currentCp += 1;
            storyLength += 1;

            papxWriter?.AddParagraph(paragraphStart, currentCp, Array.Empty<byte>());
        }

        private static void AppendNoteReferenceMarker(
            StringBuilder textBuilder,
            ChpxFkpWriter chpxWriter,
            ref int currentCp,
            string? customMarkText)
        {
            string? normalizedCustomMark = NormalizeCustomMarkText(customMarkText);
            if (normalizedCustomMark != null)
            {
                textBuilder.Append(normalizedCustomMark);
                currentCp += normalizedCustomMark.Length;
                return;
            }

            textBuilder.Append('\x0002');
            chpxWriter.AddRun(currentCp, currentCp + 1, BuildSpecialCharacterSprms());
            currentCp += 1;
        }

        private static void AppendNoteReferenceMarker(
            StringBuilder textBuilder,
            ChpxFkpWriter chpxWriter,
            ref int currentCp,
            ref int storyLength,
            string? customMarkText)
        {
            string? normalizedCustomMark = NormalizeCustomMarkText(customMarkText);
            if (normalizedCustomMark != null)
            {
                textBuilder.Append(normalizedCustomMark);
                currentCp += normalizedCustomMark.Length;
                storyLength += normalizedCustomMark.Length;
                return;
            }

            textBuilder.Append('\x0002');
            chpxWriter.AddRun(currentCp, currentCp + 1, BuildSpecialCharacterSprms());
            currentCp += 1;
            storyLength += 1;
        }

        private static string? NormalizeCustomMarkText(string? customMarkText)
        {
            if (string.IsNullOrWhiteSpace(customMarkText))
            {
                return null;
            }

            foreach (char ch in customMarkText)
            {
                if (char.IsControl(ch))
                {
                    return null;
                }
            }

            return customMarkText;
        }

        private static void AppendHeaderStoryText(
            StringBuilder textBuilder,
            ref int currentCp,
            ref int storyLength,
            List<int> storyStarts,
            string? text,
            PapxFkpWriter papxWriter)
        {
            storyStarts.Add(storyLength);
            if (text == null)
            {
                return;
            }

            AppendHeaderStoryParagraphs(textBuilder, ref currentCp, ref storyLength, text, appendGuardParagraph: false, papxWriter);
        }

        private static void AppendHeaderFooterStoryText(
            StringBuilder textBuilder,
            ref int currentCp,
            ref int storyLength,
            List<int> storyStarts,
            string? text,
            PapxFkpWriter papxWriter)
        {
            storyStarts.Add(storyLength);
            if (text == null)
            {
                return;
            }

            AppendHeaderStoryParagraphs(textBuilder, ref currentCp, ref storyLength, text, appendGuardParagraph: true, papxWriter);
        }

        private static void AppendHeaderStoryParagraphs(
            StringBuilder textBuilder,
            ref int currentCp,
            ref int storyLength,
            string? text,
            bool appendGuardParagraph,
            PapxFkpWriter papxWriter)
        {
            string storyText = text ?? string.Empty;
            int paragraphStart = currentCp;

            for (int index = 0; index < storyText.Length; index++)
            {
                char ch = storyText[index];
                textBuilder.Append(ch);
                currentCp += 1;
                storyLength += 1;

                if (ch == '\r')
                {
                    papxWriter.AddParagraph(paragraphStart, currentCp, Array.Empty<byte>());
                    paragraphStart = currentCp;
                }
            }

            if (storyText.Length == 0 || storyText[^1] != '\r')
            {
                textBuilder.Append('\r');
                currentCp += 1;
                storyLength += 1;
                papxWriter.AddParagraph(paragraphStart, currentCp, Array.Empty<byte>());
                paragraphStart = currentCp;
            }

            if (appendGuardParagraph)
            {
                textBuilder.Append('\r');
                currentCp += 1;
                storyLength += 1;
                papxWriter.AddParagraph(paragraphStart, currentCp, Array.Empty<byte>());
            }
        }

        private static ushort BuildFieldDescriptorCore(char marker, FieldModel? fieldModel, int nestingDepth)
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

            descriptor |= (ushort)((Math.Min(Math.Max(nestingDepth, 0), 15) & 0x0F) << 12);
            return descriptor;
        }

        private static bool HasExplicitFieldBoundaryCore(IList<RunModel> runs, int startIndex, FieldModel fieldModel)
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

        private static string ResolveFieldInstructionCore(FieldModel fieldModel)
        {
            if (!string.IsNullOrWhiteSpace(fieldModel.Instruction))
            {
                return fieldModel.Instruction;
            }

            return fieldModel.Type switch
            {
                FieldType.Page => "PAGE",
                FieldType.NumPages => "NUMPAGES",
                FieldType.SectionPages => "SECTIONPAGES",
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

        private static string BuildHyperlinkInstructionCore(HyperlinkModel hyperlinkModel)
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

        private static void RemoveLastOpenFieldCore(List<FieldModel> openFieldList, FieldModel fieldModel)
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

        private static int GetFieldDepthCore(List<FieldModel> openFieldList, FieldModel? fieldModel)
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

        private static bool HasExplicitHeaderStories(DocumentModel model, IReadOnlyList<SectionModel> sections)
        {
            return model.FootnoteSeparatorText != null ||
                   model.FootnoteContinuationSeparatorText != null ||
                   model.FootnoteContinuationNoticeText != null ||
                   model.EndnoteSeparatorText != null ||
                   model.EndnoteContinuationSeparatorText != null ||
                   model.EndnoteContinuationNoticeText != null ||
                   HasSectionHeaderFooterStories(model, sections);
        }

        private static bool HasSectionHeaderFooterStories(DocumentModel model, IReadOnlyList<SectionModel> sections)
        {
            foreach (var section in sections)
            {
                if (ResolveDefaultHeaderStory(section) != null ||
                    ResolveEvenPagesHeaderStory(model, section) != null ||
                    ResolveDefaultFooterStory(section) != null ||
                    ResolveEvenPagesFooterStory(model, section) != null ||
                    ResolveFirstPageHeaderStory(section) != null ||
                    ResolveFirstPageFooterStory(section) != null ||
                    ResolveDefaultHeaderText(section) != null ||
                    ResolveEvenPagesHeaderText(model, section) != null ||
                    ResolveDefaultFooterText(section) != null ||
                    ResolveEvenPagesFooterText(model, section) != null ||
                    ResolveFirstPageHeaderText(section) != null ||
                    ResolveFirstPageFooterText(section) != null)
                {
                    return true;
                }
            }

            return false;
        }

        private static HeaderFooterStoryModel? ResolveDefaultHeaderStory(SectionModel section)
        {
            return section.DefaultHeaderStory;
        }

        private static string? ResolveDefaultHeaderText(SectionModel section)
        {
            if (section.DefaultHeaderText != null)
            {
                return section.DefaultHeaderText;
            }

            return section.FirstPageHeaderStory == null &&
                   section.EvenPagesHeaderStory == null &&
                   section.FirstPageHeaderText == null &&
                   section.EvenPagesHeaderText == null
                ? section.HeaderText
                : null;
        }

        private static HeaderFooterStoryModel? ResolveEvenPagesHeaderStory(DocumentModel model, SectionModel section)
        {
            return model.DifferentOddAndEvenPages ? section.EvenPagesHeaderStory : null;
        }

        private static string? ResolveEvenPagesHeaderText(DocumentModel model, SectionModel section)
        {
            return model.DifferentOddAndEvenPages ? section.EvenPagesHeaderText : null;
        }

        private static HeaderFooterStoryModel? ResolveDefaultFooterStory(SectionModel section)
        {
            return section.DefaultFooterStory;
        }

        private static string? ResolveDefaultFooterText(SectionModel section)
        {
            if (section.DefaultFooterText != null)
            {
                return section.DefaultFooterText;
            }

            return section.FirstPageFooterStory == null &&
                   section.EvenPagesFooterStory == null &&
                   section.FirstPageFooterText == null &&
                   section.EvenPagesFooterText == null
                ? section.FooterText
                : null;
        }

        private static HeaderFooterStoryModel? ResolveEvenPagesFooterStory(DocumentModel model, SectionModel section)
        {
            return model.DifferentOddAndEvenPages ? section.EvenPagesFooterStory : null;
        }

        private static string? ResolveEvenPagesFooterText(DocumentModel model, SectionModel section)
        {
            return model.DifferentOddAndEvenPages ? section.EvenPagesFooterText : null;
        }

        private static HeaderFooterStoryModel? ResolveFirstPageHeaderStory(SectionModel section)
        {
            return HasDifferentFirstPage(section) ? section.FirstPageHeaderStory : null;
        }

        private static string? ResolveFirstPageHeaderText(SectionModel section)
        {
            return HasDifferentFirstPage(section) ? section.FirstPageHeaderText : null;
        }

        private static HeaderFooterStoryModel? ResolveFirstPageFooterStory(SectionModel section)
        {
            return HasDifferentFirstPage(section) ? section.FirstPageFooterStory : null;
        }

        private static string? ResolveFirstPageFooterText(SectionModel section)
        {
            return HasDifferentFirstPage(section) ? section.FirstPageFooterText : null;
        }

        private static bool HasDifferentFirstPage(SectionModel section)
        {
            return section.DifferentFirstPage ||
                   section.FirstPageHeaderStory != null ||
                   section.FirstPageFooterStory != null ||
                   section.FirstPageHeaderText != null ||
                   section.FirstPageFooterText != null;
        }

        private static byte[] BuildImageSprms(int pictureOffset)
        {
            var sprms = new List<byte>(BuildSpecialCharacterSprms())
            {
                0x03, 0x6A
            };

            sprms.AddRange(BitConverter.GetBytes(pictureOffset));

            return sprms.ToArray();
        }

        private static void WriteAtrdPre10(BinaryWriter writer, CommentModel comment)
        {
            string initials = ResolveCommentInitials(comment);
            int length = Math.Min(initials.Length, 10);
            for (int index = 0; index < 10; index++)
            {
                char value = index < length ? initials[index] : '\0';
                writer.Write((ushort)value);
            }

            writer.Write(unchecked((short)-1));
            writer.Write((short)0);
            writer.Write((ushort)0);
            writer.Write(-1);
        }

        private static bool CanWriteFootnote(FootnoteModel footnote)
        {
            return CanWriteNoteReference(footnote.ReferenceCp, footnote.Id);
        }

        private static bool CanWriteEndnote(EndnoteModel endnote)
        {
            return CanWriteNoteReference(endnote.ReferenceCp, endnote.Id);
        }

        private static List<(CommentModel comment, int index, string storyText)> BuildSupportedComments(IReadOnlyList<CommentModel> comments)
        {
            var commentEntries = comments
                .Select((comment, index) => (comment, index))
                .ToList();
            var commentsById = commentEntries
                .Where(static entry => !string.IsNullOrEmpty(entry.comment.Id))
                .ToDictionary(entry => entry.comment.Id, entry => entry, StringComparer.Ordinal);
            var replyCommentsByParentId = BuildReplyCommentsByParentId(commentEntries);
            var rootIndexByCommentIndex = new Dictionary<int, int>();
            var supportedComments = new List<(CommentModel comment, int index, string storyText)>();

            foreach (var entry in commentEntries)
            {
                rootIndexByCommentIndex[entry.index] = ResolveCommentEmissionRootIndex(entry.index);
            }

            foreach (var entry in commentEntries)
            {
                if (!CanAnchorComment(entry.comment))
                {
                    continue;
                }

                int rootIndex = rootIndexByCommentIndex[entry.index];
                if (rootIndex != entry.index)
                {
                    continue;
                }

                string storyText = BuildCommentStoryText(entry, replyCommentsByParentId, rootIndexByCommentIndex);
                supportedComments.Add((entry.comment, entry.index, storyText));
            }

            return supportedComments;

            int ResolveCommentEmissionRootIndex(int commentIndex)
            {
                if (rootIndexByCommentIndex.TryGetValue(commentIndex, out int cachedRootIndex))
                {
                    return cachedRootIndex;
                }

                var visited = new HashSet<int>();
                int resolvedRootIndex = ResolveCommentEmissionRootIndexCore(commentIndex, visited);
                rootIndexByCommentIndex[commentIndex] = resolvedRootIndex;
                return resolvedRootIndex;
            }

            int ResolveCommentEmissionRootIndexCore(int commentIndex, HashSet<int> visited)
            {
                if (rootIndexByCommentIndex.TryGetValue(commentIndex, out int cachedRootIndex))
                {
                    return cachedRootIndex;
                }

                if (!visited.Add(commentIndex))
                {
                    return commentIndex;
                }

                var entry = commentEntries[commentIndex];
                if (!CanAnchorComment(entry.comment))
                {
                    if (entry.comment.IsReply &&
                        !string.IsNullOrEmpty(entry.comment.ParentId) &&
                        commentsById.TryGetValue(entry.comment.ParentId, out var parentEntry))
                    {
                        int parentRootIndex = ResolveCommentEmissionRootIndexCore(parentEntry.index, visited);
                        rootIndexByCommentIndex[commentIndex] = parentRootIndex;
                        return parentRootIndex;
                    }

                    rootIndexByCommentIndex[commentIndex] = -1;
                    return -1;
                }

                if (!entry.comment.IsReply || string.IsNullOrEmpty(entry.comment.ParentId))
                {
                    rootIndexByCommentIndex[commentIndex] = commentIndex;
                    return commentIndex;
                }

                if (!commentsById.TryGetValue(entry.comment.ParentId, out var directParentEntry))
                {
                    rootIndexByCommentIndex[commentIndex] = commentIndex;
                    return commentIndex;
                }

                int directParentRootIndex = ResolveCommentEmissionRootIndexCore(directParentEntry.index, visited);
                int resolvedRootIndex = directParentRootIndex >= 0 ? directParentRootIndex : commentIndex;
                rootIndexByCommentIndex[commentIndex] = resolvedRootIndex;
                return resolvedRootIndex;
            }
        }

        private static Dictionary<string, List<(CommentModel comment, int index)>> BuildReplyCommentsByParentId(IReadOnlyList<(CommentModel comment, int index)> commentEntries)
        {
            var repliesByParentId = new Dictionary<string, List<(CommentModel comment, int index)>>(StringComparer.Ordinal);
            foreach (var entry in commentEntries)
            {
                var comment = entry.comment;
                if (!comment.IsReply || string.IsNullOrEmpty(comment.ParentId))
                {
                    continue;
                }

                if (!repliesByParentId.TryGetValue(comment.ParentId, out var replies))
                {
                    replies = new List<(CommentModel comment, int index)>();
                    repliesByParentId[comment.ParentId] = replies;
                }

                replies.Add(entry);
            }

            foreach (var replies in repliesByParentId.Values)
            {
                replies.Sort(static (left, right) => left.index.CompareTo(right.index));
            }

            return repliesByParentId;
        }

        private static string BuildCommentStoryText(
            (CommentModel comment, int index) rootEntry,
            IReadOnlyDictionary<string, List<(CommentModel comment, int index)>> replyCommentsByParentId,
            IReadOnlyDictionary<int, int> rootIndexByCommentIndex)
        {
            var builder = new StringBuilder();
            var visitedReplyIndexes = new HashSet<int>();

            AppendCommentBody(rootEntry.comment, includeReplyHeader: rootEntry.comment.IsReply);
            AppendReplies(rootEntry.comment.Id);
            return builder.ToString();

            void AppendReplies(string? parentId)
            {
                if (string.IsNullOrEmpty(parentId) || !replyCommentsByParentId.TryGetValue(parentId, out var replies))
                {
                    return;
                }

                foreach (var (reply, index) in replies)
                {
                    if (!visitedReplyIndexes.Add(index))
                    {
                        continue;
                    }

                    if (!rootIndexByCommentIndex.TryGetValue(index, out int rootIndex) || rootIndex != rootEntry.index)
                    {
                        continue;
                    }

                    AppendCommentBody(reply, includeReplyHeader: true);
                    AppendReplies(reply.Id);
                }
            }

            void AppendCommentBody(CommentModel comment, bool includeReplyHeader)
            {
                if (builder.Length > 0)
                {
                    builder.Append('\r');
                }

                if (includeReplyHeader)
                {
                    builder.Append(BuildReplyHeader(comment));
                    if (!string.IsNullOrEmpty(comment.Text))
                    {
                        builder.Append('\r');
                    }
                }

                if (!string.IsNullOrEmpty(comment.Text))
                {
                    builder.Append(comment.Text);
                }
            }
        }

        private static bool CanAnchorComment(CommentModel comment)
        {
            return comment.StartCp >= 0 &&
                   comment.EndCp >= comment.StartCp;
        }

        private static string BuildReplyHeader(CommentModel comment)
        {
            var header = new StringBuilder("[Reply");
            string displayName = ResolveCommentDisplayName(comment);
            if (!string.IsNullOrEmpty(displayName))
            {
                header.Append(" by ");
                header.Append(displayName);
            }

            if (comment.Date.HasValue)
            {
                header.Append(" at ");
                header.Append(comment.Date.Value.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture));
            }

            header.Append(']');
            return header.ToString();
        }

        private static string ResolveCommentDisplayName(CommentModel comment)
        {
            if (!string.IsNullOrWhiteSpace(comment.Author))
            {
                return comment.Author.Trim();
            }

            if (!string.IsNullOrWhiteSpace(comment.Initials))
            {
                return comment.Initials.Trim();
            }

            return string.Empty;
        }

        private static string ResolveCommentInitials(CommentModel comment)
        {
            if (!string.IsNullOrWhiteSpace(comment.Initials))
            {
                return comment.Initials.Trim();
            }

            if (string.IsNullOrWhiteSpace(comment.Author))
            {
                return string.Empty;
            }

            var builder = new StringBuilder(10);
            string[] tokens = comment.Author.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries);
            foreach (string token in tokens)
            {
                if (builder.Length >= 10)
                {
                    break;
                }

                builder.Append(char.ToUpperInvariant(token[0]));
            }

            if (builder.Length == 0)
            {
                foreach (char value in comment.Author)
                {
                    if (!char.IsWhiteSpace(value))
                    {
                        builder.Append(char.ToUpperInvariant(value));
                        break;
                    }
                }
            }

            return builder.ToString();
        }

        private static bool IsSpecialNoteId(string? id)
        {
            return int.TryParse(id, out int parsedId) && parsedId <= 1;
        }

        private static bool CanWriteNoteReference(int referenceCp, string? id)
        {
            return referenceCp >= 0 &&
                   !string.IsNullOrWhiteSpace(id) &&
                   !IsSpecialNoteId(id);
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
            List<(int cp, byte[] data, string contentType, int widthTwips, int heightTwips, int leftTwips, int topTwips, ImageWrapType wrapType, bool behindText, bool allowOverlap, string? horizontalRelativeTo, string? verticalRelativeTo)> officeArtBlips,
            int pictureIndexOffset)
        {
            var pictures = new List<OfficeArtPictureDescriptor>(officeArtBlips.Count);
            for (int index = 0; index < officeArtBlips.Count; index++)
            {
                var picture = officeArtBlips[index];
                int pictureIndex = pictureIndexOffset + index;
                pictures.Add(new OfficeArtPictureDescriptor(
                    picture.cp,
                    GetPictureShapeId(pictureIndex),
                    pictureIndex + 1,
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

        private static void WritePlcfspa(
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
            ushort propertyId = (ushort)(unchecked((ushort)propertyNumber) | (isBlipId ? (ushort)0x4000 : (ushort)0));
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

            for (int cellIndex = 0; cellIndex < row.Cells.Count; cellIndex++)
            {
                var cell = row.Cells[cellIndex];
                int span = Math.Max(1, cell.GridSpan);
                bool isFirstColumn = gridColumnIndex == 0;
                bool isLastColumn = gridColumnIndex + span >= totalColumnCount;
                var previousCell = cellIndex > 0 ? row.Cells[cellIndex - 1] : null;
                int explicitWidthTwips = ResolveExplicitCellWidthTwips(cell, targetTableWidthTwips);
                int widthTwips = explicitWidthTwips;
                bool resolvedFromExplicitWidth = explicitWidthTwips > 0;
                int horizontalOverheadTwips = ResolveAutoCellHorizontalOverheadTwips(table, cell, previousCell, isFirstColumn, isLastColumn);
                if (widthTwips <= 0 && table.GridColumnWidths.Count > 0)
                {
                    widthTwips = ResolveTableCellWidth(table, gridColumnIndex, span, 0);
                }

                widthsTwips.Add(Math.Max(0, widthTwips));
                spans.Add(span);
                isExplicitResolvedWidth.Add(resolvedFromExplicitWidth);
                explicitWidthUnits.Add(resolvedFromExplicitWidth ? cell.WidthUnit : TableWidthUnit.Auto);
                minimumAutoWidthsTwips.Add(Math.Max(720, 720 * span) + horizontalOverheadTwips);

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

        private static int ResolveAutoCellHorizontalOverheadTwips(TableModel table, TableCellModel cell, TableCellModel? previousCell, bool isFirstColumn, bool isLastColumn)
        {
            int leftBorderTwips = ResolveTableCellLeftBorderTwips(table, cell, previousCell, isFirstColumn);
            int rightBorderTwips = ResolveTableCellRightBorderTwips(table, cell, isLastColumn);
            int horizontalPaddingTwips = ResolveTableCellHorizontalPaddingTwips(table, cell);
            int horizontalSpacingTwips = ResolveTableCellHorizontalSpacingTwips(table);

            return Math.Max(0, leftBorderTwips)
                + Math.Max(0, rightBorderTwips)
                + Math.Max(0, horizontalPaddingTwips)
                + Math.Max(0, horizontalSpacingTwips);
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
                return Math.Max(276, contentHeightTwips);
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

        private static IEnumerable<object> EnumerateTableCellBlocks(TableCellModel cell)
        {
            if (cell.Content.Count > 0)
            {
                foreach (var block in cell.Content)
                {
                    yield return block;
                }

                yield break;
            }

            foreach (var paragraph in cell.Paragraphs)
            {
                yield return paragraph;
            }
        }

        private static IEnumerable<ParagraphModel> EnumerateTableBlockParagraphs(object block)
        {
            if (block is ParagraphModel paragraph)
            {
                yield return paragraph;
                yield break;
            }

            if (block is TableModel table)
            {
                foreach (var row in table.Rows)
                {
                    foreach (var cell in row.Cells)
                    {
                        foreach (var cellBlock in EnumerateTableCellBlocks(cell))
                        {
                            foreach (var cellParagraph in EnumerateTableBlockParagraphs(cellBlock))
                            {
                                yield return cellParagraph;
                            }
                        }
                    }
                }
            }
        }

        private static IEnumerable<ParagraphModel> EnumerateTableCellParagraphs(TableCellModel cell)
        {
            foreach (var block in EnumerateTableCellBlocks(cell))
            {
                foreach (var paragraph in EnumerateTableBlockParagraphs(block))
                {
                    yield return paragraph;
                }
            }
        }

        private static int EstimateTableCellContentHeightTwips(TableCellModel cell, int cellAvailableWidthTwips)
        {
            int contentHeightTwips = 0;
            foreach (var paragraph in EnumerateTableCellParagraphs(cell))
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
            foreach (var paragraph in EnumerateTableCellParagraphs(cell))
            {
                int paragraphMaxLineWidth = 0;
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
                        if (IsParagraphBreakCharacter(c))
                        {
                            paragraphMaxLineWidth = Math.Max(paragraphMaxLineWidth, currentWidth);
                            currentWidth = 0;
                            continue;
                        }

                        double baseWidth = EstimateCharacterWidthTwips(c, runFontSizeTwips, isMonospace);
                        currentWidth += (int)Math.Round(baseWidth * runStyleMultiplier, MidpointRounding.AwayFromZero);
                    }

                    if (run.Image != null && run.Image.LayoutType == ImageLayoutType.Inline)
                    {
                        currentWidth += Math.Max(960, run.Image.Width * 15);
                    }
                }

                paragraphMaxLineWidth = Math.Max(paragraphMaxLineWidth, currentWidth);
                paragraphMaxLineWidth += Math.Max(0, paragraph.Properties.LeftIndentTwips) + Math.Max(0, paragraph.Properties.RightIndentTwips);
                maxContentWidth = Math.Max(maxContentWidth, paragraphMaxLineWidth);
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
                    if (IsParagraphBreakCharacter(c))
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
                        if (IsParagraphBreakCharacter(wc)) break;

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

        private static bool IsParagraphBreakCharacter(char character)
        {
            return character == '\n' || character == '\r' || character == '\v' || character == '\f';
        }

        private static double EstimateCharacterWidthTwips(char character, int fontSizeTwips, bool isMonospace)
        {
            if (IsParagraphBreakCharacter(character))
            {
                return 0;
            }

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

            flags |= EncodeRelativeTo(picture.VerticalRelativeTo, isHorizontal: false) << 5;
            flags |= EncodeRelativeTo(picture.HorizontalRelativeTo, isHorizontal: true) << 7;

            return flags;
        }

        private static int EncodeRelativeTo(string? relativeTo, bool isHorizontal)
        {
            string normalized = NormalizeRelativeTo(relativeTo, isHorizontal);

            if (normalized == "margin")
            {
                return 1;
            }

            if (normalized == "paragraph")
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
            string normalizedRelativeTo = NormalizeRelativeTo(relativeTo, isHorizontal);
            if (normalizedRelativeTo == "paragraph")
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
                return ResolveAlignmentPositionTwips(alignment, normalizedRelativeTo, sizeTwips, pageExtentTwips, leadingMarginTwips, trailingMarginTwips, isHorizontal);
            }

            if (normalizedRelativeTo == "margin")
            {
                return leadingMarginTwips + offsetTwips;
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
