using System;
using System.IO;
using System.IO.Compression;
using System.Xml;
using System.Text;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Net;
using System.Text.RegularExpressions;

namespace Nedev.FileConverters.DocxToDoc.Format
{
    /// <summary>
    /// Reads and coordinates the extraction of features from OpenXML (.docx) files.
    /// Optimized for low-memory, forward-only reading.
    /// </summary>
    public class DocxReader : IDisposable
    {
        private enum NoteCaptureKind
        {
            Ignore,
            Regular,
            Separator,
            ContinuationSeparator,
            ContinuationNotice
        }

        private readonly ZipArchive _archive;
        private readonly Dictionary<string, string> _relationships;
        private bool _disposedValue;

        public DocxReader(Stream docxStream)
        {
            // We assume the caller manages the base stream's lifecycle.
            _archive = new ZipArchive(docxStream, ZipArchiveMode.Read, leaveOpen: true);
            _relationships = LoadRelationships();
        }

        private Dictionary<string, string> LoadRelationships()
        {
            return LoadRelationships("word/_rels/document.xml.rels");
        }

        private Dictionary<string, string> LoadRelationships(string relationshipsEntryPath)
        {
            var rels = new Dictionary<string, string>(StringComparer.Ordinal);
            var relsEntry = _archive.GetEntry(relationshipsEntryPath);
            if (relsEntry == null) return rels;

            using var stream = relsEntry.Open();
            using var reader = XmlReader.Create(stream);
            while (reader.Read())
            {
                if (reader.NodeType == XmlNodeType.Element && reader.LocalName == "Relationship")
                {
                    string? id = reader.GetAttribute("Id");
                    string? target = reader.GetAttribute("Target");
                    if (!string.IsNullOrEmpty(id) && !string.IsNullOrEmpty(target))
                    {
                        rels[id] = target;
                    }
                }
            }
            return rels;
        }

        private Dictionary<string, string> LoadRelationshipsForPart(string partEntryPath)
        {
            string fileName = Path.GetFileName(partEntryPath);
            string? directory = Path.GetDirectoryName(partEntryPath)?.Replace('\\', '/');
            string relationshipsEntryPath = string.IsNullOrEmpty(directory)
                ? $"_rels/{fileName}.rels"
                : $"{directory}/_rels/{fileName}.rels";

            return LoadRelationships(relationshipsEntryPath);
        }

        public Nedev.FileConverters.DocxToDoc.Model.DocumentModel ReadDocument()
        {
            var documentEntry = _archive.GetEntry("word/document.xml");
            if (documentEntry == null)
            {
                throw new FileNotFoundException("word/document.xml not found in the docx file.");
            }

            using var stream = documentEntry.Open();
            using var xmlReader = XmlReader.Create(stream, new XmlReaderSettings 
            { 
                IgnoreComments = true, 
                IgnoreWhitespace = true 
            });

            var docModel = new Nedev.FileConverters.DocxToDoc.Model.DocumentModel();
            
            // Parse Styles
            var stylesEntry = _archive.GetEntry("word/styles.xml");
            if (stylesEntry != null)
            {
                using var stylesStream = stylesEntry.Open();
                ParseStyles(stylesStream, docModel);
            }

            // Parse Font Table
            var fontTableEntry = _archive.GetEntry("word/fontTable.xml");
            if (fontTableEntry != null)
            {
                using var fontsStream = fontTableEntry.Open();
                ParseFonts(fontsStream, docModel);
            }

            // Parse Numbering
            var numberingEntry = _archive.GetEntry("word/numbering.xml");
            if (numberingEntry != null)
            {
                using var numberingStream = numberingEntry.Open();
                ParseNumbering(numberingStream, docModel);
            }

            // Parse Document Properties
            var propsEntry = _archive.GetEntry("docProps/core.xml");
            if (propsEntry != null)
            {
                using var propsStream = propsEntry.Open();
                ParseDocumentProperties(propsStream, docModel);
            }

            var extendedPropsEntry = _archive.GetEntry("docProps/app.xml");
            if (extendedPropsEntry != null)
            {
                using var extendedPropsStream = extendedPropsEntry.Open();
                ParseDocumentProperties(extendedPropsStream, docModel);
            }

            var settingsEntry = _archive.GetEntry("word/settings.xml");
            if (settingsEntry != null)
            {
                using var settingsStream = settingsEntry.Open();
                ParseDocumentSettings(settingsStream, docModel);
            }

            // Parse Footnotes
            var footnotesEntry = _archive.GetEntry("word/footnotes.xml");
            if (footnotesEntry != null)
            {
                using var footnotesStream = footnotesEntry.Open();
                ParseFootnotes(footnotesStream, docModel);
            }

            // Parse Endnotes
            var endnotesEntry = _archive.GetEntry("word/endnotes.xml");
            if (endnotesEntry != null)
            {
                using var endnotesStream = endnotesEntry.Open();
                ParseEndnotes(endnotesStream, docModel);
            }

            // Parse Comments
            var commentsEntry = _archive.GetEntry("word/comments.xml");
            var commentParagraphIds = new Dictionary<string, string>(StringComparer.Ordinal);
            if (commentsEntry != null)
            {
                using var commentsStream = commentsEntry.Open();
                ParseComments(commentsStream, docModel, commentParagraphIds);
            }

            var commentsExtendedEntry = _archive.GetEntry("word/commentsExtended.xml");
            if (commentsExtendedEntry != null && commentParagraphIds.Count > 0 && docModel.Comments.Count > 0)
            {
                using var commentsExtendedStream = commentsExtendedEntry.Open();
                ParseCommentsExtended(commentsExtendedStream, docModel, commentParagraphIds);
            }

            // Parse VBA Project
            var vbaEntry = _archive.GetEntry("word/vbaProject.bin");
            if (vbaEntry != null)
            {
                using var vbaStream = vbaEntry.Open();
                using var ms = new MemoryStream();
                vbaStream.CopyTo(ms);
                docModel.VbaProjectData = ms.ToArray();
            }

            StringBuilder textBuffer = new StringBuilder();

            Nedev.FileConverters.DocxToDoc.Model.ParagraphModel? currentParagraph = null;
            Nedev.FileConverters.DocxToDoc.Model.RunModel? currentRun = null;
            Nedev.FileConverters.DocxToDoc.Model.RunModel.CharacterProperties? currentRunBaseProperties = null;
            Nedev.FileConverters.DocxToDoc.Model.SectionModel? currentSection = null;
            Nedev.FileConverters.DocxToDoc.Model.TableModel? currentTable = null;
            Nedev.FileConverters.DocxToDoc.Model.TableRowModel? currentRow = null;
            Nedev.FileConverters.DocxToDoc.Model.TableCellModel? currentCell = null;
            Nedev.FileConverters.DocxToDoc.Model.TableCellModel? currentRowHorizontalMergeAnchor = null;
            int currentRowGridColumnIndex = 0;
            bool insideTableCellMargins = false;
            bool insideCellMargins = false;
            bool insideTableBorders = false;
            bool insideCellBorders = false;
            var openFields = new Stack<Nedev.FileConverters.DocxToDoc.Model.FieldModel>();
            var simpleFields = new Stack<(Nedev.FileConverters.DocxToDoc.Model.ParagraphModel Paragraph, Nedev.FileConverters.DocxToDoc.Model.FieldModel Field)>();
            bool inIns = false;
            bool inDel = false;
            bool inOMath = false;

            while (xmlReader.Read())
            {
                if (xmlReader.NodeType == XmlNodeType.Element)
                {
                    string localName = xmlReader.LocalName;
                    if (localName == "tbl")
                    {
                        currentTable = new Nedev.FileConverters.DocxToDoc.Model.TableModel();
                        docModel.Content.Add(currentTable);
                    }
                    else if (localName == "tr" && currentTable != null)
                    {
                        currentRow = new Nedev.FileConverters.DocxToDoc.Model.TableRowModel();
                        currentTable.Rows.Add(currentRow);
                        currentRowHorizontalMergeAnchor = null;
                        currentRowGridColumnIndex = 0;
                    }
                    else if (localName == "tc" && currentRow != null)
                    {
                        currentCell = new Nedev.FileConverters.DocxToDoc.Model.TableCellModel();
                        currentRow.Cells.Add(currentCell);
                    }
                    else if (localName == "trHeight" && currentRow != null)
                    {
                        if (int.TryParse(xmlReader.GetAttribute("w:val"), out int rowHeightTwips))
                        {
                            currentRow.HeightTwips = rowHeightTwips;
                        }

                        string? heightRule = xmlReader.GetAttribute("w:hRule");
                        currentRow.HeightRule = heightRule switch
                        {
                            "exact" => Nedev.FileConverters.DocxToDoc.Model.TableRowHeightRule.Exact,
                            "atLeast" => Nedev.FileConverters.DocxToDoc.Model.TableRowHeightRule.AtLeast,
                            _ => Nedev.FileConverters.DocxToDoc.Model.TableRowHeightRule.Auto
                        };
                    }
                    else if (localName == "tblHeader" && currentRow != null)
                    {
                        currentRow.IsHeader = !IsFalseValue(xmlReader.GetAttribute("w:val"));
                    }
                    else if (localName == "cantSplit" && currentRow != null)
                    {
                        currentRow.CannotSplit = !IsFalseValue(xmlReader.GetAttribute("w:val"));
                    }
                    else if (localName == "tcW" && currentCell != null)
                    {
                        string? type = xmlReader.GetAttribute("w:type");
                        currentCell.WidthUnit = type switch
                        {
                            "pct" => Nedev.FileConverters.DocxToDoc.Model.TableWidthUnit.Pct,
                            "auto" => Nedev.FileConverters.DocxToDoc.Model.TableWidthUnit.Auto,
                            "nil" => Nedev.FileConverters.DocxToDoc.Model.TableWidthUnit.Auto,
                            _ => Nedev.FileConverters.DocxToDoc.Model.TableWidthUnit.Dxa
                        };

                        if (int.TryParse(xmlReader.GetAttribute("w:w"), out int width) &&
                            (currentCell.WidthUnit == Nedev.FileConverters.DocxToDoc.Model.TableWidthUnit.Dxa ||
                             currentCell.WidthUnit == Nedev.FileConverters.DocxToDoc.Model.TableWidthUnit.Pct))
                        {
                            currentCell.Width = width;
                        }
                        else
                        {
                            currentCell.Width = 0;
                        }
                    }
                    else if (localName == "gridCol" && currentTable != null)
                    {
                        if (int.TryParse(xmlReader.GetAttribute("w:w"), out int gridWidth) && gridWidth > 0)
                        {
                            currentTable.GridColumnWidths.Add(gridWidth);
                        }
                    }
                    else if (localName == "tblCellSpacing" && currentTable != null && currentCell == null)
                    {
                        if (TryReadDxaWidth(xmlReader, out int cellSpacingTwips))
                        {
                            currentTable.CellSpacingTwips = cellSpacingTwips;
                        }
                    }
                    else if (localName == "tblW" && currentTable != null && currentCell == null)
                    {
                        string? type = xmlReader.GetAttribute("w:type");
                        string? widthValue = xmlReader.GetAttribute("w:w");

                        currentTable.PreferredWidthUnit = type switch
                        {
                            "dxa" => Nedev.FileConverters.DocxToDoc.Model.TableWidthUnit.Dxa,
                            "pct" => Nedev.FileConverters.DocxToDoc.Model.TableWidthUnit.Pct,
                            _ => Nedev.FileConverters.DocxToDoc.Model.TableWidthUnit.Auto
                        };

                        if ((currentTable.PreferredWidthUnit == Nedev.FileConverters.DocxToDoc.Model.TableWidthUnit.Dxa ||
                             currentTable.PreferredWidthUnit == Nedev.FileConverters.DocxToDoc.Model.TableWidthUnit.Pct) &&
                            int.TryParse(widthValue, out int preferredWidthValue))
                        {
                            currentTable.PreferredWidthValue = preferredWidthValue;
                        }
                        else
                        {
                            currentTable.PreferredWidthValue = 0;
                        }
                    }
                    else if (localName == "gridSpan" && currentCell != null)
                    {
                        if (int.TryParse(xmlReader.GetAttribute("w:val"), out int gridSpan) && gridSpan > 0)
                        {
                            currentCell.GridSpan = gridSpan;
                        }
                    }
                    else if (localName == "hMerge" && currentCell != null)
                    {
                        currentCell.HorizontalMerge = ParseHorizontalMerge(xmlReader.GetAttribute("w:val"));
                    }
                    else if (localName == "vMerge" && currentCell != null)
                    {
                        currentCell.VerticalMerge = ParseVerticalMerge(xmlReader.GetAttribute("w:val"));
                    }
                    else if (localName == "vAlign" && currentCell != null)
                    {
                        string? value = xmlReader.GetAttribute("w:val");
                        currentCell.VerticalAlignment = value switch
                        {
                            "center" => Nedev.FileConverters.DocxToDoc.Model.TableCellVerticalAlignment.Center,
                            "both" => Nedev.FileConverters.DocxToDoc.Model.TableCellVerticalAlignment.Center,
                            "bottom" => Nedev.FileConverters.DocxToDoc.Model.TableCellVerticalAlignment.Bottom,
                            _ => Nedev.FileConverters.DocxToDoc.Model.TableCellVerticalAlignment.Top
                        };
                    }
                    else if (localName == "tblCellMar" && currentTable != null && currentCell == null)
                    {
                        insideTableCellMargins = true;
                    }
                    else if (localName == "tblBorders" && currentTable != null && currentCell == null)
                    {
                        insideTableBorders = true;
                    }
                    else if (localName == "tcMar" && currentCell != null)
                    {
                        insideCellMargins = true;
                    }
                    else if (localName == "tcBorders" && currentCell != null)
                    {
                        insideCellBorders = true;
                    }
                    else if (localName == "insideH" && insideTableBorders && currentTable != null && TryReadBorderWidthTwips(xmlReader, out int insideHorizontalBorderTwips, out var insideHStyle))
                    {
                        currentTable.DefaultInsideHorizontalBorderTwips = insideHorizontalBorderTwips;
                        currentTable.DefaultInsideHorizontalBorderStyle = insideHStyle;
                    }
                    else if (localName == "insideV" && insideTableBorders && currentTable != null && TryReadBorderWidthTwips(xmlReader, out int insideVerticalBorderTwips, out var insideVStyle))
                    {
                        currentTable.DefaultInsideVerticalBorderTwips = insideVerticalBorderTwips;
                        currentTable.DefaultInsideVerticalBorderStyle = insideVStyle;
                    }
                    else if ((localName == "left" || localName == "start") && (insideCellBorders || insideTableBorders) && TryReadBorderWidthTwips(xmlReader, out int leftBorderTwips, out var leftStyle))
                    {
                        if (insideCellBorders && currentCell != null)
                        {
                            currentCell.HasLeftBorderOverride = true;
                            currentCell.BorderLeftTwips = leftBorderTwips;
                            currentCell.BorderLeftStyle = leftStyle;
                        }
                        else if (insideTableBorders && currentTable != null)
                        {
                            currentTable.DefaultBorderLeftTwips = leftBorderTwips;
                            currentTable.DefaultBorderLeftStyle = leftStyle;
                        }
                    }
                    else if ((localName == "right" || localName == "end") && (insideCellBorders || insideTableBorders) && TryReadBorderWidthTwips(xmlReader, out int rightBorderTwips, out var rightStyle))
                    {
                        if (insideCellBorders && currentCell != null)
                        {
                            currentCell.HasRightBorderOverride = true;
                            currentCell.BorderRightTwips = rightBorderTwips;
                            currentCell.BorderRightStyle = rightStyle;
                        }
                        else if (insideTableBorders && currentTable != null)
                        {
                            currentTable.DefaultBorderRightTwips = rightBorderTwips;
                            currentTable.DefaultBorderRightStyle = rightStyle;
                        }
                    }
                    else if (localName == "top" && (insideCellBorders || insideTableBorders) && TryReadBorderWidthTwips(xmlReader, out int topBorderTwips, out var topStyle))
                    {
                        if (insideCellBorders && currentCell != null)
                        {
                            currentCell.HasTopBorderOverride = true;
                            currentCell.BorderTopTwips = topBorderTwips;
                            currentCell.BorderTopStyle = topStyle;
                        }
                        else if (insideTableBorders && currentTable != null)
                        {
                            currentTable.DefaultBorderTopTwips = topBorderTwips;
                            currentTable.DefaultBorderTopStyle = topStyle;
                        }
                    }
                    else if (localName == "bottom" && (insideCellBorders || insideTableBorders) && TryReadBorderWidthTwips(xmlReader, out int bottomBorderTwips, out var bottomStyle))
                    {
                        if (insideCellBorders && currentCell != null)
                        {
                            currentCell.HasBottomBorderOverride = true;
                            currentCell.BorderBottomTwips = bottomBorderTwips;
                            currentCell.BorderBottomStyle = bottomStyle;
                        }
                        else if (insideTableBorders && currentTable != null)
                        {
                            currentTable.DefaultBorderBottomTwips = bottomBorderTwips;
                            currentTable.DefaultBorderBottomStyle = bottomStyle;
                        }
                    }
                    else if ((localName == "left" || localName == "start") && TryReadDxaWidth(xmlReader, out int leftPaddingTwips))
                    {
                        if (insideCellMargins && currentCell != null)
                        {
                            currentCell.HasLeftPaddingOverride = true;
                            currentCell.PaddingLeftTwips = leftPaddingTwips;
                        }
                        else if (insideTableCellMargins && currentTable != null)
                        {
                            currentTable.DefaultCellPaddingLeftTwips = leftPaddingTwips;
                        }
                    }
                    else if ((localName == "right" || localName == "end") && TryReadDxaWidth(xmlReader, out int rightPaddingTwips))
                    {
                        if (insideCellMargins && currentCell != null)
                        {
                            currentCell.HasRightPaddingOverride = true;
                            currentCell.PaddingRightTwips = rightPaddingTwips;
                        }
                        else if (insideTableCellMargins && currentTable != null)
                        {
                            currentTable.DefaultCellPaddingRightTwips = rightPaddingTwips;
                        }
                    }
                    else if (localName == "top" && TryReadDxaWidth(xmlReader, out int topPaddingTwips))
                    {
                        if (insideCellMargins && currentCell != null)
                        {
                            currentCell.HasTopPaddingOverride = true;
                            currentCell.PaddingTopTwips = topPaddingTwips;
                        }
                        else if (insideTableCellMargins && currentTable != null)
                        {
                            currentTable.DefaultCellPaddingTopTwips = topPaddingTwips;
                        }
                    }
                    else if (localName == "bottom" && TryReadDxaWidth(xmlReader, out int bottomPaddingTwips))
                    {
                        if (insideCellMargins && currentCell != null)
                        {
                            currentCell.HasBottomPaddingOverride = true;
                            currentCell.PaddingBottomTwips = bottomPaddingTwips;
                        }
                        else if (insideTableCellMargins && currentTable != null)
                        {
                            currentTable.DefaultCellPaddingBottomTwips = bottomPaddingTwips;
                        }
                    }
                    else if (localName == "altChunk")
                    {
                        AppendAltChunkContent(xmlReader.GetAttribute("r:id"), docModel, currentCell, textBuffer);
                    }
                    else if (localName == "p")
                    {
                        currentParagraph = new Nedev.FileConverters.DocxToDoc.Model.ParagraphModel();
                        if (currentCell != null)
                        {
                            currentCell.Content.Add(currentParagraph);
                            currentCell.Paragraphs.Add(currentParagraph);
                        }
                        else 
                        {
                            docModel.Content.Add(currentParagraph);
                            docModel.Paragraphs.Add(currentParagraph); // Backwards compatibility? No, needed for textBuffer logic maybe.
                        }
                    }
                    else if (localName == "numId" && currentParagraph != null)
                    {
                        if (int.TryParse(xmlReader.GetAttribute("w:val"), out int id))
                            currentParagraph.Properties.NumberingId = id;
                    }
                    else if (localName == "ilvl" && currentParagraph != null)
                    {
                        if (int.TryParse(xmlReader.GetAttribute("w:val"), out int lvl))
                            currentParagraph.Properties.NumberingLevel = lvl;
                    }
                    else if (localName == "sectPr")
                    {
                        currentSection = new Nedev.FileConverters.DocxToDoc.Model.SectionModel();
                        docModel.Sections.Add(currentSection);
                    }
                    else if (localName == "pgSz" && currentSection != null)
                    {
                        if (int.TryParse(xmlReader.GetAttribute("w:w"), out int w)) currentSection.PageWidth = w;
                        if (int.TryParse(xmlReader.GetAttribute("w:h"), out int h)) currentSection.PageHeight = h;
                    }
                    else if (localName == "pgMar" && currentSection != null)
                    {
                        if (int.TryParse(xmlReader.GetAttribute("w:left"), out int l)) currentSection.MarginLeft = l;
                        if (int.TryParse(xmlReader.GetAttribute("w:right"), out int r)) currentSection.MarginRight = r;
                        if (int.TryParse(xmlReader.GetAttribute("w:top"), out int t)) currentSection.MarginTop = t;
                        if (int.TryParse(xmlReader.GetAttribute("w:bottom"), out int b)) currentSection.MarginBottom = b;
                    }
                    else if (localName == "headerReference" && currentSection != null)
                    {
                        string? type = xmlReader.GetAttribute("w:type");
                        string? id = xmlReader.GetAttribute("r:id");
                        if (id != null)
                        {
                            switch (type)
                            {
                                case "default": currentSection.HeaderReference = id; break;
                                case "first": currentSection.FirstPageHeaderReference = id; break;
                                case "even": currentSection.EvenPagesHeaderReference = id; break;
                            }
                        }
                    }
                    else if (localName == "footerReference" && currentSection != null)
                    {
                        string? type = xmlReader.GetAttribute("w:type");
                        string? id = xmlReader.GetAttribute("r:id");
                        if (id != null)
                        {
                            switch (type)
                            {
                                case "default": currentSection.FooterReference = id; break;
                                case "first": currentSection.FirstPageFooterReference = id; break;
                                case "even": currentSection.EvenPagesFooterReference = id; break;
                            }
                        }
                    }
                    else if (localName == "titlePg" && currentSection != null)
                    {
                        currentSection.DifferentFirstPage = !string.Equals(xmlReader.GetAttribute("w:val"), "false", StringComparison.OrdinalIgnoreCase) &&
                                                            !string.Equals(xmlReader.GetAttribute("w:val"), "0", StringComparison.OrdinalIgnoreCase);
                    }
                    else if (localName == "pgNumType" && currentSection != null)
                    {
                        if (int.TryParse(xmlReader.GetAttribute("w:start"), out int start))
                        {
                            currentSection.StartPageNumber = start;
                            currentSection.RestartPageNumbering = true;
                        }
                    }
                    else if (localName == "jc" && currentParagraph != null)
                    {
                        string? val = xmlReader.GetAttribute("w:val");
                        currentParagraph.Properties.Alignment = val switch
                        {
                            "center" => Nedev.FileConverters.DocxToDoc.Model.ParagraphModel.Justification.Center,
                            "right" => Nedev.FileConverters.DocxToDoc.Model.ParagraphModel.Justification.Right,
                            "both" => Nedev.FileConverters.DocxToDoc.Model.ParagraphModel.Justification.Both,
                            _ => Nedev.FileConverters.DocxToDoc.Model.ParagraphModel.Justification.Left
                        };
                    }
                    else if (localName == "spacing" && currentParagraph != null)
                    {
                        if (int.TryParse(xmlReader.GetAttribute("w:before"), out int before))
                        {
                            currentParagraph.Properties.SpacingBeforeTwips = before;
                        }

                        if (int.TryParse(xmlReader.GetAttribute("w:after"), out int after))
                        {
                            currentParagraph.Properties.SpacingAfterTwips = after;
                        }

                        if (int.TryParse(xmlReader.GetAttribute("w:line"), out int line))
                        {
                            currentParagraph.Properties.LineSpacing = line;
                        }

                        string? lineRule = xmlReader.GetAttribute("w:lineRule");
                        if (!string.IsNullOrWhiteSpace(lineRule))
                        {
                            currentParagraph.Properties.LineSpacingRule = lineRule;
                        }
                    }
                    else if (localName == "ind" && currentParagraph != null)
                    {
                        if (int.TryParse(xmlReader.GetAttribute("w:left"), out int left))
                        {
                            currentParagraph.Properties.LeftIndentTwips = left;
                        }

                        if (int.TryParse(xmlReader.GetAttribute("w:right"), out int right))
                        {
                            currentParagraph.Properties.RightIndentTwips = right;
                        }

                        if (int.TryParse(xmlReader.GetAttribute("w:firstLine"), out int firstLine))
                        {
                            currentParagraph.Properties.FirstLineIndentTwips = firstLine;
                        }
                        else if (int.TryParse(xmlReader.GetAttribute("w:hanging"), out int hanging))
                        {
                            currentParagraph.Properties.FirstLineIndentTwips = -hanging;
                        }
                    }
                    else if (localName == "ins")
                    {
                        inIns = true;
                    }
                    else if (localName == "del")
                    {
                        inDel = true;
                    }
                    else if (localName == "oMath" || localName == "oMathPara")
                    {
                        inOMath = true;
                    }
                    else if (localName == "r" && currentParagraph != null)
                    {
                        currentRun = new Nedev.FileConverters.DocxToDoc.Model.RunModel();
                        if (inIns)
                        {
                            currentRun.Properties.Underline = Nedev.FileConverters.DocxToDoc.Model.UnderlineType.Single;
                        }
                        if (inDel)
                        {
                            currentRun.Properties.IsStrike = true;
                        }
                        if (inOMath)
                        {
                            currentRun.Properties.IsItalic = true;
                            currentRun.Properties.FontName = "Cambria Math";
                        }
                        currentRunBaseProperties = CloneCharacterProperties(currentRun.Properties);
                        currentParagraph.Runs.Add(currentRun);
                    }
                    else if (currentRun != null && TryApplyRunFormattingElement(xmlReader, currentRun))
                    {
                        if (currentRunBaseProperties != null)
                        {
                            TryApplyRunFormattingElement(xmlReader, currentRunBaseProperties);
                        }
                    }
                    else if (localName == "hyperlink" && currentParagraph != null)
                    {
                        string? relId = xmlReader.GetAttribute("r:id");
                        string? anchor = xmlReader.GetAttribute("w:anchor");
                        string? tooltip = xmlReader.GetAttribute("w:tooltip");

                        // Read the hyperlink content
                        var hyperlink = new Nedev.FileConverters.DocxToDoc.Model.HyperlinkModel
                        {
                            RelationshipId = relId,
                            Anchor = anchor,
                            Tooltip = tooltip
                        };

                        // Parse runs within hyperlink
                        while (xmlReader.Read())
                        {
                            if (xmlReader.NodeType == XmlNodeType.Element && xmlReader.LocalName == "r")
                            {
                                var run = new Nedev.FileConverters.DocxToDoc.Model.RunModel();
                                run.Hyperlink = hyperlink;
                                var runBaseProperties = CloneCharacterProperties(run.Properties);
                                currentParagraph.Runs.Add(run);

                                // Parse run properties and text
                                while (xmlReader.Read())
                                {
                                    if (xmlReader.NodeType == XmlNodeType.Element)
                                    {
                                        if (xmlReader.LocalName == "t" || xmlReader.LocalName == "delText" || xmlReader.LocalName == "tab" || xmlReader.LocalName == "ptab" || xmlReader.LocalName == "br" || xmlReader.LocalName == "cr" || xmlReader.LocalName == "noBreakHyphen" || xmlReader.LocalName == "softHyphen" || xmlReader.LocalName == "sym")
                                        {
                                            AppendRunTextFragment(currentParagraph, ref run, runBaseProperties, textBuffer, xmlReader, hyperlink);
                                        }
                                        else if (xmlReader.LocalName == "drawing")
                                        {
                                            var (image, textBoxText) = ParseDrawing(xmlReader);
                                            if (image != null)
                                            {
                                                run.Image = image;
                                                LoadImageData(image);
                                            }
                                            if (!string.IsNullOrEmpty(textBoxText))
                                            {
                                                AppendRunTextSegment(currentParagraph, ref run, runBaseProperties, textBuffer, textBoxText, hyperlink);
                                            }
                                        }
                                        else if (xmlReader.LocalName == "pict")
                                        {
                                            var (image, textBoxText) = ParsePict(xmlReader);
                                            if (image != null)
                                            {
                                                run.Image = image;
                                                LoadImageData(image);
                                            }
                                            if (!string.IsNullOrEmpty(textBoxText))
                                            {
                                                AppendRunTextSegment(currentParagraph, ref run, runBaseProperties, textBuffer, textBoxText, hyperlink);
                                            }
                                        }
                                        else if (TryApplyRunFormattingElement(xmlReader, run))
                                        {
                                            TryApplyRunFormattingElement(xmlReader, runBaseProperties);
                                        }
                                    }
                                    else if (xmlReader.NodeType == XmlNodeType.EndElement && xmlReader.LocalName == "r")
                                    {
                                        break;
                                    }
                                    if (xmlReader.Depth < 4) break;
                                }
                            }
                            else if (xmlReader.NodeType == XmlNodeType.EndElement && xmlReader.LocalName == "hyperlink")
                            {
                                break;
                            }
                            if (xmlReader.Depth < 3) break;
                        }

                        // Resolve URL from relationships
                        if (!string.IsNullOrEmpty(relId) && _relationships.TryGetValue(relId, out string? target))
                        {
                            hyperlink.TargetUrl = target;
                        }

                        // Skip normal processing since we handled the content
                        continue;
                    }
                    else if (localName == "fldSimple" && currentParagraph != null)
                    {
                        var field = CreateFieldModel(
                            xmlReader.GetAttribute("w:instr") ?? string.Empty,
                            xmlReader.GetAttribute("w:fldLock"),
                            xmlReader.GetAttribute("w:dirty"));

                        currentParagraph.Runs.Add(CreateFieldMarkerRun(field, isFieldBegin: true));
                        currentParagraph.Runs.Add(CreateFieldMarkerRun(field, isFieldSeparate: true));

                        if (xmlReader.IsEmptyElement)
                        {
                            currentParagraph.Runs.Add(CreateFieldMarkerRun(field, isFieldEnd: true));
                        }
                        else
                        {
                            simpleFields.Push((currentParagraph, field));
                        }
                    }
                    else if ((localName == "t" || localName == "delText" || localName == "tab" || localName == "ptab" || localName == "br" || localName == "cr" || localName == "noBreakHyphen" || localName == "softHyphen" || localName == "sym") && currentRun != null && currentParagraph != null && currentRunBaseProperties != null)
                    {
                        AppendRunTextFragment(currentParagraph, ref currentRun, currentRunBaseProperties, textBuffer, xmlReader);
                    }
                    else if (localName == "drawing" && currentRun != null)
                    {
                        // Parse inline or anchored image
                        var (image, textBoxText) = ParseDrawing(xmlReader);
                        if (image != null)
                        {
                            currentRun.Image = image;
                            // Load actual image data
                            LoadImageData(image);
                        }
                        if (!string.IsNullOrEmpty(textBoxText) && currentParagraph != null && currentRunBaseProperties != null)
                        {
                            AppendRunTextSegment(currentParagraph, ref currentRun, currentRunBaseProperties, textBuffer, textBoxText);
                        }
                    }
                    else if (localName == "pict" && currentRun != null)
                    {
                        // Parse fallback VML images (e.g., from SmartArt)
                        var (image, textBoxText) = ParsePict(xmlReader);
                        if (image != null)
                        {
                            currentRun.Image = image;
                            LoadImageData(image);
                        }
                        if (!string.IsNullOrEmpty(textBoxText) && currentParagraph != null && currentRunBaseProperties != null)
                        {
                            AppendRunTextSegment(currentParagraph, ref currentRun, currentRunBaseProperties, textBuffer, textBoxText);
                        }
                    }
                    else if (localName == "fldChar" && currentRun != null)
                    {
                        // Field character (begin/separate/end)
                        string? fldCharType = xmlReader.GetAttribute("w:fldCharType");
                        string? fldLock = xmlReader.GetAttribute("w:fldLock");
                        string? fldDirty = xmlReader.GetAttribute("w:fldDirty");

                        switch (fldCharType)
                        {
                            case "begin":
                                currentRun.IsFieldBegin = true;
                                currentRun.Field = CreateFieldModel(string.Empty, fldLock, fldDirty);
                                openFields.Push(currentRun.Field);
                                break;
                            case "separate":
                                currentRun.IsFieldSeparate = true;
                                if (openFields.Count > 0)
                                {
                                    currentRun.Field = openFields.Peek();
                                }
                                break;
                            case "end":
                                currentRun.IsFieldEnd = true;
                                if (openFields.Count > 0)
                                {
                                    currentRun.Field = openFields.Pop();
                                }
                                break;
                        }
                    }
                    else if (localName == "instrText" && currentRun != null)
                    {
                        // Field instruction text
                        string instruction = ReadCurrentElementString(xmlReader);
                        if (openFields.Count > 0)
                        {
                            var activeField = openFields.Peek();
                            currentRun.Field = activeField;
                            activeField.Instruction += instruction;
                            activeField.Type = ParseFieldType(activeField.Instruction);
                        }
                    }
                }
                else if (xmlReader.NodeType == XmlNodeType.EndElement)
                {
                    string localName = xmlReader.LocalName;
                    if (localName == "ins") inIns = false;
                    else if (localName == "del") inDel = false;
                    else if (localName == "oMath" || localName == "oMathPara") inOMath = false;

                    if (localName == "tbl")
                    {
                        currentTable = null;
                        insideTableCellMargins = false;
                    }
                    else if (localName == "tr") currentRow = null;
                    else if (localName == "tc")
                    {
                        if (currentCell != null)
                        {
                            if (currentCell.HorizontalMerge == Nedev.FileConverters.DocxToDoc.Model.TableCellHorizontalMerge.Continue &&
                                currentRow != null &&
                                currentRowHorizontalMergeAnchor != null)
                            {
                                int mergedWidth = 0;
                                if (currentTable != null)
                                {
                                    mergedWidth = ResolveGridWidth(currentTable.GridColumnWidths, currentRowGridColumnIndex, currentCell.GridSpan);
                                }
                                if (mergedWidth <= 0 && currentCell.WidthUnit == Nedev.FileConverters.DocxToDoc.Model.TableWidthUnit.Dxa)
                                {
                                    mergedWidth = currentCell.Width;
                                }

                                currentRowHorizontalMergeAnchor.GridSpan += Math.Max(1, currentCell.GridSpan);
                                if (mergedWidth > 0)
                                {
                                    currentRowHorizontalMergeAnchor.Width = Math.Max(0, currentRowHorizontalMergeAnchor.Width) + mergedWidth;
                                }
                                else if (currentCell.WidthUnit == Nedev.FileConverters.DocxToDoc.Model.TableWidthUnit.Pct &&
                                         currentRowHorizontalMergeAnchor.WidthUnit == Nedev.FileConverters.DocxToDoc.Model.TableWidthUnit.Pct &&
                                         currentCell.Width > 0)
                                {
                                    currentRowHorizontalMergeAnchor.Width = Math.Max(0, currentRowHorizontalMergeAnchor.Width) + currentCell.Width;
                                }

                                currentRowGridColumnIndex += Math.Max(1, currentCell.GridSpan);
                                if (currentRow.Cells.Count > 0 && ReferenceEquals(currentRow.Cells[currentRow.Cells.Count - 1], currentCell))
                                {
                                    currentRow.Cells.RemoveAt(currentRow.Cells.Count - 1);
                                }
                            }
                            else
                            {
                                if (currentCell.Width <= 0 && currentTable != null)
                                {
                                    currentCell.Width = ResolveGridWidth(currentTable.GridColumnWidths, currentRowGridColumnIndex, currentCell.GridSpan);
                                }

                                currentRowGridColumnIndex += Math.Max(1, currentCell.GridSpan);
                            }

                            if (currentCell.HorizontalMerge == Nedev.FileConverters.DocxToDoc.Model.TableCellHorizontalMerge.Restart)
                            {
                                currentRowHorizontalMergeAnchor = currentCell;
                            }
                            else if (currentCell.HorizontalMerge == Nedev.FileConverters.DocxToDoc.Model.TableCellHorizontalMerge.None)
                            {
                                currentRowHorizontalMergeAnchor = null;
                            }
                        }

                        currentCell = null;
                        insideCellMargins = false;
                    }
                    else if (localName == "tblCellMar")
                    {
                        insideTableCellMargins = false;
                    }
                    else if (localName == "tblBorders")
                    {
                        insideTableBorders = false;
                    }
                    else if (localName == "tcMar")
                    {
                        insideCellMargins = false;
                    }
                    else if (localName == "tcBorders")
                    {
                        insideCellBorders = false;
                    }
                    else if (localName == "p")
                    {
                        textBuffer.Append('\r');
                        currentParagraph = null;
                    }
                    else if (localName == "r")
                    {
                        currentRun = null;
                        currentRunBaseProperties = null;
                    }
                    else if (localName == "fldSimple" && simpleFields.Count > 0)
                    {
                        var (paragraph, field) = simpleFields.Pop();
                        paragraph.Runs.Add(CreateFieldMarkerRun(field, isFieldEnd: true));
                    }
                }
            }

            docModel.TextBuffer = textBuffer.ToString();

            // Parse bookmarks after document is fully read
            ParseBookmarks(docModel);
            ParseFootnoteReferences(docModel);
            ParseEndnoteReferences(docModel);
            ParseCommentRanges(docModel);
            ParseSectionHeaderFooterStories(docModel);

            return docModel;
        }

        private void ParseSectionHeaderFooterStories(Nedev.FileConverters.DocxToDoc.Model.DocumentModel docModel)
        {
            foreach (var section in docModel.Sections)
            {
                section.DefaultHeaderStory = ReadHeaderFooterStoryModel(section.HeaderReference);
                section.DefaultFooterStory = ReadHeaderFooterStoryModel(section.FooterReference);
                section.FirstPageHeaderStory = ReadHeaderFooterStoryModel(section.FirstPageHeaderReference);
                section.FirstPageFooterStory = ReadHeaderFooterStoryModel(section.FirstPageFooterReference);
                section.EvenPagesHeaderStory = ReadHeaderFooterStoryModel(section.EvenPagesHeaderReference);
                section.EvenPagesFooterStory = ReadHeaderFooterStoryModel(section.EvenPagesFooterReference);

                section.DefaultHeaderText = section.DefaultHeaderStory?.Text;
                section.DefaultFooterText = section.DefaultFooterStory?.Text;
                section.FirstPageHeaderText = section.FirstPageHeaderStory?.Text;
                section.FirstPageFooterText = section.FirstPageFooterStory?.Text;
                section.EvenPagesHeaderText = section.EvenPagesHeaderStory?.Text;
                section.EvenPagesFooterText = section.EvenPagesFooterStory?.Text;

                section.HeaderText = ReadHeaderFooterSummaryStory(
                    docModel.DifferentOddAndEvenPages,
                    section.DefaultHeaderStory,
                    section.FirstPageHeaderStory,
                    section.EvenPagesHeaderStory);

                section.FooterText = ReadHeaderFooterSummaryStory(
                    docModel.DifferentOddAndEvenPages,
                    section.DefaultFooterStory,
                    section.FirstPageFooterStory,
                    section.EvenPagesFooterStory);
            }
        }

        private string? ReadHeaderFooterSummaryStory(
            bool includeEvenPagesStory,
            Nedev.FileConverters.DocxToDoc.Model.HeaderFooterStoryModel? defaultStory,
            Nedev.FileConverters.DocxToDoc.Model.HeaderFooterStoryModel? firstStory,
            Nedev.FileConverters.DocxToDoc.Model.HeaderFooterStoryModel? evenStory)
        {
            if (defaultStory != null)
            {
                return defaultStory.Text;
            }

            if (firstStory != null)
            {
                return firstStory.Text;
            }

            return includeEvenPagesStory && evenStory != null
                ? evenStory.Text
                : null;
        }

        private static void ParseDocumentSettings(
            Stream settingsStream,
            Nedev.FileConverters.DocxToDoc.Model.DocumentModel docModel)
        {
            using var reader = XmlReader.Create(settingsStream, new XmlReaderSettings { IgnoreWhitespace = true });
            while (reader.Read())
            {
                if (reader.NodeType != XmlNodeType.Element || reader.LocalName != "evenAndOddHeaders")
                {
                    continue;
                }

                string? value = reader.GetAttribute("w:val");
                docModel.DifferentOddAndEvenPages =
                    !string.Equals(value, "false", StringComparison.OrdinalIgnoreCase) &&
                    !string.Equals(value, "0", StringComparison.OrdinalIgnoreCase);
                return;
            }
        }

        private Nedev.FileConverters.DocxToDoc.Model.HeaderFooterStoryModel? ReadHeaderFooterStoryModel(string? relationshipId)
        {
            if (string.IsNullOrEmpty(relationshipId) || !_relationships.TryGetValue(relationshipId, out string? target) || string.IsNullOrEmpty(target))
            {
                return null;
            }

            string entryPath = ResolvePartTargetPath("word/document.xml", target);

            var entry = _archive.GetEntry(entryPath);
            if (entry == null)
            {
                return null;
            }

            Dictionary<string, string> storyRelationships = LoadRelationshipsForPart(entryPath);
            using var stream = entry.Open();
            var story = ReadHeaderFooterStory(stream, entryPath, storyRelationships);
            ResolveHeaderFooterStoryHyperlinks(story, storyRelationships);
            return story;
        }

        private Nedev.FileConverters.DocxToDoc.Model.HeaderFooterStoryModel ReadHeaderFooterStory(
            Stream storyStream,
            string storyEntryPath,
            IReadOnlyDictionary<string, string> storyRelationships)
        {
            using var reader = XmlReader.Create(storyStream, new XmlReaderSettings { IgnoreWhitespace = true });

            var story = new Nedev.FileConverters.DocxToDoc.Model.HeaderFooterStoryModel();
            var storyText = new StringBuilder();
            int storyParagraphCount = 0;
            bool pendingSummaryWhitespaceCollapse = false;
            var openFields = new Stack<Nedev.FileConverters.DocxToDoc.Model.FieldModel>();
            var simpleFields = new Stack<(Nedev.FileConverters.DocxToDoc.Model.ParagraphModel Paragraph, Nedev.FileConverters.DocxToDoc.Model.FieldModel Field)>();
            Nedev.FileConverters.DocxToDoc.Model.ParagraphModel? currentParagraph = null;
            Nedev.FileConverters.DocxToDoc.Model.RunModel? currentRun = null;
            Nedev.FileConverters.DocxToDoc.Model.RunModel.CharacterProperties? currentRunBaseProperties = null;
            Nedev.FileConverters.DocxToDoc.Model.TableModel? currentTable = null;
            Nedev.FileConverters.DocxToDoc.Model.TableRowModel? currentRow = null;
            Nedev.FileConverters.DocxToDoc.Model.TableCellModel? currentCell = null;
            Nedev.FileConverters.DocxToDoc.Model.TableCellModel? currentRowHorizontalMergeAnchor = null;
            int currentRowGridColumnIndex = 0;
            bool insideTableCellMargins = false;
            bool insideCellMargins = false;
            bool insideTableBorders = false;
            bool insideCellBorders = false;

            while (reader.Read())
            {
                if (reader.NodeType == XmlNodeType.Element)
                {
                    string localName = reader.LocalName;
                    if (localName == "tbl" && currentCell == null)
                    {
                        currentTable = new Nedev.FileConverters.DocxToDoc.Model.TableModel();
                        story.Content.Add(currentTable);
                    }
                    else if (localName == "tr" && currentTable != null)
                    {
                        currentRow = new Nedev.FileConverters.DocxToDoc.Model.TableRowModel();
                        currentTable.Rows.Add(currentRow);
                        currentRowHorizontalMergeAnchor = null;
                        currentRowGridColumnIndex = 0;
                    }
                    else if (localName == "tc" && currentRow != null)
                    {
                        currentCell = new Nedev.FileConverters.DocxToDoc.Model.TableCellModel();
                        currentRow.Cells.Add(currentCell);
                    }
                    else if (localName == "trHeight" && currentRow != null)
                    {
                        if (int.TryParse(reader.GetAttribute("w:val"), out int rowHeightTwips))
                        {
                            currentRow.HeightTwips = rowHeightTwips;
                        }

                        string? heightRule = reader.GetAttribute("w:hRule");
                        currentRow.HeightRule = heightRule switch
                        {
                            "exact" => Nedev.FileConverters.DocxToDoc.Model.TableRowHeightRule.Exact,
                            "atLeast" => Nedev.FileConverters.DocxToDoc.Model.TableRowHeightRule.AtLeast,
                            _ => Nedev.FileConverters.DocxToDoc.Model.TableRowHeightRule.Auto
                        };
                    }
                    else if (localName == "tblHeader" && currentRow != null)
                    {
                        currentRow.IsHeader = !IsFalseValue(reader.GetAttribute("w:val"));
                    }
                    else if (localName == "cantSplit" && currentRow != null)
                    {
                        currentRow.CannotSplit = !IsFalseValue(reader.GetAttribute("w:val"));
                    }
                    else if (localName == "tcW" && currentCell != null)
                    {
                        string? type = reader.GetAttribute("w:type");
                        currentCell.WidthUnit = type switch
                        {
                            "pct" => Nedev.FileConverters.DocxToDoc.Model.TableWidthUnit.Pct,
                            "auto" => Nedev.FileConverters.DocxToDoc.Model.TableWidthUnit.Auto,
                            "nil" => Nedev.FileConverters.DocxToDoc.Model.TableWidthUnit.Auto,
                            _ => Nedev.FileConverters.DocxToDoc.Model.TableWidthUnit.Dxa
                        };

                        if (int.TryParse(reader.GetAttribute("w:w"), out int width) &&
                            (currentCell.WidthUnit == Nedev.FileConverters.DocxToDoc.Model.TableWidthUnit.Dxa ||
                             currentCell.WidthUnit == Nedev.FileConverters.DocxToDoc.Model.TableWidthUnit.Pct))
                        {
                            currentCell.Width = width;
                        }
                        else
                        {
                            currentCell.Width = 0;
                        }
                    }
                    else if (localName == "gridCol" && currentTable != null)
                    {
                        if (int.TryParse(reader.GetAttribute("w:w"), out int gridWidth) && gridWidth > 0)
                        {
                            currentTable.GridColumnWidths.Add(gridWidth);
                        }
                    }
                    else if (localName == "tblCellSpacing" && currentTable != null && currentCell == null)
                    {
                        if (TryReadDxaWidth(reader, out int cellSpacingTwips))
                        {
                            currentTable.CellSpacingTwips = cellSpacingTwips;
                        }
                    }
                    else if (localName == "tblW" && currentTable != null && currentCell == null)
                    {
                        string? type = reader.GetAttribute("w:type");
                        string? widthValue = reader.GetAttribute("w:w");

                        currentTable.PreferredWidthUnit = type switch
                        {
                            "dxa" => Nedev.FileConverters.DocxToDoc.Model.TableWidthUnit.Dxa,
                            "pct" => Nedev.FileConverters.DocxToDoc.Model.TableWidthUnit.Pct,
                            _ => Nedev.FileConverters.DocxToDoc.Model.TableWidthUnit.Auto
                        };

                        if ((currentTable.PreferredWidthUnit == Nedev.FileConverters.DocxToDoc.Model.TableWidthUnit.Dxa ||
                             currentTable.PreferredWidthUnit == Nedev.FileConverters.DocxToDoc.Model.TableWidthUnit.Pct) &&
                            int.TryParse(widthValue, out int preferredWidthValue))
                        {
                            currentTable.PreferredWidthValue = preferredWidthValue;
                        }
                        else
                        {
                            currentTable.PreferredWidthValue = 0;
                        }
                    }
                    else if (localName == "gridSpan" && currentCell != null)
                    {
                        if (int.TryParse(reader.GetAttribute("w:val"), out int gridSpan) && gridSpan > 0)
                        {
                            currentCell.GridSpan = gridSpan;
                        }
                    }
                    else if (localName == "hMerge" && currentCell != null)
                    {
                        currentCell.HorizontalMerge = ParseHorizontalMerge(reader.GetAttribute("w:val"));
                    }
                    else if (localName == "vMerge" && currentCell != null)
                    {
                        currentCell.VerticalMerge = ParseVerticalMerge(reader.GetAttribute("w:val"));
                    }
                    else if (localName == "vAlign" && currentCell != null)
                    {
                        string? value = reader.GetAttribute("w:val");
                        currentCell.VerticalAlignment = value switch
                        {
                            "center" => Nedev.FileConverters.DocxToDoc.Model.TableCellVerticalAlignment.Center,
                            "both" => Nedev.FileConverters.DocxToDoc.Model.TableCellVerticalAlignment.Center,
                            "bottom" => Nedev.FileConverters.DocxToDoc.Model.TableCellVerticalAlignment.Bottom,
                            _ => Nedev.FileConverters.DocxToDoc.Model.TableCellVerticalAlignment.Top
                        };
                    }
                    else if (localName == "tblCellMar" && currentTable != null && currentCell == null)
                    {
                        insideTableCellMargins = true;
                    }
                    else if (localName == "tblBorders" && currentTable != null && currentCell == null)
                    {
                        insideTableBorders = true;
                    }
                    else if (localName == "tcMar" && currentCell != null)
                    {
                        insideCellMargins = true;
                    }
                    else if (localName == "tcBorders" && currentCell != null)
                    {
                        insideCellBorders = true;
                    }
                    else if (localName == "insideH" && insideTableBorders && currentTable != null && TryReadBorderWidthTwips(reader, out int insideHorizontalBorderTwips, out var insideHStyle))
                    {
                        currentTable.DefaultInsideHorizontalBorderTwips = insideHorizontalBorderTwips;
                        currentTable.DefaultInsideHorizontalBorderStyle = insideHStyle;
                    }
                    else if (localName == "insideV" && insideTableBorders && currentTable != null && TryReadBorderWidthTwips(reader, out int insideVerticalBorderTwips, out var insideVStyle))
                    {
                        currentTable.DefaultInsideVerticalBorderTwips = insideVerticalBorderTwips;
                        currentTable.DefaultInsideVerticalBorderStyle = insideVStyle;
                    }
                    else if ((localName == "left" || localName == "start") && (insideCellBorders || insideTableBorders) && TryReadBorderWidthTwips(reader, out int leftBorderTwips, out var leftStyle))
                    {
                        if (insideCellBorders && currentCell != null)
                        {
                            currentCell.HasLeftBorderOverride = true;
                            currentCell.BorderLeftTwips = leftBorderTwips;
                            currentCell.BorderLeftStyle = leftStyle;
                        }
                        else if (insideTableBorders && currentTable != null)
                        {
                            currentTable.DefaultBorderLeftTwips = leftBorderTwips;
                            currentTable.DefaultBorderLeftStyle = leftStyle;
                        }
                    }
                    else if ((localName == "right" || localName == "end") && (insideCellBorders || insideTableBorders) && TryReadBorderWidthTwips(reader, out int rightBorderTwips, out var rightStyle))
                    {
                        if (insideCellBorders && currentCell != null)
                        {
                            currentCell.HasRightBorderOverride = true;
                            currentCell.BorderRightTwips = rightBorderTwips;
                            currentCell.BorderRightStyle = rightStyle;
                        }
                        else if (insideTableBorders && currentTable != null)
                        {
                            currentTable.DefaultBorderRightTwips = rightBorderTwips;
                            currentTable.DefaultBorderRightStyle = rightStyle;
                        }
                    }
                    else if (localName == "top" && (insideCellBorders || insideTableBorders) && TryReadBorderWidthTwips(reader, out int topBorderTwips, out var topStyle))
                    {
                        if (insideCellBorders && currentCell != null)
                        {
                            currentCell.HasTopBorderOverride = true;
                            currentCell.BorderTopTwips = topBorderTwips;
                            currentCell.BorderTopStyle = topStyle;
                        }
                        else if (insideTableBorders && currentTable != null)
                        {
                            currentTable.DefaultBorderTopTwips = topBorderTwips;
                            currentTable.DefaultBorderTopStyle = topStyle;
                        }
                    }
                    else if (localName == "bottom" && (insideCellBorders || insideTableBorders) && TryReadBorderWidthTwips(reader, out int bottomBorderTwips, out var bottomStyle))
                    {
                        if (insideCellBorders && currentCell != null)
                        {
                            currentCell.HasBottomBorderOverride = true;
                            currentCell.BorderBottomTwips = bottomBorderTwips;
                            currentCell.BorderBottomStyle = bottomStyle;
                        }
                        else if (insideTableBorders && currentTable != null)
                        {
                            currentTable.DefaultBorderBottomTwips = bottomBorderTwips;
                            currentTable.DefaultBorderBottomStyle = bottomStyle;
                        }
                    }
                    else if ((localName == "left" || localName == "start") && TryReadDxaWidth(reader, out int leftPaddingTwips))
                    {
                        if (insideCellMargins && currentCell != null)
                        {
                            currentCell.HasLeftPaddingOverride = true;
                            currentCell.PaddingLeftTwips = leftPaddingTwips;
                        }
                        else if (insideTableCellMargins && currentTable != null)
                        {
                            currentTable.DefaultCellPaddingLeftTwips = leftPaddingTwips;
                        }
                    }
                    else if ((localName == "right" || localName == "end") && TryReadDxaWidth(reader, out int rightPaddingTwips))
                    {
                        if (insideCellMargins && currentCell != null)
                        {
                            currentCell.HasRightPaddingOverride = true;
                            currentCell.PaddingRightTwips = rightPaddingTwips;
                        }
                        else if (insideTableCellMargins && currentTable != null)
                        {
                            currentTable.DefaultCellPaddingRightTwips = rightPaddingTwips;
                        }
                    }
                    else if (localName == "top" && TryReadDxaWidth(reader, out int topPaddingTwips))
                    {
                        if (insideCellMargins && currentCell != null)
                        {
                            currentCell.HasTopPaddingOverride = true;
                            currentCell.PaddingTopTwips = topPaddingTwips;
                        }
                        else if (insideTableCellMargins && currentTable != null)
                        {
                            currentTable.DefaultCellPaddingTopTwips = topPaddingTwips;
                        }
                    }
                    else if (localName == "bottom" && TryReadDxaWidth(reader, out int bottomPaddingTwips))
                    {
                        if (insideCellMargins && currentCell != null)
                        {
                            currentCell.HasBottomPaddingOverride = true;
                            currentCell.PaddingBottomTwips = bottomPaddingTwips;
                        }
                        else if (insideTableCellMargins && currentTable != null)
                        {
                            currentTable.DefaultCellPaddingBottomTwips = bottomPaddingTwips;
                        }
                    }
                    else if (localName == "p")
                    {
                        if (storyParagraphCount > 0)
                        {
                            storyText.Append('\r');
                        }

                        currentParagraph = new Nedev.FileConverters.DocxToDoc.Model.ParagraphModel();
                        storyParagraphCount++;
                        if (currentCell != null)
                        {
                            currentCell.Content.Add(currentParagraph);
                            currentCell.Paragraphs.Add(currentParagraph);
                        }
                        else
                        {
                            story.Paragraphs.Add(currentParagraph);
                            story.Content.Add(currentParagraph);
                        }
                        currentRun = new Nedev.FileConverters.DocxToDoc.Model.RunModel();
                        currentParagraph.Runs.Add(currentRun);
                        currentRunBaseProperties = CloneCharacterProperties(currentRun.Properties);
                    }
                    else if (localName == "altChunk")
                    {
                        AppendHeaderFooterAltChunk(reader.GetAttribute("r:id"));
                    }
                    else if (localName == "r" && currentParagraph != null)
                    {
                        currentRun = new Nedev.FileConverters.DocxToDoc.Model.RunModel();
                        currentParagraph.Runs.Add(currentRun);
                        currentRunBaseProperties = CloneCharacterProperties(currentRun.Properties);
                    }
                    else if (currentRun != null && TryApplyRunFormattingElement(reader, currentRun))
                    {
                        if (currentRunBaseProperties != null)
                        {
                            TryApplyRunFormattingElement(reader, currentRunBaseProperties);
                        }
                    }
                    else if (localName == "hyperlink" && currentParagraph != null)
                    {
                        string? relId = reader.GetAttribute("r:id");
                        string? anchor = reader.GetAttribute("w:anchor");
                        string? tooltip = reader.GetAttribute("w:tooltip");

                        var hyperlink = new Nedev.FileConverters.DocxToDoc.Model.HyperlinkModel
                        {
                            RelationshipId = relId,
                            Anchor = anchor,
                            Tooltip = tooltip
                        };

                        while (reader.Read())
                        {
                            if (reader.NodeType == XmlNodeType.Element && reader.LocalName == "r")
                            {
                                var run = new Nedev.FileConverters.DocxToDoc.Model.RunModel();
                                run.Hyperlink = hyperlink;
                                var runBaseProperties = CloneCharacterProperties(run.Properties);
                                currentParagraph.Runs.Add(run);

                                while (reader.Read())
                                {
                                    if (reader.NodeType == XmlNodeType.Element)
                                    {
                                        if (reader.LocalName == "t" || reader.LocalName == "delText" || reader.LocalName == "tab" || reader.LocalName == "ptab" || reader.LocalName == "br" || reader.LocalName == "cr" || reader.LocalName == "noBreakHyphen" || reader.LocalName == "softHyphen" || reader.LocalName == "sym")
                                        {
                                            AppendHeaderFooterRunTextFragment(currentParagraph, ref run, runBaseProperties, reader, hyperlink);
                                        }
                                        else if (reader.LocalName == "drawing")
                                        {
                                            var (image, textBoxText) = ParseDrawing(reader);
                                            if (image != null)
                                            {
                                                run.Image = image;
                                                LoadImageData(image, storyRelationships, storyEntryPath);
                                                pendingSummaryWhitespaceCollapse = true;
                                            }
                                            if (!string.IsNullOrEmpty(textBoxText))
                                            {
                                                AppendRunTextSegment(currentParagraph, ref run, runBaseProperties, storyText, textBoxText, hyperlink);
                                            }
                                        }
                                        else if (reader.LocalName == "pict")
                                        {
                                            var (image, textBoxText) = ParsePict(reader);
                                            if (image != null)
                                            {
                                                run.Image = image;
                                                LoadImageData(image, storyRelationships, storyEntryPath);
                                                pendingSummaryWhitespaceCollapse = true;
                                            }
                                            if (!string.IsNullOrEmpty(textBoxText))
                                            {
                                                AppendRunTextSegment(currentParagraph, ref run, runBaseProperties, storyText, textBoxText, hyperlink);
                                            }
                                        }
                                        else if (TryApplyRunFormattingElement(reader, run))
                                        {
                                            TryApplyRunFormattingElement(reader, runBaseProperties);
                                        }
                                    }
                                    else if (reader.NodeType == XmlNodeType.EndElement && reader.LocalName == "r")
                                    {
                                        break;
                                    }

                                    if (reader.Depth < 4)
                                    {
                                        break;
                                    }
                                }
                            }
                            else if (reader.NodeType == XmlNodeType.EndElement && reader.LocalName == "hyperlink")
                            {
                                break;
                            }

                            if (reader.Depth < 3)
                            {
                                break;
                            }
                        }
                    }
                    else if (localName == "fldSimple" && currentParagraph != null)
                    {
                        var field = CreateFieldModel(
                            reader.GetAttribute("w:instr") ?? string.Empty,
                            reader.GetAttribute("w:fldLock"),
                            reader.GetAttribute("w:dirty"));

                        currentParagraph.Runs.Add(CreateFieldMarkerRun(field, isFieldBegin: true));
                        currentParagraph.Runs.Add(CreateFieldMarkerRun(field, isFieldSeparate: true));

                        if (reader.IsEmptyElement)
                        {
                            currentParagraph.Runs.Add(CreateFieldMarkerRun(field, isFieldEnd: true));
                        }
                        else
                        {
                            simpleFields.Push((currentParagraph, field));
                        }
                    }
                    else if ((localName == "t" || localName == "delText" || localName == "tab" || localName == "ptab" || localName == "br" || localName == "cr" || localName == "noBreakHyphen" || localName == "softHyphen" || localName == "sym") &&
                             currentRun != null &&
                             currentParagraph != null)
                    {
                        AppendHeaderFooterRunTextFragment(
                            currentParagraph,
                            ref currentRun,
                            currentRunBaseProperties ?? currentRun.Properties,
                            reader);
                    }
                    else if (localName == "drawing" && currentRun != null)
                    {
                        var (image, textBoxText) = ParseDrawing(reader);
                        if (image != null)
                        {
                            currentRun.Image = image;
                            LoadImageData(image, storyRelationships, storyEntryPath);
                            pendingSummaryWhitespaceCollapse = true;
                        }
                        if (!string.IsNullOrEmpty(textBoxText) && currentParagraph != null && currentRunBaseProperties != null)
                        {
                            AppendRunTextSegment(currentParagraph, ref currentRun, currentRunBaseProperties, storyText, textBoxText);
                        }
                    }
                    else if (localName == "pict" && currentRun != null)
                    {
                        var (image, textBoxText) = ParsePict(reader);
                        if (image != null)
                        {
                            currentRun.Image = image;
                            LoadImageData(image, storyRelationships, storyEntryPath);
                            pendingSummaryWhitespaceCollapse = true;
                        }
                        if (!string.IsNullOrEmpty(textBoxText) && currentParagraph != null && currentRunBaseProperties != null)
                        {
                            AppendRunTextSegment(currentParagraph, ref currentRun, currentRunBaseProperties, storyText, textBoxText);
                        }
                    }
                    else if (localName == "fldChar" && currentRun != null)
                    {
                        string? fldCharType = reader.GetAttribute("w:fldCharType");
                        string? fldLock = reader.GetAttribute("w:fldLock");
                        string? fldDirty = reader.GetAttribute("w:dirty");

                        switch (fldCharType)
                        {
                            case "begin":
                                currentRun.IsFieldBegin = true;
                                currentRun.Field = CreateFieldModel(string.Empty, fldLock, fldDirty);
                                openFields.Push(currentRun.Field);
                                break;
                            case "separate":
                                currentRun.IsFieldSeparate = true;
                                if (openFields.Count > 0)
                                {
                                    currentRun.Field = openFields.Peek();
                                }
                                break;
                            case "end":
                                currentRun.IsFieldEnd = true;
                                if (openFields.Count > 0)
                                {
                                    currentRun.Field = openFields.Pop();
                                }
                                break;
                        }
                    }
                    else if (localName == "instrText" && currentRun != null && openFields.Count > 0)
                    {
                        string instruction = ReadCurrentElementString(reader);
                        var activeField = openFields.Peek();
                        currentRun.Field = activeField;
                        activeField.Instruction += instruction;
                        activeField.Type = ParseFieldType(activeField.Instruction);
                    }
                }
                else if (reader.NodeType == XmlNodeType.EndElement)
                {
                    if (reader.LocalName == "r")
                    {
                        currentRunBaseProperties = null;
                    }
                    else if (reader.LocalName == "fldSimple" && simpleFields.Count > 0)
                    {
                        var (paragraph, field) = simpleFields.Pop();
                        paragraph.Runs.Add(CreateFieldMarkerRun(field, isFieldEnd: true));
                    }
                    else if (reader.LocalName == "tc")
                    {
                        if (currentCell != null)
                        {
                            if (currentCell.HorizontalMerge == Nedev.FileConverters.DocxToDoc.Model.TableCellHorizontalMerge.Continue &&
                                currentRow != null &&
                                currentRowHorizontalMergeAnchor != null)
                            {
                                int mergedWidth = 0;
                                if (currentTable != null)
                                {
                                    mergedWidth = ResolveGridWidth(currentTable.GridColumnWidths, currentRowGridColumnIndex, currentCell.GridSpan);
                                }
                                if (mergedWidth <= 0 && currentCell.WidthUnit == Nedev.FileConverters.DocxToDoc.Model.TableWidthUnit.Dxa)
                                {
                                    mergedWidth = currentCell.Width;
                                }

                                currentRowHorizontalMergeAnchor.GridSpan += Math.Max(1, currentCell.GridSpan);
                                if (mergedWidth > 0)
                                {
                                    currentRowHorizontalMergeAnchor.Width = Math.Max(0, currentRowHorizontalMergeAnchor.Width) + mergedWidth;
                                }
                                else if (currentCell.WidthUnit == Nedev.FileConverters.DocxToDoc.Model.TableWidthUnit.Pct &&
                                         currentRowHorizontalMergeAnchor.WidthUnit == Nedev.FileConverters.DocxToDoc.Model.TableWidthUnit.Pct &&
                                         currentCell.Width > 0)
                                {
                                    currentRowHorizontalMergeAnchor.Width = Math.Max(0, currentRowHorizontalMergeAnchor.Width) + currentCell.Width;
                                }

                                currentRowGridColumnIndex += Math.Max(1, currentCell.GridSpan);
                                if (currentRow.Cells.Count > 0 && ReferenceEquals(currentRow.Cells[currentRow.Cells.Count - 1], currentCell))
                                {
                                    currentRow.Cells.RemoveAt(currentRow.Cells.Count - 1);
                                }
                            }
                            else
                            {
                                if (currentCell.Width <= 0 && currentTable != null)
                                {
                                    currentCell.Width = ResolveGridWidth(currentTable.GridColumnWidths, currentRowGridColumnIndex, currentCell.GridSpan);
                                }

                                currentRowGridColumnIndex += Math.Max(1, currentCell.GridSpan);
                            }

                            if (currentCell.HorizontalMerge == Nedev.FileConverters.DocxToDoc.Model.TableCellHorizontalMerge.Restart)
                            {
                                currentRowHorizontalMergeAnchor = currentCell;
                            }
                            else if (currentCell.HorizontalMerge == Nedev.FileConverters.DocxToDoc.Model.TableCellHorizontalMerge.None)
                            {
                                currentRowHorizontalMergeAnchor = null;
                            }
                        }

                        currentCell = null;
                    }
                    else if (reader.LocalName == "tr")
                    {
                        currentRow = null;
                    }
                    else if (reader.LocalName == "tbl")
                    {
                        currentTable = null;
                    }
                    else if (reader.LocalName == "tblCellMar")
                    {
                        insideTableCellMargins = false;
                    }
                    else if (reader.LocalName == "tblBorders")
                    {
                        insideTableBorders = false;
                    }
                    else if (reader.LocalName == "tcMar")
                    {
                        insideCellMargins = false;
                    }
                    else if (reader.LocalName == "tcBorders")
                    {
                        insideCellBorders = false;
                    }
                }
            }

            story.Text = storyText.ToString();
            return story;

            void AppendHeaderFooterAltChunk(string? relationshipId)
            {
                foreach (var block in ReadAltChunkBlocks(relationshipId, storyRelationships, storyEntryPath))
                {
                    if (currentCell != null)
                    {
                        AppendAltChunkBlockToCell(currentCell, block);
                        foreach (var paragraph in EnumerateAltChunkBlockParagraphs(block))
                        {
                            AppendAltChunkParagraph(paragraph, addToStoryContent: false);
                        }
                    }
                    else
                    {
                        foreach (var paragraph in EnumerateAltChunkBlockParagraphs(block))
                        {
                            AppendAltChunkParagraph(paragraph, addToStoryContent: false);
                        }

                        story.Content.Add(block);
                        if (block is Nedev.FileConverters.DocxToDoc.Model.ParagraphModel blockParagraph)
                        {
                            story.Paragraphs.Add(blockParagraph);
                        }
                    }
                }

                currentParagraph = null;
                currentRun = null;
                currentRunBaseProperties = null;
                pendingSummaryWhitespaceCollapse = false;
            }

            void AppendAltChunkParagraph(
                Nedev.FileConverters.DocxToDoc.Model.ParagraphModel paragraph,
                bool addToStoryContent)
            {
                if (storyParagraphCount > 0)
                {
                    storyText.Append('\r');
                }

                storyParagraphCount++;
                if (addToStoryContent)
                {
                    story.Paragraphs.Add(paragraph);
                    story.Content.Add(paragraph);
                }

                storyText.Append(GetParagraphVisibleText(paragraph));
            }

            void AppendHeaderFooterRunTextFragment(
                Nedev.FileConverters.DocxToDoc.Model.ParagraphModel paragraph,
                ref Nedev.FileConverters.DocxToDoc.Model.RunModel run,
                Nedev.FileConverters.DocxToDoc.Model.RunModel.CharacterProperties baseProperties,
                XmlReader xmlReader,
                Nedev.FileConverters.DocxToDoc.Model.HyperlinkModel? hyperlink = null)
            {
                string text = ReadRunTextFragment(xmlReader);
                if (string.IsNullOrEmpty(text))
                {
                    return;
                }

                if (pendingSummaryWhitespaceCollapse &&
                    storyText.Length > 0 &&
                    char.IsWhiteSpace(storyText[^1]) &&
                    char.IsWhiteSpace(text[0]))
                {
                    text = text.Substring(1);
                    if (text.Length == 0)
                    {
                        pendingSummaryWhitespaceCollapse = false;
                        return;
                    }
                }

                pendingSummaryWhitespaceCollapse = false;

                string? fragmentFontName = ResolveFragmentFontName(xmlReader, baseProperties);
                if (ShouldStartNewTextSegment(run, fragmentFontName))
                {
                    run = CreateTextSegmentRun(paragraph, baseProperties, hyperlink, fragmentFontName);
                }
                else if (run.Text.Length == 0 && xmlReader.LocalName == "sym" && !string.IsNullOrEmpty(fragmentFontName))
                {
                    run.Properties.FontName = fragmentFontName;
                }

                AppendRunText(run, storyText, text, hyperlink);
            }
        }

        private static string ResolvePartTargetPath(string sourcePartEntryPath, string target)
        {
            if (string.IsNullOrEmpty(target))
            {
                return target;
            }

            if (target.StartsWith("/", StringComparison.Ordinal))
            {
                return target.Substring(1);
            }

            string? sourceDirectory = Path.GetDirectoryName(sourcePartEntryPath)?.Replace('\\', '/');
            string combined = string.IsNullOrEmpty(sourceDirectory)
                ? target
                : $"{sourceDirectory}/{target}";

            var normalizedSegments = new List<string>();
            foreach (string segment in combined.Split('/', StringSplitOptions.RemoveEmptyEntries))
            {
                if (segment == ".")
                {
                    continue;
                }

                if (segment == "..")
                {
                    if (normalizedSegments.Count > 0)
                    {
                        normalizedSegments.RemoveAt(normalizedSegments.Count - 1);
                    }

                    continue;
                }

                normalizedSegments.Add(segment);
            }

            return string.Join("/", normalizedSegments);
        }

        private static void ResolveHeaderFooterStoryHyperlinks(
            Nedev.FileConverters.DocxToDoc.Model.HeaderFooterStoryModel story,
            IReadOnlyDictionary<string, string> relationships)
        {
            if (relationships.Count == 0)
            {
                return;
            }

            foreach (var paragraph in EnumerateHeaderFooterStoryParagraphs(story))
            {
                foreach (var run in paragraph.Runs)
                {
                    var hyperlink = run.Hyperlink;
                    if (hyperlink == null ||
                        string.IsNullOrEmpty(hyperlink.RelationshipId) ||
                        !string.IsNullOrEmpty(hyperlink.TargetUrl))
                    {
                        continue;
                    }

                    if (relationships.TryGetValue(hyperlink.RelationshipId, out string? target) &&
                        !string.IsNullOrEmpty(target))
                    {
                        hyperlink.TargetUrl = target;
                    }
                }
            }
        }

        private void AppendAltChunkContent(
            string? relationshipId,
            Nedev.FileConverters.DocxToDoc.Model.DocumentModel docModel,
            Nedev.FileConverters.DocxToDoc.Model.TableCellModel? currentCell,
            StringBuilder textBuffer)
        {
            foreach (var block in ReadAltChunkBlocks(relationshipId, _relationships, "word/document.xml"))
            {
                if (currentCell != null)
                {
                    AppendAltChunkBlockToCell(currentCell, block);
                    AppendAltChunkBlockTextBuffer(block, textBuffer);
                }
                else
                {
                    docModel.Content.Add(block);
                    if (block is Nedev.FileConverters.DocxToDoc.Model.ParagraphModel paragraph)
                    {
                        docModel.Paragraphs.Add(paragraph);
                    }

                    AppendAltChunkBlockTextBuffer(block, textBuffer);
                }
            }
        }

        private static void AppendAltChunkBlockToCell(Nedev.FileConverters.DocxToDoc.Model.TableCellModel cell, object block)
        {
            cell.Content.Add(block);
            foreach (var paragraph in EnumerateAltChunkBlockParagraphs(block))
            {
                cell.Paragraphs.Add(paragraph);
            }
        }

        private static IEnumerable<object> EnumerateTableCellBlocks(Nedev.FileConverters.DocxToDoc.Model.TableCellModel cell)
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

        private IEnumerable<object> ReadAltChunkBlocks(
            string? relationshipId,
            IReadOnlyDictionary<string, string> relationships,
            string sourcePartEntryPath)
        {
            if (string.IsNullOrEmpty(relationshipId) || !relationships.TryGetValue(relationshipId, out string? target) || string.IsNullOrEmpty(target))
            {
                yield break;
            }

            string entryPath = ResolvePartTargetPath(sourcePartEntryPath, target);
            if (!IsSupportedAltChunkEntryPath(entryPath))
            {
                yield break;
            }

            var entry = _archive.GetEntry(entryPath);
            if (entry == null)
            {
                yield break;
            }

            if (Path.GetExtension(entryPath).Equals(".docx", StringComparison.OrdinalIgnoreCase))
            {
                foreach (var block in ReadEmbeddedDocxAltChunkBlocks(entry))
                {
                    yield return block;
                }

                yield break;
            }

            using var stream = entry.Open();
            using var reader = new StreamReader(stream, Encoding.UTF8, detectEncodingFromByteOrderMarks: true);
            string content = reader.ReadToEnd();
            if (string.IsNullOrWhiteSpace(content))
            {
                yield break;
            }

            IEnumerable<string> paragraphTexts = LooksLikeRtfContent(entryPath, content)
                ? ExtractAltChunkRtfParagraphTexts(content)
                : LooksLikeMarkupContent(entryPath, content)
                    ? ExtractAltChunkMarkupParagraphTexts(content)
                    : ExtractPlainTextParagraphTexts(content);

            foreach (string paragraphText in paragraphTexts)
            {
                var paragraph = new Nedev.FileConverters.DocxToDoc.Model.ParagraphModel();
                if (!string.IsNullOrEmpty(paragraphText))
                {
                    paragraph.Runs.Add(new Nedev.FileConverters.DocxToDoc.Model.RunModel
                    {
                        Text = paragraphText
                    });
                }

                yield return paragraph;
            }
        }

        private IEnumerable<object> ReadEmbeddedDocxAltChunkBlocks(ZipArchiveEntry entry)
        {
            using var chunkStream = entry.Open();
            using var chunkCopy = new MemoryStream();
            chunkStream.CopyTo(chunkCopy);
            chunkCopy.Position = 0;

            Nedev.FileConverters.DocxToDoc.Model.DocumentModel nestedModel;
            try
            {
                using var nestedReader = new DocxReader(chunkCopy);
                nestedModel = nestedReader.ReadDocument();
            }
            catch (InvalidDataException)
            {
                yield break;
            }
            catch (FileNotFoundException)
            {
                yield break;
            }
            catch (XmlException)
            {
                yield break;
            }

            foreach (var block in nestedModel.Content)
            {
                if (block is Nedev.FileConverters.DocxToDoc.Model.ParagraphModel paragraph)
                {
                    if (!string.IsNullOrEmpty(GetParagraphVisibleText(paragraph)))
                    {
                        yield return paragraph;
                    }
                }
                else if (block is Nedev.FileConverters.DocxToDoc.Model.TableModel table &&
                         EnumerateAltChunkBlockParagraphs(table).Any(paragraph => !string.IsNullOrEmpty(GetParagraphVisibleText(paragraph))))
                {
                    yield return table;
                }
            }
        }

        private static IEnumerable<Nedev.FileConverters.DocxToDoc.Model.ParagraphModel> EnumerateAltChunkBlockParagraphs(object block)
        {
            if (block is Nedev.FileConverters.DocxToDoc.Model.ParagraphModel paragraph)
            {
                yield return paragraph;
                yield break;
            }

            if (block is Nedev.FileConverters.DocxToDoc.Model.TableModel table)
            {
                foreach (var row in table.Rows)
                {
                    foreach (var cell in row.Cells)
                    {
                        foreach (var cellBlock in EnumerateTableCellBlocks(cell))
                        {
                            foreach (var cellParagraph in EnumerateAltChunkBlockParagraphs(cellBlock))
                            {
                                yield return cellParagraph;
                            }
                        }
                    }
                }
            }
        }

        private static void AppendAltChunkBlockTextBuffer(object block, StringBuilder textBuffer)
        {
            foreach (var paragraph in EnumerateAltChunkBlockParagraphs(block))
            {
                AppendParagraphTextBuffer(paragraph, textBuffer);
            }
        }

        private static string GetParagraphVisibleText(Nedev.FileConverters.DocxToDoc.Model.ParagraphModel paragraph)
        {
            var text = new StringBuilder();
            foreach (var run in paragraph.Runs)
            {
                if (!string.IsNullOrEmpty(run.Text))
                {
                    text.Append(run.Text);
                }
            }

            return text.ToString();
        }

        private static void AppendParagraphTextBuffer(
            Nedev.FileConverters.DocxToDoc.Model.ParagraphModel paragraph,
            StringBuilder textBuffer)
        {
            textBuffer.Append(GetParagraphVisibleText(paragraph));
            textBuffer.Append('\r');
        }

        private static bool IsSupportedAltChunkEntryPath(string entryPath)
        {
            string extension = Path.GetExtension(entryPath);
            return extension.Equals(".txt", StringComparison.OrdinalIgnoreCase) ||
                   extension.Equals(".text", StringComparison.OrdinalIgnoreCase) ||
                   extension.Equals(".docx", StringComparison.OrdinalIgnoreCase) ||
                   extension.Equals(".rtf", StringComparison.OrdinalIgnoreCase) ||
                   extension.Equals(".htm", StringComparison.OrdinalIgnoreCase) ||
                   extension.Equals(".html", StringComparison.OrdinalIgnoreCase) ||
                   extension.Equals(".xhtml", StringComparison.OrdinalIgnoreCase) ||
                   extension.Equals(".xml", StringComparison.OrdinalIgnoreCase);
        }

        private static bool LooksLikeRtfContent(string entryPath, string content)
        {
            return content.AsSpan().TrimStart().StartsWith(@"{\rtf", StringComparison.OrdinalIgnoreCase);
        }

        private static bool LooksLikeMarkupContent(string entryPath, string content)
        {
            string extension = Path.GetExtension(entryPath);
            if (extension.Equals(".html", StringComparison.OrdinalIgnoreCase) ||
                extension.Equals(".htm", StringComparison.OrdinalIgnoreCase) ||
                extension.Equals(".xhtml", StringComparison.OrdinalIgnoreCase) ||
                extension.Equals(".xml", StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }

            return content.AsSpan().TrimStart().StartsWith("<", StringComparison.Ordinal);
        }

        private static IEnumerable<string> ExtractPlainTextParagraphTexts(string content)
        {
            foreach (string paragraph in content.Replace("\r\n", "\n").Replace('\r', '\n').Split('\n'))
            {
                string normalized = paragraph.Trim();
                if (!string.IsNullOrEmpty(normalized))
                {
                    yield return normalized;
                }
            }
        }

        private static IEnumerable<string> ExtractAltChunkMarkupParagraphTexts(string markup)
        {
            string text = Regex.Replace(markup, "<\\s*(script|style)\\b[^>]*>.*?<\\s*/\\s*\\1\\s*>", string.Empty, RegexOptions.IgnoreCase | RegexOptions.Singleline);
            text = Regex.Replace(text, "<\\s*br\\b[^>]*>", "\n", RegexOptions.IgnoreCase);
            text = Regex.Replace(text, "<\\s*/\\s*(p|div|li|tr|table|section|article|header|footer|h[1-6]|ul|ol)\\s*>", "\n", RegexOptions.IgnoreCase);
            text = Regex.Replace(text, "<\\s*/\\s*(td|th)\\s*>", " ", RegexOptions.IgnoreCase);
            text = Regex.Replace(text, "<[^>]+>", string.Empty, RegexOptions.Singleline);
            text = WebUtility.HtmlDecode(text).Replace('\u00A0', ' ');
            text = text.Replace("\r\n", "\n").Replace('\r', '\n');

            foreach (string paragraph in text.Split('\n'))
            {
                string normalized = Regex.Replace(paragraph, "\\s+", " ").Trim();
                if (!string.IsNullOrEmpty(normalized))
                {
                    yield return normalized;
                }
            }
        }

        private static IEnumerable<string> ExtractAltChunkRtfParagraphTexts(string rtf)
        {
            var text = new StringBuilder();
            var stateStack = new Stack<RtfGroupState>();
            var currentState = new RtfGroupState
            {
                UnicodeFallbackSkipCount = 1
            };
            int pendingUnicodeFallbackChars = 0;

            for (int i = 0; i < rtf.Length; i++)
            {
                char current = rtf[i];
                if (current == '{')
                {
                    stateStack.Push(currentState);
                    currentState.AtStartOfGroup = true;
                    pendingUnicodeFallbackChars = 0;
                    continue;
                }

                if (current == '}')
                {
                    currentState = stateStack.Count > 0
                        ? stateStack.Pop()
                        : CreateDefaultRtfGroupState();
                    pendingUnicodeFallbackChars = 0;
                    continue;
                }

                if (current == '\r' || current == '\n')
                {
                    continue;
                }

                if (current == '\\')
                {
                    if (i + 1 >= rtf.Length)
                    {
                        break;
                    }

                    i++;
                    char escaped = rtf[i];
                    if (escaped == '\\' || escaped == '{' || escaped == '}')
                    {
                        AppendRtfVisibleCharacter(text, escaped, currentState.Skip, ref pendingUnicodeFallbackChars);
                        currentState.AtStartOfGroup = false;
                        continue;
                    }

                    if (escaped == '\'')
                    {
                        if (i + 2 < rtf.Length &&
                            IsHexDigit(rtf[i + 1]) &&
                            IsHexDigit(rtf[i + 2]))
                        {
                            byte value = byte.Parse(rtf.Substring(i + 1, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture);
                            AppendRtfVisibleCharacter(text, (char)value, currentState.Skip, ref pendingUnicodeFallbackChars);
                            i += 2;
                        }

                        currentState.AtStartOfGroup = false;
                        continue;
                    }

                    if (escaped == '*')
                    {
                        currentState.Skip = true;
                        currentState.AtStartOfGroup = false;
                        continue;
                    }

                    if (!char.IsLetter(escaped))
                    {
                        HandleRtfControlSymbol(text, escaped, currentState.Skip, ref pendingUnicodeFallbackChars);
                        currentState.AtStartOfGroup = false;
                        continue;
                    }

                    int controlStart = i;
                    while (i < rtf.Length && char.IsLetter(rtf[i]))
                    {
                        i++;
                    }

                    string controlWord = rtf.Substring(controlStart, i - controlStart);
                    int? numericArgument = null;
                    int numericStart = i;
                    if (i < rtf.Length && (rtf[i] == '-' || char.IsDigit(rtf[i])))
                    {
                        i++;
                        while (i < rtf.Length && char.IsDigit(rtf[i]))
                        {
                            i++;
                        }

                        if (int.TryParse(rtf.Substring(numericStart, i - numericStart), NumberStyles.Integer, CultureInfo.InvariantCulture, out int parsedValue))
                        {
                            numericArgument = parsedValue;
                        }
                    }

                    bool hasDelimiterSpace = i < rtf.Length && rtf[i] == ' ';
                    if (!hasDelimiterSpace)
                    {
                        i--;
                    }

                    HandleRtfControlWord(text, controlWord, numericArgument, ref currentState, ref pendingUnicodeFallbackChars);
                    currentState.AtStartOfGroup = false;
                    continue;
                }

                AppendRtfVisibleCharacter(text, current, currentState.Skip, ref pendingUnicodeFallbackChars);
                currentState.AtStartOfGroup = false;
            }

            foreach (string paragraph in text.ToString().Replace("\r\n", "\n").Replace('\r', '\n').Split('\n'))
            {
                string normalized = Regex.Replace(paragraph, "\\s+", " ").Trim();
                if (!string.IsNullOrEmpty(normalized))
                {
                    yield return normalized;
                }
            }
        }

        private static void HandleRtfControlWord(
            StringBuilder text,
            string controlWord,
            int? numericArgument,
            ref RtfGroupState currentState,
            ref int pendingUnicodeFallbackChars)
        {
            if (currentState.AtStartOfGroup && IsIgnorableRtfDestination(controlWord))
            {
                currentState.Skip = true;
                return;
            }

            switch (controlWord)
            {
                case "uc":
                    if (numericArgument.HasValue && numericArgument.Value >= 0)
                    {
                        currentState.UnicodeFallbackSkipCount = numericArgument.Value;
                    }

                    return;
                case "u":
                    if (!currentState.Skip && numericArgument.HasValue)
                    {
                        int value = numericArgument.Value;
                        if (value < 0)
                        {
                            value += 65536;
                        }

                        text.Append(char.ConvertFromUtf32(ClampRtfUnicodeCodePoint(value)));
                        pendingUnicodeFallbackChars = currentState.UnicodeFallbackSkipCount;
                    }

                    return;
                case "par":
                case "line":
                    if (!currentState.Skip)
                    {
                        text.Append('\n');
                    }

                    pendingUnicodeFallbackChars = 0;
                    return;
                case "tab":
                    if (!currentState.Skip)
                    {
                        text.Append('\t');
                    }

                    return;
                case "emdash":
                    if (!currentState.Skip)
                    {
                        text.Append('\u2014');
                    }

                    return;
                case "endash":
                    if (!currentState.Skip)
                    {
                        text.Append('\u2013');
                    }

                    return;
                case "bullet":
                    if (!currentState.Skip)
                    {
                        text.Append('\u2022');
                    }

                    return;
                case "lquote":
                    if (!currentState.Skip)
                    {
                        text.Append('\u2018');
                    }

                    return;
                case "rquote":
                    if (!currentState.Skip)
                    {
                        text.Append('\u2019');
                    }

                    return;
                case "ldblquote":
                    if (!currentState.Skip)
                    {
                        text.Append('\u201C');
                    }

                    return;
                case "rdblquote":
                    if (!currentState.Skip)
                    {
                        text.Append('\u201D');
                    }

                    return;
                case "cell":
                    if (!currentState.Skip)
                    {
                        text.Append('\t');
                    }

                    return;
                case "row":
                    if (!currentState.Skip)
                    {
                        text.Append('\n');
                    }

                    return;
            }
        }

        private static void HandleRtfControlSymbol(
            StringBuilder text,
            char controlSymbol,
            bool skip,
            ref int pendingUnicodeFallbackChars)
        {
            if (skip)
            {
                return;
            }

            switch (controlSymbol)
            {
                case '~':
                    text.Append(' ');
                    return;
                case '_':
                    text.Append('\u2011');
                    return;
                case '-':
                    return;
            }

            AppendRtfVisibleCharacter(text, controlSymbol, skip, ref pendingUnicodeFallbackChars);
        }

        private static void AppendRtfVisibleCharacter(
            StringBuilder text,
            char character,
            bool skip,
            ref int pendingUnicodeFallbackChars)
        {
            if (skip)
            {
                return;
            }

            if (pendingUnicodeFallbackChars > 0)
            {
                pendingUnicodeFallbackChars--;
                return;
            }

            text.Append(character);
        }

        private static bool IsIgnorableRtfDestination(string controlWord)
        {
            return controlWord.Equals("fonttbl", StringComparison.OrdinalIgnoreCase) ||
                   controlWord.Equals("colortbl", StringComparison.OrdinalIgnoreCase) ||
                   controlWord.Equals("stylesheet", StringComparison.OrdinalIgnoreCase) ||
                   controlWord.Equals("info", StringComparison.OrdinalIgnoreCase) ||
                   controlWord.Equals("pict", StringComparison.OrdinalIgnoreCase) ||
                   controlWord.Equals("object", StringComparison.OrdinalIgnoreCase) ||
                   controlWord.Equals("header", StringComparison.OrdinalIgnoreCase) ||
                   controlWord.Equals("footer", StringComparison.OrdinalIgnoreCase) ||
                   controlWord.Equals("headerl", StringComparison.OrdinalIgnoreCase) ||
                   controlWord.Equals("headerr", StringComparison.OrdinalIgnoreCase) ||
                   controlWord.Equals("headerf", StringComparison.OrdinalIgnoreCase) ||
                   controlWord.Equals("footerl", StringComparison.OrdinalIgnoreCase) ||
                   controlWord.Equals("footerr", StringComparison.OrdinalIgnoreCase) ||
                   controlWord.Equals("footerf", StringComparison.OrdinalIgnoreCase) ||
                   controlWord.Equals("annotation", StringComparison.OrdinalIgnoreCase);
        }

        private static bool IsHexDigit(char c)
        {
            return (c >= '0' && c <= '9') ||
                   (c >= 'A' && c <= 'F') ||
                   (c >= 'a' && c <= 'f');
        }

        private static int ClampRtfUnicodeCodePoint(int value)
        {
            if (value < 0)
            {
                return 0;
            }

            if (value > 0x10FFFF)
            {
                return 0x10FFFF;
            }

            return value;
        }

        private static RtfGroupState CreateDefaultRtfGroupState()
        {
            return new RtfGroupState
            {
                UnicodeFallbackSkipCount = 1
            };
        }

        private struct RtfGroupState
        {
            public bool Skip;
            public int UnicodeFallbackSkipCount;
            public bool AtStartOfGroup;
        }

        private static IEnumerable<Nedev.FileConverters.DocxToDoc.Model.ParagraphModel> EnumerateHeaderFooterStoryParagraphs(
            Nedev.FileConverters.DocxToDoc.Model.HeaderFooterStoryModel story)
        {
            if (story.Content.Count == 0)
            {
                foreach (var paragraph in story.Paragraphs)
                {
                    yield return paragraph;
                }

                yield break;
            }

            foreach (var block in story.Content)
            {
                if (block is Nedev.FileConverters.DocxToDoc.Model.ParagraphModel paragraph)
                {
                    yield return paragraph;
                }
                else if (block is Nedev.FileConverters.DocxToDoc.Model.TableModel table)
                {
                    foreach (var row in table.Rows)
                    {
                        foreach (var cell in row.Cells)
                        {
                            foreach (var cellBlock in EnumerateTableCellBlocks(cell))
                            {
                                foreach (var cellParagraph in EnumerateAltChunkBlockParagraphs(cellBlock))
                                {
                                    yield return cellParagraph;
                                }
                            }
                        }
                    }
                }
            }
        }

        private static string ReadSecondaryStoryText(Stream storyStream)
        {
            using var reader = XmlReader.Create(storyStream, new XmlReaderSettings { IgnoreWhitespace = true });
            var text = new StringBuilder();
            int paragraphCount = 0;

            while (reader.Read())
            {
                if (reader.NodeType != XmlNodeType.Element)
                {
                    continue;
                }

                string localName = reader.LocalName;
                if (localName == "p")
                {
                    if (paragraphCount > 0)
                    {
                        text.Append('\r');
                    }

                    paragraphCount++;
                }
                else if (IsNoteTextFragmentElement(localName))
                {
                    text.Append(ReadRunTextFragment(reader));
                }
            }

            return text.ToString();
        }

        private static int ResolveGridWidth(IReadOnlyList<int> gridColumnWidths, int startIndex, int gridSpan)
        {
            if (gridColumnWidths.Count == 0)
            {
                return 0;
            }

            int span = Math.Max(1, gridSpan);
            int width = 0;
            for (int index = 0; index < span; index++)
            {
                int gridIndex = startIndex + index;
                if (gridIndex >= 0 && gridIndex < gridColumnWidths.Count)
                {
                    width += gridColumnWidths[gridIndex];
                }
            }

            return width;
        }

        private static bool TryReadDxaWidth(XmlReader xmlReader, out int width)
        {
            string? type = xmlReader.GetAttribute("w:type");
            if (string.Equals(type, "nil", StringComparison.OrdinalIgnoreCase))
            {
                width = 0;
                return true;
            }

            if (!string.IsNullOrEmpty(type) && !string.Equals(type, "dxa", StringComparison.OrdinalIgnoreCase))
            {
                width = 0;
                return false;
            }

            return int.TryParse(xmlReader.GetAttribute("w:w"), out width);
        }

        private static Nedev.FileConverters.DocxToDoc.Model.TableCellVerticalMerge ParseVerticalMerge(string? value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return Nedev.FileConverters.DocxToDoc.Model.TableCellVerticalMerge.Continue;
            }

            return value.ToLowerInvariant() switch
            {
                "restart" => Nedev.FileConverters.DocxToDoc.Model.TableCellVerticalMerge.Restart,
                "continue" => Nedev.FileConverters.DocxToDoc.Model.TableCellVerticalMerge.Continue,
                "false" => Nedev.FileConverters.DocxToDoc.Model.TableCellVerticalMerge.None,
                "0" => Nedev.FileConverters.DocxToDoc.Model.TableCellVerticalMerge.None,
                "off" => Nedev.FileConverters.DocxToDoc.Model.TableCellVerticalMerge.None,
                "nil" => Nedev.FileConverters.DocxToDoc.Model.TableCellVerticalMerge.None,
                "none" => Nedev.FileConverters.DocxToDoc.Model.TableCellVerticalMerge.None,
                _ => Nedev.FileConverters.DocxToDoc.Model.TableCellVerticalMerge.Continue
            };
        }

        private static Nedev.FileConverters.DocxToDoc.Model.TableCellHorizontalMerge ParseHorizontalMerge(string? value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return Nedev.FileConverters.DocxToDoc.Model.TableCellHorizontalMerge.Continue;
            }

            return value.ToLowerInvariant() switch
            {
                "restart" => Nedev.FileConverters.DocxToDoc.Model.TableCellHorizontalMerge.Restart,
                "continue" => Nedev.FileConverters.DocxToDoc.Model.TableCellHorizontalMerge.Continue,
                "false" => Nedev.FileConverters.DocxToDoc.Model.TableCellHorizontalMerge.None,
                "0" => Nedev.FileConverters.DocxToDoc.Model.TableCellHorizontalMerge.None,
                "off" => Nedev.FileConverters.DocxToDoc.Model.TableCellHorizontalMerge.None,
                "nil" => Nedev.FileConverters.DocxToDoc.Model.TableCellHorizontalMerge.None,
                "none" => Nedev.FileConverters.DocxToDoc.Model.TableCellHorizontalMerge.None,
                _ => Nedev.FileConverters.DocxToDoc.Model.TableCellHorizontalMerge.Continue
            };
        }

        private static bool TryReadBorderWidthTwips(XmlReader xmlReader, out int width, out Nedev.FileConverters.DocxToDoc.Model.BorderStyle style)
        {
            string? borderValue = xmlReader.GetAttribute("w:val");
            style = ParseBorderStyle(borderValue);

            if (style == Nedev.FileConverters.DocxToDoc.Model.BorderStyle.Nil || style == Nedev.FileConverters.DocxToDoc.Model.BorderStyle.None)
            {
                width = 0;
                return true;
            }

            if (!int.TryParse(xmlReader.GetAttribute("w:sz"), out int eighthPointWidth))
            {
                eighthPointWidth = 4;
            }

            width = (int)Math.Round(Math.Max(0, eighthPointWidth) * 2.5d, MidpointRounding.AwayFromZero);
            return true;
        }

        private static Nedev.FileConverters.DocxToDoc.Model.BorderStyle ParseBorderStyle(string? value)
        {
            if (string.IsNullOrWhiteSpace(value)) return Nedev.FileConverters.DocxToDoc.Model.BorderStyle.None;
            return value.ToLowerInvariant() switch
            {
                "nil" => Nedev.FileConverters.DocxToDoc.Model.BorderStyle.Nil,
                "none" => Nedev.FileConverters.DocxToDoc.Model.BorderStyle.None,
                "single" => Nedev.FileConverters.DocxToDoc.Model.BorderStyle.Single,
                "thick" => Nedev.FileConverters.DocxToDoc.Model.BorderStyle.Single,
                "dotted" => Nedev.FileConverters.DocxToDoc.Model.BorderStyle.Dotted,
                "dashed" => Nedev.FileConverters.DocxToDoc.Model.BorderStyle.Dashed,
                "dotdash" => Nedev.FileConverters.DocxToDoc.Model.BorderStyle.Dashed,
                "dotdotdash" => Nedev.FileConverters.DocxToDoc.Model.BorderStyle.Dashed,
                "double" => Nedev.FileConverters.DocxToDoc.Model.BorderStyle.Double,
                _ => Nedev.FileConverters.DocxToDoc.Model.BorderStyle.Other
            };
        }

        private static void AppendRunText(Nedev.FileConverters.DocxToDoc.Model.RunModel run, StringBuilder textBuffer, string text, Nedev.FileConverters.DocxToDoc.Model.HyperlinkModel? hyperlink = null)
        {
            if (string.IsNullOrEmpty(text))
            {
                return;
            }

            run.Text += text;
            if (hyperlink != null)
            {
                hyperlink.DisplayText += text;
            }

            textBuffer.Append(text);
        }

        private static void AppendRunTextFragment(
            Nedev.FileConverters.DocxToDoc.Model.ParagraphModel paragraph,
            ref Nedev.FileConverters.DocxToDoc.Model.RunModel run,
            Nedev.FileConverters.DocxToDoc.Model.RunModel.CharacterProperties baseProperties,
            StringBuilder textBuffer,
            XmlReader reader,
            Nedev.FileConverters.DocxToDoc.Model.HyperlinkModel? hyperlink = null)
        {
            string text = ReadRunTextFragment(reader);
            if (string.IsNullOrEmpty(text))
            {
                return;
            }

            string? fragmentFontName = ResolveFragmentFontName(reader, baseProperties);
            if (ShouldStartNewTextSegment(run, fragmentFontName))
            {
                run = CreateTextSegmentRun(paragraph, baseProperties, hyperlink, fragmentFontName);
            }
            else if (run.Text.Length == 0 && reader.LocalName == "sym" && !string.IsNullOrEmpty(fragmentFontName))
            {
                run.Properties.FontName = fragmentFontName;
            }

            AppendRunText(run, textBuffer, text, hyperlink);
        }

        private static void AppendRunTextSegment(
            Nedev.FileConverters.DocxToDoc.Model.ParagraphModel paragraph,
            ref Nedev.FileConverters.DocxToDoc.Model.RunModel run,
            Nedev.FileConverters.DocxToDoc.Model.RunModel.CharacterProperties baseProperties,
            StringBuilder textBuffer,
            string text,
            Nedev.FileConverters.DocxToDoc.Model.HyperlinkModel? hyperlink = null)
        {
            if (string.IsNullOrEmpty(text))
            {
                return;
            }

            string? fragmentFontName = baseProperties.FontName;
            if (ShouldStartNewTextSegment(run, fragmentFontName))
            {
                run = CreateTextSegmentRun(paragraph, baseProperties, hyperlink, fragmentFontName);
            }

            AppendRunText(run, textBuffer, text, hyperlink);
        }

        private static Nedev.FileConverters.DocxToDoc.Model.FieldModel CreateFieldModel(string instruction, string? fldLock, string? fldDirty)
        {
            return new Nedev.FileConverters.DocxToDoc.Model.FieldModel
            {
                Instruction = instruction,
                Type = ParseFieldType(instruction),
                IsLocked = IsTrueValue(fldLock),
                IsDirty = IsTrueValue(fldDirty)
            };
        }

        private static Nedev.FileConverters.DocxToDoc.Model.RunModel CreateFieldMarkerRun(
            Nedev.FileConverters.DocxToDoc.Model.FieldModel field,
            bool isFieldBegin = false,
            bool isFieldSeparate = false,
            bool isFieldEnd = false)
        {
            return new Nedev.FileConverters.DocxToDoc.Model.RunModel
            {
                Field = field,
                IsFieldBegin = isFieldBegin,
                IsFieldSeparate = isFieldSeparate,
                IsFieldEnd = isFieldEnd
            };
        }

        private static Nedev.FileConverters.DocxToDoc.Model.RunModel CreateTextSegmentRun(
            Nedev.FileConverters.DocxToDoc.Model.ParagraphModel paragraph,
            Nedev.FileConverters.DocxToDoc.Model.RunModel.CharacterProperties baseProperties,
            Nedev.FileConverters.DocxToDoc.Model.HyperlinkModel? hyperlink,
            string? fontNameOverride)
        {
            var run = new Nedev.FileConverters.DocxToDoc.Model.RunModel();
            CopyCharacterProperties(baseProperties, run.Properties);
            if (!string.IsNullOrEmpty(fontNameOverride))
            {
                run.Properties.FontName = fontNameOverride;
            }
            run.Hyperlink = hyperlink;
            paragraph.Runs.Add(run);
            return run;
        }

        private static bool ShouldStartNewTextSegment(Nedev.FileConverters.DocxToDoc.Model.RunModel run, string? fragmentFontName)
        {
            if (run.Image != null || run.IsFieldBegin || run.IsFieldSeparate || run.IsFieldEnd)
            {
                return true;
            }

            if (run.Text.Length == 0)
            {
                return false;
            }

            return !string.Equals(run.Properties.FontName, fragmentFontName, StringComparison.Ordinal);
        }

        private static string? ResolveFragmentFontName(XmlReader reader, Nedev.FileConverters.DocxToDoc.Model.RunModel.CharacterProperties baseProperties)
        {
            if (reader.LocalName == "sym")
            {
                return reader.GetAttribute("w:font") ?? baseProperties.FontName;
            }

            return baseProperties.FontName;
        }

        private static Nedev.FileConverters.DocxToDoc.Model.RunModel.CharacterProperties CloneCharacterProperties(Nedev.FileConverters.DocxToDoc.Model.RunModel.CharacterProperties source)
        {
            var clone = new Nedev.FileConverters.DocxToDoc.Model.RunModel.CharacterProperties();
            CopyCharacterProperties(source, clone);
            return clone;
        }

        private static void CopyCharacterProperties(
            Nedev.FileConverters.DocxToDoc.Model.RunModel.CharacterProperties source,
            Nedev.FileConverters.DocxToDoc.Model.RunModel.CharacterProperties target)
        {
            target.IsBold = source.IsBold;
            target.IsItalic = source.IsItalic;
            target.IsStrike = source.IsStrike;
            target.FontSize = source.FontSize;
            target.FontName = source.FontName;
            target.Underline = source.Underline;
            target.Color = source.Color;
        }

        private static bool TryApplyRunFormattingElement(XmlReader reader, Nedev.FileConverters.DocxToDoc.Model.RunModel run)
        {
            return TryApplyRunFormattingElement(reader, run.Properties);
        }

        private static bool TryApplyRunFormattingElement(XmlReader reader, Nedev.FileConverters.DocxToDoc.Model.RunModel.CharacterProperties properties)
        {
            switch (reader.LocalName)
            {
                case "b":
                    properties.IsBold = !IsFalseValue(reader.GetAttribute("w:val"));
                    return true;
                case "bCs":
                    properties.IsBold = !IsFalseValue(reader.GetAttribute("w:val"));
                    return true;
                case "i":
                    properties.IsItalic = !IsFalseValue(reader.GetAttribute("w:val"));
                    return true;
                case "iCs":
                    properties.IsItalic = !IsFalseValue(reader.GetAttribute("w:val"));
                    return true;
                case "strike":
                    properties.IsStrike = !IsFalseValue(reader.GetAttribute("w:val"));
                    return true;
                case "dstrike":
                    properties.IsStrike = !IsFalseValue(reader.GetAttribute("w:val"));
                    return true;
                case "u":
                    string underlineValue = reader.GetAttribute("w:val") ?? "none";
                    properties.Underline = underlineValue switch
                    {
                        "words" => Nedev.FileConverters.DocxToDoc.Model.UnderlineType.Single,
                        "single" => Nedev.FileConverters.DocxToDoc.Model.UnderlineType.Single,
                        "double" => Nedev.FileConverters.DocxToDoc.Model.UnderlineType.Double,
                        "thick" => Nedev.FileConverters.DocxToDoc.Model.UnderlineType.Thick,
                        "dash" => Nedev.FileConverters.DocxToDoc.Model.UnderlineType.Dashed,
                        "dotDash" => Nedev.FileConverters.DocxToDoc.Model.UnderlineType.Dashed,
                        "dotDotDash" => Nedev.FileConverters.DocxToDoc.Model.UnderlineType.Dashed,
                        "dotted" => Nedev.FileConverters.DocxToDoc.Model.UnderlineType.Dotted,
                        "dottedHeavy" => Nedev.FileConverters.DocxToDoc.Model.UnderlineType.Dotted,
                        "dashed" => Nedev.FileConverters.DocxToDoc.Model.UnderlineType.Dashed,
                        "dashLong" => Nedev.FileConverters.DocxToDoc.Model.UnderlineType.Dashed,
                        "dashLongHeavy" => Nedev.FileConverters.DocxToDoc.Model.UnderlineType.Dashed,
                        "wavy" => Nedev.FileConverters.DocxToDoc.Model.UnderlineType.Wave,
                        "wavyDouble" => Nedev.FileConverters.DocxToDoc.Model.UnderlineType.Wave,
                        "wavyHeavy" => Nedev.FileConverters.DocxToDoc.Model.UnderlineType.Wave,
                        "wave" => Nedev.FileConverters.DocxToDoc.Model.UnderlineType.Wave,
                        _ => Nedev.FileConverters.DocxToDoc.Model.UnderlineType.None
                    };
                    return true;
                case "color":
                    string? colorValue = reader.GetAttribute("w:val");
                    if (!string.IsNullOrEmpty(colorValue) && colorValue != "auto")
                    {
                        properties.Color = colorValue;
                    }
                    return true;
                case "rFonts":
                    string? fontName = reader.GetAttribute("w:ascii")
                        ?? reader.GetAttribute("w:hAnsi")
                        ?? reader.GetAttribute("w:cs")
                        ?? reader.GetAttribute("w:eastAsia");
                    if (!string.IsNullOrEmpty(fontName))
                    {
                        properties.FontName = fontName;
                    }
                    return true;
                case "sz":
                    if (int.TryParse(reader.GetAttribute("w:val"), out int size))
                    {
                        properties.FontSize = size;
                    }
                    return true;
                case "szCs":
                    if (int.TryParse(reader.GetAttribute("w:val"), out int csSize))
                    {
                        properties.FontSize = csSize;
                    }
                    return true;
                default:
                    return false;
            }
        }

        private static string ReadRunTextFragment(XmlReader reader)
        {
            return reader.LocalName switch
            {
                "t" or "delText" => ReadCurrentElementString(reader),
                "tab" => "\t",
                "ptab" => "\t",
                "cr" => "\v",
                "br" => string.Equals(reader.GetAttribute("w:type"), "page", StringComparison.OrdinalIgnoreCase) ? "\f" : "\v",
                "noBreakHyphen" => "\u2011",
                "softHyphen" => "\u00AD",
                "sym" => ReadSymbolFragment(reader),
                _ => string.Empty
            };
        }

        private static string ReadSymbolFragment(XmlReader reader)
        {
            string? charValue = reader.GetAttribute("w:char");
            if (string.IsNullOrWhiteSpace(charValue))
            {
                return string.Empty;
            }

            if (!uint.TryParse(charValue, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out uint codePoint))
            {
                return string.Empty;
            }

            if (codePoint > 0x10FFFF)
            {
                return string.Empty;
            }

            return char.ConvertFromUtf32((int)codePoint);
        }

        private static string ReadCurrentElementString(XmlReader reader)
        {
            if (reader.IsEmptyElement)
            {
                return string.Empty;
            }

            int elementDepth = reader.Depth;
            var content = new StringBuilder();

            while (reader.Read())
            {
                if (reader.NodeType == XmlNodeType.Text ||
                    reader.NodeType == XmlNodeType.CDATA ||
                    reader.NodeType == XmlNodeType.SignificantWhitespace ||
                    reader.NodeType == XmlNodeType.Whitespace)
                {
                    content.Append(reader.Value);
                }
                else if (reader.NodeType == XmlNodeType.EndElement && reader.Depth == elementDepth)
                {
                    break;
                }
            }

            return content.ToString();
        }

        private void ParseFootnoteReferences(Nedev.FileConverters.DocxToDoc.Model.DocumentModel docModel)
        {
            ParseNoteReferences(
                docModel.Footnotes,
                "footnoteReference",
                static footnote => footnote.Id,
                static (footnote, cp) => footnote.ReferenceCp = cp,
                static (footnote, customMarkText) => footnote.CustomMarkText = customMarkText);
        }

        private void ParseEndnoteReferences(Nedev.FileConverters.DocxToDoc.Model.DocumentModel docModel)
        {
            ParseNoteReferences(
                docModel.Endnotes,
                "endnoteReference",
                static endnote => endnote.Id,
                static (endnote, cp) => endnote.ReferenceCp = cp,
                static (endnote, customMarkText) => endnote.CustomMarkText = customMarkText);
        }

        private void ParseNoteReferences<TNote>(
            IReadOnlyList<TNote> notes,
            string referenceElementName,
            Func<TNote, string> getId,
            Action<TNote, int> setReferenceCp,
            Action<TNote, string>? setCustomMarkText = null)
        {
            if (notes.Count == 0)
            {
                return;
            }

            var docEntry = _archive.GetEntry("word/document.xml");
            if (docEntry == null)
            {
                return;
            }

            var notesById = notes
                .Select(note => (Note: note, Id: getId(note)))
                .Where(entry => !string.IsNullOrEmpty(entry.Id))
                .GroupBy(entry => entry.Id!, StringComparer.Ordinal)
                .ToDictionary(group => group.Key, group => group.First().Note, StringComparer.Ordinal);
            if (notesById.Count == 0)
            {
                return;
            }

            using var stream = docEntry.Open();
            using var reader = XmlReader.Create(stream, new XmlReaderSettings { IgnoreWhitespace = true });

            int currentCp = 0;
            var anchoredNotes = new HashSet<string>(StringComparer.Ordinal);
            TNote? pendingCustomMarkNote = default;
            StringBuilder? pendingCustomMarkText = null;
            bool currentRunCanCarryCustomMark = false;

            void CommitPendingCustomMark()
            {
                if (pendingCustomMarkNote != null &&
                    pendingCustomMarkText != null &&
                    pendingCustomMarkText.Length > 0 &&
                    setCustomMarkText != null)
                {
                    setCustomMarkText(pendingCustomMarkNote, pendingCustomMarkText.ToString());
                }

                pendingCustomMarkNote = default;
                pendingCustomMarkText = null;
            }

            while (reader.Read())
            {
                if (reader.NodeType == XmlNodeType.Element)
                {
                    string localName = reader.LocalName;
                    if (localName == "r")
                    {
                        currentRunCanCarryCustomMark = false;
                    }
                    else if (localName == "rStyle")
                    {
                        string? styleValue = reader.GetAttribute("w:val");
                        if (string.Equals(styleValue, "FootnoteReference", StringComparison.OrdinalIgnoreCase) ||
                            string.Equals(styleValue, "EndnoteReference", StringComparison.OrdinalIgnoreCase))
                        {
                            currentRunCanCarryCustomMark = true;
                        }
                    }
                    else if (localName == "vertAlign")
                    {
                        string? verticalAlign = reader.GetAttribute("w:val");
                        if (string.Equals(verticalAlign, "superscript", StringComparison.OrdinalIgnoreCase))
                        {
                            currentRunCanCarryCustomMark = true;
                        }
                    }
                    else if (localName == referenceElementName)
                    {
                        CommitPendingCustomMark();
                        string? id = reader.GetAttribute("w:id");
                        if (!string.IsNullOrEmpty(id) &&
                            !anchoredNotes.Contains(id) &&
                            notesById.TryGetValue(id, out var note))
                        {
                            setReferenceCp(note, currentCp);
                            anchoredNotes.Add(id);
                            if (setCustomMarkText != null && IsTrueValue(reader.GetAttribute("w:customMarkFollows")))
                            {
                                pendingCustomMarkNote = note;
                                pendingCustomMarkText = new StringBuilder();
                            }
                        }
                    }
                    else if (IsNoteTextFragmentElement(localName))
                    {
                        string text = ReadRunTextFragment(reader);
                        if (pendingCustomMarkNote != null && text.Length > 0 && setCustomMarkText != null)
                        {
                            if (currentRunCanCarryCustomMark)
                            {
                                string? customMarkTextFragment = NormalizeCustomMarkTextFragment(text);
                                if (customMarkTextFragment != null)
                                {
                                    pendingCustomMarkText ??= new StringBuilder();
                                    pendingCustomMarkText.Append(customMarkTextFragment);
                                }
                                else if (pendingCustomMarkText != null && pendingCustomMarkText.Length > 0)
                                {
                                    CommitPendingCustomMark();
                                }
                                else
                                {
                                    pendingCustomMarkNote = default;
                                    pendingCustomMarkText = null;
                                }
                            }
                            else
                            {
                                if (pendingCustomMarkText != null && pendingCustomMarkText.Length > 0)
                                {
                                    CommitPendingCustomMark();
                                }
                                else
                                {
                                    pendingCustomMarkNote = default;
                                    pendingCustomMarkText = null;
                                }
                            }
                        }

                        currentCp += text.Length;
                    }
                }
                else if (reader.NodeType == XmlNodeType.EndElement && reader.LocalName == "r")
                {
                    currentRunCanCarryCustomMark = false;
                }
                else if (reader.NodeType == XmlNodeType.EndElement && reader.LocalName == "p")
                {
                    CommitPendingCustomMark();
                    currentCp += 1;
                }
            }

            CommitPendingCustomMark();
        }

        private void ParseBookmarks(Nedev.FileConverters.DocxToDoc.Model.DocumentModel docModel)
        {
            // Try to load the document.xml again to find bookmarks
            var docEntry = _archive.GetEntry("word/document.xml");
            if (docEntry == null) return;

            using var stream = docEntry.Open();
            using var reader = XmlReader.Create(stream, new XmlReaderSettings { IgnoreWhitespace = true });

            int currentCp = 0;
            var bookmarkStarts = new Dictionary<string, int>();

            while (reader.Read())
            {
                if (reader.NodeType == XmlNodeType.Element)
                {
                    string localName = reader.LocalName;

                    if (localName == "bookmarkStart")
                    {
                        string? id = reader.GetAttribute("w:id");
                        string? name = reader.GetAttribute("w:name");
                        if (!string.IsNullOrEmpty(id) && !string.IsNullOrEmpty(name))
                        {
                            bookmarkStarts[id] = currentCp;
                        }
                    }
                    else if (localName == "bookmarkEnd")
                    {
                        string? id = reader.GetAttribute("w:id");
                        if (!string.IsNullOrEmpty(id) && bookmarkStarts.TryGetValue(id, out int startCp))
                        {
                            // Find the bookmark name from the start
                            var bookmarkEntry = docEntry;
                            string? name = null;

                            // Re-read to find the name
                            using var stream2 = _archive.GetEntry("word/document.xml")?.Open();
                            if (stream2 != null)
                            {
                                using var reader2 = XmlReader.Create(stream2, new XmlReaderSettings { IgnoreWhitespace = true });
                                while (reader2.Read())
                                {
                                    if (reader2.NodeType == XmlNodeType.Element &&
                                        reader2.LocalName == "bookmarkStart" &&
                                        reader2.GetAttribute("w:id") == id)
                                    {
                                        name = reader2.GetAttribute("w:name");
                                        break;
                                    }
                                }
                            }

                            if (!string.IsNullOrEmpty(name))
                            {
                                docModel.Bookmarks.Add(new Nedev.FileConverters.DocxToDoc.Model.BookmarkModel
                                {
                                    Id = id,
                                    Name = name,
                                    StartCp = startCp,
                                    EndCp = currentCp,
                                    IsCollapsed = startCp == currentCp
                                });
                            }

                            bookmarkStarts.Remove(id);
                        }
                    }
                    else if (localName == "t" || localName == "delText" || localName == "tab" || localName == "ptab" || localName == "br" || localName == "cr" || localName == "noBreakHyphen" || localName == "softHyphen" || localName == "sym")
                    {
                        currentCp += ReadRunTextFragment(reader).Length;
                    }
                }
                else if (reader.NodeType == XmlNodeType.EndElement && reader.LocalName == "p")
                {
                    currentCp += 1;
                }
            }
        }

        private static string? NormalizeCustomMarkTextFragment(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
            {
                return null;
            }

            foreach (char ch in text)
            {
                if (char.IsControl(ch))
                {
                    return null;
                }
            }

            return text;
        }

        private void ParseCommentRanges(Nedev.FileConverters.DocxToDoc.Model.DocumentModel docModel)
        {
            if (docModel.Comments.Count == 0)
            {
                return;
            }

            var docEntry = _archive.GetEntry("word/document.xml");
            if (docEntry == null)
            {
                return;
            }

            var commentsById = docModel.Comments
                .Where(comment => !string.IsNullOrEmpty(comment.Id))
                .GroupBy(comment => comment.Id, StringComparer.Ordinal)
                .ToDictionary(group => group.Key, group => group.First(), StringComparer.Ordinal);

            if (commentsById.Count == 0)
            {
                return;
            }

            using var stream = docEntry.Open();
            using var reader = XmlReader.Create(stream, new XmlReaderSettings { IgnoreWhitespace = true });

            int currentCp = 0;
            var commentStarts = new Dictionary<string, int>(StringComparer.Ordinal);
            var anchoredComments = new HashSet<string>(StringComparer.Ordinal);

            while (reader.Read())
            {
                if (reader.NodeType == XmlNodeType.Element)
                {
                    string localName = reader.LocalName;
                    if (localName == "commentRangeStart")
                    {
                        string? id = reader.GetAttribute("w:id");
                        if (!string.IsNullOrEmpty(id) && commentsById.ContainsKey(id))
                        {
                            commentStarts[id] = currentCp;
                        }
                    }
                    else if (localName == "commentRangeEnd")
                    {
                        string? id = reader.GetAttribute("w:id");
                        if (!string.IsNullOrEmpty(id) &&
                            commentStarts.TryGetValue(id, out int startCp) &&
                            commentsById.TryGetValue(id, out var comment))
                        {
                            comment.StartCp = startCp;
                            comment.EndCp = currentCp;
                            commentStarts.Remove(id);
                            anchoredComments.Add(id);
                        }
                    }
                    else if (localName == "commentReference")
                    {
                        string? id = reader.GetAttribute("w:id");
                        if (!string.IsNullOrEmpty(id) &&
                            !anchoredComments.Contains(id) &&
                            commentsById.TryGetValue(id, out var comment))
                        {
                            if (commentStarts.TryGetValue(id, out int startCp))
                            {
                                comment.StartCp = startCp;
                                comment.EndCp = currentCp;
                                commentStarts.Remove(id);
                            }
                            else
                            {
                                comment.StartCp = currentCp;
                                comment.EndCp = currentCp;
                            }

                            anchoredComments.Add(id);
                        }
                    }
                    else if (localName == "t" || localName == "delText" || localName == "tab" || localName == "ptab" || localName == "br" || localName == "cr" || localName == "noBreakHyphen" || localName == "softHyphen" || localName == "sym")
                    {
                        currentCp += ReadRunTextFragment(reader).Length;
                    }
                }
                else if (reader.NodeType == XmlNodeType.EndElement && reader.LocalName == "p")
                {
                    currentCp += 1;
                }
            }

            foreach (var entry in commentStarts)
            {
                if (!anchoredComments.Contains(entry.Key) &&
                    commentsById.TryGetValue(entry.Key, out var comment))
                {
                    comment.StartCp = entry.Value;
                    comment.EndCp = currentCp;
                }
            }
        }

        private void ParseNumbering(Stream numberingStream, Nedev.FileConverters.DocxToDoc.Model.DocumentModel docModel)
        {
            using var reader = XmlReader.Create(numberingStream, new XmlReaderSettings { IgnoreWhitespace = true });
            Nedev.FileConverters.DocxToDoc.Model.AbstractNumberingModel? currentAbstract = null;
            Nedev.FileConverters.DocxToDoc.Model.NumberingLevelModel? currentLevel = null;
            Nedev.FileConverters.DocxToDoc.Model.NumberingInstanceModel? currentInstance = null;

            while (reader.Read())
            {
                if (reader.NodeType == XmlNodeType.Element)
                {
                    string localName = reader.LocalName;
                    if (localName == "abstractNum")
                    {
                        currentAbstract = new Nedev.FileConverters.DocxToDoc.Model.AbstractNumberingModel
                        {
                            Id = int.Parse(reader.GetAttribute("w:abstractNumId") ?? "0")
                        };
                        docModel.AbstractNumbering.Add(currentAbstract);
                    }
                    else if (localName == "lvl" && currentAbstract != null)
                    {
                        currentLevel = new Nedev.FileConverters.DocxToDoc.Model.NumberingLevelModel
                        {
                            Level = int.Parse(reader.GetAttribute("w:ilvl") ?? "0")
                        };
                        currentAbstract.Levels.Add(currentLevel);
                    }
                    else if (localName == "start" && currentLevel != null)
                    {
                        currentLevel.Start = int.Parse(reader.GetAttribute("w:val") ?? "1");
                    }
                    else if (localName == "numFmt" && currentLevel != null)
                    {
                        currentLevel.NumberFormat = reader.GetAttribute("w:val") ?? "decimal";
                    }
                    else if (localName == "lvlText" && currentLevel != null)
                    {
                        currentLevel.LevelText = reader.GetAttribute("w:val") ?? string.Empty;
                    }
                    else if (localName == "num")
                    {
                        currentInstance = new Nedev.FileConverters.DocxToDoc.Model.NumberingInstanceModel
                        {
                            Id = int.Parse(reader.GetAttribute("w:numId") ?? "0")
                        };
                        docModel.NumberingInstances.Add(currentInstance);
                    }
                    else if (localName == "abstractNumId" && currentInstance != null)
                    {
                        currentInstance.AbstractNumberId = int.Parse(reader.GetAttribute("w:val") ?? "0");
                    }
                }
                else if (reader.NodeType == XmlNodeType.EndElement)
                {
                    if (reader.LocalName == "abstractNum") currentAbstract = null;
                    else if (reader.LocalName == "lvl") currentLevel = null;
                    else if (reader.LocalName == "num") currentInstance = null;
                }
            }
        }

        private void ParseStyles(Stream stylesStream, Nedev.FileConverters.DocxToDoc.Model.DocumentModel docModel)
        {
            using var reader = XmlReader.Create(stylesStream, new XmlReaderSettings { IgnoreWhitespace = true });
            Nedev.FileConverters.DocxToDoc.Model.StyleModel? currentStyle = null;

            while (reader.Read())
            {
                if (reader.NodeType == XmlNodeType.Element)
                {
                    if (reader.LocalName == "style")
                    {
                        currentStyle = new Nedev.FileConverters.DocxToDoc.Model.StyleModel
                        {
                            Id = reader.GetAttribute("w:styleId") ?? string.Empty,
                            IsParagraphStyle = reader.GetAttribute("w:type") == "paragraph"
                        };
                        docModel.Styles.Add(currentStyle);
                    }
                    else if (reader.LocalName == "name" && currentStyle != null)
                    {
                        currentStyle.Name = reader.GetAttribute("w:val") ?? string.Empty;
                    }
                }
                else if (reader.NodeType == XmlNodeType.EndElement && reader.LocalName == "style")
                {
                    currentStyle = null;
                }
            }
        }

        private void ParseFonts(Stream fontStream, Nedev.FileConverters.DocxToDoc.Model.DocumentModel docModel)
        {
            using var reader = XmlReader.Create(fontStream, new XmlReaderSettings { IgnoreWhitespace = true });
            Nedev.FileConverters.DocxToDoc.Model.FontModel? currentFont = null;

            while (reader.Read())
            {
                if (reader.NodeType == XmlNodeType.Element)
                {
                    if (reader.LocalName == "font")
                    {
                        string name = reader.GetAttribute("w:name") ?? string.Empty;
                        if (!string.IsNullOrEmpty(name))
                        {
                            currentFont = new Nedev.FileConverters.DocxToDoc.Model.FontModel { Name = name };
                            docModel.Fonts.Add(currentFont);
                        }
                    }
                    else if (reader.LocalName == "family" && currentFont != null)
                    {
                        string family = reader.GetAttribute("w:val") ?? "auto";
                        currentFont.Family = family switch
                        {
                            "roman" => Nedev.FileConverters.DocxToDoc.Model.FontFamily.Roman,
                            "swiss" => Nedev.FileConverters.DocxToDoc.Model.FontFamily.Swiss,
                            "modern" => Nedev.FileConverters.DocxToDoc.Model.FontFamily.Modern,
                            "script" => Nedev.FileConverters.DocxToDoc.Model.FontFamily.Script,
                            "decorative" => Nedev.FileConverters.DocxToDoc.Model.FontFamily.Decorative,
                            _ => Nedev.FileConverters.DocxToDoc.Model.FontFamily.Auto
                        };
                    }
                    else if (reader.LocalName == "pitch" && currentFont != null)
                    {
                        string pitch = reader.GetAttribute("w:val") ?? "default";
                        currentFont.Pitch = pitch switch
                        {
                            "fixed" => Nedev.FileConverters.DocxToDoc.Model.FontPitch.Fixed,
                            "variable" => Nedev.FileConverters.DocxToDoc.Model.FontPitch.Variable,
                            _ => Nedev.FileConverters.DocxToDoc.Model.FontPitch.Default
                        };
                    }
                    else if (reader.LocalName == "charset" && currentFont != null)
                    {
                        if (byte.TryParse(reader.GetAttribute("w:val"), out byte charset))
                        {
                            currentFont.Charset = charset;
                        }
                    }
                }
                else if (reader.NodeType == XmlNodeType.EndElement && reader.LocalName == "font")
                {
                    currentFont = null;
                }
            }
        }

        private (Nedev.FileConverters.DocxToDoc.Model.ImageModel? image, string textBoxText) ParseDrawing(XmlReader reader)
        {
            var image = new Nedev.FileConverters.DocxToDoc.Model.ImageModel();
            string? relId = null;
            int width = 0;
            int height = 0;
            string? currentPositionAxis = null;
            var textBoxText = new StringBuilder();

            // Read the entire drawing element subtree
            while (reader.Read())
            {
                if (reader.NodeType == XmlNodeType.Element)
                {
                    string localName = reader.LocalName;
                    if (localName == "inline")
                    {
                        image.LayoutType = Nedev.FileConverters.DocxToDoc.Model.ImageLayoutType.Inline;
                        image.WrapType = Nedev.FileConverters.DocxToDoc.Model.ImageWrapType.Inline;
                    }
                    else if (localName == "anchor")
                    {
                        image.LayoutType = Nedev.FileConverters.DocxToDoc.Model.ImageLayoutType.Floating;
                        image.BehindText = IsTrueValue(reader.GetAttribute("behindDoc"));
                        image.AllowOverlap = !IsFalseValue(reader.GetAttribute("allowOverlap"));
                        image.DistanceTopTwips = ConvertEmuToTwips(reader.GetAttribute("distT"));
                        image.DistanceBottomTwips = ConvertEmuToTwips(reader.GetAttribute("distB"));
                        image.DistanceLeftTwips = ConvertEmuToTwips(reader.GetAttribute("distL"));
                        image.DistanceRightTwips = ConvertEmuToTwips(reader.GetAttribute("distR"));
                    }
                    else if (localName == "blip")
                    {
                        relId = reader.GetAttribute("r:embed");
                    }
                    else if (localName == "extent" || localName == "extents")
                    {
                        // CX/CY are in EMUs (English Metric Units), 1 inch = 914400 EMUs
                        if (long.TryParse(reader.GetAttribute("cx"), out long cx))
                        {
                            width = ConvertEmuToPixels(cx);
                        }
                        if (long.TryParse(reader.GetAttribute("cy"), out long cy))
                        {
                            height = ConvertEmuToPixels(cy);
                        }
                    }
                    else if (localName == "positionH")
                    {
                        currentPositionAxis = "H";
                        image.HorizontalRelativeTo = reader.GetAttribute("relativeFrom");
                    }
                    else if (localName == "positionV")
                    {
                        currentPositionAxis = "V";
                        image.VerticalRelativeTo = reader.GetAttribute("relativeFrom");
                    }
                    else if (localName == "posOffset")
                    {
                        int offsetTwips = ConvertEmuToTwips(reader.ReadElementContentAsString());
                        if (currentPositionAxis == "H")
                        {
                            image.PositionXTwips = offsetTwips;
                        }
                        else if (currentPositionAxis == "V")
                        {
                            image.PositionYTwips = offsetTwips;
                        }

                        continue;
                    }
                    else if (localName == "align")
                    {
                        string alignment = reader.ReadElementContentAsString();
                        if (currentPositionAxis == "H")
                        {
                            image.HorizontalAlignment = alignment;
                        }
                        else if (currentPositionAxis == "V")
                        {
                            image.VerticalAlignment = alignment;
                        }

                        continue;
                    }
                    else if (localName == "wrapNone")
                    {
                        image.WrapType = Nedev.FileConverters.DocxToDoc.Model.ImageWrapType.None;
                    }
                    else if (localName == "wrapSquare")
                    {
                        image.WrapType = Nedev.FileConverters.DocxToDoc.Model.ImageWrapType.Square;
                    }
                    else if (localName == "wrapTight")
                    {
                        image.WrapType = Nedev.FileConverters.DocxToDoc.Model.ImageWrapType.Tight;
                    }
                    else if (localName == "wrapThrough")
                    {
                        image.WrapType = Nedev.FileConverters.DocxToDoc.Model.ImageWrapType.Through;
                    }
                    else if (localName == "wrapTopAndBottom")
                    {
                        image.WrapType = Nedev.FileConverters.DocxToDoc.Model.ImageWrapType.TopAndBottom;
                    }
                    else if (localName == "txbxContent")
                    {
                        AppendTextBoxVisibleText(reader, textBoxText);
                        continue;
                    }
                }
                else if (reader.NodeType == XmlNodeType.EndElement && (reader.LocalName == "positionH" || reader.LocalName == "positionV"))
                {
                    currentPositionAxis = null;
                }
                else if (reader.NodeType == XmlNodeType.EndElement && reader.LocalName == "drawing")
                {
                    break;
                }

                // Prevent infinite loop on malformed XML
                if (reader.Depth < 3) break;
            }

            if (string.IsNullOrEmpty(relId))
            {
                return (null, textBoxText.ToString());
            }

            image.RelationshipId = relId;
            image.Width = width;
            image.Height = height;

            return (image, textBoxText.ToString());
        }

        private (Nedev.FileConverters.DocxToDoc.Model.ImageModel? image, string textBoxText) ParsePict(XmlReader reader)
        {
            var image = new Nedev.FileConverters.DocxToDoc.Model.ImageModel
            {
                LayoutType = Nedev.FileConverters.DocxToDoc.Model.ImageLayoutType.Inline,
                WrapType = Nedev.FileConverters.DocxToDoc.Model.ImageWrapType.Inline
            };
            string? relId = null;
            var textBoxText = new StringBuilder();

            while (reader.Read())
            {
                if (reader.NodeType == XmlNodeType.Element)
                {
                    string localName = reader.LocalName;
                    if (localName == "shape")
                    {
                        string? style = reader.GetAttribute("style");
                        ParseShapeStyle(style, image);
                    }
                    else if (localName == "imagedata")
                    {
                        relId = reader.GetAttribute("r:id");
                    }
                    else if (localName == "txbxContent")
                    {
                        AppendTextBoxVisibleText(reader, textBoxText);
                        continue;
                    }
                }
                else if (reader.NodeType == XmlNodeType.EndElement && reader.LocalName == "pict")
                {
                    break;
                }
                
                if (reader.Depth < 3) break;
            }

            if (string.IsNullOrEmpty(relId))
            {
                return (null, textBoxText.ToString());
            }
            image.RelationshipId = relId;
            return (image, textBoxText.ToString());
        }

        private static void AppendTextBoxVisibleText(XmlReader reader, StringBuilder textBoxText)
        {
            if (reader.IsEmptyElement)
            {
                return;
            }

            int textBoxDepth = reader.Depth;
            bool seenParagraph = false;

            while (reader.Read())
            {
                if (reader.NodeType == XmlNodeType.Element)
                {
                    string localName = reader.LocalName;
                    if (localName == "p")
                    {
                        if (seenParagraph)
                        {
                            AppendTextBoxSeparator(textBoxText, ' ');
                        }

                        seenParagraph = true;
                    }
                    else if (localName == "t" || localName == "delText" || localName == "tab" || localName == "ptab" || localName == "br" || localName == "cr" || localName == "noBreakHyphen" || localName == "softHyphen" || localName == "sym")
                    {
                        string text = ReadRunTextFragment(reader);
                        if (!string.IsNullOrEmpty(text))
                        {
                            textBoxText.Append(text);
                        }

                        continue;
                    }
                }
                else if (reader.NodeType == XmlNodeType.EndElement && reader.LocalName == "txbxContent")
                {
                    break;
                }

                if (reader.Depth < textBoxDepth)
                {
                    break;
                }
            }
        }

        private static void AppendTextBoxSeparator(StringBuilder textBoxText, char separator)
        {
            if (textBoxText.Length == 0)
            {
                return;
            }

            char lastChar = textBoxText[textBoxText.Length - 1];
            if (!char.IsWhiteSpace(lastChar))
            {
                textBoxText.Append(separator);
            }
        }

        private static void ParseShapeStyle(string? style, Nedev.FileConverters.DocxToDoc.Model.ImageModel image)
        {
            if (string.IsNullOrWhiteSpace(style)) return;
            var parts = style.Split(';', StringSplitOptions.RemoveEmptyEntries);
            foreach (var part in parts)
            {
                var kv = part.Split(':', 2);
                if (kv.Length == 2)
                {
                    string key = kv[0].Trim().ToLowerInvariant();
                    string val = kv[1].Trim().ToLowerInvariant();
                    if (key == "width" && val.EndsWith("pt") && double.TryParse(val.TrimEnd('p', 't'), out double w))
                    {
                        image.Width = (int)Math.Round(w * 96.0 / 72.0);
                    }
                    else if (key == "height" && val.EndsWith("pt") && double.TryParse(val.TrimEnd('p', 't'), out double h))
                    {
                        image.Height = (int)Math.Round(h * 96.0 / 72.0);
                    }
                }
            }
        }

        private static int ConvertEmuToPixels(long value)
        {
            return (int)Math.Round(value / 914400.0 * 96.0, MidpointRounding.AwayFromZero);
        }

        private static int ConvertEmuToTwips(string? value)
        {
            return long.TryParse(value, out long parsedValue) ? ConvertEmuToTwips(parsedValue) : 0;
        }

        private static int ConvertEmuToTwips(long value)
        {
            return (int)Math.Round(value / 635.0, MidpointRounding.AwayFromZero);
        }

        private static bool IsTrueValue(string? value)
        {
            return value == "1" || string.Equals(value, "true", StringComparison.OrdinalIgnoreCase);
        }

        private static bool IsFalseValue(string? value)
        {
            return value == "0" || string.Equals(value, "false", StringComparison.OrdinalIgnoreCase);
        }

        private void LoadImageData(Nedev.FileConverters.DocxToDoc.Model.ImageModel image)
        {
            LoadImageData(image, _relationships, "word/document.xml");
        }

        private void LoadImageData(
            Nedev.FileConverters.DocxToDoc.Model.ImageModel image,
            IReadOnlyDictionary<string, string> relationships,
            string sourcePartEntryPath)
        {
            if (!relationships.TryGetValue(image.RelationshipId, out string? target))
                return;

            string imagePath = ResolvePartTargetPath(sourcePartEntryPath, target);

            var entry = _archive.GetEntry(imagePath);
            if (entry == null) return;

            using var stream = entry.Open();
            using var ms = new MemoryStream();
            stream.CopyTo(ms);
            image.Data = ms.ToArray();
            image.FileName = Path.GetFileName(target);

            // Determine content type from extension
            string ext = Path.GetExtension(target).ToLowerInvariant();
            image.ContentType = ext switch
            {
                ".png" => "image/png",
                ".jpg" or ".jpeg" => "image/jpeg",
                ".gif" => "image/gif",
                ".bmp" => "image/bmp",
                ".tiff" or ".tif" => "image/tiff",
                ".wmf" => "image/x-wmf",
                ".emf" => "image/x-emf",
                _ => "application/octet-stream"
            };
        }

        private static Nedev.FileConverters.DocxToDoc.Model.FieldType ParseFieldType(string instruction)
        {
            if (string.IsNullOrWhiteSpace(instruction))
                return Nedev.FileConverters.DocxToDoc.Model.FieldType.Unknown;

            // Extract the field name from the instruction
            string fieldName = instruction.Trim().Split(' ')[0].ToUpperInvariant();

            return fieldName switch
            {
                "PAGE" => Nedev.FileConverters.DocxToDoc.Model.FieldType.Page,
                "NUMPAGES" => Nedev.FileConverters.DocxToDoc.Model.FieldType.NumPages,
                "SECTIONPAGES" => Nedev.FileConverters.DocxToDoc.Model.FieldType.SectionPages,
                "DATE" => Nedev.FileConverters.DocxToDoc.Model.FieldType.Date,
                "TIME" => Nedev.FileConverters.DocxToDoc.Model.FieldType.Time,
                "AUTHOR" => Nedev.FileConverters.DocxToDoc.Model.FieldType.Author,
                "TITLE" => Nedev.FileConverters.DocxToDoc.Model.FieldType.Title,
                "SUBJECT" => Nedev.FileConverters.DocxToDoc.Model.FieldType.Subject,
                "FILENAME" => Nedev.FileConverters.DocxToDoc.Model.FieldType.FileName,
                "HYPERLINK" => Nedev.FileConverters.DocxToDoc.Model.FieldType.Hyperlink,
                "BOOKMARK" => Nedev.FileConverters.DocxToDoc.Model.FieldType.Bookmark,
                "INDEX" => Nedev.FileConverters.DocxToDoc.Model.FieldType.Index,
                "SEQ" => Nedev.FileConverters.DocxToDoc.Model.FieldType.Seq,
                "REF" => Nedev.FileConverters.DocxToDoc.Model.FieldType.Ref,
                "MERGEFIELD" => Nedev.FileConverters.DocxToDoc.Model.FieldType.MergeField,
                _ => Nedev.FileConverters.DocxToDoc.Model.FieldType.Unknown
            };
        }

        private void ParseDocumentProperties(Stream propsStream, Nedev.FileConverters.DocxToDoc.Model.DocumentModel docModel)
        {
            using var reader = XmlReader.Create(propsStream, new XmlReaderSettings { IgnoreWhitespace = true });
            var props = docModel.Properties;

            while (!reader.EOF)
            {
                if (reader.NodeType != XmlNodeType.Element)
                {
                    reader.Read();
                    continue;
                }

                string normalizedLocalName = reader.LocalName.ToLowerInvariant();

                // Core properties namespace: http://purl.org/dc/elements/1.1/ or http://schemas.openxmlformats.org/package/2006/metadata/core-properties
                if (normalizedLocalName == "title")
                {
                    props.Title = reader.ReadElementContentAsString();
                }
                else if (normalizedLocalName == "subject")
                {
                    props.Subject = reader.ReadElementContentAsString();
                }
                else if (normalizedLocalName == "creator" || normalizedLocalName == "author")
                {
                    props.Author = reader.ReadElementContentAsString();
                }
                else if (normalizedLocalName == "keywords")
                {
                    props.Keywords = reader.ReadElementContentAsString();
                }
                else if (normalizedLocalName == "description" || normalizedLocalName == "comments")
                {
                    props.Comments = reader.ReadElementContentAsString();
                }
                else if (normalizedLocalName == "created")
                {
                    if (DateTime.TryParse(reader.ReadElementContentAsString(), out DateTime created))
                    {
                        props.Created = created;
                    }
                }
                else if (normalizedLocalName == "modified")
                {
                    if (DateTime.TryParse(reader.ReadElementContentAsString(), out DateTime modified))
                    {
                        props.Modified = modified;
                    }
                }
                else if (normalizedLocalName == "lastprinted")
                {
                    if (DateTime.TryParse(reader.ReadElementContentAsString(), out DateTime printed))
                    {
                        props.LastPrinted = printed;
                    }
                }
                else if (normalizedLocalName == "revision")
                {
                    if (int.TryParse(reader.ReadElementContentAsString(), out int revision))
                    {
                        props.Revision = revision;
                    }
                }
                else if (normalizedLocalName == "category")
                {
                    props.Category = reader.ReadElementContentAsString();
                }
                else if (normalizedLocalName == "manager")
                {
                    props.Manager = reader.ReadElementContentAsString();
                }
                else if (normalizedLocalName == "company")
                {
                    props.Company = reader.ReadElementContentAsString();
                }
                else if (normalizedLocalName == "totaltime")
                {
                    if (int.TryParse(reader.ReadElementContentAsString(), out int totalTime))
                    {
                        props.TotalEditingTime = totalTime;
                    }
                }
                else if (normalizedLocalName == "pages")
                {
                    if (int.TryParse(reader.ReadElementContentAsString(), out int pages))
                    {
                        props.Pages = pages;
                    }
                }
                else if (normalizedLocalName == "words")
                {
                    if (int.TryParse(reader.ReadElementContentAsString(), out int words))
                    {
                        props.Words = words;
                    }
                }
                else if (normalizedLocalName == "characters")
                {
                    if (int.TryParse(reader.ReadElementContentAsString(), out int characters))
                    {
                        props.Characters = characters;
                    }
                }
                else
                {
                    reader.Read();
                }
            }
        }

        private static void ParseFootnotes(Stream footnotesStream, Nedev.FileConverters.DocxToDoc.Model.DocumentModel docModel)
        {
            ParseNotes(
                footnotesStream,
                docModel,
                docModel.Footnotes,
                "footnote",
                static id => new Nedev.FileConverters.DocxToDoc.Model.FootnoteModel { Id = id },
                static (footnote, text) => footnote.Text = text,
                static (document, text) => document.FootnoteSeparatorText = text,
                static (document, text) => document.FootnoteContinuationSeparatorText = text,
                static (document, text) => document.FootnoteContinuationNoticeText = text);
        }

        private static void ParseEndnotes(Stream endnotesStream, Nedev.FileConverters.DocxToDoc.Model.DocumentModel docModel)
        {
            ParseNotes(
                endnotesStream,
                docModel,
                docModel.Endnotes,
                "endnote",
                static id => new Nedev.FileConverters.DocxToDoc.Model.EndnoteModel { Id = id },
                static (endnote, text) => endnote.Text = text,
                static (document, text) => document.EndnoteSeparatorText = text,
                static (document, text) => document.EndnoteContinuationSeparatorText = text,
                static (document, text) => document.EndnoteContinuationNoticeText = text);
        }

        private static void ParseNotes<TNote>(
            Stream notesStream,
            Nedev.FileConverters.DocxToDoc.Model.DocumentModel docModel,
            ICollection<TNote> notes,
            string noteElementName,
            Func<string, TNote> createNote,
            Action<TNote, string> setText,
            Action<Nedev.FileConverters.DocxToDoc.Model.DocumentModel, string> setSeparatorText,
            Action<Nedev.FileConverters.DocxToDoc.Model.DocumentModel, string> setContinuationSeparatorText,
            Action<Nedev.FileConverters.DocxToDoc.Model.DocumentModel, string> setContinuationNoticeText)
            where TNote : class
        {
            using var reader = XmlReader.Create(notesStream, new XmlReaderSettings { IgnoreWhitespace = true });
            TNote? currentNote = null;
            NoteCaptureKind currentNoteKind = NoteCaptureKind.Ignore;
            var currentText = new StringBuilder();
            int currentNoteParagraphCount = 0;

            while (reader.Read())
            {
                if (reader.NodeType == XmlNodeType.Element)
                {
                    string localName = reader.LocalName;
                    if (localName == noteElementName)
                    {
                        string? id = reader.GetAttribute("w:id");
                        string? type = reader.GetAttribute("w:type");
                        currentText.Clear();
                        currentNoteParagraphCount = 0;
                        currentNoteKind = ResolveNoteCaptureKind(id, type);
                        if (currentNoteKind == NoteCaptureKind.Regular && !string.IsNullOrEmpty(id))
                        {
                            currentNote = createNote(id);
                        }
                        else
                        {
                            currentNote = null;
                        }
                    }
                    else if (localName == "p" && currentNoteKind != NoteCaptureKind.Ignore)
                    {
                        if (currentNoteParagraphCount > 0)
                        {
                            currentText.Append('\r');
                        }

                        currentNoteParagraphCount++;
                    }
                    else if (IsNoteTextFragmentElement(localName) && currentNoteKind != NoteCaptureKind.Ignore)
                    {
                        currentText.Append(ReadRunTextFragment(reader));
                    }
                }
                else if (reader.NodeType == XmlNodeType.EndElement && reader.LocalName == noteElementName)
                {
                    string capturedText = currentText.ToString();
                    switch (currentNoteKind)
                    {
                        case NoteCaptureKind.Regular when currentNote != null:
                            setText(currentNote, capturedText);
                            notes.Add(currentNote);
                            break;
                        case NoteCaptureKind.Separator:
                            setSeparatorText(docModel, capturedText);
                            break;
                        case NoteCaptureKind.ContinuationSeparator:
                            setContinuationSeparatorText(docModel, capturedText);
                            break;
                        case NoteCaptureKind.ContinuationNotice:
                            setContinuationNoticeText(docModel, capturedText);
                            break;
                    }

                    currentNote = null;
                    currentNoteKind = NoteCaptureKind.Ignore;
                    currentText.Clear();
                    currentNoteParagraphCount = 0;
                }
            }
        }

        private static NoteCaptureKind ResolveNoteCaptureKind(string? id, string? type)
        {
            return type switch
            {
                "separator" => NoteCaptureKind.Separator,
                "continuationSeparator" => NoteCaptureKind.ContinuationSeparator,
                "continuationNotice" => NoteCaptureKind.ContinuationNotice,
                _ when !string.IsNullOrEmpty(id) && string.IsNullOrEmpty(type) && !IsSpecialNoteId(id) => NoteCaptureKind.Regular,
                _ => NoteCaptureKind.Ignore
            };
        }

        private static bool IsNoteTextFragmentElement(string localName)
        {
            return localName == "t" ||
                   localName == "delText" ||
                   localName == "tab" ||
                   localName == "ptab" ||
                   localName == "br" ||
                   localName == "cr" ||
                   localName == "noBreakHyphen" ||
                   localName == "softHyphen" ||
                   localName == "sym";
        }

        private static bool IsSpecialNoteId(string? id)
        {
            return int.TryParse(id, NumberStyles.Integer, CultureInfo.InvariantCulture, out int parsedId) &&
                   parsedId <= 1;
        }

        private void ParseComments(
            Stream commentsStream,
            Nedev.FileConverters.DocxToDoc.Model.DocumentModel docModel,
            IDictionary<string, string> commentParagraphIds)
        {
            using var reader = XmlReader.Create(commentsStream, new XmlReaderSettings { IgnoreWhitespace = true });
            Nedev.FileConverters.DocxToDoc.Model.CommentModel? currentComment = null;
            string? currentCommentLastParagraphId = null;
            int currentCommentParagraphCount = 0;

            while (reader.Read())
            {
                if (reader.NodeType == XmlNodeType.Element)
                {
                    string localName = reader.LocalName;

                    if (localName == "comment")
                    {
                        currentComment = new Nedev.FileConverters.DocxToDoc.Model.CommentModel();
                        currentCommentLastParagraphId = null;
                        currentCommentParagraphCount = 0;

                        string? id = reader.GetAttribute("w:id");
                        if (!string.IsNullOrEmpty(id))
                            currentComment.Id = id;

                        string? author = reader.GetAttribute("w:author");
                        if (!string.IsNullOrEmpty(author))
                            currentComment.Author = author;

                        string? initials = reader.GetAttribute("w:initials");
                        if (!string.IsNullOrEmpty(initials))
                            currentComment.Initials = initials;

                        string? dateStr = reader.GetAttribute("w:date");
                        if (!string.IsNullOrEmpty(dateStr) && DateTime.TryParse(dateStr, out DateTime date))
                            currentComment.Date = date;

                        string? doneStr = reader.GetAttribute("w:done");
                        if (!string.IsNullOrEmpty(doneStr))
                            currentComment.IsDone = doneStr == "1" || doneStr.Equals("true", StringComparison.OrdinalIgnoreCase);
                    }
                    else if (localName == "p" && currentComment != null)
                    {
                        if (currentCommentParagraphCount > 0)
                        {
                            currentComment.Text += "\r";
                        }

                        currentCommentParagraphCount++;

                        string? paragraphId = GetAttributeByLocalName(reader, "paraId");
                        if (!string.IsNullOrEmpty(paragraphId))
                        {
                            currentCommentLastParagraphId = paragraphId;
                        }
                    }
                    else if ((localName == "t" || localName == "delText" || localName == "tab" || localName == "ptab" || localName == "br" || localName == "cr" || localName == "noBreakHyphen" || localName == "softHyphen" || localName == "sym") && currentComment != null)
                    {
                        currentComment.Text += ReadRunTextFragment(reader);
                    }
                }
                else if (reader.NodeType == XmlNodeType.EndElement)
                {
                    if (reader.LocalName == "comment" && currentComment != null)
                    {
                        docModel.Comments.Add(currentComment);
                        if (!string.IsNullOrEmpty(currentComment.Id) && !string.IsNullOrEmpty(currentCommentLastParagraphId))
                        {
                            commentParagraphIds[currentComment.Id] = currentCommentLastParagraphId;
                        }
                        currentComment = null;
                        currentCommentLastParagraphId = null;
                        currentCommentParagraphCount = 0;
                    }
                }
            }
        }

        private static void ParseCommentsExtended(
            Stream commentsExtendedStream,
            Nedev.FileConverters.DocxToDoc.Model.DocumentModel docModel,
            IReadOnlyDictionary<string, string> commentParagraphIds)
        {
            var commentsById = docModel.Comments
                .Where(comment => !string.IsNullOrEmpty(comment.Id))
                .ToDictionary(comment => comment.Id, StringComparer.Ordinal);
            if (commentsById.Count == 0)
            {
                return;
            }

            var commentIdsByParagraphId = new Dictionary<string, string>(StringComparer.Ordinal);
            foreach (var entry in commentParagraphIds)
            {
                if (!string.IsNullOrEmpty(entry.Key) && !string.IsNullOrEmpty(entry.Value))
                {
                    commentIdsByParagraphId[entry.Value] = entry.Key;
                }
            }

            using var reader = XmlReader.Create(commentsExtendedStream, new XmlReaderSettings { IgnoreWhitespace = true });
            while (reader.Read())
            {
                if (reader.NodeType != XmlNodeType.Element || reader.LocalName != "commentEx")
                {
                    continue;
                }

                string? paragraphId = GetAttributeByLocalName(reader, "paraId");
                if (string.IsNullOrEmpty(paragraphId) ||
                    !commentIdsByParagraphId.TryGetValue(paragraphId, out string? commentId) ||
                    !commentsById.TryGetValue(commentId, out var comment))
                {
                    continue;
                }

                string? done = GetAttributeByLocalName(reader, "done");
                if (!string.IsNullOrEmpty(done))
                {
                    comment.IsDone = IsTrueValue(done);
                }

                string? parentParagraphId = GetAttributeByLocalName(reader, "paraIdParent");
                if (!string.IsNullOrEmpty(parentParagraphId))
                {
                    comment.IsReply = true;
                    if (commentIdsByParagraphId.TryGetValue(parentParagraphId, out string? parentId))
                    {
                        comment.ParentId = parentId;
                    }
                }
            }
        }

        private static string? GetAttributeByLocalName(XmlReader reader, string localName)
        {
            if (!reader.HasAttributes)
            {
                return null;
            }

            for (int index = 0; index < reader.AttributeCount; index++)
            {
                reader.MoveToAttribute(index);
                if (reader.LocalName == localName)
                {
                    string value = reader.Value;
                    reader.MoveToElement();
                    return value;
                }
            }

            reader.MoveToElement();
            return null;
        }

        protected virtual void Dispose(bool disposing)
        {
            if (!_disposedValue)
            {
                if (disposing)
                {
                    _archive?.Dispose();
                }
                _disposedValue = true;
            }
        }

        public void Dispose()
        {
            Dispose(disposing: true);
            GC.SuppressFinalize(this);
        }
    }
}
