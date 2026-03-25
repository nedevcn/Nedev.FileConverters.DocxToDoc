using System;
using System.IO;
using System.IO.Compression;
using System.Xml;
using System.Text;
using System.Collections.Generic;
using System.Linq;

namespace Nedev.FileConverters.DocxToDoc.Format
{
    /// <summary>
    /// Reads and coordinates the extraction of features from OpenXML (.docx) files.
    /// Optimized for low-memory, forward-only reading.
    /// </summary>
    public class DocxReader : IDisposable
    {
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
            var rels = new Dictionary<string, string>();
            var relsEntry = _archive.GetEntry("word/_rels/document.xml.rels");
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

            // Parse Comments
            var commentsEntry = _archive.GetEntry("word/comments.xml");
            if (commentsEntry != null)
            {
                using var commentsStream = commentsEntry.Open();
                ParseComments(commentsStream, docModel);
            }

            StringBuilder textBuffer = new StringBuilder();

            Nedev.FileConverters.DocxToDoc.Model.ParagraphModel currentParagraph = null;
            Nedev.FileConverters.DocxToDoc.Model.RunModel currentRun = null;
            Nedev.FileConverters.DocxToDoc.Model.SectionModel currentSection = null;
            Nedev.FileConverters.DocxToDoc.Model.TableModel currentTable = null;
            Nedev.FileConverters.DocxToDoc.Model.TableRowModel currentRow = null;
            Nedev.FileConverters.DocxToDoc.Model.TableCellModel currentCell = null;
            int currentRowGridColumnIndex = 0;
            bool insideTableCellMargins = false;
            bool insideCellMargins = false;
            bool insideTableBorders = false;
            bool insideCellBorders = false;
            var openFields = new Stack<Nedev.FileConverters.DocxToDoc.Model.FieldModel>();

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
                    else if (localName == "tcW" && currentCell != null)
                    {
                        string? type = xmlReader.GetAttribute("w:type");
                        if ((string.IsNullOrEmpty(type) || string.Equals(type, "dxa", StringComparison.OrdinalIgnoreCase)) &&
                            int.TryParse(xmlReader.GetAttribute("w:w"), out int width))
                        {
                            currentCell.Width = width;
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
                    else if (localName == "gridSpan" && currentCell != null)
                    {
                        if (int.TryParse(xmlReader.GetAttribute("w:val"), out int gridSpan) && gridSpan > 0)
                        {
                            currentCell.GridSpan = gridSpan;
                        }
                    }
                    else if (localName == "vAlign" && currentCell != null)
                    {
                        string? value = xmlReader.GetAttribute("w:val");
                        currentCell.VerticalAlignment = value switch
                        {
                            "center" => Nedev.FileConverters.DocxToDoc.Model.TableCellVerticalAlignment.Center,
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
                    else if (localName == "insideH" && insideTableBorders && currentTable != null && TryReadBorderWidthTwips(xmlReader, out int insideHorizontalBorderTwips))
                    {
                        currentTable.DefaultInsideHorizontalBorderTwips = insideHorizontalBorderTwips;
                    }
                    else if (localName == "insideV" && insideTableBorders && currentTable != null && TryReadBorderWidthTwips(xmlReader, out int insideVerticalBorderTwips))
                    {
                        currentTable.DefaultInsideVerticalBorderTwips = insideVerticalBorderTwips;
                    }
                    else if ((localName == "left" || localName == "start") && TryReadBorderWidthTwips(xmlReader, out int leftBorderTwips))
                    {
                        if (insideCellBorders && currentCell != null)
                        {
                            currentCell.BorderLeftTwips = leftBorderTwips;
                        }
                        else if (insideTableBorders && currentTable != null)
                        {
                            currentTable.DefaultBorderLeftTwips = leftBorderTwips;
                        }
                    }
                    else if ((localName == "right" || localName == "end") && TryReadBorderWidthTwips(xmlReader, out int rightBorderTwips))
                    {
                        if (insideCellBorders && currentCell != null)
                        {
                            currentCell.BorderRightTwips = rightBorderTwips;
                        }
                        else if (insideTableBorders && currentTable != null)
                        {
                            currentTable.DefaultBorderRightTwips = rightBorderTwips;
                        }
                    }
                    else if (localName == "top" && TryReadBorderWidthTwips(xmlReader, out int topBorderTwips))
                    {
                        if (insideCellBorders && currentCell != null)
                        {
                            currentCell.BorderTopTwips = topBorderTwips;
                        }
                        else if (insideTableBorders && currentTable != null)
                        {
                            currentTable.DefaultBorderTopTwips = topBorderTwips;
                        }
                    }
                    else if (localName == "bottom" && TryReadBorderWidthTwips(xmlReader, out int bottomBorderTwips))
                    {
                        if (insideCellBorders && currentCell != null)
                        {
                            currentCell.BorderBottomTwips = bottomBorderTwips;
                        }
                        else if (insideTableBorders && currentTable != null)
                        {
                            currentTable.DefaultBorderBottomTwips = bottomBorderTwips;
                        }
                    }
                    else if ((localName == "left" || localName == "start") && TryReadDxaWidth(xmlReader, out int leftPaddingTwips))
                    {
                        if (insideCellMargins && currentCell != null)
                        {
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
                            currentCell.PaddingBottomTwips = bottomPaddingTwips;
                        }
                        else if (insideTableCellMargins && currentTable != null)
                        {
                            currentTable.DefaultCellPaddingBottomTwips = bottomPaddingTwips;
                        }
                    }
                    else if (localName == "p")
                    {
                        currentParagraph = new Nedev.FileConverters.DocxToDoc.Model.ParagraphModel();
                        if (currentCell != null) currentCell.Paragraphs.Add(currentParagraph);
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
                        string val = xmlReader.GetAttribute("w:val");
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
                    else if (localName == "r" && currentParagraph != null)
                    {
                        currentRun = new Nedev.FileConverters.DocxToDoc.Model.RunModel();
                        currentParagraph.Runs.Add(currentRun);
                    }
                    else if (localName == "b" && currentRun != null)
                    {
                        string val = xmlReader.GetAttribute("w:val");
                        currentRun.Properties.IsBold = (val != "0" && val != "false");
                    }
                    else if (localName == "i" && currentRun != null)
                    {
                        string val = xmlReader.GetAttribute("w:val");
                        currentRun.Properties.IsItalic = (val != "0" && val != "false");
                    }
                    else if (localName == "strike" && currentRun != null)
                    {
                        string val = xmlReader.GetAttribute("w:val");
                        currentRun.Properties.IsStrike = (val != "0" && val != "false");
                    }
                    else if (localName == "u" && currentRun != null)
                    {
                        string val = xmlReader.GetAttribute("w:val") ?? "none";
                        currentRun.Properties.Underline = val switch
                        {
                            "single" => Nedev.FileConverters.DocxToDoc.Model.UnderlineType.Single,
                            "double" => Nedev.FileConverters.DocxToDoc.Model.UnderlineType.Double,
                            "thick" => Nedev.FileConverters.DocxToDoc.Model.UnderlineType.Thick,
                            "dotted" => Nedev.FileConverters.DocxToDoc.Model.UnderlineType.Dotted,
                            "dashed" => Nedev.FileConverters.DocxToDoc.Model.UnderlineType.Dashed,
                            "wave" => Nedev.FileConverters.DocxToDoc.Model.UnderlineType.Wave,
                            _ => Nedev.FileConverters.DocxToDoc.Model.UnderlineType.None
                        };
                    }
                    else if (localName == "color" && currentRun != null)
                    {
                        string? val = xmlReader.GetAttribute("w:val");
                        if (!string.IsNullOrEmpty(val) && val != "auto")
                        {
                            currentRun.Properties.Color = val;
                        }
                    }
                    else if (localName == "rFonts" && currentRun != null)
                    {
                        string? ascii = xmlReader.GetAttribute("w:ascii");
                        if (!string.IsNullOrEmpty(ascii))
                        {
                            currentRun.Properties.FontName = ascii;
                        }
                    }
                    else if (localName == "sz" && currentRun != null)
                    {
                        string val = xmlReader.GetAttribute("w:val");
                        if (int.TryParse(val, out int size))
                        {
                            currentRun.Properties.FontSize = size;
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
                                currentParagraph.Runs.Add(run);

                                // Parse run properties and text
                                while (xmlReader.Read())
                                {
                                    if (xmlReader.NodeType == XmlNodeType.Element)
                                    {
                                        if (xmlReader.LocalName == "t")
                                        {
                                            string text = xmlReader.ReadElementContentAsString();
                                            run.Text = text;
                                            hyperlink.DisplayText += text;
                                            textBuffer.Append(text);
                                        }
                                        else if (xmlReader.LocalName == "b")
                                        {
                                            string val = xmlReader.GetAttribute("w:val");
                                            run.Properties.IsBold = (val != "0" && val != "false");
                                        }
                                        else if (xmlReader.LocalName == "i")
                                        {
                                            string val = xmlReader.GetAttribute("w:val");
                                            run.Properties.IsItalic = (val != "0" && val != "false");
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
                    else if (localName == "t" && currentRun != null)
                    {
                        string text = xmlReader.ReadElementContentAsString();
                        currentRun.Text = text;
                        textBuffer.Append(text);
                    }
                    else if (localName == "drawing" && currentRun != null)
                    {
                        // Parse inline or anchored image
                        var image = ParseDrawing(xmlReader);
                        if (image != null)
                        {
                            currentRun.Image = image;
                            // Load actual image data
                            LoadImageData(image);
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
                                currentRun.Field = new Nedev.FileConverters.DocxToDoc.Model.FieldModel
                                {
                                    IsLocked = fldLock == "1" || fldLock == "true",
                                    IsDirty = fldDirty == "1" || fldDirty == "true"
                                };
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
                        string instruction = xmlReader.ReadElementContentAsString();
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
                            if (currentCell.Width <= 0 && currentTable != null)
                            {
                                currentCell.Width = ResolveGridWidth(currentTable.GridColumnWidths, currentRowGridColumnIndex, currentCell.GridSpan);
                            }

                            currentRowGridColumnIndex += Math.Max(1, currentCell.GridSpan);
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
                    }
                }
            }

            docModel.TextBuffer = textBuffer.ToString();

            // Parse bookmarks after document is fully read
            ParseBookmarks(docModel);

            return docModel;
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
            if (!string.IsNullOrEmpty(type) && !string.Equals(type, "dxa", StringComparison.OrdinalIgnoreCase))
            {
                width = 0;
                return false;
            }

            return int.TryParse(xmlReader.GetAttribute("w:w"), out width);
        }

        private static bool TryReadBorderWidthTwips(XmlReader xmlReader, out int width)
        {
            string? borderValue = xmlReader.GetAttribute("w:val");
            if (string.Equals(borderValue, "nil", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(borderValue, "none", StringComparison.OrdinalIgnoreCase))
            {
                width = 0;
                return true;
            }

            if (!int.TryParse(xmlReader.GetAttribute("w:sz"), out int eighthPointWidth))
            {
                width = 0;
                return false;
            }

            width = (int)Math.Round(Math.Max(0, eighthPointWidth) * 2.5d, MidpointRounding.AwayFromZero);
            return true;
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
                    else if (localName == "t")
                    {
                        // Add text length to current CP
                        string text = reader.ReadElementContentAsString();
                        currentCp += text.Length;
                    }
                    else if (localName == "p")
                    {
                        // Paragraph mark
                        currentCp += 1;
                    }
                }
            }
        }

        private void ParseNumbering(Stream numberingStream, Nedev.FileConverters.DocxToDoc.Model.DocumentModel docModel)
        {
            using var reader = XmlReader.Create(numberingStream, new XmlReaderSettings { IgnoreWhitespace = true });
            Nedev.FileConverters.DocxToDoc.Model.AbstractNumberingModel currentAbstract = null;
            Nedev.FileConverters.DocxToDoc.Model.NumberingLevelModel currentLevel = null;
            Nedev.FileConverters.DocxToDoc.Model.NumberingInstanceModel currentInstance = null;

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
            Nedev.FileConverters.DocxToDoc.Model.StyleModel currentStyle = null;

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

        private Nedev.FileConverters.DocxToDoc.Model.ImageModel? ParseDrawing(XmlReader reader)
        {
            var image = new Nedev.FileConverters.DocxToDoc.Model.ImageModel();
            string? relId = null;
            int width = 0;
            int height = 0;
            string? currentPositionAxis = null;

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

            if (string.IsNullOrEmpty(relId)) return null;

            image.RelationshipId = relId;
            image.Width = width;
            image.Height = height;

            return image;
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
            if (!_relationships.TryGetValue(image.RelationshipId, out string? target))
                return;

            // Resolve target path
            string imagePath = target.StartsWith("/") ? target.Substring(1) : $"word/{target}";

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

            while (reader.Read())
            {
                if (reader.NodeType == XmlNodeType.Element)
                {
                    string localName = reader.LocalName;
                    string? ns = reader.NamespaceURI;

                    // Core properties namespace: http://purl.org/dc/elements/1.1/ or http://schemas.openxmlformats.org/package/2006/metadata/core-properties
                    if (localName == "title")
                    {
                        props.Title = reader.ReadElementContentAsString();
                    }
                    else if (localName == "subject")
                    {
                        props.Subject = reader.ReadElementContentAsString();
                    }
                    else if (localName == "creator" || localName == "author")
                    {
                        props.Author = reader.ReadElementContentAsString();
                    }
                    else if (localName == "keywords")
                    {
                        props.Keywords = reader.ReadElementContentAsString();
                    }
                    else if (localName == "description" || localName == "comments")
                    {
                        props.Comments = reader.ReadElementContentAsString();
                    }
                    else if (localName == "created")
                    {
                        if (DateTime.TryParse(reader.ReadElementContentAsString(), out DateTime created))
                        {
                            props.Created = created;
                        }
                    }
                    else if (localName == "modified")
                    {
                        if (DateTime.TryParse(reader.ReadElementContentAsString(), out DateTime modified))
                        {
                            props.Modified = modified;
                        }
                    }
                    else if (localName == "lastPrinted")
                    {
                        if (DateTime.TryParse(reader.ReadElementContentAsString(), out DateTime printed))
                        {
                            props.LastPrinted = printed;
                        }
                    }
                    else if (localName == "revision")
                    {
                        if (int.TryParse(reader.ReadElementContentAsString(), out int revision))
                        {
                            props.Revision = revision;
                        }
                    }
                    else if (localName == "category")
                    {
                        props.Category = reader.ReadElementContentAsString();
                    }
                    else if (localName == "manager")
                    {
                        props.Manager = reader.ReadElementContentAsString();
                    }
                    else if (localName == "company")
                    {
                        props.Company = reader.ReadElementContentAsString();
                    }
                    else if (localName == "totalTime")
                    {
                        if (int.TryParse(reader.ReadElementContentAsString(), out int totalTime))
                        {
                            props.TotalEditingTime = totalTime;
                        }
                    }
                    else if (localName == "pages")
                    {
                        if (int.TryParse(reader.ReadElementContentAsString(), out int pages))
                        {
                            props.Pages = pages;
                        }
                    }
                    else if (localName == "words")
                    {
                        if (int.TryParse(reader.ReadElementContentAsString(), out int words))
                        {
                            props.Words = words;
                        }
                    }
                    else if (localName == "characters")
                    {
                        if (int.TryParse(reader.ReadElementContentAsString(), out int characters))
                        {
                            props.Characters = characters;
                        }
                    }
                }
            }
        }

        private void ParseComments(Stream commentsStream, Nedev.FileConverters.DocxToDoc.Model.DocumentModel docModel)
        {
            using var reader = XmlReader.Create(commentsStream, new XmlReaderSettings { IgnoreWhitespace = true });
            Nedev.FileConverters.DocxToDoc.Model.CommentModel? currentComment = null;

            while (reader.Read())
            {
                if (reader.NodeType == XmlNodeType.Element)
                {
                    string localName = reader.LocalName;

                    if (localName == "comment")
                    {
                        currentComment = new Nedev.FileConverters.DocxToDoc.Model.CommentModel();

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
                    else if (localName == "t" && currentComment != null)
                    {
                        // Read comment text
                        string text = reader.ReadElementContentAsString();
                        currentComment.Text += text;
                    }
                }
                else if (reader.NodeType == XmlNodeType.EndElement)
                {
                    if (reader.LocalName == "comment" && currentComment != null)
                    {
                        docModel.Comments.Add(currentComment);
                        currentComment = null;
                    }
                }
            }
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
