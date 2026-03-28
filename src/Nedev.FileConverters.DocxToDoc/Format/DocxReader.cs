using System;
using System.IO;
using System.IO.Compression;
using System.Xml;
using System.Text;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

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

            var extendedPropsEntry = _archive.GetEntry("docProps/app.xml");
            if (extendedPropsEntry != null)
            {
                using var extendedPropsStream = extendedPropsEntry.Open();
                ParseDocumentProperties(extendedPropsStream, docModel);
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

            Nedev.FileConverters.DocxToDoc.Model.ParagraphModel currentParagraph = null;
            Nedev.FileConverters.DocxToDoc.Model.RunModel currentRun = null;
            Nedev.FileConverters.DocxToDoc.Model.RunModel.CharacterProperties currentRunBaseProperties = null;
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
                        currentCell.WidthUnit = type switch
                        {
                            "pct" => Nedev.FileConverters.DocxToDoc.Model.TableWidthUnit.Pct,
                            "auto" => Nedev.FileConverters.DocxToDoc.Model.TableWidthUnit.Auto,
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

                        if (int.TryParse(widthValue, out int preferredWidthValue))
                        {
                            currentTable.PreferredWidthValue = preferredWidthValue;
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
                                            var image = ParseDrawing(xmlReader);
                                            if (image != null)
                                            {
                                                run.Image = image;
                                                LoadImageData(image);
                                            }
                                        }
                                        else if (xmlReader.LocalName == "pict")
                                        {
                                            var image = ParsePict(xmlReader);
                                            if (image != null)
                                            {
                                                run.Image = image;
                                                LoadImageData(image);
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
                    else if ((localName == "t" || localName == "delText" || localName == "tab" || localName == "ptab" || localName == "br" || localName == "cr" || localName == "noBreakHyphen" || localName == "softHyphen" || localName == "sym") && currentRun != null)
                    {
                        AppendRunTextFragment(currentParagraph, ref currentRun, currentRunBaseProperties, textBuffer, xmlReader);
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
                    else if (localName == "pict" && currentRun != null)
                    {
                        // Parse fallback VML images (e.g., from SmartArt)
                        var image = ParsePict(xmlReader);
                        if (image != null)
                        {
                            currentRun.Image = image;
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
                width = 0;
                return false;
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
                "dotted" => Nedev.FileConverters.DocxToDoc.Model.BorderStyle.Dotted,
                "dashed" => Nedev.FileConverters.DocxToDoc.Model.BorderStyle.Dashed,
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

            string fragmentFontName = ResolveFragmentFontName(reader, baseProperties);
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
                case "i":
                    properties.IsItalic = !IsFalseValue(reader.GetAttribute("w:val"));
                    return true;
                case "strike":
                    properties.IsStrike = !IsFalseValue(reader.GetAttribute("w:val"));
                    return true;
                case "u":
                    string underlineValue = reader.GetAttribute("w:val") ?? "none";
                    properties.Underline = underlineValue switch
                    {
                        "single" => Nedev.FileConverters.DocxToDoc.Model.UnderlineType.Single,
                        "double" => Nedev.FileConverters.DocxToDoc.Model.UnderlineType.Double,
                        "thick" => Nedev.FileConverters.DocxToDoc.Model.UnderlineType.Thick,
                        "dotted" => Nedev.FileConverters.DocxToDoc.Model.UnderlineType.Dotted,
                        "dashed" => Nedev.FileConverters.DocxToDoc.Model.UnderlineType.Dashed,
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

        private Nedev.FileConverters.DocxToDoc.Model.ImageModel? ParsePict(XmlReader reader)
        {
            var image = new Nedev.FileConverters.DocxToDoc.Model.ImageModel
            {
                LayoutType = Nedev.FileConverters.DocxToDoc.Model.ImageLayoutType.Inline,
                WrapType = Nedev.FileConverters.DocxToDoc.Model.ImageWrapType.Inline
            };
            string? relId = null;

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
                }
                else if (reader.NodeType == XmlNodeType.EndElement && reader.LocalName == "pict")
                {
                    break;
                }
                
                if (reader.Depth < 3) break;
            }

            if (string.IsNullOrEmpty(relId)) return null;
            image.RelationshipId = relId;
            return image;
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
