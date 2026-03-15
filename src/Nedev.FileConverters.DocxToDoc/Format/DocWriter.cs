using System;
using System.IO;
using System.Text;
using System.Collections.Generic;
using Nedev.FileConverters.DocxToDoc.Model;

namespace Nedev.FileConverters.DocxToDoc.Format
{
    /// <summary>
    /// Writes the various streams required by the MS-DOC File Format (.doc)
    /// such as the WordDocument stream, 1Table, Data, etc.
    /// </summary>
    public class DocWriter
    {
        public void WriteDocBlocks(DocumentModel model, Stream outputStream)
        {
            // 1. Initialize streams needed for the MS-DOC file
            using var wordDocumentStream = new MemoryStream();
            using var tableStream = new MemoryStream();
            using var dataStream = new MemoryStream();

            // Track embedded objects and images
            var embeddedObjects = new List<(int cp, byte[] data, string contentType)>();
            int dataStreamOffset = 0;

            // 1. Build the text buffer and formatting structures in one pass
            var textBuilder = new StringBuilder();
            var chpxWriter = new ChpxFkpWriter();
            var papxWriter = new PapxFkpWriter();
            var tapxWriter = new TapxFkpWriter();
            
            int currentCp = 0;
            var tableWriter = new BinaryWriter(tableStream);

            void ProcessParagraph(ParagraphModel para)
            {
                int paraStart = currentCp;
                foreach (var run in para.Runs)
                {
                    // Handle images
                    if (run.Image != null && run.Image.Data != null)
                    {
                        // Add a placeholder character for the image
                        // In MS-DOC, embedded objects use special characters
                        textBuilder.Append('\x0001'); // Object placeholder
                        
                        // Track the image data
                        embeddedObjects.Add((currentCp, run.Image.Data, run.Image.ContentType));
                        
                        // Add CHPX with sprmCFSpec = 1 (special character)
                        List<byte> imageSprms = new List<byte>();
                        imageSprms.Add(0x55); imageSprms.Add(0x08); imageSprms.Add(1); // sprmCFSpec
                        chpxWriter.AddRun(currentCp, currentCp + 1, imageSprms.ToArray());
                        
                        currentCp += 1;
                        continue;
                    }

                    if (run.Text.Length == 0) continue;
                    
                    // Build Runs
                    List<byte> runSprms = new List<byte>();
                    if (run.Properties.IsBold) { runSprms.Add(0x35); runSprms.Add(0x08); runSprms.Add(1); }
                    if (run.Properties.IsItalic) { runSprms.Add(0x36); runSprms.Add(0x08); runSprms.Add(1); }
                    if (run.Properties.IsStrike) { runSprms.Add(0x37); runSprms.Add(0x08); runSprms.Add(1); }
                    if (run.Properties.FontSize.HasValue) 
                    { 
                        runSprms.Add(0x43); runSprms.Add(0x4A); 
                        runSprms.Add(BitConverter.GetBytes((short)run.Properties.FontSize.Value)[0]); 
                        runSprms.Add(BitConverter.GetBytes((short)run.Properties.FontSize.Value)[1]); 
                    }

                    if (runSprms.Count > 0)
                        chpxWriter.AddRun(currentCp, currentCp + run.Text.Length, runSprms.ToArray());
                    
                    textBuilder.Append(run.Text);
                    currentCp += run.Text.Length;
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

                textBuilder.Append('\r');
                papxWriter.AddParagraph(paraStart, currentCp + 1, paraSprms.ToArray());
                currentCp += 1;
            }

            foreach (var item in model.Content)
            {
                if (item is ParagraphModel para)
                {
                    ProcessParagraph(para);
                }
                else if (item is TableModel table)
                {
                    foreach (var row in table.Rows)
                    {
                        int rowStart = currentCp;
                        foreach (var cell in row.Cells)
                        {
                            foreach (var cellPara in cell.Paragraphs)
                            {
                                ProcessParagraph(cellPara);
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
            WriteStyleSheet(tableWriter, model.Styles);
            int lcbStshf = (int)tableStream.Position - fcStshf;

            // 8. Write Numbering (SttbLst and PlcfLfo)
            int fcSttbLst = (int)tableStream.Position;
            WriteNumbering(tableWriter, model);
            int lcbSttbLst = (int)tableStream.Position - fcSttbLst;

            int fcPlfLfo = (int)tableStream.Position;
            WriteLfo(tableWriter, model);
            int lcbPlfLfo = (int)tableStream.Position - fcPlfLfo;

            // 8.5. Write embedded objects/images to Data stream
            int fcData = 0;
            int lcbData = 0;
            if (embeddedObjects.Count > 0)
            {
                fcData = (int)dataStream.Position;
                using var dataWriter = new BinaryWriter(dataStream, System.Text.Encoding.GetEncoding(1252), leaveOpen: true);
                
                foreach (var (cp, data, contentType) in embeddedObjects)
                {
                    // Write object header (simplified)
                    // In a full implementation, this would be a proper OLE object header
                    dataWriter.Write(data.Length); // Size
                    dataWriter.Write(data); // Data
                }
                
                lcbData = (int)dataStream.Position - fcData;
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
                fcSttbLst = fcSttbLst,
                lcbSttbLst = lcbSttbLst,
                fcPlfLfo = fcPlfLfo,
                lcbPlfLfo = lcbPlfLfo,
                fcData = fcData,
                lcbData = lcbData,
                fcPlcfBkmkf = fcPlcfBkmkf,
                lcbPlcfBkmkf = lcbPlcfBkmkf,
                fcPlcfBkmkl = fcPlcfBkmkl,
                lcbPlcfBkmkl = lcbPlcfBkmkl,
                fcSttbfbkmk = fcSttbfbkmk,
                lcbSttbfbkmk = lcbSttbfbkmk,
                ccpText = currentCp
            };
            fib.WriteTo(new BinaryWriter(wordDocumentStream, System.Text.Encoding.GetEncoding(1252), leaveOpen: true));

            // 8. Wrap the streams into OLE Compound File Binary (CFB) format
            using var cfbWriter = new CfbWriter();
            cfbWriter.AddStream("WordDocument", wordDocumentStream.ToArray());
            cfbWriter.AddStream("1Table", tableStream.ToArray());
            cfbWriter.AddStream("Data", dataStream.ToArray());
            
            // 9. Write out the final CFB to the destination
            cfbWriter.WriteTo(outputStream);
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

        private void WriteStyleSheet(BinaryWriter writer, List<Nedev.FileConverters.DocxToDoc.Model.StyleModel> styles)
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
                    chpxData = BuildChpxFromStyle(style.CharacterProps);
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

            return sprms.ToArray();
        }

        private byte[] BuildChpxFromStyle(RunModel.CharacterProperties props)
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

            return sprms.ToArray();
        }
    }
}
