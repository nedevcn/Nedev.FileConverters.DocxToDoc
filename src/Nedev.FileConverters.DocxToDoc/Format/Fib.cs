using System.IO;

namespace Nedev.FileConverters.DocxToDoc.Format
{
    /// <summary>
    /// Represents the File Information Block (FIB) in a Word 97-2003 binary file.
    /// Found at offset 0 in the WordDocument stream.
    /// </summary>
    public class Fib
    {
        private const ushort HasPicturesMask = 0x0008;
        public const int FootnoteReferencePairIndex = 2;
        public const int FootnoteTextPairIndex = 3;
        public const int CommentReferencePairIndex = 4;
        public const int CommentTextPairIndex = 5;
        public const int HeaderStoryPairIndex = 11;
        public const int HeaderFieldPairIndex = 17;
        public const int EndnoteReferencePairIndex = 46;
        public const int EndnoteTextPairIndex = 47;
        public const int ChpxPairIndex = 12;
        public const int PapxPairIndex = 13;
        public const int ClxPairIndex = 33;
        public const int HeaderShapePairIndex = 39;
        public const int TapxPairIndex = 54;

        public ushort wIdent { get; set; } = 0xA5EC;
        public ushort nFib { get; set; } = 0x00C1; // Microsoft Word 97-2003
        public ushort unused { get; set; } = 0x0000;
        public ushort lid { get; set; } = 0x0409; // English US (Default)
        public short pnNext { get; set; } = 0;
        public ushort fFlags { get; set; } = 0x0060; // Has FIB extended, is complex (PieceTable)
        public bool HasPictures { get; set; }

        // Basic FIB ends here in this simplified version, followed by csw, rgsw, cslw, rglw, cbRgFcLcb, etc.

        public int fcClx { get; set; } // Offset to Piece Table in 1Table stream
        public int lcbClx { get; set; } // Size of Piece Table

        public int fcStshf { get; set; } // Offset to STSH in 1Table stream
        public int lcbStshf { get; set; } // Size of STSH

        public int fcPlcffndRef { get; set; } // Footnote reference PLCF
        public int lcbPlcffndRef { get; set; }

        public int fcPlcffndTxt { get; set; } // Footnote text PLCF
        public int lcbPlcffndTxt { get; set; }

        public int fcPlcfHdd { get; set; } // Header document PLCF
        public int lcbPlcfHdd { get; set; }
        public int fcPlcffldHdr { get; set; } // Header document field table
        public int lcbPlcffldHdr { get; set; }

        public int fcPlcfbteChpx { get; set; }
        public int lcbPlcfbteChpx { get; set; }

        public int fcPlcfbteTapx { get; set; }
        public int lcbPlcfbteTapx { get; set; }

        public int fcPlcfbtePapx { get; set; }
        public int lcbPlcfbtePapx { get; set; }

        public int fcPlcfsed { get; set; } // Section Descriptors
        public int lcbPlcfsed { get; set; }

        public int fcSttbfffn { get; set; } // Font Table
        public int lcbSttbfffn { get; set; }

        public int fcPlcffldMom { get; set; } // Main document field table
        public int lcbPlcffldMom { get; set; }

        public int fcPlcfspaMom { get; set; } // Main document shape positions
        public int lcbPlcfspaMom { get; set; }
        public int fcPlcSpaHdr { get; set; } // Header document shape positions
        public int lcbPlcSpaHdr { get; set; }

        public int fcDggInfo { get; set; } // OfficeArt drawing group information
        public int lcbDggInfo { get; set; }

        public int fcSttbLst { get; set; } // List Data (LST)
        public int lcbSttbLst { get; set; }

        public int fcPlfLfo { get; set; } // List Format Override (LFO)
        public int lcbPlfLfo { get; set; }

        public int fcPlcfBkmkf { get; set; } // Bookmark first CPs
        public int lcbPlcfBkmkf { get; set; }
        public int fcPlcfBkmkl { get; set; } // Bookmark last CPs
        public int lcbPlcfBkmkl { get; set; }
        public int fcSttbfbkmk { get; set; } // Bookmark names
        public int lcbSttbfbkmk { get; set; }

        // The writer still uses a simplified Fc/Lcb map. Footnotes/comments now occupy
        // their early story slots, CHPX/PAPX use their dedicated pairs, and TAPX
        // remains parked in a custom slot until the broader FIB mapping is normalized.
        public int fcPlcfandRef { get; set; }
        public int lcbPlcfandRef { get; set; }
        public int fcPlcfandTxt { get; set; }
        public int lcbPlcfandTxt { get; set; }
        public int fcPlcfendRef { get; set; }
        public int lcbPlcfendRef { get; set; }
        public int fcPlcfendTxt { get; set; }
        public int lcbPlcfendTxt { get; set; }

        public int ccpText { get; set; } // Length of plain text document
        public int ccpFtn { get; set; } // Length of footnotes
        public int ccpHdd { get; set; } // Length of headers
        public int ccpAtn { get; set; } // Length of comments
        public int ccpEdn { get; set; } // Length of endnotes

        public void WriteTo(BinaryWriter writer)
        {
            writer.Write(wIdent);          // 0-1 (2)
            writer.Write(nFib);            // 2-3 (2)
            writer.Write(unused);          // 4-5 (2)
            writer.Write(lid);             // 6-7 (2)
            writer.Write(pnNext);          // 8-9 (2)
            writer.Write((ushort)(fFlags | (HasPictures ? HasPicturesMask : 0))); // 10-11 (2)
            writer.Write((ushort)0);       // 12-13 (2) (nFibBack)
            writer.Write((int)0);          // 14-17 (4) (lKey)
            writer.Write((byte)0);         // 18 (1) (envr)
            writer.Write((byte)1);         // 19 (1) (fMac)
            writer.Write((ushort)0);       // 20-21 (2) (reserved1)
            writer.Write((ushort)0);       // 22-23 (2) (reserved2)
            writer.Write((ushort)0);       // 24-25 (2) (reserved3)
            writer.Write((ushort)0);       // 26-27 (2) (reserved4)
            writer.Write((ushort)0);       // 28-29 (2) (reserved5)
            writer.Write((ushort)0);       // 30-31 (2) (reserved6)

            // FibRgW97 (2-byte fields)
            writer.Write((ushort)14);      // csw (offset 32)
            writer.Write(new byte[28]);    // 14 * 2 bytes = 28 (offset 34-61)

            // FibRgLw97 (4-byte fields)
            writer.Write((ushort)22);      // cslw (offset 62)
            byte[] rgLw97 = new byte[88];  // 22 * 4 bytes = 88 (offset 64-151)
            
            // ccpText is index 0 of rgLw97 (offset 64)
            BitConverter.GetBytes(ccpText).CopyTo(rgLw97, 0);
            BitConverter.GetBytes(ccpFtn).CopyTo(rgLw97, 4);
            BitConverter.GetBytes(ccpHdd).CopyTo(rgLw97, 8);
            BitConverter.GetBytes(ccpAtn).CopyTo(rgLw97, 12);
            BitConverter.GetBytes(ccpEdn).CopyTo(rgLw97, 16);
            
            writer.Write(rgLw97);

            // FibRgFcLcb (starts at offset 154)
            writer.Write((ushort)93); // cbRgFcLcb = 93 pairs (offset 152-153)
            
            // We need to write 93 pairs of (fc, lcb) = 93 * 8 = 744 bytes (offset 154-897)
            byte[] rgFcLcb = new byte[744];

            void SetPair(int index, int fc, int lcb)
            {
                BitConverter.GetBytes(fc).CopyTo(rgFcLcb, index * 8);
                BitConverter.GetBytes(lcb).CopyTo(rgFcLcb, index * 8 + 4);
            }

            SetPair(0, fcStshf, lcbStshf);
            SetPair(FootnoteReferencePairIndex, fcPlcffndRef, lcbPlcffndRef);
            SetPair(FootnoteTextPairIndex, fcPlcffndTxt, lcbPlcffndTxt);
            SetPair(TapxPairIndex, fcPlcfbteTapx, lcbPlcfbteTapx);
            SetPair(6, fcPlcfsed, lcbPlcfsed);
            SetPair(10, fcSttbLst, lcbSttbLst);
            SetPair(HeaderStoryPairIndex, fcPlcfHdd, lcbPlcfHdd);
            SetPair(ChpxPairIndex, fcPlcfbteChpx, lcbPlcfbteChpx);
            SetPair(PapxPairIndex, fcPlcfbtePapx, lcbPlcfbtePapx);
            SetPair(14, fcSttbfffn, lcbSttbfffn);
            SetPair(15, fcPlcffldMom, lcbPlcffldMom);
            SetPair(HeaderFieldPairIndex, fcPlcffldHdr, lcbPlcffldHdr);
            SetPair(ClxPairIndex, fcClx, lcbClx);
            SetPair(HeaderShapePairIndex, fcPlcSpaHdr, lcbPlcSpaHdr);
            SetPair(40, fcPlcfspaMom, lcbPlcfspaMom);
            SetPair(50, fcDggInfo, lcbDggInfo);
            SetPair(53, fcPlfLfo, lcbPlfLfo);
            SetPair(CommentReferencePairIndex, fcPlcfandRef, lcbPlcfandRef);
            SetPair(CommentTextPairIndex, fcPlcfandTxt, lcbPlcfandTxt);
            SetPair(EndnoteReferencePairIndex, fcPlcfendRef, lcbPlcfendRef);
            SetPair(EndnoteTextPairIndex, fcPlcfendTxt, lcbPlcfendTxt);

            writer.Write(rgFcLcb);
        }
    }
}
