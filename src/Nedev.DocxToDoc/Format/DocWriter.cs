using System;
using System.IO;

namespace Nedev.DocxToDoc.Format
{
    /// <summary>
    /// Writes the various streams required by the MS-DOC File Format (.doc)
    /// such as the WordDocument stream, 1Table, Data, etc.
    /// </summary>
    internal class DocWriter
    {
        public void WriteDocBlocks(DocxReader reader, Stream outputStream)
        {
            // 1. Initialize streams needed for the MS-DOC file
            using var wordDocumentStream = new MemoryStream();
            using var tableStream = new MemoryStream();
            using var dataStream = new MemoryStream();

            // 2. Write File Information Block (FIB) - placeholder
            WriteFib(wordDocumentStream);

            // 3. Process reader contents 
            // e.g., stream text and paragraph properties to tableStream & wordDocumentStream
            
            // 4. Wrap the streams into OLE Compound File Binary (CFB) format
            using var cfbWriter = new CfbWriter();
            cfbWriter.AddStream("WordDocument", wordDocumentStream.ToArray());
            cfbWriter.AddStream("1Table", tableStream.ToArray());
            cfbWriter.AddStream("Data", dataStream.ToArray());
            
            // Optional/Standard metadata streams
            // cfbWriter.AddStream("\x05SummaryInformation", SummaryInfoData);
            // cfbWriter.AddStream("\x05DocumentSummaryInformation", DocSummaryInfoData);

            // 5. Write out the final CFB to the destination
            cfbWriter.WriteTo(outputStream);
        }

        private void WriteFib(Stream stream)
        {
            // The File Information Block starts at offset 0 in the WordDocument stream.
            // Minimum size of base FIB is 32 bytes (FIB base), then more for FibRgW97, FibRgLw97, etc.
            // Here we just write a skeleton representing a valid base to be expanded later.
            using var writer = new BinaryWriter(stream, System.Text.Encoding.GetEncoding(1252), leaveOpen: true);
            
            // magic number
            writer.Write((ushort)0xA5EC); // wIdent
            writer.Write((ushort)0x00C1); // nFib (Word 97 = 193)
            writer.Write((ushort)0x0000); // unused / nProduct
            writer.Write((ushort)0x0000); // lid
            writer.Write((short)0); // pnNext
            
            // And so forth... to be fully implemented based on MS-DOC specifications.
        }
    }
}
