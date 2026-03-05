using System;
using System.IO;
using System.Text;

namespace Nedev.DocxToDoc.Utils
{
    /// <summary>
    /// High-performance extensions for binary operations over Streams.
    /// Eliminates multiple buffer allocations.
    /// </summary>
    internal static class BinaryWriterExtensions
    {
        // MS-DOC strings are commonly saved in zero-terminated 16-bit unicode format.
        public static void WriteDocString(this BinaryWriter writer, string value)
        {
            if (string.IsNullOrEmpty(value))
            {
                writer.Write((short)0);
                return;
            }

            var length = (short)value.Length;
            writer.Write(length);
            
            // Prefer straightforward loop or Span encoding avoiding string conversions.
            // Using Encoding.Unicode will emit Little-Endian 16-bit which matches standard MS-DOC
            var bytes = Encoding.Unicode.GetBytes(value);
            writer.Write(bytes);
        }

        public static void WriteZeroBytes(this BinaryWriter writer, int count)
        {
            if (count <= 0) return;
            // Write chunks of 0s efficiently
            Span<byte> zeros = count <= 1024 ? stackalloc byte[count] : new byte[count];
            writer.Write(zeros);
        }
    }
}
