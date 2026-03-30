using System;
using System.Collections.Generic;
using System.IO;

namespace Nedev.FileConverters.DocxToDoc.Format
{
    /// <summary>
    /// Helps pack Table Properties (TAPX) into a Formatted Disk Page (FKP).
    /// </summary>
    public class TapxFkpWriter
    {
        private const int PageSize = 512;
        
        private readonly List<(int startCp, int endCp, byte[] sprms)> _runs = new();

        public void AddRow(int startCp, int endCp, byte[] sprms)
        {
            if (endCp <= startCp)
            {
                throw new ArgumentOutOfRangeException(nameof(endCp), "TAPX row ranges must have positive length.");
            }

            if (_runs.Count > 0 && startCp < _runs[^1].endCp)
            {
                throw new InvalidOperationException("TAPX row ranges must be added in non-overlapping CP order.");
            }

            _runs.Add((startCp, endCp, sprms ?? Array.Empty<byte>()));
        }

        public byte[] GeneratePage()
        {
            byte[] page = new byte[PageSize];
            
            if (_runs.Count == 0) return page;

            var rgfc = new List<int> { _runs[0].startCp };
            var sprmRuns = new List<byte[]>();
            int cursor = _runs[0].startCp;

            foreach (var run in _runs)
            {
                if (run.startCp > cursor)
                {
                    rgfc.Add(run.startCp);
                    sprmRuns.Add(Array.Empty<byte>());
                    cursor = run.startCp;
                }

                rgfc.Add(run.endCp);
                sprmRuns.Add(run.sprms);
                cursor = run.endCp;
            }

            byte cRun = (byte)sprmRuns.Count;
            
            // Last byte of page is cRun
            page[PageSize - 1] = cRun;

            // Write rgfc (cRun + 1 ints)
            int offset = 0;
            foreach (var fc in rgfc)
            {
                BitConverter.GetBytes(fc).CopyTo(page, offset);
                offset += 4;
            }

            // In TAPX FKP, following rgfc is an array of bx (offset)
            // bx is 1 byte (offset to TAPX in the page / 2? No, for TAPX it might be different.)
            // Actually, for TAPX, the offset is also 1 byte (divided by 2).
            
            int currentTail = PageSize - 1; 
            
            for (int i = 0; i < cRun; i++)
            {
                byte[] sprm = sprmRuns[i];
                
                // TAPX structure in FKP: [cb (2 bytes)] [grpprl]
                int grpprlLen = sprm?.Length ?? 0;
                int tapxSize = 2 + grpprlLen; 
                
                // Align to 2-byte boundary
                if (tapxSize % 2 != 0) tapxSize++;

                currentTail -= tapxSize;
                
                // cb (2 bytes)
                BitConverter.GetBytes((ushort)grpprlLen).CopyTo(page, currentTail);

                if (sprm != null)
                {
                    Array.Copy(sprm, 0, page, currentTail + 2, sprm.Length);
                }
                
                // Offset to TAPX is at (rgfc.Count * 4) + i
                page[offset + i] = (byte)(currentTail / 2);
            }

            return page;
        }
    }
}
