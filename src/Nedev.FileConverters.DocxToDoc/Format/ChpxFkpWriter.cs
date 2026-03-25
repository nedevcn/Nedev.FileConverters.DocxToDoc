using System;
using System.Collections.Generic;
using System.IO;

namespace Nedev.FileConverters.DocxToDoc.Format
{
    /// <summary>
    /// Helps pack Single Property Modifiers (Sprm) into a Formatted Disk Page (FKP) 
    /// for character formatting (CHPX). 
    /// </summary>
    public class ChpxFkpWriter
    {
        private const int PageSize = 512;
        
        // CPs (character positions) arrays
        private readonly List<int> _rgfc = new();
        // Sprm byte arrays for each CP interval
        private readonly List<byte[]> _rgbx = new();

        /// <summary>
        /// Adds a run of formatting to the FKP.
        /// </summary>
        public void AddRun(int startCp, int endCp, byte[] sprms)
        {
            if (_rgfc.Count == 0)
            {
                _rgfc.Add(startCp);
            }

            if (_rgfc[^1] < startCp)
            {
                _rgfc.Add(startCp);
                _rgbx.Add(System.Array.Empty<byte>());
            }

            _rgfc.Add(endCp);
            _rgbx.Add(sprms);
        }


        public byte[] GeneratePage()
        {
            byte[] page = new byte[PageSize];
            
            if (_rgfc.Count == 0) return page;

            byte cRun = (byte)(_rgfc.Count - 1);
            
            // Last byte of page is cRun
            page[PageSize - 1] = cRun;

            // Write rgfc (cRun + 1 ints)
            int offset = 0;
            foreach (var fc in _rgfc)
            {
                BitConverter.GetBytes(fc).CopyTo(page, offset);
                offset += 4;
            }

            // Write offset array rgbx
            int currentTail = PageSize - 1; // End of page, before cRun
            
            for (int i = 0; i < cRun; i++)
            {
                byte[] sprm = _rgbx[i];
                if (sprm == null || sprm.Length == 0)
                {
                    page[offset + i] = 0;
                    continue;
                }

                byte cb = (byte)sprm.Length;
                int chpxSize = 1 + cb;
                
                currentTail -= chpxSize;
                
                page[currentTail] = cb;
                Array.Copy(sprm, 0, page, currentTail + 1, cb);
                
                page[offset + i] = (byte)(currentTail / 2);
            }

            return page;
        }
    }
}
