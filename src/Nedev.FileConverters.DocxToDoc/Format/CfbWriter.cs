using System;
using System.Collections.Generic;
using System.IO;
using OpenMcdf;

namespace Nedev.FileConverters.DocxToDoc.Format
{
    /// <summary>
    /// Writes Compound File Binary (OLE2) format streams into a single structure.
    /// Essential for MS-DOC binary files.
    /// </summary>
    public class CfbWriter : IDisposable
    {
        private readonly CompoundFile _compoundFile;

        public CfbWriter()
        {
            _compoundFile = new CompoundFile();
        }

        public void AddStream(string name, byte[] data)
        {
            if (string.IsNullOrWhiteSpace(name))
                throw new ArgumentNullException(nameof(name));
            if (data == null)
                throw new ArgumentNullException(nameof(data));

            var stream = _compoundFile.RootStorage.AddStream(name);
            stream.SetData(data);
        }

        public void EmbedStorage(string name, byte[] cfbData)
        {
            if (string.IsNullOrWhiteSpace(name))
                throw new ArgumentNullException(nameof(name));
            if (cfbData == null)
                throw new ArgumentNullException(nameof(cfbData));

            var targetStorage = _compoundFile.RootStorage.AddStorage(name);

            using var ms = new MemoryStream(cfbData);
            using var sourceCfb = new CompoundFile(ms);
            
            CopyStorage(sourceCfb.RootStorage, targetStorage);
        }

        private void CopyStorage(CFStorage source, CFStorage target)
        {
            source.VisitEntries(item =>
            {
                if (item is CFStream sourceStream)
                {
                    var targetStream = target.AddStream(item.Name);
                    targetStream.SetData(sourceStream.GetData());
                }
                else if (item is CFStorage sourceSubStorage)
                {
                    var targetSubStorage = target.AddStorage(item.Name);
                    CopyStorage(sourceSubStorage, targetSubStorage);
                }
            }, false);
        }

        public void WriteTo(Stream outputStream)
        {
            if (outputStream == null)
                throw new ArgumentNullException(nameof(outputStream));

            _compoundFile.Save(outputStream);
        }

        public void Dispose()
        {
            _compoundFile?.Close();
        }
    }
}
