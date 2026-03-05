using System;
using System.IO;
using System.IO.Compression;
using System.Xml;

namespace Nedev.DocxToDoc.Format
{
    /// <summary>
    /// Reads and coordinates the extraction of features from OpenXML (.docx) files.
    /// Optimized for low-memory, forward-only reading.
    /// </summary>
    internal class DocxReader : IDisposable
    {
        private readonly ZipArchive _archive;
        private readonly Stream _baseStream;
        private bool _disposedValue;

        public DocxReader(Stream docxStream)
        {
            _baseStream = docxStream;
            // Leave open is true usually, or false depending on ownership.
            // We assume the caller manages the base stream's lifecycle.
            _archive = new ZipArchive(docxStream, ZipArchiveMode.Read, leaveOpen: true);
        }

        public void ParseDocument(Action<XmlReader> documentPartCallback)
        {
            var documentEntry = _archive.GetEntry("word/document.xml");
            if (documentEntry == null)
            {
                // Fallback: proper way is to read _rels/.rels then find the document part, 
                // but for maximum speed, we first try the standard path.
                throw new FileNotFoundException("word/document.xml not found in the docx file.");
            }

            using var stream = documentEntry.Open();
            // Use XmlReader for high-performance, forward-only parsing without loading the DOM
            using var xmlReader = XmlReader.Create(stream, new XmlReaderSettings 
            { 
                IgnoreComments = true, 
                IgnoreWhitespace = true 
            });

            documentPartCallback(xmlReader);
        }

        // Additional methods to parse styles.xml, numbering.xml, theme/theme1.xml, etc.

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
