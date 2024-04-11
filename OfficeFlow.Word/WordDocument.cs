using System;
using System.IO;
using OfficeFlow.DocumentObjectModel;
using OfficeFlow.Word.Core.Elements;
using OfficeFlow.Word.Core.Enums;
using OfficeFlow.Word.Core.Interfaces;

namespace OfficeFlow.Word
{
    public sealed class WordDocument : CompositeElement, IDisposable
    {
        private bool _isDisposed;

        public WordDocument(WordDocumentType documentType)
            : this(documentType, DocumentSettings.Default)
        { }
        
        public WordDocument(WordDocumentType documentType, DocumentSettings settings)
        { }
        
        public WordDocument(Stream stream)
            : this(stream, DocumentSettings.Default)
        { }

        public WordDocument(Stream stream, DocumentSettings settings)
        { }

        public WordDocument(string filePath)
            : this(filePath, DocumentSettings.Default)
        { }

        public WordDocument(string filePath, DocumentSettings settings)
        { }

        public Section LastSection
        {
            get
            {
                throw new NotImplementedException();
            }
        }

        public Section AddSection()
        {
            throw new NotImplementedException();
        }
        
        public void Save()
            => throw new NotImplementedException();

        public void SaveTo(Stream stream)
            => throw new NotImplementedException();

        public void SaveTo(string filePath)
            => throw new NotImplementedException();

        public void Close()
            => Dispose();

        /// <inheritdoc />
        public void Dispose()
        {
            if (_isDisposed)
            {
                return;
            }
            
            throw new NotImplementedException();

            _isDisposed = true;
        }
    }
}