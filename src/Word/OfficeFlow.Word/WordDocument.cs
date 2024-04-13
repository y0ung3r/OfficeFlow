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
        private readonly IWordProcessor _processor;
        private readonly DocumentSettings _settings;
     
        public WordDocument(WordDocumentType documentType)
            : this(documentType, DocumentSettings.Default)
        { }
        
        public WordDocument(WordDocumentType documentType, DocumentSettings settings)
            : this(WordProcessorFactory.Instance.Create(documentType, settings), settings)
        { }
        
        public WordDocument(Stream stream)
            : this(stream, DocumentSettings.Default)
        { }

        public WordDocument(Stream stream, DocumentSettings settings)
            : this(WordProcessorFactory.Instance.Open(stream, settings), settings)
        { }

        public WordDocument(string filePath)
            : this(filePath, DocumentSettings.Default)
        { }

        public WordDocument(string filePath, DocumentSettings settings)
            : this(WordProcessorFactory.Instance.Open(filePath, settings), settings)
        { }
        
        private WordDocument(IWordProcessor processor, DocumentSettings settings)
        {
            _processor = processor;
            _settings = settings;
        }

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
            => _processor.Save();

        public void SaveTo(Stream stream)
            => _processor.SaveTo(stream);

        public void SaveTo(string filePath)
            => _processor.SaveTo(filePath);

        public void Close()
            => Dispose();

        /// <inheritdoc />
        public void Dispose()
        {
            if (_isDisposed)
            {
                return;
            }
            
            _processor.Dispose();

            _isDisposed = true;
        }
    }
}