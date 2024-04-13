using System;
using System.IO;
using System.IO.Packaging;
using OfficeFlow.OpenXml.Packaging;
using OfficeFlow.OpenXml.Resources;
using OfficeFlow.OpenXml.Resources.Interfaces;
using OfficeFlow.Word.Core.Elements;
using OfficeFlow.Word.Core.Interfaces;
using OfficeFlow.Word.OpenXml.Enums;

namespace OfficeFlow.Word.OpenXml
{
    public sealed class OpenXmlWordProcessor : IWordProcessor
    {
        private bool _isDisposed;
        private readonly OpenXmlPackage _package;
        private readonly OpenXmlWordDocumentType _documentType;
        private readonly IResourceDecompressor _resourceDecompressor;
        private Body? _body;

        public OpenXmlWordProcessor(
            OpenXmlPackage package,
            OpenXmlWordDocumentType documentType = OpenXmlWordDocumentType.Document)
        {
            _package = package;
            _documentType = documentType;
            _resourceDecompressor = new EmbeddedResourceDecompressor();
        }

        private OpenXmlPackagePart MainDocumentPart
        {
            get
            {
                var uri = new Uri("/word/document.xml", UriKind.Relative);

                if (_package.TryGetPart(uri, out var packagePart))
                    return packagePart;

                var content = _resourceDecompressor.Decompress("wordDocument.xml.gz");
                
                return OpenXmlPackagePart.Create(
                    uri,
                    OpenXmlWordContentTypes.GetContentType(_documentType),
                    CompressionOption.Normal,
                    OpenXmlWordRelationships.Document,
                    content);
            }
        }
        
        /// <inheritdoc />
        public Body Body
        {
            get
            {
                if (_body != null)
                    return _body;

                var body = new Body();

                new OpenXmlElementReader(MainDocumentPart.Root)
                    .Visit(body);

                return _body = body;
            }
        }

        /// <inheritdoc />
        public void Save()
            => Save(() => _package.Save());

        /// <inheritdoc />
        public void SaveTo(Stream stream)
            => Save(() => _package.SaveTo(stream));

        /// <inheritdoc />
        public void SaveTo(string filePath)
            => Save(() => _package.SaveTo(filePath));

        private void Save(Action saveStrategy)
        {
            new OpenXmlElementWriter(MainDocumentPart.Root)
                .Visit(Body);
            
            saveStrategy.Invoke();
        }
        
        /// <inheritdoc />
        public void Dispose()
        {
            if (_isDisposed)
            {
                return;
            }

            _package.Dispose();

            _isDisposed = true;
        }
    }
}