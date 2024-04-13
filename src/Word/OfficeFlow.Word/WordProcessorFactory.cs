using System;
using System.IO;
using OfficeFlow.Formats;
using OfficeFlow.OpenXml.Packaging;
using OfficeFlow.Word.Core.Enums;
using OfficeFlow.Word.Core.Interfaces;
using OfficeFlow.Word.Extensions;
using OfficeFlow.Word.OpenXml;

namespace OfficeFlow.Word
{
    internal sealed class WordProcessorFactory
    {
        private static readonly Lazy<WordProcessorFactory> InstanceFactory
            = new Lazy<WordProcessorFactory>(() => new WordProcessorFactory());

        public static WordProcessorFactory Instance
            => InstanceFactory.Value;
        
        public IWordProcessor Create(WordDocumentType documentType, DocumentSettings settings)
        {
            return new OpenXmlWordProcessor(
                OpenXmlPackage.Create(),
                documentType.ToOpenXmlType());
        }

        public IWordProcessor Open(Stream stream, DocumentSettings settings)
        {
            var format = OfficeFormatDetector.Detect(stream);

            if (format is OpenXmlFormat)
                return new OpenXmlWordProcessor(
                    OpenXmlPackage.Open(stream));

            if (format is BinaryFormat)
                throw new NotSupportedException();

            throw new NotSupportedException();
        }

        public IWordProcessor Open(string filePath, DocumentSettings settings)
        {
            var format = OfficeFormatDetector.Detect(filePath);
            
            if (format is OpenXmlFormat)
                return new OpenXmlWordProcessor(
                    OpenXmlPackage.Open(filePath));

            if (format is BinaryFormat)
                throw new NotSupportedException();
            
            throw new NotSupportedException();
        }
    }
}