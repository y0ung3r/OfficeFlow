using System;
using OfficeFlow.Formats;
using OfficeFlow.Formats.Interfaces;
using OfficeFlow.Word.Core.Enums;
using OfficeFlow.Word.OpenXml.Enums;

namespace OfficeFlow.Word.Extensions
{
    internal static class WordDocumentTypeExtensions
    {
        internal static IOfficeFormat ToOfficeFormat(this WordDocumentType documentType)
        {
            switch (documentType)
            {
                case WordDocumentType.Doc:
                case WordDocumentType.Dot:
                    return new BinaryFormat();
                
                case WordDocumentType.Docx:
                case WordDocumentType.Docm:
                case WordDocumentType.Dotx:
                case WordDocumentType.Dotm:
                    return new OpenXmlFormat();
                
                default:
                    throw new NotSupportedException();
            }
        }
        
        internal static OpenXmlWordDocumentType ToOpenXmlType(this WordDocumentType documentType)
        {
            switch (documentType)
            {
                case WordDocumentType.Docx:
                    return OpenXmlWordDocumentType.Document;
                
                case WordDocumentType.Docm:
                    return OpenXmlWordDocumentType.MacroEnabledDocument;
                
                case WordDocumentType.Dotx:
                    return OpenXmlWordDocumentType.Template;
                
                case WordDocumentType.Dotm:
                    return OpenXmlWordDocumentType.MacroEnabledTemplate;
                
                default:
                    throw new NotSupportedException();
            }
        }
    }
}