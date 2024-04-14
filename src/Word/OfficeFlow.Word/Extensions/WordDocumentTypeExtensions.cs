using System;
using OfficeFlow.Word.Core.Enums;
using OfficeFlow.Word.OpenXml.Enums;

namespace OfficeFlow.Word.Extensions
{
    internal static class WordDocumentTypeExtensions
    {
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