using System;
using OfficeFlow.Word.Core.Enums;
using OfficeFlow.Word.OpenXml.Enums;

namespace OfficeFlow.Word.Extensions
{
    internal static class WordDocumentTypeExtensions
    {
        internal static OpenXmlWordDocumentType ToOpenXmlType(this WordDocumentType documentType)
            => documentType switch
            {
                WordDocumentType.Docx => OpenXmlWordDocumentType.Document,
                WordDocumentType.Docm => OpenXmlWordDocumentType.MacroEnabledDocument,
                WordDocumentType.Dotx => OpenXmlWordDocumentType.Template,
                WordDocumentType.Dotm => OpenXmlWordDocumentType.MacroEnabledTemplate,
                _ => throw new NotSupportedException()
            };
    }
}