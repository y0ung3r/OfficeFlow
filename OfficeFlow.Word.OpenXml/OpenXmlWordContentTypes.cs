using System;
using OfficeFlow.Word.OpenXml.Enums;

namespace OfficeFlow.Word.OpenXml
{
    internal static class OpenXmlWordContentTypes
    {
        /// <summary>
        /// Word Document (*.docx)
        /// </summary>
        internal const string Document
            = "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml";

        /// <summary>
        /// Word Template (*.dotx)
        /// </summary>
        internal const string Template
            = "application/vnd.openxmlformats-officedocument.wordprocessingml.template.main+xml";

        /// <summary>
        /// Word Macro-Enabled Document (*.docm)
        /// </summary>
        internal const string MacroEnabledDocument
            = "application/vnd.ms-word.document.macroEnabled.main+xml";

        /// <summary>
        /// Word Macro-Enabled Document (*.dotm)
        /// </summary>
        internal const string MacroEnabledTemplate
            = "application/vnd.ms-word.template.macroEnabledTemplate.main+xml";

        /// <summary>
        /// Styles
        /// </summary>
        internal const string Styles
            = "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml";

        /// <summary>
        /// Numbering
        /// </summary>
        internal const string Numbering
            = "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml";

        internal static string GetContentType(OpenXmlWordDocumentType documentType)
            => documentType switch
            {
                OpenXmlWordDocumentType.Document => Document,
                OpenXmlWordDocumentType.Template => Template,
                OpenXmlWordDocumentType.MacroEnabledDocument => MacroEnabledDocument,
                OpenXmlWordDocumentType.MacroEnabledTemplate => MacroEnabledTemplate,
                _ => throw new NotSupportedException(
                    $"Specified document type \"{documentType}\" is not supported")
            };

        internal static OpenXmlWordDocumentType GetDocumentType(string contentType)
            => contentType switch
            {
                Document => OpenXmlWordDocumentType.Document,
                Template => OpenXmlWordDocumentType.Template,
                MacroEnabledDocument => OpenXmlWordDocumentType.MacroEnabledDocument,
                MacroEnabledTemplate => OpenXmlWordDocumentType.MacroEnabledTemplate,
                _ => throw new NotSupportedException(
                    $"Specified content type \"{contentType}\" is not supported")
            };
    }
}