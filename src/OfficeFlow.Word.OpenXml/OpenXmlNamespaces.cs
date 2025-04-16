using System.Xml.Linq;

namespace OfficeFlow.Word.OpenXml;

/// <summary>
/// Представляет пространства имен OpenXml
/// </summary>
internal static class OpenXmlNamespaces
{
    /// <summary>
    /// Word (W)
    /// </summary>
    internal static readonly XNamespace Word
        = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    /// <summary>
    /// Office (O)
    /// </summary>
    internal static readonly XNamespace Office
        = "urn:schemas-microsoft-com:office:office";

    /// <summary>
    /// Package Relationships (Rel)
    /// </summary>
    internal static readonly XNamespace PackageRelationship
        = "http://schemas.openxmlformats.org/package/2006/relationships";

    /// <summary>
    /// Office Relationships (R)
    /// </summary>
    internal static readonly XNamespace OfficeRelationship
        = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

    /// <summary>
    /// Custom Properties
    /// </summary>
    internal static readonly XNamespace CustomProperties
        = "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties";

    /// <summary>
    /// Custom VTypes
    /// </summary>
    internal static readonly XNamespace CustomVTypes
        = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes";

    /// <summary>
    /// Word Drawing (Wp)
    /// </summary>
    internal static XNamespace Drawing
        = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";

    /// <summary>
    /// Word (W10)
    /// </summary>
    internal static readonly XNamespace Word10
        = "urn:schemas-microsoft-com:office:word";

    /// <summary>
    /// WordML (Wne)
    /// </summary>
    internal static readonly XNamespace WordMl
        = "http://schemas.microsoft.com/office/word/2006/wordml";

    /// <summary>
    /// Drawing Main (A)
    /// </summary>
    internal static readonly XNamespace DrawingMain
        = "http://schemas.openxmlformats.org/drawingml/2006/main";

    /// <summary>
    /// Drawing Chart (C)
    /// </summary>
    internal static readonly XNamespace DrawingChart
        = "http://schemas.openxmlformats.org/drawingml/2006/chart";

    /// <summary>
    /// VML (V)
    /// </summary>
    internal static readonly XNamespace Vml
        = "urn:schemas-microsoft-com:vml";

    /// <summary>
    /// Markup (Ve)
    /// </summary>
    internal static readonly XNamespace Markup
        = "http://schemas.openxmlformats.org/markup-compatibility/2006";

    /// <summary>
    /// Numbering (N)
    /// </summary>
    internal static readonly XNamespace Numbering
        = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering";

    /// <summary>
    /// Math (M)
    /// </summary>
    internal static readonly XNamespace Math
        = "http://schemas.openxmlformats.org/officeDocument/2006/math";
}