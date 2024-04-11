using System.Xml.Linq;

namespace OfficeFlow.Word.OpenXml
{
	/// <summary>
	/// Представляет пространства имен OpenXml
	/// </summary>
	internal static class OpenXmlNamespaces
	{
		/// <summary>
		/// Word (W)
		/// </summary>
		internal static XNamespace Word
			=> "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

		/// <summary>
		/// Office (O)
		/// </summary>
		internal static XNamespace Office
			=> "urn:schemas-microsoft-com:office:office";

		/// <summary>
		/// Package Relationships (Rel)
		/// </summary>
		internal static XNamespace PackageRelationship
			=> "http://schemas.openxmlformats.org/package/2006/relationships";

		/// <summary>
		/// Office Relationships (R)
		/// </summary>
		internal static XNamespace OfficeRelationship
			=> "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

		/// <summary>
		/// Custom Properties
		/// </summary>
		internal static XNamespace CustomProperties
			=> "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties";

		/// <summary>
		/// Custom VTypes
		/// </summary>
		internal static XNamespace CustomVTypes
			=> "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes";

		/// <summary>
		/// Word Drawing (Wp)
		/// </summary>
		internal static XNamespace Drawing
			=> "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";

		/// <summary>
		/// Word (W10)
		/// </summary>
		internal static XNamespace Word10
			=> "urn:schemas-microsoft-com:office:word";

		/// <summary>
		/// WordML (Wne)
		/// </summary>
		internal static XNamespace WordMl
			=> "http://schemas.microsoft.com/office/word/2006/wordml";

		/// <summary>
		/// Word Drawing Main (A)
		/// </summary>
		internal static XNamespace DrawingMain
			=> "http://schemas.openxmlformats.org/drawingml/2006/main";

		/// <summary>
		/// Word Drawing Chart (C)
		/// </summary>
		internal static XNamespace DrawingChart
			=> "http://schemas.openxmlformats.org/drawingml/2006/chart";

		/// <summary>
		/// VML (V)
		/// </summary>
		internal static XNamespace Vml
			=> "urn:schemas-microsoft-com:vml";

		/// <summary>
		/// Markup (Ve)
		/// </summary>
		internal static XNamespace Markup
			=> "http://schemas.openxmlformats.org/markup-compatibility/2006";

		/// <summary>
		/// Numbering (N)
		/// </summary>
		internal static XNamespace Numbering
			=> "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering";

		/// <summary>
		/// Math (M)
		/// </summary>
		internal static XNamespace Math
			=> "http://schemas.openxmlformats.org/officeDocument/2006/math";
	}
}