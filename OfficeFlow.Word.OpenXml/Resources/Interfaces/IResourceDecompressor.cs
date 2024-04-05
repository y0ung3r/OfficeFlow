using System.Xml.Linq;

namespace OfficeFlow.Word.OpenXml.Resources.Interfaces
{
    public interface IResourceDecompressor
    {
        XDocument Decompress(string resourceName);
    }
}