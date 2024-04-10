using System.Xml.Linq;

namespace OfficeFlow.OpenXml.Resources.Interfaces
{
    public interface IResourceDecompressor
    {
        XDocument Decompress(string resourceName);
    }
}