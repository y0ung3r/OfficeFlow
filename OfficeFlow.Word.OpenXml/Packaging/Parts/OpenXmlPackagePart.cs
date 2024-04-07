using System.IO;
using System.IO.Packaging;
using System.Xml.Linq;

namespace OfficeFlow.Word.OpenXml.Packaging.Parts
{
    public class OpenXmlPackagePart
    {
        public static OpenXmlPackagePart Load(PackagePart source)
        {
            using var contentStream =
                source.GetStream(FileMode.Open, FileAccess.Read);
            
            var xml = XDocument.Load(contentStream);

            return new OpenXmlPackagePart(source, xml);
        }
        
        private readonly PackagePart _source;
        private readonly XDocument _xml;

        private OpenXmlPackagePart(PackagePart source, XDocument xml)
        {
            _source = source;
            _xml = xml;
        }
        
        public void Flush()
        {
            using var contentStream = 
                _source.GetStream(FileMode.Create, FileAccess.Write);
            
            using var synchronizedStream = 
                Stream.Synchronized(contentStream);
            
            using var contentWriter = 
                new StreamWriter(synchronizedStream);
            
            _xml.Save(contentWriter, SaveOptions.None);
        }
    }
}