using System;
using System.IO.Packaging;
using System.Xml.Linq;
using OfficeFlow.OpenXml.Packaging;

namespace OfficeFlow.OpenXml.Tests.Packaging
{
    public static class TestOpenXmlPackagePartFactory
    {
        public static OpenXmlPackagePart Create(string uri, XDocument? content = null)
            => OpenXmlPackagePart.Create(
                uri: new Uri(uri, UriKind.Relative),
                contentType: "content/type",
                compressionMode: CompressionOption.Normal,
                relationshipType: nameof(OpenXmlPackagePart.RelationshipType),
                content: content ?? CreateDefaultPackagePartContent());
        
        private static XDocument CreateDefaultPackagePartContent()
        {
            var xml = new XElement
            (
                "document",
                new XElement
                (
                    "body",
                    new XText("Content")
                )
            );

            return new XDocument(xml);
        }
    }
}