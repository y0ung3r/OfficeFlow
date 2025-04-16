using System;
using System.IO.Packaging;
using System.Xml.Linq;
using OfficeFlow.OpenXml.Packaging;

namespace OfficeFlow.OpenXml.Tests.Packaging;

public static class TestOpenXmlPackagePartFactory
{
    private static readonly XDocument DefaultContent
        = new(
            new XElement
            (
                "document",
                new XElement
                (
                    "body",
                    new XText("Content")
                )
            ));

    public static OpenXmlPackagePart Create(string uri, XDocument? content = null)
        => Create(
            uri, 
            relationshipType: nameof(OpenXmlPackagePart.RelationshipType),
            content);
    
    public static OpenXmlPackagePart Create(string uri, string relationshipType, XDocument? content = null)
        => OpenXmlPackagePart.Create(
            uri: new Uri(uri, UriKind.Relative),
            contentType: "content/type",
            compressionMode: CompressionOption.Normal,
            relationshipType: relationshipType,
            content: content ?? DefaultContent);
}