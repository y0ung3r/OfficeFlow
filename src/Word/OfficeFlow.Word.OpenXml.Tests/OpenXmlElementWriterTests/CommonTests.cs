using System.Xml.Linq;
using FluentAssertions;
using Xunit;

namespace OfficeFlow.Word.OpenXml.Tests.OpenXmlElementWriterTests;

public sealed class CommonTests
{
    [Fact]
    public void Should_clear_xml_before_writing()
    {
        // Arrange
        var expectedXml = new XElement(OpenXmlNamespaces.Word + "document");
        
        // Act
        var sut = new OpenXmlElementWriter(
            new XElement(OpenXmlNamespaces.Word + "document",
                new XElement(OpenXmlNamespaces.Word + "body")));
        
        // Assert
        sut.Xml
            .Should()
            .Be(expectedXml);
    }
}