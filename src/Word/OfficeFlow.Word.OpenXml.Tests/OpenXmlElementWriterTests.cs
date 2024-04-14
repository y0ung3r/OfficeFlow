using System.Xml.Linq;
using FluentAssertions;
using OfficeFlow.Word.Core.Elements;
using OfficeFlow.Word.Core.Elements.Paragraphs;
using Xunit;

namespace OfficeFlow.Word.OpenXml.Tests;

public sealed class OpenXmlElementWriterTests
{
    [Fact]
    public void Should_write_body_element_properly()
    {
        // Arrange
        var expectedXml = new XElement(
            new XElement(OpenXmlNamespaces.Word + "document",
                new XElement(OpenXmlNamespaces.Word + "body",
                    new XElement(OpenXmlNamespaces.Word + "p"),
                    new XElement(OpenXmlNamespaces.Word + "sectPr"))));

        var sut = new OpenXmlElementWriter(
            new XElement(OpenXmlNamespaces.Word + "document"));
            
        var bodyElement = new Body();
            
        bodyElement.AppendChild(
            new Paragraph());
            
        bodyElement.AppendChild(
            new Section());

        // Act
        sut.Visit(bodyElement);
            
        // Assert
        sut.Xml
            .Should()
            .Be(expectedXml);
    }
}