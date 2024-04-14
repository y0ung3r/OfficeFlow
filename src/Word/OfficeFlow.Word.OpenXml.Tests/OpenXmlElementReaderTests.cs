using System.Linq;
using System.Xml.Linq;
using FluentAssertions;
using OfficeFlow.Word.Core.Elements;
using OfficeFlow.Word.Core.Elements.Paragraphs;
using Xunit;

namespace OfficeFlow.Word.OpenXml.Tests;

public sealed class OpenXmlElementReaderTests
{
    [Fact]
    public void Should_read_body_element_properly()
    {
        // Arrange
        var xml = new XElement(
            new XElement(OpenXmlNamespaces.Word + "document",
                new XElement(OpenXmlNamespaces.Word + "body",
                    new XElement(OpenXmlNamespaces.Word + "p"),
                    new XElement(OpenXmlNamespaces.Word + "sectPr"))));
            
        var bodyElement = new Body();
        var sut = new OpenXmlElementReader(xml);
        
        // Act
        sut.Visit(bodyElement);
        
        // Assert
        bodyElement
            .Select(child => child?.GetType())
            .Should()
            .ContainInOrder(
                typeof(Paragraph),
                typeof(Section));
    }
}