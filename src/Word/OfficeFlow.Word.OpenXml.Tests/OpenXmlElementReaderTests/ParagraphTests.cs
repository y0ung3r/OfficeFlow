using System.Xml.Linq;
using FluentAssertions;
using OfficeFlow.Word.Core.Elements.Paragraphs;
using OfficeFlow.Word.Core.Elements.Paragraphs.Text;
using Xunit;

namespace OfficeFlow.Word.OpenXml.Tests.OpenXmlElementReaderTests;

public sealed class ParagraphTests
{
    [Fact]
    public void Should_read_runs_properly()
    {
        // Arrange
        var xml = new XElement(OpenXmlNamespaces.Word + "p",
            new XElement(OpenXmlNamespaces.Word + "pPr"),
            new XElement(OpenXmlNamespaces.Word + "r"),
            new XElement(OpenXmlNamespaces.Word + "r"),
            new XElement(OpenXmlNamespaces.Word + "r"),
            new XElement(OpenXmlNamespaces.Word + "r"),
            new XElement(OpenXmlNamespaces.Word + "r"));

        var paragraph = new Paragraph();
        var sut = new OpenXmlElementReader(xml);
        
        // Act
        sut.Visit(paragraph);
        
        // Assert
        paragraph
            .Should()
            .ContainItemsAssignableTo<Run>()
            .And
            .HaveCount(expected: 5);
    }

    [Fact]
    public void Should_read_empty_paragraph()
    {
        // Arrange
        var xml = new XElement(OpenXmlNamespaces.Word + "p",
            new XElement(OpenXmlNamespaces.Word + "pPr"));
        
        var paragraph = new Paragraph();
        var sut = new OpenXmlElementReader(xml);
        
        // Act
        sut.Visit(paragraph);
        
        // Assert
        paragraph
            .Should()
            .BeEmpty();
    }
}