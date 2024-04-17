using System.Xml.Linq;
using FluentAssertions;
using OfficeFlow.Word.Core.Elements.Paragraphs;
using OfficeFlow.Word.Core.Elements.Paragraphs.Enums;
using Xunit;

namespace OfficeFlow.Word.OpenXml.Tests.OpenXmlElementReaderTests;

public sealed class ParagraphFormatTests
{
    [Theory]
    [InlineData("left", HorizontalAlignment.Left)]
    [InlineData("start", HorizontalAlignment.Left)]
    [InlineData("right", HorizontalAlignment.Right)]
    [InlineData("end", HorizontalAlignment.Right)]
    [InlineData("center", HorizontalAlignment.Center)]
    [InlineData("both", HorizontalAlignment.Both)]
    [InlineData("distribute", HorizontalAlignment.Distribute)]
    public void Should_read_horizontal_alignment_properly(string actualValue, HorizontalAlignment expectedValue)
    {
        // Arrange
        var xml = new XElement(OpenXmlNamespaces.Word + "pPr",
            new XElement(OpenXmlNamespaces.Word + "jc", 
                new XAttribute(OpenXmlNamespaces.Word + "val", actualValue)));

        var paragraphFormat = new ParagraphFormat();
        var sut = new OpenXmlElementReader(xml);

        // Act
        sut.Visit(paragraphFormat);
        
        // Assert
        paragraphFormat
            .HorizontalAlignment
            .Should()
            .Be(expectedValue);
    }

    [Fact]
    public void Should_read_empty_paragraph_format_properly()
    {
        // Arrange
        var xml = new XElement(OpenXmlNamespaces.Word + "pPr");
        
        var paragraphFormat = new ParagraphFormat();
        var sut = new OpenXmlElementReader(xml);
        
        // Act
        sut.Visit(paragraphFormat);
        
        // Assert
        paragraphFormat
            .HorizontalAlignment
            .Should()
            .Be(ParagraphFormat.Default.HorizontalAlignment);
    }
}