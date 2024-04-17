using System.Xml.Linq;
using FluentAssertions;
using OfficeFlow.Word.Core.Elements.Paragraphs;
using OfficeFlow.Word.Core.Elements.Paragraphs.Enums;
using Xunit;

namespace OfficeFlow.Word.OpenXml.Tests.OpenXmlElementWriterTests;

public sealed class ParagraphFormatTests
{
    [Theory]
    [InlineData("end", HorizontalAlignment.Right)]
    [InlineData("center", HorizontalAlignment.Center)]
    [InlineData("both", HorizontalAlignment.Both)]
    [InlineData("distribute", HorizontalAlignment.Distribute)]
    public void Should_write_horizontal_alignment_properly(string expectedValue, HorizontalAlignment actualValue)
    {
        // Arrange
        var expectedXml = new XElement(OpenXmlNamespaces.Word + "pPr",
            new XElement(OpenXmlNamespaces.Word + "jc", 
                new XAttribute(OpenXmlNamespaces.Word + "val", expectedValue)));

        var sut = new OpenXmlElementWriter(
            new XElement(OpenXmlNamespaces.Word + "pPr"));

        var paragraphFormat = new ParagraphFormat
        {
            HorizontalAlignment = actualValue
        };
        
        // Act
        sut.Visit(paragraphFormat);
        
        // Assert
        sut.Xml
            .Should()
            .Be(expectedXml);
    }

    [Fact]
    public void Default_value_of_horizontal_alignment_should_not_be_written()
    {
        // Arrange
        var expectedXml = new XElement(OpenXmlNamespaces.Word + "pPr");
        
        var sut = new OpenXmlElementWriter(
            new XElement(OpenXmlNamespaces.Word + "pPr"));

        var paragraphFormat = new ParagraphFormat
        {
            HorizontalAlignment = HorizontalAlignment.Left
        };
        
        // Act
        sut.Visit(paragraphFormat);
        
        // Assert
        sut.Xml
            .Should()
            .Be(expectedXml);
    }
}