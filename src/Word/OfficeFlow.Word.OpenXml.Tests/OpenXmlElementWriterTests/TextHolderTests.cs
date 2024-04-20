using System.Xml.Linq;
using FluentAssertions;
using OfficeFlow.Word.Core.Elements.Paragraphs.Text;
using Xunit;

namespace OfficeFlow.Word.OpenXml.Tests.OpenXmlElementWriterTests;

public sealed class TextHolderTests
{
    [Fact]
    public void Should_write_text_properly()
    {
        // Assert
        var expectedXml = new XElement(OpenXmlNamespaces.Word + "t", "Text");
        
        var sut = new OpenXmlElementWriter(
            new XElement(OpenXmlNamespaces.Word + "t"));

        var textHolder = new TextHolder("Text");
        
        // Act
        sut.Visit(textHolder);
        
        // Assert
        sut.Xml
            .Should()
            .Be(expectedXml);
    }

    
    [Theory]
    [InlineData("    Text")]
    [InlineData(" Text")]     
    [InlineData("Text ")]
    [InlineData("Text    ")]
    [InlineData("Te xt")]
    [InlineData("Te     xt")]
    public void Should_preserve_spaces(string text)
    {
        // Assert
        var expectedXml = new XElement(OpenXmlNamespaces.Word + "t",
            new XAttribute(XNamespace.Xml + "space", "preserve"),
            text);
        
        var sut = new OpenXmlElementWriter(
            new XElement(OpenXmlNamespaces.Word + "t"));

        var textHolder = new TextHolder(text);
        
        // Act
        sut.Visit(textHolder);
        
        // Assert
        sut.Xml
            .Should()
            .Be(expectedXml);
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    public void Should_not_write_when_value_is_empty_or_null(string? text)
    {
        // Assert
        var expectedXml = new XElement(OpenXmlNamespaces.Word + "t");
        
        var sut = new OpenXmlElementWriter(
            new XElement(OpenXmlNamespaces.Word + "t"));

        var textHolder = new TextHolder(text!);
        
        // Act
        sut.Visit(textHolder);
        
        // Assert
        sut.Xml
            .Should()
            .Be(expectedXml);
    }
}