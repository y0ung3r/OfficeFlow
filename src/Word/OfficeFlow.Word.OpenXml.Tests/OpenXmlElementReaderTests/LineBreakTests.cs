using System.Xml.Linq;
using FluentAssertions;
using OfficeFlow.Word.Core.Elements.Paragraphs.Text;
using OfficeFlow.Word.Core.Elements.Paragraphs.Text.Enums;
using Xunit;

namespace OfficeFlow.Word.OpenXml.Tests.OpenXmlElementReaderTests;

public sealed class LineBreakTests
{
    [Theory]
    [InlineData("textWrapping", LineBreakType.TextWrapping)]
    [InlineData("column", LineBreakType.Column)]
    [InlineData("page", LineBreakType.Page)]
    public void Should_read_line_break_type_properly(string actualType, LineBreakType expectedType)
    {
        // Arrange
        var xml = new XElement(OpenXmlNamespaces.Word + "br",
            new XAttribute(OpenXmlNamespaces.Word + "type", actualType));

        var lineBreak = new LineBreak();
        var sut = new OpenXmlElementReader(xml);
        
        // Act
        sut.Visit(lineBreak);
        
        // Assert
        lineBreak
            .Type
            .Should()
            .Be(expectedType);
    }

    [Fact]
    public void Break_type_should_be_default_if_attribute_is_not_specified()
    {
        // Arrange
        var xml = new XElement(OpenXmlNamespaces.Word + "br");
        
        var lineBreak = new LineBreak();
        var sut = new OpenXmlElementReader(xml);
        
        // Act
        sut.Visit(lineBreak);
        
        // Assert
        lineBreak
            .Type
            .Should()
            .Be(LineBreakType.TextWrapping);
    }
    
    [Fact]
    public void Break_type_should_be_default_if_provided_attribute_value_is_invalid()
    {
        // Arrange
        var xml = new XElement(OpenXmlNamespaces.Word + "br",
            new XAttribute(OpenXmlNamespaces.Word + "type", "invalid"));
        
        var lineBreak = new LineBreak();
        var sut = new OpenXmlElementReader(xml);
        
        // Act
        sut.Visit(lineBreak);
        
        // Assert
        lineBreak
            .Type
            .Should()
            .Be(LineBreakType.TextWrapping);
    }
}