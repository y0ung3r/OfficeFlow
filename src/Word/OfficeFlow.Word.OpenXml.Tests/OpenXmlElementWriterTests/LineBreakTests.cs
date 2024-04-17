using System.Xml.Linq;
using FluentAssertions;
using OfficeFlow.Word.Core.Elements.Paragraphs.Text;
using OfficeFlow.Word.Core.Elements.Paragraphs.Text.Enums;
using Xunit;

namespace OfficeFlow.Word.OpenXml.Tests.OpenXmlElementWriterTests;

public sealed class LineBreakTests
{
    [Theory]
    [InlineData("column", LineBreakType.Column)]
    [InlineData("page", LineBreakType.Page)]
    public void Should_write_line_break_type_properly(string expectedType, LineBreakType actualType)
    {
        // Assert
        var expectedXml = new XElement(OpenXmlNamespaces.Word + "br",
            new XAttribute(OpenXmlNamespaces.Word + "type", expectedType));
        
        var sut = new OpenXmlElementWriter(
            new XElement(OpenXmlNamespaces.Word + "br"));

        var @break = new LineBreak
        {
            Type = actualType
        };

        // Act
        sut.Visit(@break);
        
        // Assert
        sut.Xml
            .Should()
            .Be(expectedXml);
    }

    [Fact]
    public void Default_value_of_break_type_should_not_be_written()
    {
        // Assert
        var expectedXml = new XElement(OpenXmlNamespaces.Word + "br");
        
        var sut = new OpenXmlElementWriter(
            new XElement(OpenXmlNamespaces.Word + "br"));

        var @break = new LineBreak
        {
            Type = LineBreakType.TextWrapping
        };

        // Act
        sut.Visit(@break);
        
        // Assert
        sut.Xml
            .Should()
            .Be(expectedXml);
    }
}