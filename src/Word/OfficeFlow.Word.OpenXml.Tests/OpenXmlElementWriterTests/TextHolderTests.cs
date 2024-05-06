using System.Xml.Linq;
using FluentAssertions;
using OfficeFlow.TestFramework.Extensions;
using OfficeFlow.Word.Core.Elements.Paragraphs.Text;
using Xunit;

namespace OfficeFlow.Word.OpenXml.Tests.OpenXmlElementWriterTests;

public sealed class TextHolderTests
{
    [Fact]
    public void Should_write_text_properly()
    {
        // Assert
        var sut = new OpenXmlElementWriter(
            new XElement(OpenXmlNamespaces.Word + "t"));

        var textHolder = new TextHolder("Text");
        
        // Act
        sut.Visit(textHolder);
        
        // Assert
        sut.Xml
            .Should()
            .HaveName(OpenXmlNamespaces.Word + "t")
            .And
            .HaveValue("Text");
    }

    
    [Theory]
    [InlineData("    Text")]
    [InlineData(" Text")]     
    [InlineData("Text ")]
    [InlineData("Text    ")]
    [InlineData("Te xt")]
    [InlineData("Te     xt")]
    public void Should_preserve_spaces(string expectedText)
    {
        // Assert
        var sut = new OpenXmlElementWriter(
            new XElement(OpenXmlNamespaces.Word + "t"));

        var textHolder = new TextHolder(expectedText);
        
        // Act
        sut.Visit(textHolder);
        
        // Assert
        sut.Xml
            .Should()
            .HaveName(OpenXmlNamespaces.Word + "t")
            .And
            .HaveAttribute(XNamespace.Xml + "space", "preserve")
            .And
            .HaveValue(expectedText);
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    public void Should_not_write_when_value_is_empty_or_null(string? expectedText)
    {
        // Assert
        var sut = new OpenXmlElementWriter(
            new XElement(OpenXmlNamespaces.Word + "t"));

        var textHolder = new TextHolder(expectedText!);
        
        // Act
        sut.Visit(textHolder);
        
        // Assert
        sut.Xml
            .Should()
            .HaveName(OpenXmlNamespaces.Word + "t")
            .And
            .Subject
            .Elements()
            .Should()
            .BeEmpty()
            .And
            .Subject
            .Attributes()
            .Should()
            .BeEmpty();
    }
}