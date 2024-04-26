using System.Xml.Linq;
using FluentAssertions;
using OfficeFlow.MeasureUnits.Absolute;
using OfficeFlow.Word.Core.Elements.Paragraphs;
using OfficeFlow.Word.Core.Elements.Paragraphs.Enums;
using OfficeFlow.Word.Core.Elements.Paragraphs.Spacing;
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
                new XAttribute(OpenXmlNamespaces.Word + "val", expectedValue)),
            new XElement(OpenXmlNamespaces.Word + "spacing",
                new XAttribute(OpenXmlNamespaces.Word + "before", "0"),
                new XAttribute(OpenXmlNamespaces.Word + "after", "160")));

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
    public void Should_write_paragraph_spacing_properly()
    {
        // Arrange
        var expectedXml = new XElement(OpenXmlNamespaces.Word + "pPr",
            new XElement(OpenXmlNamespaces.Word + "spacing",
                new XAttribute(OpenXmlNamespaces.Word + "beforeAutospacing", "true"),
                new XAttribute(OpenXmlNamespaces.Word + "after", "240")));
        
        var sut = new OpenXmlElementWriter(
            new XElement(OpenXmlNamespaces.Word + "pPr"));
        
        var paragraphFormat = new ParagraphFormat
        {
            SpacingBefore = ParagraphSpacing.Auto,
            SpacingAfter = ParagraphSpacing.Exactly<Points>(12.0)
        };
        
        // Act
        sut.Visit(paragraphFormat);
        
        // Assert
        sut.Xml
            .Should()
            .Be(expectedXml);
    }
    
    [Fact]
    public void Spacing_before_should_be_calculated_automatically()
    {
        // Arrange
        var expectedXml = new XElement(OpenXmlNamespaces.Word + "pPr",
            new XElement(OpenXmlNamespaces.Word + "spacing",
                new XAttribute(OpenXmlNamespaces.Word + "beforeAutospacing", "true"),
                new XAttribute(OpenXmlNamespaces.Word + "after", "160")));
        
        var sut = new OpenXmlElementWriter(
            new XElement(OpenXmlNamespaces.Word + "pPr"));
        
        var paragraphFormat = new ParagraphFormat
        {
            SpacingBefore = ParagraphSpacing.Auto
        };
        
        // Act
        sut.Visit(paragraphFormat);
        
        // Assert
        sut.Xml
            .Should()
            .Be(expectedXml);
    }

    [Fact]
    public void Should_write_exact_value_of_spacing_before_properly()
    {
        // Arrange
        var expectedXml = new XElement(OpenXmlNamespaces.Word + "pPr",
            new XElement(OpenXmlNamespaces.Word + "spacing",
                new XAttribute(OpenXmlNamespaces.Word + "before", "240"),
                new XAttribute(OpenXmlNamespaces.Word + "after", "160")));
        
        var sut = new OpenXmlElementWriter(
            new XElement(OpenXmlNamespaces.Word + "pPr"));
        
        var paragraphFormat = new ParagraphFormat
        {
            SpacingBefore = ParagraphSpacing.Exactly<Points>(12.0)
        };
        
        // Act
        sut.Visit(paragraphFormat);
        
        // Assert
        sut.Xml
            .Should()
            .Be(expectedXml);
    }
    
    [Fact]
    public void Spacing_after_should_be_calculated_automatically()
    {
        // Arrange
        var expectedXml = new XElement(OpenXmlNamespaces.Word + "pPr",
            new XElement(OpenXmlNamespaces.Word + "spacing",
                new XAttribute(OpenXmlNamespaces.Word + "before", "0"),
                new XAttribute(OpenXmlNamespaces.Word + "afterAutospacing", "true")));
        
        var sut = new OpenXmlElementWriter(
            new XElement(OpenXmlNamespaces.Word + "pPr"));
        
        var paragraphFormat = new ParagraphFormat
        {
            SpacingAfter = ParagraphSpacing.Auto
        };
        
        // Act
        sut.Visit(paragraphFormat);
        
        // Assert
        sut.Xml
            .Should()
            .Be(expectedXml);
    }

    [Fact]
    public void Should_write_exact_value_of_spacing_after_properly()
    {
        // Arrange
        var expectedXml = new XElement(OpenXmlNamespaces.Word + "pPr",
            new XElement(OpenXmlNamespaces.Word + "spacing",
                new XAttribute(OpenXmlNamespaces.Word + "before", "0"),
                new XAttribute(OpenXmlNamespaces.Word + "after", "240")));
        
        var sut = new OpenXmlElementWriter(
            new XElement(OpenXmlNamespaces.Word + "pPr"));
        
        var paragraphFormat = new ParagraphFormat
        {
            SpacingAfter = ParagraphSpacing.Exactly<Points>(12.0)
        };
        
        // Act
        sut.Visit(paragraphFormat);
        
        // Assert
        sut.Xml
            .Should()
            .Be(expectedXml);
    }

    [Fact]
    public void Should_write_keep_lines_properly()
    {
        // Arrange
        var expectedXml = new XElement(OpenXmlNamespaces.Word + "pPr",
            new XElement(OpenXmlNamespaces.Word + "spacing",
                new XAttribute(OpenXmlNamespaces.Word + "before", "0"),
                new XAttribute(OpenXmlNamespaces.Word + "after", "160")),
            new XElement(OpenXmlNamespaces.Word + "keepLines"));
        
        var sut = new OpenXmlElementWriter(
            new XElement(OpenXmlNamespaces.Word + "pPr"));

        var paragraphFormat = new ParagraphFormat
        {
            KeepLines = true
        };
        
        // Act
        sut.Visit(paragraphFormat);
        
        // Assert
        sut.Xml
            .Should()
            .Be(expectedXml);
    }

    [Fact]
    public void Should_not_write_keep_lines()
    {
        // Arrange
        var expectedXml = new XElement(OpenXmlNamespaces.Word + "pPr",
            new XElement(OpenXmlNamespaces.Word + "spacing",
                new XAttribute(OpenXmlNamespaces.Word + "before", "0"),
                new XAttribute(OpenXmlNamespaces.Word + "after", "160")));
        
        var sut = new OpenXmlElementWriter(
            new XElement(OpenXmlNamespaces.Word + "pPr"));

        var paragraphFormat = new ParagraphFormat
        {
            KeepLines = false
        };
        
        // Act
        sut.Visit(paragraphFormat);
        
        // Assert
        sut.Xml
            .Should()
            .Be(expectedXml);
    }
    
    [Fact]
    public void Should_write_keep_next_properly()
    {
        // Arrange
        var expectedXml = new XElement(OpenXmlNamespaces.Word + "pPr",
            new XElement(OpenXmlNamespaces.Word + "spacing",
                new XAttribute(OpenXmlNamespaces.Word + "before", "0"),
                new XAttribute(OpenXmlNamespaces.Word + "after", "160")),
            new XElement(OpenXmlNamespaces.Word + "keepNext"));
        
        var sut = new OpenXmlElementWriter(
            new XElement(OpenXmlNamespaces.Word + "pPr"));

        var paragraphFormat = new ParagraphFormat
        {
            KeepNext = true
        };
        
        // Act
        sut.Visit(paragraphFormat);
        
        // Assert
        sut.Xml
            .Should()
            .Be(expectedXml);
    }

    [Fact]
    public void Should_not_write_keep_next()
    {
        // Arrange
        var expectedXml = new XElement(OpenXmlNamespaces.Word + "pPr",
            new XElement(OpenXmlNamespaces.Word + "spacing",
                new XAttribute(OpenXmlNamespaces.Word + "before", "0"),
                new XAttribute(OpenXmlNamespaces.Word + "after", "160")));
        
        var sut = new OpenXmlElementWriter(
            new XElement(OpenXmlNamespaces.Word + "pPr"));

        var paragraphFormat = new ParagraphFormat
        {
            KeepNext = false
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
        var expectedXml = new XElement(OpenXmlNamespaces.Word + "pPr",
            new XElement(OpenXmlNamespaces.Word + "spacing",
                new XAttribute(OpenXmlNamespaces.Word + "before", "0"),
                new XAttribute(OpenXmlNamespaces.Word + "after", "160")));
        
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