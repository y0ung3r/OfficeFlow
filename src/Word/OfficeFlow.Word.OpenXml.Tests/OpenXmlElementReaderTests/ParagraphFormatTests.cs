using System.Xml.Linq;
using FluentAssertions;
using OfficeFlow.MeasureUnits.Absolute;
using OfficeFlow.Word.Core.Elements.Paragraphs;
using OfficeFlow.Word.Core.Elements.Paragraphs.Enums;
using OfficeFlow.Word.Core.Elements.Paragraphs.Spacing;
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
    public void Should_read_paragraph_spacing_properly()
    {
        // Arrange
        var xml = new XElement(OpenXmlNamespaces.Word + "pPr",
            new XElement(OpenXmlNamespaces.Word + "spacing",
                new XAttribute(OpenXmlNamespaces.Word + "before", "240"),
                new XAttribute(OpenXmlNamespaces.Word + "beforeAutospacing", "true"),
                new XAttribute(OpenXmlNamespaces.Word + "after", "240"),
                new XAttribute(OpenXmlNamespaces.Word + "afterAutospacing", "false")));

        var paragraphFormat = new ParagraphFormat();
        var sut = new OpenXmlElementReader(xml);
        
        // Act
        sut.Visit(paragraphFormat);
        
        // Assert
        paragraphFormat
            .SpacingBefore
            .Should()
            .BeOfType<AutoSpacing>();
        
        paragraphFormat
            .SpacingAfter
            .Should()
            .BeOfType<ExactSpacing>()
            .And
            .Be(ParagraphSpacing.Exactly<Points>(12.0));
    }

    [Theory]
    [InlineData("1")]
    [InlineData("true")]
    public void Spacing_after_should_be_calculated_automatically(string value)
    {
        // Arrange
        var xml = new XElement(OpenXmlNamespaces.Word + "pPr",
            new XElement(OpenXmlNamespaces.Word + "spacing",
                new XAttribute(OpenXmlNamespaces.Word + "afterAutospacing", value)));

        var paragraphFormat = new ParagraphFormat();
        var sut = new OpenXmlElementReader(xml);
        
        // Act
        sut.Visit(paragraphFormat);
        
        // Assert
        paragraphFormat
            .SpacingAfter
            .Should()
            .BeOfType<AutoSpacing>();
    }
    
    [Theory]
    [InlineData("1")]
    [InlineData("true")]
    public void Spacing_before_should_be_calculated_automatically(string value)
    {
        // Arrange
        var xml = new XElement(OpenXmlNamespaces.Word + "pPr",
            new XElement(OpenXmlNamespaces.Word + "spacing",
                new XAttribute(OpenXmlNamespaces.Word + "beforeAutospacing", value)));

        var paragraphFormat = new ParagraphFormat();
        var sut = new OpenXmlElementReader(xml);
        
        // Act
        sut.Visit(paragraphFormat);
        
        // Assert
        paragraphFormat
            .SpacingBefore
            .Should()
            .BeOfType<AutoSpacing>();
    }

    [Fact]
    public void Should_read_exact_value_of_spacing_after_properly()
    {
        // Arrange
        var xml = new XElement(OpenXmlNamespaces.Word + "pPr",
            new XElement(OpenXmlNamespaces.Word + "spacing",
                new XAttribute(OpenXmlNamespaces.Word + "after", "240")));

        var paragraphFormat = new ParagraphFormat();
        var sut = new OpenXmlElementReader(xml);
        
        // Act
        sut.Visit(paragraphFormat);
        
        // Assert
        paragraphFormat
            .SpacingAfter
            .Should()
            .BeOfType<ExactSpacing>()
            .And
            .Be(ParagraphSpacing.Exactly<Points>(12.0));
    }
    
    [Fact]
    public void Should_read_exact_value_of_spacing_before_properly()
    {
        // Arrange
        var xml = new XElement(OpenXmlNamespaces.Word + "pPr",
            new XElement(OpenXmlNamespaces.Word + "spacing",
                new XAttribute(OpenXmlNamespaces.Word + "before", "240")));

        var paragraphFormat = new ParagraphFormat();
        var sut = new OpenXmlElementReader(xml);
        
        // Act
        sut.Visit(paragraphFormat);
        
        // Assert
        paragraphFormat
            .SpacingBefore
            .Should()
            .BeOfType<ExactSpacing>()
            .And
            .Be(ParagraphSpacing.Exactly<Points>(12.0));
    }

    [Fact]
    public void Should_read_keep_lines_properly()
    {
        // Arrange
        var xml = new XElement(OpenXmlNamespaces.Word + "pPr",
            new XElement(OpenXmlNamespaces.Word + "keepLines"));
        
        var paragraphFormat = new ParagraphFormat();
        var sut = new OpenXmlElementReader(xml);
        
        // Act
        sut.Visit(paragraphFormat);
        
        // Assert
        paragraphFormat
            .KeepLines
            .Should()
            .BeTrue();
    }
    
    [Fact]
    public void Should_read_keep_next_properly()
    {
        // Arrange
        var xml = new XElement(OpenXmlNamespaces.Word + "pPr",
            new XElement(OpenXmlNamespaces.Word + "keepNext"));
        
        var paragraphFormat = new ParagraphFormat();
        var sut = new OpenXmlElementReader(xml);
        
        // Act
        sut.Visit(paragraphFormat);
        
        // Assert
        paragraphFormat
            .KeepNext
            .Should()
            .BeTrue();
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

        paragraphFormat
            .SpacingBefore
            .Should()
            .Be(ParagraphFormat.Default.SpacingBefore);
        
        paragraphFormat
            .SpacingAfter
            .Should()
            .Be(ParagraphFormat.Default.SpacingAfter);

        paragraphFormat
            .KeepLines
            .Should()
            .BeFalse();
        
        paragraphFormat
            .KeepNext
            .Should()
            .BeFalse();
    }
}