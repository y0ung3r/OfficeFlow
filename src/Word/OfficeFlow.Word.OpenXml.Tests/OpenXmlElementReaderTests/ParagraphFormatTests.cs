using System.Xml.Linq;
using FluentAssertions;
using OfficeFlow.MeasureUnits.Absolute;
using OfficeFlow.Word.Core.Elements.Paragraphs.Formatting;
using OfficeFlow.Word.Core.Elements.Paragraphs.Formatting.Enums;
using OfficeFlow.Word.Core.Elements.Paragraphs.Formatting.Spacing;
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
    
    [Theory]
    [InlineData("auto", TextAlignment.Auto)]
    [InlineData("baseline", TextAlignment.Baseline)]
    [InlineData("bottom", TextAlignment.Bottom)]
    [InlineData("center", TextAlignment.Center)]
    [InlineData("top", TextAlignment.Top)]
    public void Should_read_text_alignment_properly(string actualValue, TextAlignment expectedValue)
    {
        // Arrange
        var xml = new XElement(OpenXmlNamespaces.Word + "pPr",
            new XElement(OpenXmlNamespaces.Word + "textAlignment", 
                new XAttribute(OpenXmlNamespaces.Word + "val", actualValue)));

        var paragraphFormat = new ParagraphFormat();
        var sut = new OpenXmlElementReader(xml);

        // Act
        sut.Visit(paragraphFormat);
        
        // Assert
        paragraphFormat
            .TextAlignment
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
            .Be(ParagraphSpacing.Exactly(12.0, AbsoluteUnits.Points));
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
            .Be(ParagraphSpacing.Exactly(12.0, AbsoluteUnits.Points));
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
            .Be(ParagraphSpacing.Exactly(12.0, AbsoluteUnits.Points));
    }
    
    [Fact]
    public void Line_spacing_should_be_calculated_automatically_if_rule_is_not_specified()
    {
        // Arrange
        var xml = new XElement(OpenXmlNamespaces.Word + "pPr",
            new XElement(OpenXmlNamespaces.Word + "spacing",
                new XAttribute(OpenXmlNamespaces.Word + "line", "360")));

        var paragraphFormat = new ParagraphFormat();
        var sut = new OpenXmlElementReader(xml);
        
        // Act
        sut.Visit(paragraphFormat);
        
        // Assert
        paragraphFormat
            .SpacingBetweenLines
            .Should()
            .BeOfType<MultipleSpacing>()
            .And
            .Be(ParagraphFormat.Default.SpacingBetweenLines);
    }
    
    [Fact]
    public void Line_spacing_should_be_calculated_automatically_if_value_is_not_specified()
    {
        // Arrange
        var xml = new XElement(OpenXmlNamespaces.Word + "pPr",
            new XElement(OpenXmlNamespaces.Word + "spacing",
                new XAttribute(OpenXmlNamespaces.Word + "lineRule", "auto")));

        var paragraphFormat = new ParagraphFormat();
        var sut = new OpenXmlElementReader(xml);
        
        // Act
        sut.Visit(paragraphFormat);
        
        // Assert
        paragraphFormat
            .SpacingBetweenLines
            .Should()
            .BeOfType<MultipleSpacing>()
            .And
            .Be(ParagraphFormat.Default.SpacingBetweenLines);
    }

    [Fact]
    public void Should_read_exact_value_of_line_spacing_properly()
    {
        // Arrange
        var xml = new XElement(OpenXmlNamespaces.Word + "pPr",
            new XElement(OpenXmlNamespaces.Word + "spacing",
                new XAttribute(OpenXmlNamespaces.Word + "lineRule", "exactly"),
                new XAttribute(OpenXmlNamespaces.Word + "line", "360")));

        var paragraphFormat = new ParagraphFormat();
        var sut = new OpenXmlElementReader(xml);
        
        // Act
        sut.Visit(paragraphFormat);
        
        // Assert
        paragraphFormat
            .SpacingBetweenLines
            .Should()
            .BeOfType<ExactSpacing>()
            .And
            .Be(LineSpacing.Exactly(18.0, AbsoluteUnits.Points));
    }
    
    [Fact]
    public void Should_read_at_least_value_of_line_spacing_properly()
    {
        // Arrange
        var xml = new XElement(OpenXmlNamespaces.Word + "pPr",
            new XElement(OpenXmlNamespaces.Word + "spacing",
                new XAttribute(OpenXmlNamespaces.Word + "lineRule", "atLeast"),
                new XAttribute(OpenXmlNamespaces.Word + "line", "360")));

        var paragraphFormat = new ParagraphFormat();
        var sut = new OpenXmlElementReader(xml);
        
        // Act
        sut.Visit(paragraphFormat);
        
        // Assert
        paragraphFormat
            .SpacingBetweenLines
            .Should()
            .BeOfType<AtLeastSpacing>()
            .And
            .Be(LineSpacing.AtLeast(18.0, AbsoluteUnits.Points));
    }
    
    [Fact]
    public void Should_read_multiple_value_of_line_spacing_properly()
    {
        // Arrange
        var xml = new XElement(OpenXmlNamespaces.Word + "pPr",
            new XElement(OpenXmlNamespaces.Word + "spacing",
                new XAttribute(OpenXmlNamespaces.Word + "lineRule", "auto"),
                new XAttribute(OpenXmlNamespaces.Word + "line", "360")));

        var paragraphFormat = new ParagraphFormat();
        var sut = new OpenXmlElementReader(xml);
        
        // Act
        sut.Visit(paragraphFormat);
        
        // Assert
        paragraphFormat
            .SpacingBetweenLines
            .Should()
            .BeOfType<MultipleSpacing>()
            .And
            .Be(LineSpacing.Multiple(1.5));
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
            .TextAlignment
            .Should()
            .Be(ParagraphFormat.Default.TextAlignment);
        
        paragraphFormat
            .SpacingBefore
            .Should()
            .Be(ParagraphFormat.Default.SpacingBefore);
        
        paragraphFormat
            .SpacingAfter
            .Should()
            .Be(ParagraphFormat.Default.SpacingAfter);
        
        paragraphFormat
            .SpacingBetweenLines
            .Should()
            .Be(ParagraphFormat.Default.SpacingBetweenLines);

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