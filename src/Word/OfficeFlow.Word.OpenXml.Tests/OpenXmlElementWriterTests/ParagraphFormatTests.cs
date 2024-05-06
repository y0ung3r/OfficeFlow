using System.Xml.Linq;
using FluentAssertions;
using OfficeFlow.MeasureUnits.Absolute;
using OfficeFlow.TestFramework.Extensions;
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
            .HaveName(OpenXmlNamespaces.Word + "pPr")
            .And
            .HaveElement(OpenXmlNamespaces.Word + "jc")
            .Which
            .Should()
            .HaveAttribute(OpenXmlNamespaces.Word + "val", expectedValue);
    }
    
    [Fact]
    public void Default_value_of_horizontal_alignment_should_not_be_written()
    {
        // Arrange
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
            .HaveName(OpenXmlNamespaces.Word + "pPr")
            .And
            .Subject
            .Elements()
            .Should()
            .AllSatisfy(elementXml => 
                elementXml
                    .Should()
                    .NotHaveName(OpenXmlNamespaces.Word + "jc"));
    }
    
    [Theory]
    [InlineData("baseline", TextAlignment.Baseline)]
    [InlineData("bottom", TextAlignment.Bottom)]
    [InlineData("center", TextAlignment.Center)]
    [InlineData("top", TextAlignment.Top)]
    public void Should_write_text_alignment_properly(string expectedValue, TextAlignment actualValue)
    {
        // Arrange
        var sut = new OpenXmlElementWriter(
            new XElement(OpenXmlNamespaces.Word + "pPr"));

        var paragraphFormat = new ParagraphFormat
        {
            TextAlignment = actualValue
        };
        
        // Act
        sut.Visit(paragraphFormat);
        
        // Assert
        sut.Xml
            .Should()
            .HaveName(OpenXmlNamespaces.Word + "pPr")
            .And
            .HaveElement(OpenXmlNamespaces.Word + "textAlignment")
            .Which
            .Should()
            .HaveAttribute(OpenXmlNamespaces.Word + "val", expectedValue);
    }
    
    [Fact]
    public void Default_value_of_text_alignment_should_not_be_written()
    {
        // Arrange
        var sut = new OpenXmlElementWriter(
            new XElement(OpenXmlNamespaces.Word + "pPr"));

        var paragraphFormat = new ParagraphFormat
        {
            TextAlignment = TextAlignment.Auto
        };
        
        // Act
        sut.Visit(paragraphFormat);
        
        // Assert
        sut.Xml
            .Should()
            .HaveName(OpenXmlNamespaces.Word + "pPr")
            .And
            .Subject
            .Elements()
            .Should()
            .AllSatisfy(elementXml => 
                elementXml
                    .Should()
                    .NotHaveName(OpenXmlNamespaces.Word + "textAlignment"));
    }

    [Fact]
    public void Should_write_paragraph_spacing_properly()
    {
        // Arrange
        var sut = new OpenXmlElementWriter(
            new XElement(OpenXmlNamespaces.Word + "pPr"));
        
        var paragraphFormat = new ParagraphFormat
        {
            SpacingBefore = ParagraphSpacing.Auto,
            SpacingAfter = ParagraphSpacing.Exactly(12.0, AbsoluteUnits.Points)
        };
        
        // Act
        sut.Visit(paragraphFormat);
        
        // Assert
        sut.Xml
            .Should()
            .HaveName(OpenXmlNamespaces.Word + "pPr")
            .And
            .HaveElement(OpenXmlNamespaces.Word + "spacing")
            .Which
            .Should()
            .HaveAttribute(OpenXmlNamespaces.Word + "beforeAutospacing", "true")
            .And
            .HaveAttribute(OpenXmlNamespaces.Word + "after", "240");
    }
    
    [Fact]
    public void Spacing_before_should_be_calculated_automatically()
    {
        // Arrange
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
            .HaveName(OpenXmlNamespaces.Word + "pPr")
            .And
            .HaveElement(OpenXmlNamespaces.Word + "spacing")
            .Which
            .Should()
            .HaveAttribute(OpenXmlNamespaces.Word + "beforeAutospacing", "true")
            .And
            .HaveAttribute(OpenXmlNamespaces.Word + "after", "160");
    }

    [Fact]
    public void Should_write_exact_value_of_spacing_before_properly()
    {
        // Arrange
        var sut = new OpenXmlElementWriter(
            new XElement(OpenXmlNamespaces.Word + "pPr"));
        
        var paragraphFormat = new ParagraphFormat
        {
            SpacingBefore = ParagraphSpacing.Exactly(12.0, AbsoluteUnits.Points)
        };
        
        // Act
        sut.Visit(paragraphFormat);
        
        // Assert
        sut.Xml
            .Should()
            .HaveName(OpenXmlNamespaces.Word + "pPr")
            .And
            .HaveElement(OpenXmlNamespaces.Word + "spacing")
            .Which
            .Should()
            .HaveAttribute(OpenXmlNamespaces.Word + "before", "240")
            .And
            .HaveAttribute(OpenXmlNamespaces.Word + "after", "160");
    }
    
    [Fact]
    public void Spacing_after_should_be_calculated_automatically()
    {
        // Arrange
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
            .HaveName(OpenXmlNamespaces.Word + "pPr")
            .And
            .HaveElement(OpenXmlNamespaces.Word + "spacing")
            .Which
            .Should()
            .HaveAttribute(OpenXmlNamespaces.Word + "before", "0")
            .And
            .HaveAttribute(OpenXmlNamespaces.Word + "afterAutospacing", "true");
    }

    [Fact]
    public void Should_write_exact_value_of_spacing_after_properly()
    {
        // Arrange
        var sut = new OpenXmlElementWriter(
            new XElement(OpenXmlNamespaces.Word + "pPr"));
        
        var paragraphFormat = new ParagraphFormat
        {
            SpacingAfter = ParagraphSpacing.Exactly(12.0, AbsoluteUnits.Points)
        };
        
        // Act
        sut.Visit(paragraphFormat);
        
        // Assert
        sut.Xml
            .Should()
            .Be(expectedXml);
            .HaveName(OpenXmlNamespaces.Word + "pPr")
            .And
            .HaveElement(OpenXmlNamespaces.Word + "spacing")
            .Which
            .Should()
            .HaveAttribute(OpenXmlNamespaces.Word + "before", "0")
            .And
            .HaveAttribute(OpenXmlNamespaces.Word + "after", "240");
    }

    [Fact]
    public void Should_write_keep_lines_properly()
    {
        // Arrange
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
            .HaveName(OpenXmlNamespaces.Word + "pPr")
            .And
            .HaveElement(OpenXmlNamespaces.Word + "spacing")
            .And
            .HaveElement(OpenXmlNamespaces.Word + "keepLines");
    }

    [Fact]
    public void Should_not_write_keep_lines()
    {
        // Arrange
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
            .HaveName(OpenXmlNamespaces.Word + "pPr")
            .And
            .Subject
            .Elements()
            .Should()
            .AllSatisfy(elementXml => 
                elementXml
                    .Should()
                    .NotHaveName(OpenXmlNamespaces.Word + "keepLines"));
    }
    
    [Fact]
    public void Should_write_keep_next_properly()
    {
        // Arrange
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
            .HaveName(OpenXmlNamespaces.Word + "pPr")
            .And
            .HaveElement(OpenXmlNamespaces.Word + "spacing")
            .And
            .HaveElement(OpenXmlNamespaces.Word + "keepNext");
    }

    [Fact]
    public void Should_not_write_keep_next()
    {
        // Arrange
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
            .HaveName(OpenXmlNamespaces.Word + "pPr")
            .And
            .Subject
            .Elements()
            .Should()
            .AllSatisfy(elementXml => 
                elementXml
                    .Should()
                    .NotHaveName(OpenXmlNamespaces.Word + "keepNext"));
    }
}