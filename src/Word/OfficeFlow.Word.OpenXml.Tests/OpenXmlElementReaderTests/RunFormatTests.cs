using System.Xml.Linq;
using FluentAssertions;
using OfficeFlow.Word.Core.Elements.Paragraphs.Text;
using OfficeFlow.Word.Core.Elements.Paragraphs.Text.Enums;
using Xunit;

namespace OfficeFlow.Word.OpenXml.Tests.OpenXmlElementReaderTests;

public sealed class RunFormatTests
{
    [Fact]
    public void Should_read_italic_properly()
    {
        // Arrange
        var xml = new XElement(OpenXmlNamespaces.Word + "rPr",
            new XElement(OpenXmlNamespaces.Word + "i"));
        
        var runFormat = new RunFormat();
        var sut = new OpenXmlElementReader(xml);
        
        // Act
        sut.Visit(runFormat);

        // Assert
        runFormat
            .IsItalic
            .Should()
            .BeTrue();
    }
    
    [Fact]
    public void Should_read_bold_properly()
    {
        // Arrange
        var xml = new XElement(OpenXmlNamespaces.Word + "rPr",
            new XElement(OpenXmlNamespaces.Word + "b"));
        
        var runFormat = new RunFormat();
        var sut = new OpenXmlElementReader(xml);
        
        // Act
        sut.Visit(runFormat);

        // Assert
        runFormat
            .IsBold
            .Should()
            .BeTrue();
    }
    
    [Fact]
    public void Should_read_hidden_properly()
    {
        // Arrange
        var xml = new XElement(OpenXmlNamespaces.Word + "rPr",
            new XElement(OpenXmlNamespaces.Word + "vanish"));
        
        var runFormat = new RunFormat();
        var sut = new OpenXmlElementReader(xml);
        
        // Act
        sut.Visit(runFormat);

        // Assert
        runFormat
            .IsHidden
            .Should()
            .BeTrue();
    }
    
    [Theory]
    [InlineData("strike", StrikethroughType.Single)]
    [InlineData("dstrike", StrikethroughType.Double)]
    public void Should_read_strikethrough_type_properly(string actualStrikethroughType, StrikethroughType expectedStrikethroughType)
    {
        // Arrange
        var xml = new XElement(OpenXmlNamespaces.Word + "rPr",
            new XElement(OpenXmlNamespaces.Word + actualStrikethroughType));
        
        var runFormat = new RunFormat();
        var sut = new OpenXmlElementReader(xml);
        
        // Act
        sut.Visit(runFormat);

        // Assert
        runFormat
            .StrikethroughType
            .Should()
            .Be(expectedStrikethroughType);
    }
    
    [Fact]
    public void Should_read_empty_run_format_properly()
    {
        // Arrange
        var xml = new XElement(OpenXmlNamespaces.Word + "rPr");
        
        var runFormat = new RunFormat();
        var sut = new OpenXmlElementReader(xml);
        
        // Act
        sut.Visit(runFormat);
        
        // Assert
        runFormat
            .IsItalic
            .Should()
            .Be(RunFormat.Default.IsItalic);
        
        runFormat
            .IsBold
            .Should()
            .Be(RunFormat.Default.IsBold);

        runFormat
            .IsHidden
            .Should()
            .BeFalse();
        
        runFormat
            .StrikethroughType
            .Should()
            .Be(RunFormat.Default.StrikethroughType);
    }
}