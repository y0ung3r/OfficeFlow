using System.Xml.Linq;
using FluentAssertions;
using OfficeFlow.Word.Core.Elements.Paragraphs.Text;
using OfficeFlow.Word.Core.Elements.Paragraphs.Text.Enums;
using Xunit;

namespace OfficeFlow.Word.OpenXml.Tests.OpenXmlElementWriterTests;

public sealed class RunFormatTests
{
    [Fact]
    public void Should_write_italic_properly()
    {
        // Assert
        var expectedXml = new XElement(OpenXmlNamespaces.Word + "rPr",
            new XElement(OpenXmlNamespaces.Word + "i"));
        
        var sut = new OpenXmlElementWriter(
            new XElement(OpenXmlNamespaces.Word + "rPr"));

        var runFormat = new RunFormat
        {
            IsItalic = true
        };
        
        // Act
        sut.Visit(runFormat);
        
        // Assert
        sut.Xml
            .Should()
            .Be(expectedXml);
    }

    [Fact]
    public void Default_value_of_italic_should_not_be_written()
    {
        // Assert
        var expectedXml = new XElement(OpenXmlNamespaces.Word + "rPr");
        
        var sut = new OpenXmlElementWriter(
            new XElement(OpenXmlNamespaces.Word + "rPr"));

        var runFormat = new RunFormat
        {
            IsItalic = false
        };
        
        // Act
        sut.Visit(runFormat);
        
        // Assert
        sut.Xml
            .Should()
            .Be(expectedXml);
    }
    
    [Fact]
    public void Should_write_bold_properly()
    {
        // Assert
        var expectedXml = new XElement(OpenXmlNamespaces.Word + "rPr",
            new XElement(OpenXmlNamespaces.Word + "b"));
        
        var sut = new OpenXmlElementWriter(
            new XElement(OpenXmlNamespaces.Word + "rPr"));

        var runFormat = new RunFormat
        {
            IsBold = true
        };
        
        // Act
        sut.Visit(runFormat);
        
        // Assert
        sut.Xml
            .Should()
            .Be(expectedXml);
    }

    [Fact]
    public void Default_value_of_bold_should_not_be_written()
    {
        // Assert
        var expectedXml = new XElement(OpenXmlNamespaces.Word + "rPr");
        
        var sut = new OpenXmlElementWriter(
            new XElement(OpenXmlNamespaces.Word + "rPr"));

        var runFormat = new RunFormat
        {
            IsBold = false
        };
        
        // Act
        sut.Visit(runFormat);
        
        // Assert
        sut.Xml
            .Should()
            .Be(expectedXml);
    }
    
    [Fact]
    public void Should_write_hidden_properly()
    {
        // Assert
        var expectedXml = new XElement(OpenXmlNamespaces.Word + "rPr",
            new XElement(OpenXmlNamespaces.Word + "vanish"));
        
        var sut = new OpenXmlElementWriter(
            new XElement(OpenXmlNamespaces.Word + "rPr"));

        var runFormat = new RunFormat
        {
            IsHidden = true
        };
        
        // Act
        sut.Visit(runFormat);
        
        // Assert
        sut.Xml
            .Should()
            .Be(expectedXml);
    }

    [Fact]
    public void Default_value_of_hidden_should_not_be_written()
    {
        // Assert
        var expectedXml = new XElement(OpenXmlNamespaces.Word + "rPr");
        
        var sut = new OpenXmlElementWriter(
            new XElement(OpenXmlNamespaces.Word + "rPr"));

        var runFormat = new RunFormat
        {
            IsHidden = false
        };
        
        // Act
        sut.Visit(runFormat);
        
        // Assert
        sut.Xml
            .Should()
            .Be(expectedXml);
    }

    [Theory]
    [InlineData("strike", StrikethroughType.Single)]
    [InlineData("dstrike", StrikethroughType.Double)]
    public void Should_write_strikethrough_type_properly(string expectedStrikethroughType, StrikethroughType actualStrikethroughType)
    {
        // Assert
        var expectedXml = new XElement(OpenXmlNamespaces.Word + "rPr",
            new XElement(OpenXmlNamespaces.Word + expectedStrikethroughType));
        
        var sut = new OpenXmlElementWriter(
            new XElement(OpenXmlNamespaces.Word + "rPr"));

        var runFormat = new RunFormat
        {
            StrikethroughType = actualStrikethroughType
        };
        
        // Act
        sut.Visit(runFormat);
        
        // Assert
        sut.Xml
            .Should()
            .Be(expectedXml);
    }
    
    [Fact]
    public void Default_value_of_strikethrough_type_should_not_be_written()
    {
        // Assert
        var expectedXml = new XElement(OpenXmlNamespaces.Word + "rPr");
        
        var sut = new OpenXmlElementWriter(
            new XElement(OpenXmlNamespaces.Word + "rPr"));

        var runFormat = new RunFormat
        {
            StrikethroughType = StrikethroughType.None
        };
        
        // Act
        sut.Visit(runFormat);
        
        // Assert
        sut.Xml
            .Should()
            .Be(expectedXml);
    }
}