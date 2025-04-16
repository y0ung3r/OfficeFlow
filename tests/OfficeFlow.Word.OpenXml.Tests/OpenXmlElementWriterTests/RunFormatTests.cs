using System.Xml.Linq;
using FluentAssertions;
using OfficeFlow.TestFramework.Extensions;
using OfficeFlow.Word.Core.Elements.Paragraphs.Text.Formatting;
using OfficeFlow.Word.Core.Elements.Paragraphs.Text.Formatting.Enums;
using Xunit;

namespace OfficeFlow.Word.OpenXml.Tests.OpenXmlElementWriterTests;

public sealed class RunFormatTests
{
    [Fact]
    public void Should_write_italic_properly()
    {
        // Assert
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
            .HaveName(OpenXmlNamespaces.Word + "rPr")
            .And
            .HaveElement(OpenXmlNamespaces.Word + "i");
    }

    [Fact]
    public void Default_value_of_italic_should_not_be_written()
    {
        // Assert
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
            .HaveName(OpenXmlNamespaces.Word + "rPr")
            .And
            .Subject
            .Elements()
            .Should()
            .BeEmpty();
    }
    
    [Fact]
    public void Should_write_bold_properly()
    {
        // Assert
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
            .HaveName(OpenXmlNamespaces.Word + "rPr")
            .And
            .HaveElement(OpenXmlNamespaces.Word + "b");
    }

    [Fact]
    public void Default_value_of_bold_should_not_be_written()
    {
        // Assert
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
            .HaveName(OpenXmlNamespaces.Word + "rPr")
            .And
            .Subject
            .Elements()
            .Should()
            .BeEmpty();
    }
    
    [Fact]
    public void Should_write_hidden_properly()
    {
        // Assert
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
            .HaveName(OpenXmlNamespaces.Word + "rPr")
            .And
            .HaveElement(OpenXmlNamespaces.Word + "vanish");
    }

    [Fact]
    public void Default_value_of_hidden_should_not_be_written()
    {
        // Assert
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
            .HaveName(OpenXmlNamespaces.Word + "rPr")
            .And
            .Subject
            .Elements()
            .Should()
            .BeEmpty();
    }

    [Theory]
    [InlineData("strike", StrikethroughType.Single)]
    [InlineData("dstrike", StrikethroughType.Double)]
    public void Should_write_strikethrough_type_properly(string expectedStrikethroughType, StrikethroughType actualStrikethroughType)
    {
        // Assert
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
            .HaveName(OpenXmlNamespaces.Word + "rPr")
            .And
            .HaveElement(OpenXmlNamespaces.Word + expectedStrikethroughType);
    }
    
    [Fact]
    public void Default_value_of_strikethrough_type_should_not_be_written()
    {
        // Assert
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
            .HaveName(OpenXmlNamespaces.Word + "rPr")
            .And
            .Subject
            .Elements()
            .Should()
            .BeEmpty();
    }
}