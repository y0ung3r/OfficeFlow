using System.Xml.Linq;
using FluentAssertions;
using OfficeFlow.Word.Core.Elements.Paragraphs.Text;
using Xunit;

namespace OfficeFlow.Word.OpenXml.Tests.OpenXmlElementWriterTests;

public sealed class RunTests
{
    [Fact]
    public void Should_write_run_format_properly()
    {
        // Assert
        var expectedXml = new XElement(OpenXmlNamespaces.Word + "r",
            new XElement(OpenXmlNamespaces.Word + "rPr"));
        
        var sut = new OpenXmlElementWriter(
            new XElement(OpenXmlNamespaces.Word + "r"));

        var run = new Run();

        // Act
        sut.Visit(run);
        
        // Assert
        sut.Xml
            .Should()
            .Be(expectedXml);
    }
    
    [Fact]
    public void Should_write_line_breaks_properly()
    {
        // Assert
        var expectedXml = new XElement(OpenXmlNamespaces.Word + "r",
            new XElement(OpenXmlNamespaces.Word + "rPr"),
            new XElement(OpenXmlNamespaces.Word + "br"),
            new XElement(OpenXmlNamespaces.Word + "br"),
            new XElement(OpenXmlNamespaces.Word + "br"));
        
        var sut = new OpenXmlElementWriter(
            new XElement(OpenXmlNamespaces.Word + "r"));

        var run = new Run();
        
        run.AppendChild(
            new LineBreak());
        
        run.AppendChild(
            new LineBreak());
        
        run.AppendChild(
            new LineBreak());

        // Act
        sut.Visit(run);
        
        // Assert
        sut.Xml
            .Should()
            .Be(expectedXml);
    }
    
    [Fact]
    public void Should_write_vertical_tabs_properly()
    {
        // Assert
        var expectedXml = new XElement(OpenXmlNamespaces.Word + "r",
            new XElement(OpenXmlNamespaces.Word + "rPr"),
            new XElement(OpenXmlNamespaces.Word + "ptab"),
            new XElement(OpenXmlNamespaces.Word + "ptab"),
            new XElement(OpenXmlNamespaces.Word + "ptab"));
        
        var sut = new OpenXmlElementWriter(
            new XElement(OpenXmlNamespaces.Word + "r"));

        var run = new Run();
        
        run.AppendChild(
            new VerticalTabulation());
        
        run.AppendChild(
            new VerticalTabulation());
        
        run.AppendChild(
            new VerticalTabulation());

        // Act
        sut.Visit(run);
        
        // Assert
        sut.Xml
            .Should()
            .Be(expectedXml);
    }
    
    [Fact]
    public void Should_write_horizontal_tabs_properly()
    {
        // Assert
        var expectedXml = new XElement(OpenXmlNamespaces.Word + "r",
            new XElement(OpenXmlNamespaces.Word + "rPr"),
            new XElement(OpenXmlNamespaces.Word + "tab"),
            new XElement(OpenXmlNamespaces.Word + "tab"),
            new XElement(OpenXmlNamespaces.Word + "tab"));
        
        var sut = new OpenXmlElementWriter(
            new XElement(OpenXmlNamespaces.Word + "r"));

        var run = new Run();
        
        run.AppendChild(
            new HorizontalTabulation());
        
        run.AppendChild(
            new HorizontalTabulation());
        
        run.AppendChild(
            new HorizontalTabulation());

        // Act
        sut.Visit(run);
        
        // Assert
        sut.Xml
            .Should()
            .Be(expectedXml);
    }
    
    [Fact]
    public void Should_write_text_holders_properly()
    {
        // Assert
        var expectedXml = new XElement(OpenXmlNamespaces.Word + "r",
            new XElement(OpenXmlNamespaces.Word + "rPr"),
            new XElement(OpenXmlNamespaces.Word + "t"),
            new XElement(OpenXmlNamespaces.Word + "t"),
            new XElement(OpenXmlNamespaces.Word + "t"));
        
        var sut = new OpenXmlElementWriter(
            new XElement(OpenXmlNamespaces.Word + "r"));

        var run = new Run();
        
        run.AppendChild(
            new TextHolder());
        
        run.AppendChild(
            new TextHolder());
        
        run.AppendChild(
            new TextHolder());

        // Act
        sut.Visit(run);
        
        // Assert
        sut.Xml
            .Should()
            .Be(expectedXml);
    }
}