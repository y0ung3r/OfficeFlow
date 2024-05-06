using System.Xml.Linq;
using FluentAssertions;
using OfficeFlow.TestFramework.Extensions;
using OfficeFlow.Word.Core.Elements.Paragraphs.Text;
using Xunit;

namespace OfficeFlow.Word.OpenXml.Tests.OpenXmlElementWriterTests;

public sealed class RunTests
{
    [Fact]
    public void Should_write_run_format_properly()
    {
        // Assert
        var sut = new OpenXmlElementWriter(
            new XElement(OpenXmlNamespaces.Word + "r"));

        var run = new Run();

        // Act
        sut.Visit(run);
        
        // Assert
        sut.Xml
            .Should()
            .HaveName(OpenXmlNamespaces.Word + "r")
            .And
            .HaveElement(OpenXmlNamespaces.Word + "rPr");
    }
    
    [Fact]
    public void Should_write_line_breaks_properly()
    {
        // Assert
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
            .HaveName(OpenXmlNamespaces.Word + "r")
            .And
            .HaveElement(OpenXmlNamespaces.Word + "br", Exactly.Times(expected: 3));
    }
    
    [Fact]
    public void Should_write_vertical_tabs_properly()
    {
        // Assert
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
            .HaveName(OpenXmlNamespaces.Word + "r")
            .And
            .HaveElement(OpenXmlNamespaces.Word + "ptab", Exactly.Times(expected: 3));
    }
    
    [Fact]
    public void Should_write_horizontal_tabs_properly()
    {
        // Assert
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
            .HaveName(OpenXmlNamespaces.Word + "r")
            .And
            .HaveElement(OpenXmlNamespaces.Word + "tab", Exactly.Times(expected: 3));
    }
    
    [Fact]
    public void Should_write_text_holders_properly()
    {
        // Assert
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
            .HaveName(OpenXmlNamespaces.Word + "r")
            .And
            .HaveElement(OpenXmlNamespaces.Word + "t", Exactly.Times(expected: 3));
    }
}