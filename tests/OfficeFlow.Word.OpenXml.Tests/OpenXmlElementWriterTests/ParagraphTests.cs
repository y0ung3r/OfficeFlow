using System.Xml.Linq;
using FluentAssertions;
using OfficeFlow.TestFramework.Extensions;
using OfficeFlow.Word.Core.Elements.Paragraphs;
using OfficeFlow.Word.Core.Elements.Paragraphs.Text;
using Xunit;

namespace OfficeFlow.Word.OpenXml.Tests.OpenXmlElementWriterTests;

public sealed class ParagraphTests
{
    [Fact]
    public void Should_write_paragraph_format_properly()
    {
        // Assert
        var sut = new OpenXmlElementWriter(
            new XElement(OpenXmlNamespaces.Word + "p"));

        var paragraph = new Paragraph();

        // Act
        sut.Visit(paragraph);
        
        // Assert
        sut.Xml
            .Should()
            .HaveName(OpenXmlNamespaces.Word + "p")
            .And
            .HaveElement(OpenXmlNamespaces.Word + "pPr");
    }
    
    [Fact]
    public void Should_write_runs_properly()
    {
        // Assert
        var sut = new OpenXmlElementWriter(
            new XElement(OpenXmlNamespaces.Word + "p"));

        var paragraph = new Paragraph();
        
        paragraph.AppendChild(
            new Run());
        
        paragraph.AppendChild(
            new Run());

        // Act
        sut.Visit(paragraph);
        
        // Assert
        sut.Xml
            .Should()
            .HaveName(OpenXmlNamespaces.Word + "p")
            .And
            .HaveElement(OpenXmlNamespaces.Word + "r", Exactly.Times(expected: 2));
    }
}