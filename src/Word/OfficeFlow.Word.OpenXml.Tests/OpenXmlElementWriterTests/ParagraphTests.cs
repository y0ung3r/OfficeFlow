using System.Xml.Linq;
using FluentAssertions;
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
        var expectedXml = new XElement(OpenXmlNamespaces.Word + "p",
            new XElement(OpenXmlNamespaces.Word + "pPr",
                new XElement(OpenXmlNamespaces.Word + "spacing",
                    new XAttribute(OpenXmlNamespaces.Word + "before", "0"),
                    new XAttribute(OpenXmlNamespaces.Word + "after", "160"))));
        
        var sut = new OpenXmlElementWriter(
            new XElement(OpenXmlNamespaces.Word + "p"));

        var paragraph = new Paragraph();

        // Act
        sut.Visit(paragraph);
        
        // Assert
        sut.Xml
            .Should()
            .Be(expectedXml);
    }
    
    [Fact]
    public void Should_write_runs_properly()
    {
        // Assert
        var expectedXml = new XElement(OpenXmlNamespaces.Word + "p",
            new XElement(OpenXmlNamespaces.Word + "pPr",
                new XElement(OpenXmlNamespaces.Word + "spacing",
                    new XAttribute(OpenXmlNamespaces.Word + "before", "0"),
                    new XAttribute(OpenXmlNamespaces.Word + "after", "160"))),
            new XElement(OpenXmlNamespaces.Word + "r",
                new XElement(OpenXmlNamespaces.Word + "rPr")),
            new XElement(OpenXmlNamespaces.Word + "r",
                new XElement(OpenXmlNamespaces.Word + "rPr")));
        
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
            .Be(expectedXml);
    }
}