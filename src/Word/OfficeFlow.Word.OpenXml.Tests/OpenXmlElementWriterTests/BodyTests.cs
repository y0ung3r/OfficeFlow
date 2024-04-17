using System;
using System.Xml.Linq;
using FluentAssertions;
using OfficeFlow.Word.Core.Elements;
using OfficeFlow.Word.Core.Elements.Paragraphs;
using Xunit;

namespace OfficeFlow.Word.OpenXml.Tests.OpenXmlElementWriterTests;

public sealed class BodyTests
{
    [Fact]
    public void Should_throw_exception_if_provided_node_is_not_a_document_node()
    {
        // Arrange
        var sut = new OpenXmlElementWriter(
            new XElement(OpenXmlNamespaces.Word + "p"));
        
        // Act & Assert
        sut.Invoking(writer => writer.Visit(new Body()))
            .Should()
            .Throw<InvalidOperationException>();
    }
    
    [Fact]
    public void Should_write_body_element_properly()
    {
        // Arrange
        var expectedXml = new XElement(OpenXmlNamespaces.Word + "document",
            new XElement(OpenXmlNamespaces.Word + "body",
                new XElement(OpenXmlNamespaces.Word + "p",
                    new XElement(OpenXmlNamespaces.Word + "pPr")),
                new XElement(OpenXmlNamespaces.Word + "p",
                    new XElement(OpenXmlNamespaces.Word + "pPr")),
                new XElement(OpenXmlNamespaces.Word + "p",
                    new XElement(OpenXmlNamespaces.Word + "sectPr"),
                    new XElement(OpenXmlNamespaces.Word + "pPr")),
                new XElement(OpenXmlNamespaces.Word + "p",
                    new XElement(OpenXmlNamespaces.Word + "pPr")),
                new XElement(OpenXmlNamespaces.Word + "sectPr")));

        var sut = new OpenXmlElementWriter(
            new XElement(OpenXmlNamespaces.Word + "document"));

        var body = new Body();
        var firstSection = new Section();
        
        firstSection.AppendChild(
            new Paragraph());
        
        firstSection.AppendChild(
            new Paragraph());
        
        firstSection.AppendChild(
            new Paragraph());
        
        body.AppendChild(firstSection);

        var lastSection = new Section();
        
        lastSection.AppendChild(
            new Paragraph());
        
        body.AppendChild(lastSection);

        // Act
        sut.Visit(body);

        // Assert
        sut.Xml
            .Should()
            .Be(expectedXml);
    }
}