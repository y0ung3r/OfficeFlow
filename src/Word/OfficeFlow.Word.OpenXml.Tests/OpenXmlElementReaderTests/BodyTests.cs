using System.Linq;
using System.Xml.Linq;
using FluentAssertions;
using OfficeFlow.Word.Core.Elements;
using OfficeFlow.Word.Core.Elements.Paragraphs;
using Xunit;

namespace OfficeFlow.Word.OpenXml.Tests.OpenXmlElementReaderTests;

public sealed class BodyTests
{
    [Fact]
    public void Should_read_paragraphs_from_respective_sections()
    {
        // Arrange
        var xml = new XElement(OpenXmlNamespaces.Word + "document",
            new XElement(OpenXmlNamespaces.Word + "body",
                new XElement(OpenXmlNamespaces.Word + "p"),
                new XElement(OpenXmlNamespaces.Word + "p"),
                new XElement(OpenXmlNamespaces.Word + "p",
                    new XElement(OpenXmlNamespaces.Word + "sectPr")),
                new XElement(OpenXmlNamespaces.Word + "p"),
                new XElement(OpenXmlNamespaces.Word + "sectPr")));

        var body = new Body();
        var sut = new OpenXmlElementReader(xml);

        // Act
        sut.Visit(body);

        // Assert
        body.Should()
            .ContainItemsAssignableTo<Section>()
            .And
            .HaveCount(expected: 2);

        body.OfType<Section>()
            .Should()
            .SatisfyRespectively(
                section =>
                    section
                        .Should()
                        .ContainItemsAssignableTo<Paragraph>()
                        .And
                        .HaveCount(3),
                section =>
                    section
                        .Should()
                        .ContainItemsAssignableTo<Paragraph>()
                        .And
                        .HaveCount(1));
    }

    [Fact]
    public void Should_read_main_section_properly()
    {
        // Arrange
        var xml = new XElement(OpenXmlNamespaces.Word + "document",
            new XElement(OpenXmlNamespaces.Word + "body",
                new XElement(OpenXmlNamespaces.Word + "sectPr")));

        var body = new Body();
        var sut = new OpenXmlElementReader(xml);

        // Act
        sut.Visit(body);
        
        // Assert
        body.Should()
            .ContainItemsAssignableTo<Section>()
            .And
            .HaveCount(expected: 1);
    }

    [Fact]
    public void Should_be_only_one_main_section()
    {
        // Arrange
        var xml = new XElement(OpenXmlNamespaces.Word + "document",
            new XElement(OpenXmlNamespaces.Word + "body",
                new XElement(OpenXmlNamespaces.Word + "sectPr"),
                new XElement(OpenXmlNamespaces.Word + "sectPr")));

        var body = new Body();
        var sut = new OpenXmlElementReader(xml);

        // Act
        sut.Visit(body);
        
        // Assert
        body.Should()
            .ContainItemsAssignableTo<Section>()
            .And
            .HaveCount(expected: 1);
    }
    
    [Fact]
    public void Should_have_main_section()
    {
        // Arrange
        var xml = new XElement(OpenXmlNamespaces.Word + "document",
            new XElement(OpenXmlNamespaces.Word + "body"));

        var body = new Body();
        var sut = new OpenXmlElementReader(xml);

        // Act
        sut.Visit(body);
        
        // Assert
        body.Should()
            .ContainItemsAssignableTo<Section>()
            .And
            .HaveCount(expected: 1);
    }

    [Fact]
    public void Should_read_empty_body_properly()
    {
        // Arrange
        var xml = new XElement(OpenXmlNamespaces.Word + "document");

        var body = new Body();
        var sut = new OpenXmlElementReader(xml);

        // Act
        sut.Visit(body);
        
        // Assert
        body.Should()
            .ContainItemsAssignableTo<Section>()
            .And
            .HaveCount(expected: 1);
    }
}