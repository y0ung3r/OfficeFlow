using System.Xml.Linq;
using FluentAssertions;
using OfficeFlow.Word.Core.Elements.Paragraphs.Text;
using Xunit;

namespace OfficeFlow.Word.OpenXml.Tests.OpenXmlElementReaderTests;

public sealed class TextHolderTests
{
    [Fact]
    public void Should_read_text_properly()
    {
        // Arrange
        var xml = new XElement(OpenXmlNamespaces.Word + "t")
        {
            Value = "Text"
        };

        var textHolder = new TextHolder();
        var sut = new OpenXmlElementReader(xml);
        
        // Act
        sut.Visit(textHolder);
        
        // Assert
        textHolder
            .Value
            .Should()
            .Be("Text");
    }
}