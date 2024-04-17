using System.Linq;
using System.Xml.Linq;
using FluentAssertions;
using OfficeFlow.Word.Core.Elements.Paragraphs.Text;
using Xunit;

namespace OfficeFlow.Word.OpenXml.Tests.OpenXmlElementReaderTests;

public sealed class RunTests
{
    [Fact]
    public void Should_read_run_properly()
    {
        // Arrange
        var xml = new XElement(OpenXmlNamespaces.Word + "r",
            new XElement(OpenXmlNamespaces.Word + "rPr"),
            new XElement(OpenXmlNamespaces.Word + "t"),
            new XElement(OpenXmlNamespaces.Word + "br"),
            new XElement(OpenXmlNamespaces.Word + "ptab"),
            new XElement(OpenXmlNamespaces.Word + "tab"));

        var run = new Run();
        var sut = new OpenXmlElementReader(xml);
        
        // Act
        sut.Visit(run);
        
        // Assert
        run.Should()
            .HaveCount(expected: 4);

        run.Select(child => child?.GetType())
            .Should()
            .ContainInOrder(
                typeof(TextHolder),
                typeof(LineBreak),
                typeof(VerticalTabulation),
                typeof(HorizontalTabulation));
    }
}