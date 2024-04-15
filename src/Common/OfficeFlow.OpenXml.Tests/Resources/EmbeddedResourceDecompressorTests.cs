using System.Xml.Linq;
using FluentAssertions;
using OfficeFlow.OpenXml.Resources;
using OfficeFlow.OpenXml.Resources.Exceptions;
using Xunit;

namespace OfficeFlow.OpenXml.Tests.Resources;

public sealed class EmbeddedResourceDecompressorTests
{
    [Fact]
    public void Should_decompress_exists_embedded_resource_from_calling_assembly()
    {
        // Arrange
        var expectedXml = new XElement
        (
            "document",
            new XElement
            (
                "body",
                new XText("Content")
            )
        );

        var sut = new EmbeddedResourceDecompressor();

        // Act
        var actualXml = sut.Decompress("TestEmbeddedResource.xml.gz");

        // Assert
        actualXml
            .Root
            .Should()
            .BeEquivalentTo(expectedXml);
    }

    [Fact]
    public void Should_throws_exception_if_resource_not_found_in_calling_assembly()
        => new EmbeddedResourceDecompressor()
            .Invoking(decompressor => decompressor.Decompress("NotExistsResource.xml.gz"))
            .Should()
            .Throw<ResourceNotFoundException>();
}