using System;
using System.IO;
using System.Xml.Linq;
using FluentAssertions;
using Xunit;

namespace OfficeFlow.OpenXml.Tests.Packaging
{
    public sealed class OpenXmlPackagePartTests
    {
        [Fact]
        public void Should_throw_exception_if_root_is_empty()
            => TestOpenXmlPackagePartFactory
                .Create(uri: "/part", content: new XDocument())
                .Invoking(packagePart => packagePart.Root)
                .Should()
                .Throw<InvalidOperationException>();

        [Fact]
        public void Should_add_child_properly()
        {
            // Arrange
            var sut = TestOpenXmlPackagePartFactory.Create(uri: "/child");
            var parent = TestOpenXmlPackagePartFactory.Create(uri: "/parent");
            
            // Act
            parent.AddChild(sut);
            
            // Assert
            sut.Parent
                .Should()
                .Be(parent);
        }

        [Fact]
        public void Should_flush_to_stream_properly()
        {
            // Arrange
            using var expectedStream = 
                new MemoryStream();
            
            var content = new XDocument(
                new XElement
                (
                    "document",
                    new XElement
                    (
                        "body",
                        new XText("Content")
                    )
                ));
            
            var sut = TestOpenXmlPackagePartFactory.Create(uri: "/part", content);
            
            // Act
            sut.FlushTo(expectedStream);
            
            // Assert
            expectedStream
                .Length
                .Should()
                .BePositive();
        }
    }
}