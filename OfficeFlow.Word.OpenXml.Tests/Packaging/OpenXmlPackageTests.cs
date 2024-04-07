using System;
using System.IO;
using System.IO.Packaging;
using System.Xml.Linq;
using FluentAssertions;
using OfficeFlow.Word.OpenXml.Packaging;
using Xunit;

namespace OfficeFlow.Word.OpenXml.Tests.Packaging
{
    public sealed class OpenXmlPackageTests
    {
        [Fact]
        public void Should_throw_exception_if_package_is_disposed()
        {
            // Arrange
            var sut = OpenXmlPackage.Create();
            sut.Dispose();
            
            var testCases = new Action[]
            {
                () => sut.EnumerateParts(),
                () => sut.Save(),
                () => sut.SaveTo(remoteStream: new MemoryStream()),
                () => sut.SaveTo(filePath: nameof(OpenXmlPackage))
            };
            
            // Act & Assert
            testCases
                .Should()
                .AllSatisfy(testCase => 
                    testCase
                        .Should()
                        .Throw<ObjectDisposedException>());
        }

        [Fact]
        public void Should_create_new_package_properly()
            => new Action(() => OpenXmlPackage.Create())
                .Should()
                .NotThrow();

        [Fact]
        public void Should_open_package_using_stream_properly()
            => new Action(() =>
            {
                using var originalStream = 
                    PrepareTestPackageStream();
                
                OpenXmlPackage.Open(originalStream);
            })
            .Should()
            .NotThrow();
        
        [Fact]
        public void Should_open_package_using_file_path_properly()
        {
            // Arrange
            using var originalStream = 
                PrepareTestPackageStream();

            var filePath = Path.GetTempFileName();
            using (var fileStream = new FileStream(filePath, FileMode.Open, FileAccess.ReadWrite))
                originalStream.CopyTo(fileStream);

            // Act & Assert
            new Action(() => OpenXmlPackage.Open(filePath))
                .Should()
                .NotThrow();
            
            File.Delete(filePath);
        }

        [Fact]
        public void Package_parts_should_be_empty_for_new_package()
        {
            // Arrange
            using var sut = 
                OpenXmlPackage.Create();
            
            // Act & Assert
            sut.EnumerateParts()
                .Should()
                .BeEmpty();
        }
        
        [Fact]
        public void Should_enumerate_package_parts_properly()
        {
            // Arrange
            using var originalStream =
                PrepareTestPackageStream();

            using var sut = 
                OpenXmlPackage.Open(originalStream);
            
            // Act & Assert
            sut.EnumerateParts()
                .Should()
                .NotBeEmpty();
        }
        
        [Fact]
        public void Should_save_package_using_stream_properly()
        {
            // Arrange
            using var originalStream =
                PrepareTestPackageStream();

            using var sut = 
                OpenXmlPackage.Open(originalStream);
            
            // Act
            sut.Save();
            
            // Assert
            originalStream
                .Position
                .Should()
                .Be(originalStream.Length);
        }

        [Fact]
        public void Should_save_package_using_file_path_properly()
        {
            // Arrange
            using var originalStream = 
                PrepareTestPackageStream();

            var filePath = Path.GetTempFileName();
            using (var fileStream = new FileStream(filePath, FileMode.Open, FileAccess.ReadWrite))
                originalStream.CopyTo(fileStream);

            using var sut = 
                OpenXmlPackage.Open(filePath);
            
            // Act
            sut.Save();
            
            // Assert
            originalStream
                .Position
                .Should()
                .Be(originalStream.Length);
            
            File.Delete(filePath);
        }

        [Fact]
        public void Should_save_package_to_stream_properly()
        {
            // Arrange
            using var originalStream = 
                PrepareTestPackageStream();

            using var destinationStream = 
                new MemoryStream();
            
            using var sut = 
                OpenXmlPackage.Open(originalStream);
            
            // Act
            sut.SaveTo(destinationStream);
            
            // Assert
            destinationStream
                .Length
                .Should()
                .BePositive();
        }
        
        [Fact]
        public void Should_save_package_to_file_properly()
        {
            // Arrange
            using var originalStream = 
                PrepareTestPackageStream();

            var filePath = Path.GetTempFileName();
            
            using var sut = 
                OpenXmlPackage.Open(originalStream);
            
            // Act
            sut.SaveTo(filePath);
            
            // Assert
            File.ReadAllBytes(filePath)
                .Length
                .Should()
                .BePositive();
            
            File.Delete(filePath);
        }

        private static MemoryStream PrepareTestPackageStream()
        {
            var partXml = new XElement
            (
                "document",
                new XElement
                (
                    "body",
                    new XText("Content")
                )
            );

            return PrepareTestPackageStream(partXml);
        }

        private static MemoryStream PrepareTestPackageStream(XElement partXml)
        {
            var originalStream = new MemoryStream();

            using (var package = Package.Open(originalStream, FileMode.Create, FileAccess.ReadWrite))
            {
                var partUri = PackUriHelper.CreatePartUri(
                    new Uri(nameof(OpenXmlPackage), UriKind.Relative));

                var part = package.CreatePart(partUri, string.Empty);

                using var contentStream =
                    part.GetStream();

                using (var contentWriter = new StreamWriter(contentStream))
                    partXml.Save(contentWriter);
            }

            originalStream.Seek(offset: 0, SeekOrigin.Begin);
            
            return originalStream;
        }
    }
}