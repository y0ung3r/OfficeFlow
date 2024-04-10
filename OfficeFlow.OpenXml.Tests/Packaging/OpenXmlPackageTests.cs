using System;
using System.IO;
using System.IO.Packaging;
using FluentAssertions;
using OfficeFlow.OpenXml.Packaging;
using Xunit;

namespace OfficeFlow.OpenXml.Tests.Packaging
{
    public sealed class OpenXmlPackageTests : IClassFixture<TempFilePool>
    {
        private readonly TempFilePool _tempFilePool;

        public OpenXmlPackageTests(TempFilePool tempFilePool)
            => _tempFilePool = tempFilePool;

        [Fact]
        public void Should_throw_exception_if_package_is_disposed()
        {
            // Arrange
            var sut = OpenXmlPackage.Create();
            
            // Act
            sut.Dispose();
            
            // Assert
            VerifyDisposed(sut);
        }
        
        [Fact]
        public void Should_throw_exception_if_package_is_closed()
        {
            // Arrange
            var sut = OpenXmlPackage.Create();
            
            // Act
            sut.Close();
            
            // Assert
            VerifyDisposed(sut);
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

            var filePath = _tempFilePool.GetTempFilePath();
            
            using (var fileStream = new FileStream(filePath, FileMode.Open, FileAccess.ReadWrite))
                originalStream.CopyTo(fileStream);

            // Act & Assert
            new Action(() => OpenXmlPackage.Open(filePath))
                .Should()
                .NotThrow();
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
        public void Should_enumerate_pending_parts_properly()
        {
            // Arrange
            using var originalStream =
                PrepareTestPackageStream();

            using var sut = 
                OpenXmlPackage.Open(originalStream);
            
            var expectedPart = TestOpenXmlPackagePartFactory.Create(uri: "/part");
            
            sut.AddPart(expectedPart);
            
            // Act & Assert
            sut.EnumerateParts()
                .Should()
                .Contain(actualPart => 
                    expectedPart.Uri == actualPart.Uri);
        }

        [Fact]
        public void Should_enumerate_exists_parts_properly()
        {
            // Arrange
            using var originalStream =
                PrepareTestPackageStream();
            
            var expectedPart = TestOpenXmlPackagePartFactory.Create(uri: "/part");
            
            using (var sut = OpenXmlPackage.Open(originalStream))
                sut.AddPart(expectedPart);
            
            // Act & Assert
            OpenXmlPackage
                .Open(originalStream)
                .EnumerateParts()
                .Should()
                .Contain(actualPart => 
                    expectedPart.Uri == actualPart.Uri);
        }

        [Fact]
        public void Should_throw_exception_if_parent_package_part_is_not_added_to_package()
        {
            // Arrange
            var parentPart = TestOpenXmlPackagePartFactory.Create(uri: "/parent");
            var childPart = TestOpenXmlPackagePartFactory.Create(uri: "/child");
            parentPart.AddChild(childPart);
            
            // Act
            new Action(() =>
            {
                using var originalStream =
                    PrepareTestPackageStream();
                
                using var sut = 
                    OpenXmlPackage.Open(originalStream);
                
                sut.AddPart(childPart);
            })
            .Should()
            .Throw<InvalidOperationException>();
        }

        [Fact]
        public void Should_add_package_part_and_create_relationship_with_package()
        {
            // Arrange
            using var originalStream =
                PrepareTestPackageStream();

            var expectedPart = TestOpenXmlPackagePartFactory.Create(uri: "/part");
            
            // Act
            using (var sut = OpenXmlPackage.Open(originalStream))
                sut.AddPart(expectedPart);
            
            // Assert
            using var packageSource = 
                Package.Open(originalStream);

            var partSource = packageSource.GetPart(expectedPart.Uri);

            partSource
                .ContentType
                .Should()
                .Be(expectedPart.ContentType);
            
            partSource
                .CompressionOption
                .Should()
                .Be(expectedPart.CompressionMode);

            packageSource
                .GetRelationships()
                .Should()
                .OnlyContain(relationship => 
                    relationship.RelationshipType == expectedPart.RelationshipType);
        }
        
        [Fact]
        public void Should_add_package_part_and_create_relationship_with_another_part()
        {
            // Arrange
            using var originalStream =
                PrepareTestPackageStream();

            var parentPart = TestOpenXmlPackagePartFactory.Create(uri: "/parent");
            var childPart = TestOpenXmlPackagePartFactory.Create(uri: "/child");
            parentPart.AddChild(childPart);
            
            // Act
            using (var sut = OpenXmlPackage.Open(originalStream))
            {
                sut.AddPart(parentPart);
                sut.AddPart(childPart);
            }

            // Assert
            using var packageSource = 
                Package.Open(originalStream);

            var parentSource = packageSource.GetPart(parentPart.Uri);

            parentSource
                .GetRelationships()
                .Should()
                .OnlyContain(relationship => 
                    relationship.RelationshipType == childPart.RelationshipType);
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
                .Length
                .Should()
                .BePositive();
        }

        [Fact]
        public void Should_save_package_using_file_path_properly()
        {
            // Arrange
            using var originalStream = 
                PrepareTestPackageStream();

            var filePath = _tempFilePool.GetTempFilePath();
            using (var fileStream = new FileStream(filePath, FileMode.Open, FileAccess.ReadWrite))
                originalStream.CopyTo(fileStream);

            using var sut = 
                OpenXmlPackage.Open(filePath);
            
            // Act
            sut.Save();
            
            // Assert
            File.ReadAllBytes(filePath)
                .Length
                .Should()
                .BePositive();
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

            var filePath = _tempFilePool.GetTempFilePath();
            
            using var sut = 
                OpenXmlPackage.Open(originalStream);
            
            // Act
            sut.SaveTo(filePath);
            
            // Assert
            File.ReadAllBytes(filePath)
                .Length
                .Should()
                .BePositive();
        }

        private static void VerifyDisposed(OpenXmlPackage package)
        {
            var testCases = new Action[]
            {
                () => package.EnumerateParts(),
                () => package.Save(),
                () => package.SaveTo(remoteStream: new MemoryStream()),
                () => package.SaveTo(filePath: nameof(OpenXmlPackage))
            };
            
            // Act & Assert
            testCases
                .Should()
                .AllSatisfy(testCase => 
                    testCase
                        .Should()
                        .Throw<ObjectDisposedException>());
        }

        private static MemoryStream PrepareTestPackageStream()
        {
            var originalStream = new MemoryStream();

            Package
                .Open(originalStream, FileMode.Create, FileAccess.ReadWrite)
                .Close();

            originalStream.Seek(offset: 0, SeekOrigin.Begin);
            
            return originalStream;
        }
    }
}