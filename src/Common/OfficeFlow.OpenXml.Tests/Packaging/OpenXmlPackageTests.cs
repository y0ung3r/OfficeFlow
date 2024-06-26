﻿using System;
using System.IO;
using System.IO.Packaging;
using FluentAssertions;
using OfficeFlow.OpenXml.Packaging;
using OfficeFlow.TestFramework;
using Xunit;

namespace OfficeFlow.OpenXml.Tests.Packaging;

public sealed class OpenXmlPackageTests(TempFilePool tempFilePool) : IClassFixture<TempFilePool>
{
    [Fact]
    public void Should_throw_exception_if_package_is_disposed()
    {
        // Arrange
        var sut = OpenXmlPackage.Create();

        // Act
        sut.Dispose();

        // Assert
        var testCases = new Action[]
        {
            () => sut.EnumerateParts(),
            () => sut.Save(),
            () => sut.SaveTo(remoteStream: new MemoryStream()),
            () => sut.SaveTo(filePath: nameof(OpenXmlPackage))
        };

        testCases
            .Should()
            .AllSatisfy(testCase =>
                testCase
                    .Should()
                    .Throw<ObjectDisposedException>());
    }

    [Fact]
    public void Should_do_nothing_if_package_is_disposed()
    {
        // Arrange
        var sut = OpenXmlPackage.Create();
        
        sut.Dispose();
        
        // Act & Assert
        sut.Invoking(package => package.Dispose())
            .Should()
            .NotThrow();
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

        var filePath = tempFilePool.GetTempFilePath();

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
    
    [Theory]
    [InlineData(null)]
    [InlineData("")]
    public void Should_throw_exception_if_relationship_type_is_not_specified(string? relationshipType)
    {
        // Arrange
        var expectedPart = TestOpenXmlPackagePartFactory.Create(uri: "/part", relationshipType!);
        
        // Act & Assert
        new Action(() =>
        {
            using var sut =
                OpenXmlPackage.Create();
                
            sut.AddPart(expectedPart);
        })
        .Should()
        .Throw<InvalidOperationException>();
    }
    
    [Fact]
    public void Should_do_nothing_if_package_already_exists()
    {
        // Arrange
        using var sut =
            OpenXmlPackage.Create();

        var expectedPart = TestOpenXmlPackagePartFactory.Create(uri: "/part");

        sut.AddPart(expectedPart);
        
        // Act
        sut.AddPart(expectedPart);
        
        // Arrange
        sut.EnumerateParts()
            .Should()
            .ContainSingle(actualPart =>
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
    public void Should_return_package_part_properly()
    {
        // Arrange
        using var originalStream =
            PrepareTestPackageStream();

        var expectedPart =
            TestOpenXmlPackagePartFactory.Create(uri: "/part");

        using var sut =
            OpenXmlPackage.Open(originalStream);

        sut.AddPart(expectedPart);

        // Act
        var isExists = sut.TryGetPart(expectedPart.Uri, out var actualPart);

        // Assert
        isExists
            .Should()
            .BeTrue();

        actualPart
            .ContentType
            .Should()
            .Be(expectedPart.ContentType);

        actualPart
            .CompressionMode
            .Should()
            .Be(expectedPart.CompressionMode);
    }

    [Fact]
    public void Should_not_return_package_part()
    {
        // Arrange
        using var originalStream =
            PrepareTestPackageStream();

        var uri = new Uri("/part", UriKind.Relative);

        using var sut =
            OpenXmlPackage.Open(originalStream);

        // Act
        var isExists = sut.TryGetPart(uri, out _);

        // Assert
        isExists
            .Should()
            .BeFalse();
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

        var originalLength = originalStream.Length;
        
        using var sut =
            OpenXmlPackage.Open(originalStream);
        
        var expectedPart = TestOpenXmlPackagePartFactory.Create(uri: "/part");
        
        sut.AddPart(expectedPart);

        // Act
        sut.Save();

        // Assert
        originalStream
            .Length
            .Should()
            .BeGreaterThan(originalLength);
    }

    [Fact]
    public void Should_save_package_using_file_path_properly()
    {
        // Arrange
        using var originalStream =
            PrepareTestPackageStream();

        var originalLength = (int)originalStream.Length;

        var filePath = tempFilePool.GetTempFilePath();
        using (var fileStream = new FileStream(filePath, FileMode.Open, FileAccess.ReadWrite))
            originalStream.CopyTo(fileStream);

        using var sut =
            OpenXmlPackage.Open(filePath);
        
        var expectedPart = TestOpenXmlPackagePartFactory.Create(uri: "/part");
        
        sut.AddPart(expectedPart);

        // Act
        sut.Save();

        // Assert
        File.ReadAllBytes(filePath)
            .Length
            .Should()
            .BeGreaterThan(originalLength);
    }

    [Fact]
    public void Should_save_package_to_stream_properly()
    {
        // Arrange
        using var originalStream =
            PrepareTestPackageStream();

        var originalLength = originalStream.Length;
        
        using var destinationStream =
            new MemoryStream();

        using var sut =
            OpenXmlPackage.Open(originalStream);
        
        var expectedPart = TestOpenXmlPackagePartFactory.Create(uri: "/part");
        
        sut.AddPart(expectedPart);

        // Act
        sut.SaveTo(destinationStream);

        // Assert
        destinationStream
            .Length
            .Should()
            .BeGreaterThan(originalLength);
    }

    [Fact]
    public void Should_save_package_to_file_properly()
    {
        // Arrange
        using var originalStream =
            PrepareTestPackageStream();
        
        var originalLength = (int)originalStream.Length;

        var filePath = tempFilePool.GetTempFilePath();

        using var sut =
            OpenXmlPackage.Open(originalStream);
        
        var expectedPart = TestOpenXmlPackagePartFactory.Create(uri: "/part");
        
        sut.AddPart(expectedPart);

        // Act
        sut.SaveTo(filePath);

        // Assert
        File.ReadAllBytes(filePath)
            .Length
            .Should()
            .BeGreaterThan(originalLength);
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