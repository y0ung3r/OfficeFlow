using System.IO;
using FluentAssertions;
using Xunit;

namespace OfficeFlow.Word.Tests;

public sealed class DocumentSettingsTests
{
    [Fact]
    public void Should_always_create_new_default_settings()
        => DocumentSettings
            .Default
            .Should()
            .NotBeSameAs(DocumentSettings.Default);
	
    [Fact]
    public void Should_enable_read_only_mode()
    {
        // Arrange & Act
        var sut = new DocumentSettings
        {
            IsReadOnly = true
        };
		
        // Assert
        sut.AccessMode
            .Should()
            .Be(FileAccess.Read);
    }
	
    [Fact]
    public void Should_disable_read_only_mode()
    {
        // Arrange & Act
        var sut = new DocumentSettings
        {
            IsReadOnly = false
        };
		
        // Assert
        sut.AccessMode
            .Should()
            .Be(FileAccess.ReadWrite);
    }

    [Fact]
    public void Should_disable_auto_saving_in_read_only_mode()
    {
        // Arrange & Act
        var sut = new DocumentSettings
        {
            IsReadOnly = true,
            AllowAutoSaving = true
        };

        // Assert
        sut.AllowAutoSaving
            .Should()
            .Be(false);
    }
	
    [Fact]
    public void Should_enable_auto_saving()
    {
        // Arrange & Act
        var sut = new DocumentSettings
        {
            IsReadOnly = false,
            AllowAutoSaving = true
        };

        // Assert
        sut.AllowAutoSaving
            .Should()
            .Be(true);
    }
}