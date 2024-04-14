using System;
using FluentAssertions;
using OfficeFlow.Word.Core.Styling;
using Xunit;

namespace OfficeFlow.Word.Core.Tests;

public sealed class RgbColorTests
{
    [Fact]
    public void Should_returns_auto_color()
        => RgbColor.Auto
            .IsAuto
            .Should()
            .BeTrue();

    [Fact]
    public void Should_create_from_rgb_properly()
    {
        // Arrange
        const int red = 0xFF;
        const int green = 0xFF;
        const int blue = 0xFF;

        // Act
        var sut = RgbColor.FromRgb(red, green, blue);
            
        // Assert
        sut.Red
            .Should()
            .Be(0xFF);
            
        sut.Green
            .Should()
            .Be(0xFF);
            
        sut.Blue
            .Should()
            .Be(0xFF);
    }

    [Theory]
    [InlineData(256, 0x0, 0x0)]
    [InlineData(0x0, -1, 0x0)]
    [InlineData(0x0, 0x0, 256)]
    public void Should_throws_exception_while_creating_from_rgb(int red, int green, int blue)
        => new Action(() => RgbColor.FromRgb(red, green, blue))
            .Should()
            .Throw<ArgumentException>();

    [Fact]
    public void Should_convert_from_system_color()
    {
        // Arrange
        var systemColor = System.Drawing.Color.FromArgb(red: 255, green: 255, blue: 255);
            
        // Act
        var sut = RgbColor.FromSystemColor(systemColor);

        // Assert
        sut.Red
            .Should()
            .Be(0xFF);
            
        sut.Green
            .Should()
            .Be(0xFF);
            
        sut.Blue
            .Should()
            .Be(0xFF);
    }
        
    [Theory]
    [InlineData(null, 0x0, 0x0, 0x0)]
    [InlineData("", 0x0, 0x0, 0x0)]
    [InlineData("fff", 0xFF, 0xFF, 0xFF)]
    [InlineData("#fff", 0xFF, 0xFF, 0xFF)]
    [InlineData("ffffff", 0xFF, 0xFF, 0xFF)]
    [InlineData("#ffffff", 0xFF, 0xFF, 0xFF)]
    [InlineData("abc", 0xAA, 0xBB, 0xCC)]
    [InlineData("aabbcc", 0xAA, 0xBB, 0xCC)]
    [InlineData("000", 0x0, 0x0, 0x0)]
    [InlineData("#000", 0x0, 0x0, 0x0)]
    [InlineData("000000", 0x0, 0x0, 0x0)]
    [InlineData("#000000", 0x0, 0x0, 0x0)]
    public void Should_convert_from_hex_string(string? hex, byte red, byte green, byte blue)
    {
        // Arrange & Act
        var color = RgbColor.FromHex(hex);

        // Assert
        color.Red
            .Should()
            .Be(red);
            
        color.Green
            .Should()
            .Be(green);
            
        color.Blue
            .Should()
            .Be(blue);
    }

    [Theory]
    [InlineData("#")]
    [InlineData("x")]
    [InlineData("#x")]
    [InlineData("yy")]
    [InlineData("#xx")]
    [InlineData("xxyc")]
    [InlineData("#xxyc")]
    [InlineData("xxyyy")]
    [InlineData("#xxyyy")]
    [InlineData("xxxyyyz")]
    public void Should_throws_format_exception_while_converting_from_hex_string(string hex)
        => new Action(() => RgbColor.FromHex(hex))
            .Should()
            .Throw<FormatException>();
}