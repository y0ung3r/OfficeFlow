using FluentAssertions;
using OfficeFlow.MeasureUnits.Absolute;
using OfficeFlow.MeasureUnits.Tests.Fakes;
using Xunit;

namespace OfficeFlow.MeasureUnits.Tests;

public sealed class AbsoluteUnitsTests
{
    [Fact]
    public void Should_create_instance_of_centimeters()
    {
        // Arrange & Act
        var sut = AbsoluteUnits.Centimeters;

        // Assert
        sut.Should()
            .NotBeNull();

        sut.Ratio
            .Should()
            .Be(ConversionRatios.Centimeters);
    }

    [Fact]
    public void Should_create_instance_of_points()
    {
        // Arrange & Act
        var sut = AbsoluteUnits.Points;

        // Assert
        sut.Should()
            .NotBeNull();

        sut.Ratio
            .Should()
            .Be(ConversionRatios.Points);
    }

    [Fact]
    public void Should_create_instance_of_inches()
    {
        // Arrange & Act
        var sut = AbsoluteUnits.Inches;

        // Assert
        sut.Should()
            .NotBeNull();

        sut.Ratio
            .Should()
            .Be(ConversionRatios.Inches);
    }

    [Fact]
    public void Should_create_instance_of_picas()
    {
        // Arrange & Act
        var sut = AbsoluteUnits.Picas;

        // Assert
        sut.Should()
            .NotBeNull();

        sut.Ratio
            .Should()
            .Be(ConversionRatios.Picas);
    }

    [Fact]
    public void Should_create_instance_of_millimeters()
    {
        // Arrange & Act
        var sut = AbsoluteUnits.Millimeters;

        // Assert
        sut.Should()
            .NotBeNull();

        sut.Ratio
            .Should()
            .Be(ConversionRatios.Millimeters);
    }

    [Fact]
    public void Should_create_instance_of_emu()
    {
        // Arrange & Act
        var sut = AbsoluteUnits.Emu;

        // Assert
        sut.Should()
            .NotBeNull();

        sut.Ratio
            .Should()
            .Be(ConversionRatios.Emu);
    }

    [Fact]
    public void Should_create_instance_of_half_points()
    {
        // Arrange & Act
        var sut = AbsoluteUnits.HalfPoints;

        // Assert
        sut.Should()
            .NotBeNull();

        sut.Ratio
            .Should()
            .Be(ConversionRatios.HalfPoints);
    }

    [Fact]
    public void Should_create_instance_of_twips()
    {
        // Arrange & Act
        var sut = AbsoluteUnits.Twips;

        // Assert
        sut.Should()
            .NotBeNull();

        sut.Ratio
            .Should()
            .Be(ConversionRatios.Twips);
    }

    [Fact]
    public void Should_convert_an_absolute_value_to_emu()
    {
        // Arrange
        var sut = new FakeAbsoluteUnits(2.0);

        // Act
        var emu = sut.ToEmu(10.0);

        // Assert
        emu.Raw
            .Should()
            .Be(20.0);
    }

    [Fact]
    public void Should_convert_emu_to_an_absolute_value()
    {
        // Arrange
        var emu = AbsoluteValue.From(20.0, AbsoluteUnits.Emu);
        var sut = new FakeAbsoluteUnits(2.0);

        // Act
        var value = sut.FromEmu(emu);

        // Assert
        value
            .Should()
            .Be(10.0);
    }

    [Fact]
    public void Should_be_equal_to_itself()
    {
        // Arrange
        var sut = new FakeAbsoluteUnits(2.0);
        
        // Act & Assert
        sut.Equals(sut)
            .Should()
            .BeTrue();
    }

    [Fact]
    public void Two_same_units_should_be_equal()
    {
        // Arrange
        var left = new FakeAbsoluteUnits(2.0);
        var right = new FakeAbsoluteUnits(2.0);
        
        // Act & Assert
        left.Equals(right)
            .Should()
            .BeTrue();
    }
    
    [Fact]
    public void Two_same_units_with_different_ratio_should_not_be_equal()
    {
        // Arrange
        var left = new FakeAbsoluteUnits(2.0);
        var right = new FakeAbsoluteUnits(4.0);
        
        // Act & Assert
        left.Equals(right)
            .Should()
            .BeFalse();
    }
    
    [Fact]
    public void Two_different_units_should_be_equal()
    {
        // Arrange
        var left = new FakeAbsoluteUnits(2.0);
        var right = new Emu();
        
        // Act & Assert
        left.Equals(right)
            .Should()
            .BeFalse();
    }
}