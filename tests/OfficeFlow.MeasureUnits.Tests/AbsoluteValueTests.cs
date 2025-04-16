using FluentAssertions;
using OfficeFlow.MeasureUnits.Absolute;
using OfficeFlow.MeasureUnits.Tests.Fakes;
using Xunit;

namespace OfficeFlow.MeasureUnits.Tests;

public sealed class AbsoluteValueTests
{
    [Fact]
    public void Should_create_an_absolute_value()
    {
        // Arrange
        var units = new FakeAbsoluteUnits(2.0);

        // Act
        var sut = AbsoluteValue.From(10.0, units);

        // Assert
        sut.Units
            .Should()
            .BeOfType<FakeAbsoluteUnits>();

        sut.Units
            .Ratio
            .Should()
            .Be(2.0);

        sut.Raw
            .Should()
            .Be(10.0);
    }

    [Fact]
    public void Should_convert_an_absolute_value_to_emu()
    {
        // Arrange
        var units = new FakeAbsoluteUnits(2.0);
        var sut = AbsoluteValue.From(10.0, units);

        // Act
        var emu = sut.To(AbsoluteUnits.Emu);

        // Assert
        emu.Raw
            .Should()
            .Be(20.0);
    }

    [Fact]
    public void Should_wrap_to_generic_value_with_same_units()
    {
        // Arrange
        var sut = AbsoluteValue.From(20.0, AbsoluteUnits.Emu);

        // Act
        var value = sut.To(AbsoluteUnits.Emu);

        // Assert
        value.Should()
            .BeOfType<AbsoluteValue<Emu>>();
    }
    
    [Fact]
    public void Should_be_equal_to_itself()
    {
        // Arrange
        var sut = AbsoluteValue.From(20.0, new FakeAbsoluteUnits());
        
        // Act & Assert
        sut.Equals(sut)
            .Should()
            .BeTrue();
    }

    [Fact]
    public void Two_same_values_should_be_equal()
    {
        // Arrange
        var left = AbsoluteValue.From(20.0, new FakeAbsoluteUnits());
        var right = AbsoluteValue.From(20.0, new FakeAbsoluteUnits());
        
        // Act && Assert
        left.Equals(right)
            .Should()
            .BeTrue();
    }
    
    [Fact]
    public void Two_different_values_should_not_be_equal()
    {
        // Arrange
        var left = AbsoluteValue.From(20.0, new FakeAbsoluteUnits());
        var right = AbsoluteValue.From(40.0, new FakeAbsoluteUnits());
        
        // Act && Assert
        left.Equals(right)
            .Should()
            .BeFalse();
    }
    
    [Fact]
    public void Two_same_values_with_different_units_should_not_be_equal()
    {
        // Arrange
        var left = AbsoluteValue.From(20.0, new FakeAbsoluteUnits());
        var right = AbsoluteValue.From(20.0, AbsoluteUnits.Emu);
        
        // Act && Assert
        // ReSharper disable once SuspiciousTypeConversion.Global
        left.Equals(right)
            .Should()
            .BeFalse();
    }
    
    [Fact]
    public void Should_returns_string_representation_of_value()
    {
        // Arrange
        var sut = AbsoluteValue.From(10.0, new FakeAbsoluteUnits());

        // Act
        var stringRepresentation = sut.ToString();

        // Assert
        stringRepresentation
            .Should()
            .Be("10");
    }
}