using System;
using System.Globalization;
using FluentAssertions;
using OfficeFlow.MeasureUnits.Extensions;
using OfficeFlow.MeasureUnits.Tests.Fakes;
using Xunit;

namespace OfficeFlow.MeasureUnits.Tests;

public sealed class ConvertibleExtensionsTests
{
    [Theory]
    [MemberData(nameof(ValidTestCases))]
    public void Should_convert_properly(IConvertible convertible)
    {
        // Arrange
        var expectedValue = convertible.ToDouble(CultureInfo.InvariantCulture);
            
        // Act
        var value = convertible.As<FakeAbsoluteUnits>();
            
        // Assert
        value.Raw
            .Should()
            .Be(expectedValue);
    }

    [Theory]
    [InlineData("some text")]
    [InlineData("0.5va")]
    [InlineData("")]
    [InlineData("ab0.5")]
    public void Should_throws_exception_while_converting_from_string_with_invalid_format(IConvertible convertible)
        => convertible
            .Invoking(sut => sut.As<FakeAbsoluteUnits>())
            .Should()
            .Throw<FormatException>();

    [Theory]
    [MemberData(nameof(InvalidTestCases))]
    public void Should_throws_exception_while_converting_from_not_via_number(IConvertible convertible)
        => convertible
            .Invoking(sut => sut.As<FakeAbsoluteUnits>())
            .Should()
            .Throw<InvalidCastException>();

    public static readonly TheoryData<IConvertible> ValidTestCases = new()
    {
        true,
        sbyte.MaxValue,
        byte.MaxValue,
        short.MaxValue,
        ushort.MaxValue,
        int.MaxValue,
        uint.MaxValue,
        long.MaxValue,
        ulong.MaxValue,
        float.MaxValue,
        double.MaxValue,
        decimal.MaxValue
    };

    public static readonly TheoryData<IConvertible> InvalidTestCases = new()
    {
        '\0',
        DateTime.MaxValue
    };
}