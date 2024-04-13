using System;
using System.Collections.Generic;
using System.Globalization;
using FluentAssertions;
using OfficeFlow.MeasureUnits.Extensions;
using OfficeFlow.MeasureUnits.Tests.Fakes;
using Xunit;

namespace OfficeFlow.MeasureUnits.Tests
{
    public sealed class ConvertibleExtensionsTests
    {
        [Theory]
        [MemberData(nameof(GetValidTestCases))]
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
        {
            // Arrange & Act & Assert
            new Action(() => convertible.As<FakeAbsoluteUnits>())
                .Should()
                .Throw<FormatException>();
        }

        [Theory]
        [MemberData(nameof(GetInvalidTestCases))]
        public void Should_throws_exception_while_converting_from_not_via_number(IConvertible convertible)
        {
            // Arrange & Act & Assert
            new Action(() => convertible.As<FakeAbsoluteUnits>())
                .Should()
                .Throw<InvalidCastException>();
        }

        public static IEnumerable<object[]> GetValidTestCases()
        {
            yield return new object[] { true }; // because bool can be converted to number (0, 1)
            yield return new object[] { sbyte.MaxValue };
            yield return new object[] { byte.MaxValue };
            yield return new object[] { short.MaxValue };
            yield return new object[] { ushort.MaxValue };
            yield return new object[] { int.MaxValue };
            yield return new object[] { uint.MaxValue };
            yield return new object[] { long.MaxValue };
            yield return new object[] { ulong.MaxValue };
            yield return new object[] { float.MaxValue };
            yield return new object[] { double.MaxValue };
            yield return new object[] { decimal.MaxValue };
        }

        public static IEnumerable<object[]> GetInvalidTestCases()
        {
            yield return new object[] { '\0' };
            yield return new object[] { DateTime.MaxValue };
        }
    }
}