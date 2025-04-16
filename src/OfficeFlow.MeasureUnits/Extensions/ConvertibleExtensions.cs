using System;
using System.Globalization;
using OfficeFlow.MeasureUnits.Absolute;

namespace OfficeFlow.MeasureUnits.Extensions;

public static class ConvertibleExtensions
{
    public static AbsoluteValue<TUnits> ToUnits<TUnits>(this IConvertible convertible, TUnits units)
        where TUnits : AbsoluteUnits
    {
        var value = convertible.ToDouble(CultureInfo.InvariantCulture);
        return AbsoluteValue.From(value, units);
    }
}