using System;
using System.Globalization;
using OfficeFlow.MeasureUnits.Absolute;

namespace OfficeFlow.MeasureUnits.Extensions
{
    public static class ConvertibleExtensions
    {
        public static AbsoluteValue As(this IConvertible convertible, AbsoluteUnits units)
        {
            var value = convertible.ToDouble(CultureInfo.InvariantCulture);
            return AbsoluteValue.From(value, units);
        }
		
        public static AbsoluteValue<TUnits> As<TUnits>(this IConvertible convertible)
            where TUnits : AbsoluteUnits, new()
        {
            var units = new TUnits();
            
            return convertible
                .As(units)
                .To<TUnits>();
        }
    }
}