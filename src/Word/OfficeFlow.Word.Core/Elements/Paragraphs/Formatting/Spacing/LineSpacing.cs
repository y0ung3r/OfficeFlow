﻿using OfficeFlow.MeasureUnits.Absolute;
using OfficeFlow.Word.Core.Elements.Paragraphs.Formatting.Spacing.Interfaces;

namespace OfficeFlow.Word.Core.Elements.Paragraphs.Formatting.Spacing;

public static class LineSpacing
{
    public static readonly ILineSpacing Single
        = Multiple(factor: 1.0);

    public static readonly ILineSpacing OneAndHalf
        = Multiple(factor: 1.5);
    
    public static readonly ILineSpacing Double
        = Multiple(factor: 2.0);
    
    public static ILineSpacing Exactly<TUnits>(AbsoluteValue<TUnits> value)
        where TUnits : AbsoluteUnits, new()
        => new ExactSpacing(
            value.To(AbsoluteUnits.Points));

    public static ILineSpacing Exactly<TUnits>(double value, TUnits units)
        where TUnits : AbsoluteUnits, new()
        => Exactly(
            AbsoluteValue.From(value, units));
    
    public static ILineSpacing AtLeast<TUnits>(AbsoluteValue<TUnits> value)
        where TUnits : AbsoluteUnits, new()
        => new AtLeastSpacing(
            value.To(AbsoluteUnits.Points));

    public static ILineSpacing AtLeast<TUnits>(double value, TUnits units)
        where TUnits : AbsoluteUnits, new()
        => AtLeast(
            AbsoluteValue.From(value, units));
    
    // TODO[#11]: Design relative measurement units
    public static ILineSpacing Multiple(double factor)
        => new MultipleSpacing(factor);
}