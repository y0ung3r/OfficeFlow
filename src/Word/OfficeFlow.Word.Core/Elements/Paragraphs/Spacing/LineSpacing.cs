using OfficeFlow.MeasureUnits.Absolute;
using OfficeFlow.Word.Core.Elements.Paragraphs.Spacing.Interfaces;

namespace OfficeFlow.Word.Core.Elements.Paragraphs.Spacing;

public static class LineSpacing
{
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
    
    public static readonly ILineSpacing Auto
        = new AutoSpacing();
}