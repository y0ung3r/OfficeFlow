using OfficeFlow.MeasureUnits.Absolute;
using OfficeFlow.Word.Core.Elements.Paragraphs.Spacing.Interfaces;

namespace OfficeFlow.Word.Core.Elements.Paragraphs.Spacing;

public static class LineSpacing
{
    public static ILineSpacing Exactly<TUnits>(AbsoluteValue<TUnits> value)
        where TUnits : AbsoluteUnits, new()
        => new ExactSpacing(
            value.To<Points>());

    public static ILineSpacing Exactly<TUnits>(double value)
        where TUnits : AbsoluteUnits, new()
        => Exactly(
            AbsoluteValue<TUnits>.From(value));
    
    public static ILineSpacing AtLeast<TUnits>(AbsoluteValue<TUnits> value)
        where TUnits : AbsoluteUnits, new()
        => new AtLeastSpacing(
            value.To<Points>());

    public static ILineSpacing AtLeast<TUnits>(double value)
        where TUnits : AbsoluteUnits, new()
        => AtLeast(
            AbsoluteValue<TUnits>.From(value));
    
    public static readonly ILineSpacing Auto
        = new AutoSpacing();
}