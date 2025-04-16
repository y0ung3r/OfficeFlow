using OfficeFlow.MeasureUnits.Absolute;
using OfficeFlow.Word.Core.Elements.Paragraphs.Formatting.Spacing.Interfaces;

namespace OfficeFlow.Word.Core.Elements.Paragraphs.Formatting.Spacing;

public static class ParagraphSpacing
{
    public static IParagraphSpacing Exactly<TUnits>(in AbsoluteValue<TUnits> value)
        where TUnits : AbsoluteUnits
        => new ExactSpacing(
            value.To(AbsoluteUnits.Points));

    public static IParagraphSpacing Exactly<TUnits>(double value, TUnits units)
        where TUnits : AbsoluteUnits
        => Exactly(
            AbsoluteValue.From(value, units));
    
    public static readonly IParagraphSpacing Auto
        = new AutoSpacing();
}