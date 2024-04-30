﻿using OfficeFlow.MeasureUnits.Absolute;
using OfficeFlow.Word.Core.Elements.Paragraphs.Spacing.Interfaces;

namespace OfficeFlow.Word.Core.Elements.Paragraphs.Spacing;

public static class ParagraphSpacing
{
    public static IParagraphSpacing Exactly<TUnits>(AbsoluteValue<TUnits> value)
        where TUnits : AbsoluteUnits, new()
        => new ExactSpacing(
            value.To(AbsoluteUnits.Points));

    public static IParagraphSpacing Exactly<TUnits>(double value, TUnits units)
        where TUnits : AbsoluteUnits, new()
        => Exactly(
            AbsoluteValue.From(value, units));
    
    public static readonly IParagraphSpacing Auto
        = new AutoSpacing();
}