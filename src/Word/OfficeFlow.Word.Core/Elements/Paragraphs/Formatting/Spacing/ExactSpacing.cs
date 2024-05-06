using System;
using OfficeFlow.MeasureUnits.Absolute;
using OfficeFlow.Word.Core.Elements.Paragraphs.Formatting.Spacing.Interfaces;

namespace OfficeFlow.Word.Core.Elements.Paragraphs.Formatting.Spacing;

public sealed class ExactSpacing : IParagraphSpacing, ILineSpacing, IEquatable<ExactSpacing>
{
    public AbsoluteValue<Points> Value { get; }
    
    internal ExactSpacing(AbsoluteValue<Points> value)
        => Value = value;

    public bool Equals(ExactSpacing? other)
        => other is not null && Value.Equals(other.Value);

    /// <inheritdoc />
    public override bool Equals(object? other)
        => other is ExactSpacing spacing && Equals(spacing);

    /// <inheritdoc />
    public override int GetHashCode()
        => Value.GetHashCode();
}