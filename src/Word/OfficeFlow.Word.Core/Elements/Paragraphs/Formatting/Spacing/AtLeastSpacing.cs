using System;
using OfficeFlow.MeasureUnits.Absolute;
using OfficeFlow.Word.Core.Elements.Paragraphs.Formatting.Spacing.Interfaces;

namespace OfficeFlow.Word.Core.Elements.Paragraphs.Formatting.Spacing;

public sealed class AtLeastSpacing : ILineSpacing, IEquatable<AtLeastSpacing>
{
    public AbsoluteValue<Points> Value { get; }
    
    internal AtLeastSpacing(AbsoluteValue<Points> value)
        => Value = value;
    
    public bool Equals(AtLeastSpacing? other)
        => other is not null && Value.Equals(other.Value);

    /// <inheritdoc />
    public override bool Equals(object? other)
        => other is AtLeastSpacing spacing && Equals(spacing);

    /// <inheritdoc />
    public override int GetHashCode()
        => Value.GetHashCode();
}