using System;
using OfficeFlow.MeasureUnits.Absolute;
using OfficeFlow.Word.Core.Elements.Paragraphs.Spacing.Interfaces;

namespace OfficeFlow.Word.Core.Elements.Paragraphs.Spacing;

public sealed class AtLeastSpacing : ILineSpacing, IEquatable<AtLeastSpacing>
{
    public AbsoluteValue<Points> Value { get; }
    
    internal AtLeastSpacing(AbsoluteValue<Points> value)
        => Value = value;
    
    public bool Equals(AtLeastSpacing? other)
    {
        if (ReferenceEquals(null, other))
        {
            return false;
        }

        if (ReferenceEquals(this, other))
        {
            return true;
        }
        
        return Value.Equals(other.Value);
    }

    /// <inheritdoc />
    public override bool Equals(object? other)
    {
        if (ReferenceEquals(null, other))
        {
            return false;
        }
        
        if (ReferenceEquals(this, other))
        {
            return true;
        }
        
        if (other.GetType() != GetType())
        {
            return false;
        }
        
        return Equals((AtLeastSpacing)other);
    }

    /// <inheritdoc />
    public override int GetHashCode()
        => Value.GetHashCode();
}