using System;
using OfficeFlow.MeasureUnits.Absolute;
using OfficeFlow.Word.Core.Elements.Paragraphs.Spacing.Interfaces;

namespace OfficeFlow.Word.Core.Elements.Paragraphs.Spacing;

public sealed class ExactSpacing : IParagraphSpacing, ILineSpacing, IEquatable<ExactSpacing>
{
    public AbsoluteValue<Points> Value { get; }
    
    internal ExactSpacing(AbsoluteValue<Points> value)
        => Value = value;

    public bool Equals(ExactSpacing? other)
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
        
        return Equals((ExactSpacing)other);
    }

    /// <inheritdoc />
    public override int GetHashCode()
        => Value.GetHashCode();
}