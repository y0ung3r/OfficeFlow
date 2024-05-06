using System;
using OfficeFlow.Word.Core.Elements.Paragraphs.Formatting.Spacing.Interfaces;

namespace OfficeFlow.Word.Core.Elements.Paragraphs.Formatting.Spacing;

public sealed class MultipleSpacing : ILineSpacing, IEquatable<MultipleSpacing>
{
    private const double MinValue = 0.0;
    private const double MaxValue = 132.0;
    
    // TODO[#11]: Design relative measurement units
    public double Factor { get; }
    
    internal MultipleSpacing(double factor)
    {
        if (factor is < MinValue or > MaxValue)
            throw new ArgumentException(
                $"The factor of multiple spacing must be between {MinValue} and {MaxValue}", 
                nameof(factor));
            
        Factor = factor;
    }
    
    public bool Equals(MultipleSpacing? other)
        => other is not null && Factor.Equals(other.Factor);

    /// <inheritdoc />
    public override bool Equals(object? other)
        => other is MultipleSpacing spacing && Equals(spacing);

    /// <inheritdoc />
    public override int GetHashCode()
        => Factor.GetHashCode();
}