using System;

namespace OfficeFlow.MeasureUnits.Absolute;

public abstract class AbsoluteUnits : IEquatable<AbsoluteUnits>
{
    public static readonly Centimeters Centimeters 
        = new();

    public static readonly Points Points 
        = new();

    public static readonly Inches Inches 
        = new();

    public static readonly Picas Picas 
        = new();

    public static readonly Millimeters Millimeters 
        = new();

    internal static readonly Emu Emu 
        = new();

    internal static readonly HalfPoints HalfPoints 
        = new();

    public static readonly Twips Twips 
        = new();
    
    internal double Ratio { get; }

    protected AbsoluteUnits(double ratio)
        => Ratio = ratio;

    internal AbsoluteValue<Emu> ToEmu(double value)
        => AbsoluteValue.From(value * Ratio, Emu);

    internal double FromEmu(in AbsoluteValue<Emu> emu)
        => emu.Raw / Ratio;

    /// <inheritdoc />
    public bool Equals(AbsoluteUnits? other)
        => other is not null && Ratio.Equals(other.Ratio);

    /// <inheritdoc />
    public override bool Equals(object? other)
        => other is AbsoluteUnits units && Equals(units);

    /// <inheritdoc />
    public override int GetHashCode()
        => Ratio.GetHashCode();
}