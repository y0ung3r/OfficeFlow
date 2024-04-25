using System;

namespace OfficeFlow.MeasureUnits.Absolute;

public abstract class AbsoluteUnits : IEquatable<AbsoluteUnits>
{
    public static Centimeters Centimeters => new();

    public static Points Points => new();

    public static Inches Inches => new();

    public static Picas Picas => new();

    public static Millimeters Millimeters => new();

    internal static Emu Emu => new();

    internal static HalfPoints HalfPoints => new();

    public static Twips Twips => new();

    internal abstract double Ratio { get; }

    internal AbsoluteValue<Emu> ToEmu(double value)
        => AbsoluteValue<Emu>.From(value * Ratio);

    internal double FromEmu(AbsoluteValue<Emu> emu)
        => emu.Raw / Ratio;

    /// <inheritdoc />
    public bool Equals(AbsoluteUnits? other)
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

        return Ratio.Equals(other.Ratio);
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

        return Equals((AbsoluteUnits)other);
    }

    /// <inheritdoc />
    public override int GetHashCode()
        => Ratio.GetHashCode();
}