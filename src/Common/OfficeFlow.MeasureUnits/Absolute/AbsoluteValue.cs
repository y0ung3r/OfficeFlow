using System;
using System.Globalization;

namespace OfficeFlow.MeasureUnits.Absolute;

public class AbsoluteValue : IEquatable<AbsoluteValue>
{
    public static AbsoluteValue From(double value, AbsoluteUnits units) 
        => new(value, units);

    public double Raw { get; }

    public AbsoluteUnits Units { get; }

    protected AbsoluteValue(double value, AbsoluteUnits units)
    {
        if (value < 0)
        {
            throw new ArgumentException(
                $"The value \"{nameof(value)}\" cannot be less than zero");
        }

        Units = units;
        Raw = value;
    }

    public AbsoluteValue<TTargetUnits> To<TTargetUnits>()
        where TTargetUnits : AbsoluteUnits, new()
    {
        if (Units is TTargetUnits)
        {
            return AbsoluteValue<TTargetUnits>.From(Raw);
        }

        var emu = Units.ToEmu(Raw);
        var targetUnits = new TTargetUnits();
        var convertedValue = targetUnits.FromEmu(emu);

        return AbsoluteValue<TTargetUnits>.From(convertedValue);
    }

    /// <inheritdoc />
    public bool Equals(AbsoluteValue? other)
    {
        if (ReferenceEquals(null, other))
        {
            return false;
        }

        if (ReferenceEquals(this, other))
        {
            return true;
        }

        return Raw.Equals(other.Raw) && Units.Equals(other.Units);
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

        return Equals((AbsoluteValue)other);
    }

    /// <inheritdoc />
    public override int GetHashCode()
    {
        unchecked
        {
            var hashCode = (int) 2166136261;

            hashCode = hashCode * 16777619 ^ Raw.GetHashCode();
            hashCode = hashCode * 16777619 ^ Units.GetHashCode();
            
            return hashCode;
        }
    }

    public override string ToString()
        => Raw.ToString(CultureInfo.InvariantCulture);
}

public class AbsoluteValue<TUnits> : AbsoluteValue
    where TUnits : AbsoluteUnits, new()
{
    public static AbsoluteValue<TUnits> Zero
        => new(0.0);

    public static AbsoluteValue<TUnits> From(double value)
        => new(value);

    private AbsoluteValue(double value)
        : base(value, new TUnits())
    { }
}