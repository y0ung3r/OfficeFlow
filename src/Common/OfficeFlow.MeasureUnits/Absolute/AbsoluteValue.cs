using System;
using System.Globalization;

namespace OfficeFlow.MeasureUnits.Absolute;

public static class AbsoluteValue
{
    public static AbsoluteValue<TUnits> From<TUnits>(double value, TUnits units) 
        where TUnits : AbsoluteUnits, new()
        => AbsoluteValue<TUnits>.From(value, units);
}

public sealed class AbsoluteValue<TUnits> : IEquatable<AbsoluteValue<TUnits>>
    where TUnits : AbsoluteUnits, new()
{
    internal static AbsoluteValue<TUnits> From(double value, TUnits units)
        => new(value, units);
    
    public double Raw { get; }

    public AbsoluteUnits Units { get; }

    private AbsoluteValue(double value, TUnits units)
    {
        Raw = value;
        Units = units;
    }
    
    public AbsoluteValue<TTargetUnits> To<TTargetUnits>(TTargetUnits targetUnits)
        where TTargetUnits : AbsoluteUnits, new()
    {
        if (Units is TTargetUnits)
        {
            return AbsoluteValue.From(Raw, targetUnits);
        }

        var emu = Units.ToEmu(Raw);
        var convertedValue = targetUnits.FromEmu(emu);

        return AbsoluteValue.From(convertedValue, targetUnits);
    }
    
    /// <inheritdoc />
    public bool Equals(AbsoluteValue<TUnits>? other)
        => other is not null && Raw.Equals(other.Raw) && Units.Equals(other.Units);

    /// <inheritdoc />
    public override bool Equals(object? other)
        => other is AbsoluteValue<TUnits> value && Equals(value);

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