namespace OfficeFlow.MeasureUnits.Absolute;

public abstract class AbsoluteUnits
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
}