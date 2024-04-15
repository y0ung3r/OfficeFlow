using System;

namespace OfficeFlow.Word.Core.Styling;

public readonly struct RgbColor : IEquatable<RgbColor>
{
    private const int RedShift = 16;

    private const int GreenShift = 8;

    private const int BlueShift = 0;

    public static readonly RgbColor Auto = new();

    public static RgbColor FromSystemColor(System.Drawing.Color color) 
        => new(color.ToArgb());

    public static RgbColor FromRgb(int red, int green, int blue)
    {
        CheckByte(red, nameof(red));
        CheckByte(green, nameof(green));
        CheckByte(blue, nameof(blue));
        
        return new RgbColor((uint)red << RedShift | (uint)green << GreenShift | (uint)blue << BlueShift);
    }

    public static RgbColor FromHex(string? hex)
    {
        if (hex is null || hex.Length is 0)
        {
            return Auto;
        }

        hex = hex.Replace("#", string.Empty);

        switch (hex.Length)
        {
            case 3:
                var red = char.ToString(hex[0]);
                var green = char.ToString(hex[1]);
                var blue = char.ToString(hex[2]);

                return FromRgb(
                    Convert.ToInt16(red + red, fromBase: 16),
                    Convert.ToInt16(green + green, fromBase: 16),
                    Convert.ToInt16(blue + blue, fromBase: 16));

            case 6:
                var value = Convert.ToInt64(hex, fromBase: 16);
                return new RgbColor(value);

            default:
                throw new FormatException(
                    "The value must have the following format: #XXX, #XXX, XXX or XXXXXX");
        }
    }

    private readonly long _value;

    public byte Red
        => unchecked((byte)(_value >> RedShift));

    public byte Green
        => unchecked((byte)(_value >> GreenShift));

    public byte Blue
        => unchecked((byte)(_value >> BlueShift));

    public bool IsAuto
        => _value is 0L;

    private RgbColor(long value)
        => _value = value;

    public bool Equals(RgbColor other)
        => _value == other._value;

    public override bool Equals(object? obj)
        => obj is RgbColor other && Equals(other);

    public override int GetHashCode()
        => _value.GetHashCode();

    private static void CheckByte(int value, string name)
    {
        if (unchecked((uint)value) > byte.MaxValue)
        {
            throw new ArgumentException(
                $"Component \"{name}\" of {nameof(RgbColor)} must be between {byte.MinValue} and {byte.MinValue}",
                name);
        }
    }
}