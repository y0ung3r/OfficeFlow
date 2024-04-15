namespace OfficeFlow.MeasureUnits.Absolute;

public sealed class Twips : AbsoluteUnits
{
    /// <inheritdoc />
    internal override double Ratio
        => ConversionRatios.Twips;
}