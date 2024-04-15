using OfficeFlow.MeasureUnits.Absolute;

namespace OfficeFlow.MeasureUnits.Tests.Fakes;

internal sealed class FakeAbsoluteUnits(double ratio) : AbsoluteUnits
{
    internal override double Ratio { get; } = ratio;

    public FakeAbsoluteUnits()
        : this(1.0)
    { }
}