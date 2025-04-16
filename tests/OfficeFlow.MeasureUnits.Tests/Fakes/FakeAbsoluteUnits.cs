using OfficeFlow.MeasureUnits.Absolute;

namespace OfficeFlow.MeasureUnits.Tests.Fakes;

internal sealed class FakeAbsoluteUnits(double ratio) : AbsoluteUnits(ratio)
{
    public FakeAbsoluteUnits()
        : this(1.0)
    { }
}