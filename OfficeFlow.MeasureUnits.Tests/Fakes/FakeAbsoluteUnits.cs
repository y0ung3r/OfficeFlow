using OfficeFlow.MeasureUnits.Absolute;

namespace OfficeFlow.MeasureUnits.Tests.Fakes
{
	internal sealed class FakeAbsoluteUnits : AbsoluteUnits
	{
		internal override double Ratio { get; }
		
		public FakeAbsoluteUnits()
			: this(1.0)
		{ }

		public FakeAbsoluteUnits(double ratio)
			=> Ratio = ratio;
	}
}