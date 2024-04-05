namespace OfficeFlow.MeasureUnits.Absolute
{
	/// <summary>
	/// https://en.wikipedia.org/wiki/Office_Open_XML_file_formats
	/// </summary>
	public sealed class Emu : AbsoluteUnits
	{
		internal override double Ratio
			=> ConversionRatios.Emu;
	}
}