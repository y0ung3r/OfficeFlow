namespace OfficeFlow.MeasureUnits.Absolute
{
	public abstract class AbsoluteUnits
	{
		public static Centimeters Centimeters
			=> new Centimeters();

		public static Points Points
			=> new Points();

		public static Inches Inches
			=> new Inches();

		public static Picas Picas
			=> new Picas();

		public static Millimeters Millimeters
			=> new Millimeters();

		internal static Emu Emu
			=> new Emu();

		internal static HalfPoints HalfPoints
			=> new HalfPoints();

		public static Twips Twips
			=> new Twips();

		internal abstract double Ratio { get; }

		internal AbsoluteValue<Emu> ToEmu(double value)
			=> AbsoluteValue<Emu>.From(value * Ratio);

		internal double FromEmu(AbsoluteValue<Emu> emu)
			=> emu.Raw / Ratio;
	}
}