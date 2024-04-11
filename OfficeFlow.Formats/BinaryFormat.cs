using OfficeFlow.Formats.Interfaces;

namespace OfficeFlow.Formats
{
    public sealed class BinaryFormat : IOfficeFormat
    {
        /// <inheritdoc />
        public string[] Extensions { get; } =
        {
            ".doc",
            ".dot",
            ".xls",
            ".xlt",
            ".xlm",
            ".pps",
            ".ppt",
            ".pot"
        };

        /// <inheritdoc />
        public byte[][] Hexdumps { get; } =
        {
            new byte[] { 0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1 }
        };
    }
}