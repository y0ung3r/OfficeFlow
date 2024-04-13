using OfficeFlow.Formats.Interfaces;

namespace OfficeFlow.Formats
{
    public sealed class OpenXmlFormat : IOfficeFormat
    {
        /// <inheritdoc />
        public string[] Extensions { get; } =
        {
            ".docx",
            ".docm",
            ".dotx",
            ".dotm",
            ".xlsx",
            ".xlsm",
            ".xltx",
            ".xltm",
            ".pptx",
            ".pptm",
            ".potx",
            ".potm",
            ".ppsx",
            ".ppsm"
        };

        /// <inheritdoc />
        public byte[][] Hexdumps { get; } =
        {
            new byte[] { 0x50, 0x4B, 0x03, 0x04 },
            new byte[] { 0x50, 0x4B, 0x03, 0x04, 0x14, 0x00, 0x06, 0x00 }
        };
    }
}