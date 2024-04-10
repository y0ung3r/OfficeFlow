using System.IO;
using OfficeFlow.Formats.Enums;

namespace OfficeFlow.Formats
{
    public static class FileFormatDetector
    {
        public static FileFormat Detect(Stream stream)
        {
            return FileFormat.OpenXml;
        }

        public static FileFormat Detect(string filePath)
        {
            return FileFormat.OpenXml;
        }
    }
}