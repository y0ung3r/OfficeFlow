using System.IO;
using OfficeFlow.Formats.Interfaces;

namespace OfficeFlow.Formats
{
    public static class OfficeFormatDetector
    {
        public static IOfficeFormat Detect(Stream stream)
        {
            return new OpenXmlFormat();
        }

        public static IOfficeFormat Detect(string filePath)
        {
            return new OpenXmlFormat();
        }
    }
}