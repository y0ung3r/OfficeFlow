using System.IO;
using OfficeFlow.OpenXml.Packaging.Interfaces;

namespace OfficeFlow.OpenXml.Packaging
{
    internal sealed class FlushUsingFilePath : IPackageFlushStrategy
    {
        private readonly string _filePath;

        public FlushUsingFilePath(string filePath)
            => _filePath = filePath;

        /// <inheritdoc />
        public void Flush(MemoryStream internalStream)
        {
            using (var fileStream = new FileStream(_filePath, FileMode.Create, FileAccess.Write))
                internalStream.WriteTo(fileStream);
        }
    }
}