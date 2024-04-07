using System.IO;
using OfficeFlow.Word.OpenXml.Packaging.Interfaces;

namespace OfficeFlow.Word.OpenXml.Packaging
{
    internal sealed class PreserveFlush : IPackageFlushStrategy
    {
        /// <inheritdoc />
        public void Flush(MemoryStream internalStream)
        { }
    }
}