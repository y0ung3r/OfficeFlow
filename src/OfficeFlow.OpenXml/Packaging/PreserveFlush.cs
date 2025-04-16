using System.IO;
using OfficeFlow.OpenXml.Packaging.Interfaces;

namespace OfficeFlow.OpenXml.Packaging;

internal sealed class PreserveFlush : IPackageFlushStrategy
{
    /// <inheritdoc />
    public void Flush(MemoryStream internalStream)
    { }
}