using System.IO;
using OfficeFlow.OpenXml.Packaging.Interfaces;

namespace OfficeFlow.OpenXml.Packaging;

internal sealed class FlushUsingFilePath(string filePath) : IPackageFlushStrategy
{
    /// <inheritdoc />
    public void Flush(MemoryStream internalStream)
    {
        using var fileStream = new FileStream(filePath, FileMode.Create, FileAccess.Write);
        internalStream.WriteTo(fileStream);
    }
}