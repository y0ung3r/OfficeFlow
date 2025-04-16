using System.IO;
using OfficeFlow.OpenXml.Packaging.Interfaces;

namespace OfficeFlow.OpenXml.Packaging;

internal sealed class FlushUsingStream(Stream remoteStream) : IPackageFlushStrategy
{
    /// <inheritdoc />
    public void Flush(MemoryStream internalStream)
    {
        if (remoteStream == internalStream)
        {
            return;
        }

        remoteStream.SetLength(value: 0);
        remoteStream.Seek(offset: 0, SeekOrigin.Begin);

        internalStream.WriteTo(remoteStream);
    }
}